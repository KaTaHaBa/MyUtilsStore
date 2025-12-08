import win32com.client
import os
from datetime import datetime
import shutil
import uuid
# 画像サイズ取得用
from PIL import Image

# --- 描画ライブラリ ---
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- PowerPoint定数 ---
msoShapeRectangle = 1
msoTextOrientationHorizontal = 1
msoAlignLeft = 1
msoAlignCenter = 2
msoAlignRight = 3
msoTrue = -1
msoFalse = 0
ppLayoutBlank = 12
ppSaveAsOpenXMLPresentation = 11
ppWindowNormal = 1
ppWindowMinimized = 2

# ==============================================================================
# A. 設定定義クラス群 (Config Classes)
# ==============================================================================
class PathConfig:
    def __init__(self, template_path=None, output_dir=None, temp_dir_name='temp_images_slide_gen'):
        base_dir = os.getcwd()
        if template_path:
            self.template_file = template_path
        else:
            self.template_file = os.path.join(base_dir, 'template.pptx')
        if output_dir:
            self.output_dir = output_dir
        else:
            self.output_dir = base_dir
        self.temp_img_dir = os.path.join(base_dir, temp_dir_name)

class ColorConfig:
    def __init__(self):
        # HEXコード定義 (Matplotlib/Plotly/PPT共通)
        self.palette = {
            'navy': '#002060', 'red': '#E60000', 'teal': '#008080',
            'gray_text': '#595959', 'gray_bg': '#F2F2F2',
            'white': '#FFFFFF', 'black': '#000000'
        }

class FontConfig:
    def __init__(self):
        # 日本語フォント
        self.main_font = 'Meiryo'
        plt.rcParams['font.family'] = self.main_font

class LayoutElementRatio:
    """スライドサイズに対する比率(0.0~1.0)で定義"""
    def __init__(self, left_ratio, top_ratio, width_ratio, height_ratio, font_size=None):
        self.left_ratio = left_ratio
        self.top_ratio = top_ratio
        self.width_ratio = width_ratio
        self.height_ratio = height_ratio
        self.font_size = font_size

class LayoutConfig:
    """レイアウト定義（50:50分割レイアウト）"""
    def __init__(self):
        # 表紙
        self.cover_title = LayoutElementRatio(0.1, 0.35, 0.8, 0.15, font_size=32)
        self.cover_sub   = LayoutElementRatio(0.1, 0.52, 0.8, 0.10, font_size=20)
        self.cover_date  = LayoutElementRatio(0.1, 0.85, 0.8, 0.05, font_size=18)
        
        # 本文スライドタイトル (右上に配置)
        self.content_title = LayoutElementRatio(0.40, 0.05, 0.55, 0.08, font_size=24)
        
        # --- レイアウト変更: 左右50%ずつのイメージ ---
        # グラフエリア (左半分)
        # Left: 5%, Top: 18%, Width: 45%, Height: 70%
        self.chart_area = LayoutElementRatio(0.05, 0.18, 0.45, 0.70)
        
        # 提案ポイントエリア (右半分)
        # Left: 52%, Top: 18%, Width: 43%, Height: 70%
        self.proposal_area = LayoutElementRatio(0.52, 0.18, 0.43, 0.70)

class SlideConfig:
    def __init__(self, template_path=None, output_dir=None, engine='matplotlib'):
        """
        engine: 'matplotlib' (default) or 'plotly'
        """
        self.paths = PathConfig(template_path=template_path, output_dir=output_dir)
        self.colors = ColorConfig()
        self.fonts = FontConfig()
        self.layout = LayoutConfig()
        self.engine = engine # 描画エンジンの選択


# ==============================================================================
# B. グラフ描画戦略 (Strategy Pattern)
# ==============================================================================
class ChartStrategyBase:
    """描画エンジンの基底クラス"""
    def __init__(self, config: SlideConfig):
        self.config = config
    
    def _get_color(self, color_key):
        return self.config.colors.palette.get(color_key, '#000000')

    def plot_combo_bar_line_2axis(self, df, mapping, chart_text):
        raise NotImplementedError

# --- Matplotlib 実装 ---
class MatplotlibStrategy(ChartStrategyBase):
    def __init__(self, config: SlideConfig):
        super().__init__(config)
        sns.set_theme(style="white", rc={"axes.grid": False})
        plt.rcParams['font.family'] = config.fonts.main_font

    def _apply_common_style(self, fig, ax, chart_text):
        ax.set_title(chart_text.get('title', ''), fontsize=16, fontweight='bold', 
                     color=self._get_color('navy'), pad=20)
        for axis in [ax.xaxis, ax.yaxis]:
            axis.label.set_color(self._get_color('gray_text'))
            axis.set_tick_params(colors=self._get_color('gray_text'))
        sns.despine(left=True, bottom=True)
        ax.yaxis.grid(True, color='lightgray', linestyle='--', linewidth=0.5)
        return fig, ax

    def plot_combo_bar_line_2axis(self, df, mapping, chart_text):
        print(f"  [DEBUG] Matplotlib engine selected.")
        fig, ax1 = plt.subplots(figsize=(10, 6)) # アスペクト比 5:3
        x_col = mapping['x_col']
        
        # Bar Chart
        bar_cols = [t['col'] for t in mapping.get('bar_traces', [])]
        if bar_cols:
            df_bar_melt = df.melt(id_vars=[x_col], value_vars=bar_cols, var_name='Metric', value_name='Value')
            bar_palette = {t['col']: self._get_color(t.get('color_key', 'navy')) for t in mapping.get('bar_traces', [])}
            sns.barplot(data=df_bar_melt, x=x_col, y='Value', hue='Metric', 
                        palette=bar_palette, ax=ax1, alpha=0.7, dodge=True)
            handles, labels = ax1.get_legend_handles_labels()
            trace_name_map = {t['col']: t['name'] for t in mapping.get('bar_traces', [])}
            new_labels = [trace_name_map.get(l, l) for l in labels]
            ax1.legend(handles, new_labels, loc='upper left', bbox_to_anchor=(0, -0.15), ncol=len(bar_cols), frameon=False)
            ax1.set_ylabel(chart_text.get('y1_label', ''), fontsize=12)

        # Line Chart
        line_traces = mapping.get('line_traces', [])
        if line_traces:
            ax2 = ax1.twinx()
            for trace_def in line_traces:
                color = self._get_color(trace_def.get('color_key', 'red'))
                marker_size = trace_def.get('marker_size', 8)
                line_width = trace_def.get('line_width', 2.5)
                sns.lineplot(data=df, x=x_col, y=trace_def['col'], ax=ax2,
                             color=color, marker='o', markersize=marker_size, linewidth=line_width,
                             label=trace_def['name'])
            ax2.set_ylabel(chart_text.get('y2_label', ''), fontsize=12, color=self._get_color('gray_text'))
            ax2.tick_params(axis='y', colors=self._get_color('gray_text'))
            sns.despine(ax=ax2, right=False, left=True, bottom=True)
            ax2.legend(loc='upper left', bbox_to_anchor=(0.5, -0.15), ncol=len(line_traces), frameon=False)
            ax2.grid(False)

        ax1.set_xlabel('')
        fig, ax1 = self._apply_common_style(fig, ax1, chart_text)
        plt.tight_layout()
        return fig

# --- Plotly 実装 ---
class PlotlyStrategy(ChartStrategyBase):
    def __init__(self, config: SlideConfig):
        super().__init__(config)

    def _apply_common_layout(self, fig, chart_text):
        fonts = self.config.fonts
        fig.update_layout(
            title=dict(
                text=f"<b>{chart_text.get('title', '')}</b>",
                font=dict(family=fonts.main_font, size=18, color=self._get_color('navy')),
                x=0.5, xanchor='center'
            ),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5, font=dict(family=fonts.main_font)),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family=fonts.main_font, color=self._get_color('gray_text')),
            margin=dict(l=40, r=40, t=60, b=60)
        )
        return fig

    def plot_combo_bar_line_2axis(self, df, mapping, chart_text):
        print(f"  [DEBUG] Plotly engine selected.")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        x_col = mapping['x_col']

        for trace_def in mapping.get('bar_traces', []):
            fig.add_trace(
                go.Bar(x=df[x_col], y=df[trace_def['col']], name=trace_def['name'],
                       marker_color=self._get_color(trace_def.get('color_key', 'navy')), opacity=0.8),
                secondary_y=False
            )
        for trace_def in mapping.get('line_traces', []):
            c_hex = self._get_color(trace_def.get('color_key', 'red'))
            fig.add_trace(
                go.Scatter(x=df[x_col], y=df[trace_def['col']], name=trace_def['name'],
                           mode='lines+markers',
                           line=dict(color=c_hex, width=trace_def.get('line_width', 2.5)),
                           marker=dict(size=trace_def.get('marker_size', 8), color=c_hex)),
                secondary_y=True
            )
        fig = self._apply_common_layout(fig, chart_text)
        fig.update_yaxes(title_text=chart_text.get('y1_label', ''), showgrid=True, gridcolor='lightgray', secondary_y=False)
        fig.update_yaxes(title_text=chart_text.get('y2_label', ''), showgrid=False, secondary_y=True)
        return fig


# ==============================================================================
# C. スライド生成エンジン (Core Engine)
# ==============================================================================
class PowerPointGeneratorEngine:
    def __init__(self, config: SlideConfig):
        self.config = config
        self.ppt_app = None
        self.prs = None
        
        self.idx_cover = 1
        self.idx_template_body = 2
        self.idx_back_cover = 3
        
        # エンジンの切り替えロジック
        if self.config.engine == 'plotly':
            self.strategy = PlotlyStrategy(config)
        else:
            self.strategy = MatplotlibStrategy(config)
        
        self.slide_width = 960 
        self.slide_height = 540 

        if os.path.exists(self.config.paths.temp_img_dir):
             shutil.rmtree(self.config.paths.temp_img_dir)
        os.makedirs(self.config.paths.temp_img_dir, exist_ok=True)
        os.makedirs(self.config.paths.output_dir, exist_ok=True)

    def _initialize_ppt(self):
        tpl_path = self.config.paths.template_file
        print(f"[DEBUG] テンプレートファイルを開きます: {tpl_path}")
        if not os.path.exists(tpl_path):
             raise FileNotFoundError(f"テンプレートが見つかりません: {tpl_path}")
        self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt_app.Visible = msoTrue
        try: self.ppt_app.WindowState = ppWindowMinimized
        except: pass
        self.prs = self.ppt_app.Presentations.Open(tpl_path, ReadOnly=msoTrue)
        
        # スライドサイズ取得
        self.slide_width = self.prs.PageSetup.SlideWidth
        self.slide_height = self.prs.PageSetup.SlideHeight
        print(f"[DEBUG] スライドサイズ: W={self.slide_width}, H={self.slide_height}")

    def _to_ppt_rgb(self, color_key_or_tuple):
        val = color_key_or_tuple
        if isinstance(val, str) and not val.startswith('#'):
            val = self.config.colors.palette.get(val, '#000000')
        if isinstance(val, str) and val.startswith('#'):
            hex_color = val.lstrip('#')
            if len(hex_color) == 6:
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                val = (r, g, b)
            else: val = (0, 0, 0)
        if not isinstance(val, tuple): val = (0, 0, 0)
        return val[0] + (val[1] << 8) + (val[2] << 16)

    def _calc_rect(self, layout_ratio: LayoutElementRatio):
        left = layout_ratio.left_ratio * self.slide_width
        top = layout_ratio.top_ratio * self.slide_height
        width = layout_ratio.width_ratio * self.slide_width
        height = layout_ratio.height_ratio * self.slide_height
        return left, top, width, height

    def _add_text_box(self, slide, text, layout_ratio, bold=False, color_key='navy', align=msoAlignLeft):
        left, top, width, height = self._calc_rect(layout_ratio)
        shape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
        tf = shape.TextFrame
        tf.TextRange.Text = text
        if layout_ratio.font_size: tf.TextRange.Font.Size = layout_ratio.font_size
        tf.TextRange.Font.Bold = msoTrue if bold else msoFalse
        tf.TextRange.Font.Color.RGB = self._to_ppt_rgb(color_key)
        tf.TextRange.Font.Name = self.config.fonts.main_font
        tf.TextRange.ParagraphFormat.Alignment = align
        return shape

    def _add_picture_fitted(self, slide, image_path, layout_ratio):
        """
        【重要】画像をバウンディングボックス内に、アスペクト比を維持したまま最大サイズで配置し、中央寄せする。
        """
        if not os.path.exists(image_path): return
        
        # 1. 配置エリア（バウンディングボックス）の計算
        box_left, box_top, box_width, box_height = self._calc_rect(layout_ratio)
        
        # 2. 画像の元サイズを取得
        with Image.open(image_path) as img:
            img_w, img_h = img.size
            
        # 3. 縮尺率の計算 (枠に収まるように)
        scale_w = box_width / img_w
        scale_h = box_height / img_h
        scale = min(scale_w, scale_h) # 小さい方に合わせて全体を縮小
        
        # 4. 新しい画像サイズ
        new_w = img_w * scale
        new_h = img_h * scale
        
        # 5. 中央寄せのための座標オフセット計算
        offset_x = (box_width - new_w) / 2
        offset_y = (box_height - new_h) / 2
        
        final_left = box_left + offset_x
        final_top = box_top + offset_y
        
        # 6. 貼り付け (Width/Heightを指定することでリサイズされるが、比率は維持した計算値)
        slide.Shapes.AddPicture(image_path, msoFalse, msoTrue, final_left, final_top, new_w, new_h)

    def _add_proposal_points(self, slide, points_data, section_title_text="■ 主要な論点・ご提案"):
        layout_ratio = self.config.layout.proposal_area
        if not points_data: return
        area_left, area_top, area_width, area_height = self._calc_rect(layout_ratio)

        # セクションタイトル
        self._add_text_box(slide, section_title_text, 
                           LayoutElementRatio(layout_ratio.left_ratio, layout_ratio.top_ratio, layout_ratio.width_ratio, 0), 
                           bold=True)
        
        current_top = area_top + 40
        available_height = area_height - 40
        margin = 15
        box_height = (available_height - (len(points_data)-1)*margin) / len(points_data)
        
        for point in points_data:
            box = slide.Shapes.AddShape(msoShapeRectangle, area_left, current_top, area_width, box_height)
            box.Fill.ForeColor.RGB = self._to_ppt_rgb('gray_bg')
            box.Line.Visible = msoFalse
            
            accent_height = 30
            accent_color = point.get('accent_color_key', 'red')
            accent = slide.Shapes.AddShape(msoShapeRectangle, area_left, current_top, area_width, accent_height)
            accent.Fill.ForeColor.RGB = self._to_ppt_rgb(accent_color)
            accent.Line.Visible = msoFalse
            
            # Helper for text placement relative to the box (simulating ratio based on box position)
            # 座標計算が複雑になるため、ここでは AddTextbox に直接計算済み座標を渡すヘルパーがあれば良いが、
            # 統一性のために LayoutElementRatio を逆算して渡す
            
            ct_ratio = current_top / self.slide_height
            bh_ratio = box_height / self.slide_height
            
            title_color = point.get('title_color_key', 'white')
            
            self._add_text_box(slide, point['title'], 
                               LayoutElementRatio(layout_ratio.left_ratio + 0.01, ct_ratio, layout_ratio.width_ratio - 0.02, 0.05),
                               bold=True, color_key=title_color, align=msoAlignCenter)
            
            self._add_text_box(slide, point['body'],
                               LayoutElementRatio(layout_ratio.left_ratio + 0.015, ct_ratio + (35/self.slide_height), layout_ratio.width_ratio - 0.03, bh_ratio - (40/self.slide_height)),
                               color_key='gray_text')
            
            current_top += box_height + margin

    def generate(self, df, cover_content, slides_structure, filename_prefix="Presentation"):
        try:
            print("Engine: 処理開始...")
            self._initialize_ppt()

            print("Engine: 表紙を作成中...")
            slide1 = self.prs.Slides(self.idx_cover)
            self._add_text_box(slide1, cover_content['main_title'], self.config.layout.cover_title, bold=True, align=msoAlignCenter)
            self._add_text_box(slide1, cover_content['sub_title'], self.config.layout.cover_sub, color_key='gray_text', align=msoAlignCenter)
            self._add_text_box(slide1, cover_content.get('date', datetime.now().strftime("%Y年%m月%d日")), 
                               self.config.layout.cover_date, color_key='gray_text', align=msoAlignCenter)

            template_slide = self.prs.Slides(self.idx_template_body)
            
            for i, slide_def in enumerate(slides_structure):
                print(f"Engine: スライド {i+2} ('{slide_def['slide_title']}') を作成中...")
                new_slide = template_slide.Duplicate().Item(1)
                
                # 裏表紙の前へ移動
                target_index = self.prs.Slides.Count - 1
                if target_index < 2: target_index = 2
                new_slide.MoveTo(target_index)

                # タイトル (右寄せ)
                self._add_text_box(new_slide, slide_def['slide_title'], self.config.layout.content_title, bold=True, align=msoAlignRight)

                # グラフ描画
                category = slide_def['category']
                mapping = slide_def['data_mapping']
                chart_text = slide_def['chart_text']
                
                plot_method_name = f"plot_{category.lower()}"
                plot_method = getattr(self.strategy, plot_method_name, None)
                
                if plot_method:
                    fig = plot_method(df, mapping, chart_text)
                    img_filename = f"chart_{uuid.uuid4()}.png"
                    img_path = os.path.join(self.config.paths.temp_img_dir, img_filename)
                    
                    if self.config.engine == 'plotly':
                        fig.write_image(img_path, width=1000, height=600, scale=2)
                    else:
                        # Matplotlib
                        fig.savefig(img_path, dpi=150, bbox_inches='tight')
                        plt.close(fig)
                    
                    # --- 変更点: アスペクト比維持貼り付け ---
                    self._add_picture_fitted(new_slide, img_path, self.config.layout.chart_area)
                
                if 'proposal_points' in slide_def:
                    section_title = slide_def.get('proposal_section_title', "■ 主要な論点・ご提案")
                    self._add_proposal_points(new_slide, slide_def['proposal_points'], section_title_text=section_title)

            print("Engine: 仕上げ処理中...")
            template_slide.Delete()
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{filename_prefix}_{timestamp}.pptx"
            output_path = os.path.join(self.config.paths.output_dir, filename)
            
            print(f"[DEBUG] 最終ファイルを保存します: {output_path}")
            self.prs.SaveAs(output_path, ppSaveAsOpenXMLPresentation)
            self.prs.Close()
            self.prs = None
            print(f"Engine: 完了！ファイルが出力されました: {output_path}")

        except Exception as e:
            print(f"Engine Error: {e}")
            import traceback
            traceback.print_exc()
        
        finally:
            print("Engine: PowerPointを終了します...")
            if self.prs:
                try: self.prs.Close()
                except: pass
            if self.ppt_app:
                try:
                    try: self.ppt_app.WindowState = ppWindowNormal
                    except: pass
                    self.ppt_app.Quit()
                except: pass
            
            if os.path.exists(self.config.paths.temp_img_dir):
                 try: shutil.rmtree(self.config.paths.temp_img_dir)
                 except: pass