import win32com.client
import os
from datetime import datetime
import shutil
import uuid
from PIL import Image
from typing import List, Dict, Optional, Union, Tuple, Any

# --- 描画ライブラリ ---
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.ticker as ticker
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- PowerPoint定数 ---
MSO_SHAPE_RECTANGLE = 1
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_ALIGN_LEFT = 1
MSO_ALIGN_CENTER = 2
MSO_ALIGN_RIGHT = 3
MSO_TRUE = -1
MSO_FALSE = 0
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_OPENXML_PRESENTATION = 11
PP_WINDOW_NORMAL = 1
PP_WINDOW_MINIMIZED = 2

# ==============================================================================
# A. 設定定義クラス群
# ==============================================================================

class PathConfig:
    def __init__(self, template_path: Optional[str] = None, output_dir: Optional[str] = None, temp_dir_name: str = 'temp_images_slide_gen'):
        base_dir = os.getcwd()
        self.template_file = template_path if template_path else os.path.join(base_dir, 'template.pptx')
        self.output_dir = output_dir if output_dir else base_dir
        self.temp_img_dir = os.path.join(base_dir, temp_dir_name)


class ColorConfig:
    """
    カラーパレット定義。
    汎用性を高めるため、用途ではなく「一般的な色の名前」で定義。
    """
    def __init__(self):
        self.palette: Dict[str, str] = {
            # --- Monotone / Basic ---
            'black': '#000000',
            'white': '#FFFFFF',
            'gray_dark': '#595959',       # 濃いグレー (文字用)
            'gray_medium': '#7F8C8D',     # 中間のグレー (グラフ線など)
            'gray_light': '#F2F2F2',      # 薄いグレー (背景用)
            
            # --- Blues ---
            'navy': '#002060',            # 濃紺
            'midnight_blue': '#1F3A5F',   # 深い青
            'teal': '#008080',            # 青緑
            'cadet_blue': '#5f9ea0',      # くすんだ青緑
            'sky_blue': '#a8d8ea',        # 空色 (パステル)
            'ice_blue': '#e0f2f1',        # 非常に薄い青
            
            # --- Reds / Pinks ---
            'red': '#E60000',             # 赤
            'salmon': '#f38181',          # サーモンピンク
            'pale_pink': '#fcbad3',       # 薄いピンク
            'coral': '#ff6f69',           # コーラルレッド
            
            # --- Yellows / Golds ---
            'gold': '#C2A970',            # 落ち着いた金
            'cream': '#ffffd2',           # クリーム色
            'mustard': '#ffcc5c',         # マスタードイエロー
            
            # --- Greens ---
            'spring_green': '#42e6a4',    # 明るい緑
            'mint': '#95e1d3',            # ミントグリーン
            
            # --- Purples ---
            'lavender': '#aa96da',        # 薄紫
            'dark_slate': '#2C3E50',      # ダークスレート (ほぼ黒に近い青)
        }


class FontConfig:
    def __init__(self):
        self.japanese_font: str = 'Meiryo'
        self.english_font: str = 'Segoe UI'
        plt.rcParams['font.family'] = self.japanese_font
        self.chart_title_size: int = 18
        self.chart_label_size: int = 12
        self.chart_tick_size: int = 11


class LayoutRatio:
    def __init__(self, left: float, top: float, width: float, height: float, font_size: Optional[int] = None):
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.font_size = font_size


class LayoutConfig:
    def __init__(self, split_ratio: float = 0.5):
        self.split_ratio = split_ratio
        
        self.cover_title = LayoutRatio(0.1, 0.35, 0.8, 0.15, font_size=32)
        self.cover_sub   = LayoutRatio(0.1, 0.52, 0.8, 0.10, font_size=20)
        self.cover_date  = LayoutRatio(0.1, 0.85, 0.8, 0.05, font_size=18)
        self.content_title = LayoutRatio(0.40, 0.05, 0.55, 0.08, font_size=24)
        
        margin_outer = 0.05
        margin_inner = 0.02
        title_area_h = 0.15
        
        # Horizontal Split
        h_chart_w = self.split_ratio - margin_outer - (margin_inner / 2)
        self.layout_horizontal_chart = LayoutRatio(margin_outer, title_area_h + 0.03, h_chart_w, 0.75)
        h_text_x = self.split_ratio + (margin_inner / 2)
        h_text_w = (1.0 - margin_outer) - h_text_x
        self.layout_horizontal_text  = LayoutRatio(h_text_x, title_area_h + 0.03, h_text_w, 0.75)
        
        # Vertical Split
        content_total_h = 1.0 - title_area_h - margin_outer
        v_chart_h = content_total_h * self.split_ratio
        self.layout_vertical_chart = LayoutRatio(margin_outer, title_area_h, 1.0 - (margin_outer * 2), v_chart_h)
        v_text_y = title_area_h + v_chart_h + margin_inner
        v_text_h = content_total_h * (1.0 - self.split_ratio) - margin_inner
        self.layout_vertical_text  = LayoutRatio(margin_outer, v_text_y, 1.0 - (margin_outer * 2), v_text_h)
        
        self.body_text_max_font_size = 16


class SlideConfig:
    def __init__(self, template_path: Optional[str] = None, output_dir: Optional[str] = None, engine: str = 'matplotlib', split_ratio: float = 0.5):
        self.paths = PathConfig(template_path=template_path, output_dir=output_dir)
        self.colors = ColorConfig()
        self.fonts = FontConfig()
        self.layout = LayoutConfig(split_ratio=split_ratio)
        self.engine = engine


# ==============================================================================
# B. グラフ描画戦略
# ==============================================================================

class ChartStrategyBase:
    def __init__(self, config: SlideConfig):
        self.config = config
    
    def _get_color(self, color_key: str) -> str:
        """カラーキーからHEXコードを取得。キーが見つからない場合はデフォルト黒を返す"""
        return self.config.colors.palette.get(color_key, '#000000')

    def plot_combo_bar_line_2axis(self, df: Any, mapping: Dict, chart_text: Dict) -> Any:
        raise NotImplementedError

    def plot_balance_sheet(self, df: Any, mapping: Dict, chart_text: Dict) -> Any:
        raise NotImplementedError


def jpy_currency_formatter(x: float, pos: Any) -> str:
    if x == 0: return '0'
    abs_x = abs(x)
    if abs_x >= 1e12: return f'{x/1e12:,.1f}兆'
    elif abs_x >= 1e8: return f'{x/1e8:,.0f}億'
    elif abs_x >= 1e6: return f'{x/1e6:,.0f}百万'
    elif abs_x >= 1e4: return f'{x/1e4:,.0f}万'
    else: return f'{x:,.0f}'


class MatplotlibStrategy(ChartStrategyBase):
    def __init__(self, config: SlideConfig):
        super().__init__(config)
        sns.set_theme(style="white", rc={"axes.grid": False})
        plt.rcParams['font.family'] = config.fonts.japanese_font

    def _apply_common_style(self, fig: plt.Figure, ax: plt.Axes, chart_text: Dict) -> Tuple[plt.Figure, plt.Axes]:
        fonts = self.config.fonts
        ax.set_title(chart_text.get('title', ''), 
                     fontsize=fonts.chart_title_size, fontweight='bold', 
                     color=self._get_color('navy'), pad=20)
        for axis in [ax.xaxis, ax.yaxis]:
            axis.label.set_color(self._get_color('gray_dark'))
            axis.label.set_fontsize(fonts.chart_label_size)
            axis.set_tick_params(colors=self._get_color('gray_dark'), labelsize=fonts.chart_tick_size)
        sns.despine(left=True, bottom=True)
        ax.yaxis.grid(True, color='#E0E0E0', linestyle='--', linewidth=0.5)
        return fig, ax

    def plot_combo_bar_line_2axis(self, df: Any, mapping: Dict, chart_text: Dict) -> plt.Figure:
        print(f"  [DEBUG] Matplotlib: Plotting Combo Bar-Line Chart.")
        fig, ax1 = plt.subplots(figsize=(12, 7)) 
        x_col = mapping['x_col']
        
        bar_cols = [t['col'] for t in mapping.get('bar_traces', [])]
        if bar_cols:
            df_bar_melt = df.melt(id_vars=[x_col], value_vars=bar_cols, var_name='Metric', value_name='Value')
            # デフォルト色: navy
            bar_palette = {t['col']: self._get_color(t.get('color_key', 'navy')) for t in mapping.get('bar_traces', [])}
            sns.barplot(data=df_bar_melt, x=x_col, y='Value', hue='Metric', 
                        palette=bar_palette, ax=ax1, alpha=0.85, dodge=True)
            
            handles, labels = ax1.get_legend_handles_labels()
            trace_name_map = {t['col']: t['name'] for t in mapping.get('bar_traces', [])}
            new_labels = [trace_name_map.get(l, l) for l in labels]
            ax1.legend(handles, new_labels, loc='upper left', bbox_to_anchor=(0, -0.15), 
                       ncol=len(bar_cols), frameon=False, fontsize=self.config.fonts.chart_tick_size)
            ax1.set_ylabel(chart_text.get('y1_label', ''), fontsize=self.config.fonts.chart_label_size)
            ax1.yaxis.set_major_formatter(ticker.FuncFormatter(jpy_currency_formatter))

        line_traces = mapping.get('line_traces', [])
        if line_traces:
            ax2 = ax1.twinx()
            for trace_def in line_traces:
                # デフォルト色: red
                color = self._get_color(trace_def.get('color_key', 'red'))
                marker_size = trace_def.get('marker_size', 10)
                line_width = trace_def.get('line_width', 3.5)
                sns.lineplot(data=df, x=x_col, y=trace_def['col'], ax=ax2,
                             color=color, marker='o', markersize=marker_size, linewidth=line_width,
                             label=trace_def['name'])
            ax2.set_ylabel(chart_text.get('y2_label', ''), fontsize=self.config.fonts.chart_label_size, color=self._get_color('gray_dark'))
            ax2.tick_params(axis='y', colors=self._get_color('gray_dark'), labelsize=self.config.fonts.chart_tick_size)
            sns.despine(ax=ax2, right=False, left=True, bottom=True)
            ax2.legend(loc='upper left', bbox_to_anchor=(0.5, -0.15), ncol=len(line_traces), frameon=False, fontsize=self.config.fonts.chart_tick_size)
            ax2.grid(False)

        ax1.set_xlabel('')
        fig, ax1 = self._apply_common_style(fig, ax1, chart_text)
        plt.tight_layout()
        return fig

    def plot_balance_sheet(self, df: Any, mapping: Dict, chart_text: Dict) -> plt.Figure:
        print(f"  [DEBUG] Matplotlib: Plotting Balance Sheet Chart.")
        if len(df) == 0: raise ValueError("DataFrame is empty.")
        row = df.iloc[0]
        
        unit_scale = mapping.get('unit_scale', 1.0)
        total_assets_val = row.get(mapping.get('total_assets_col'), 0)
        
        def _build_stack_data(stack_def):
            data = []
            for item in stack_def:
                val = row.get(item['col'], 0)
                if hasattr(val, 'isna') and val.isna(): val = 0
                if val is None: val = 0
                data.append({
                    'label': item['name'],
                    'value': val / unit_scale,
                    # デフォルト色: gray_medium
                    'color': self._get_color(item.get('color_key', 'gray_medium'))
                })
            return data

        left_stack_data = _build_stack_data(mapping.get('left_stack', []))
        right_stack_data = _build_stack_data(mapping.get('right_stack', []))

        fig, ax = plt.subplots(figsize=(8, 8))
        
        def _stack_bars(x_pos, data_list):
            bottom = 0
            for d in data_list:
                val = d['value']
                if val > 0:
                    ax.bar(x_pos, val, bottom=bottom, color=d['color'], edgecolor='white', width=0.6)
                    threshold = (total_assets_val / unit_scale) * 0.05
                    if val > threshold:
                        ax.text(x_pos, bottom + val/2, f"{d['label']}\n{val:,.1f}", 
                                ha='center', va='center', fontsize=self.config.fonts.chart_tick_size, color='#333333')
                    bottom += val

        _stack_bars(0, left_stack_data)
        _stack_bars(1, right_stack_data)
        
        ax.set_xticks([0, 1])
        ax.set_xticklabels([mapping.get('left_label', 'Assets'), mapping.get('right_label', 'Liab & Equity')])
        ax.set_ylabel(chart_text.get('y1_label', ''))
        
        fig, ax = self._apply_common_style(fig, ax, chart_text)
        ax.xaxis.grid(False)
        plt.tight_layout()
        return fig


class PlotlyStrategy(ChartStrategyBase):
    def __init__(self, config: SlideConfig):
        super().__init__(config)
    def _apply_common_layout(self, fig: go.Figure, chart_text: Dict) -> go.Figure:
        return fig
    def plot_combo_bar_line_2axis(self, df: Any, mapping: Dict, chart_text: Dict) -> go.Figure:
        return go.Figure()
    def plot_balance_sheet(self, df: Any, mapping: Dict, chart_text: Dict) -> go.Figure:
        return go.Figure()


# ==============================================================================
# C. スライド生成エンジン
# ==============================================================================

class PowerPointGeneratorEngine:
    def __init__(self, config: SlideConfig):
        self.config = config
        self.ppt_app = None
        self.prs = None
        self.idx_cover = 1
        self.idx_template_body = 2
        self.idx_back_cover = 3
        if self.config.engine == 'plotly': self.strategy = PlotlyStrategy(config)
        else: self.strategy = MatplotlibStrategy(config)
        self.slide_width = 960 
        self.slide_height = 540 
        if os.path.exists(self.config.paths.temp_img_dir): shutil.rmtree(self.config.paths.temp_img_dir)
        os.makedirs(self.config.paths.temp_img_dir, exist_ok=True)
        os.makedirs(self.config.paths.output_dir, exist_ok=True)

    def _initialize_ppt(self):
        tpl_path = self.config.paths.template_file
        print(f"[DEBUG] Opening template: {tpl_path}")
        if not os.path.exists(tpl_path): raise FileNotFoundError(f"Template not found: {tpl_path}")
        self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt_app.Visible = MSO_TRUE
        try: self.ppt_app.WindowState = PP_WINDOW_MINIMIZED
        except: pass
        self.prs = self.ppt_app.Presentations.Open(tpl_path, ReadOnly=MSO_TRUE)
        self.slide_width = self.prs.PageSetup.SlideWidth
        self.slide_height = self.prs.PageSetup.SlideHeight

    def _to_ppt_rgb(self, color_key_or_tuple: Union[str, Tuple[int, int, int]]) -> int:
        val = color_key_or_tuple
        if isinstance(val, str) and not val.startswith('#'):
            val = self.config.colors.palette.get(val, '#000000')
        if isinstance(val, str) and val.startswith('#'):
            hex_color = val.lstrip('#')
            if len(hex_color) == 6:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                val = (r, g, b)
            else: val = (0, 0, 0)
        if not isinstance(val, tuple): val = (0, 0, 0)
        return val[0] + (val[1] << 8) + (val[2] << 16)

    def _calc_rect(self, layout_ratio: LayoutRatio) -> Tuple[float, float, float, float]:
        left = layout_ratio.left * self.slide_width
        top = layout_ratio.top * self.slide_height
        width = layout_ratio.width * self.slide_width
        height = layout_ratio.height * self.slide_height
        return left, top, width, height

    def _add_text_box(self, slide: Any, text: str, layout_ratio: LayoutRatio, 
                      bold: bool = False, color_key: str = 'navy', align: int = MSO_ALIGN_LEFT) -> Any:
        left, top, width, height = self._calc_rect(layout_ratio)
        shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)
        tf = shape.TextFrame
        tf.TextRange.Text = text
        if layout_ratio.font_size: tf.TextRange.Font.Size = layout_ratio.font_size
        tf.TextRange.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        tf.TextRange.Font.Color.RGB = self._to_ppt_rgb(color_key)
        tf.TextRange.Font.Name = self.config.fonts.english_font
        tf.TextRange.Font.NameAscii = self.config.fonts.english_font
        tf.TextRange.Font.NameFarEast = self.config.fonts.japanese_font
        tf.TextRange.ParagraphFormat.Alignment = align
        return shape

    def _add_picture_fitted(self, slide: Any, image_path: str, layout_ratio: LayoutRatio) -> None:
        if not os.path.exists(image_path): return
        box_left, box_top, box_width, box_height = self._calc_rect(layout_ratio)
        with Image.open(image_path) as img: img_w, img_h = img.size
        scale = min(box_width / img_w, box_height / img_h)
        new_w, new_h = img_w * scale, img_h * scale
        final_left = box_left + (box_width - new_w) / 2
        final_top = box_top + (box_height - new_h) / 2
        slide.Shapes.AddPicture(image_path, MSO_FALSE, MSO_TRUE, final_left, final_top, new_w, new_h)

    def _calculate_auto_font_size(self, text: str, max_size: int) -> int:
        length = len(text)
        if length < 50: return max_size
        elif length < 100: return max_size - 2
        elif length < 200: return max_size - 4
        else: return max_size - 6

    def _add_text_content_boxes(self, slide: Any, content_data: List[Dict], layout_type: str = 'horizontal', section_header_text: str = "■ 主要な論点・ご提案"):
        if not content_data: return
        num_boxes = len(content_data)
        if num_boxes == 0: return

        if layout_type == 'vertical': layout_ratio = self.config.layout.layout_vertical_text
        else: layout_ratio = self.config.layout.layout_horizontal_text

        area_left, area_top, area_width, area_height = self._calc_rect(layout_ratio)

        # セクションヘッダー
        title_h_ratio = 0.05
        self._add_text_box(slide, section_header_text, 
                           LayoutRatio(layout_ratio.left, layout_ratio.top, layout_ratio.width, title_h_ratio), 
                           bold=True)
        
        # ボックス配置
        header_margin_px = 40
        current_top_px = area_top + header_margin_px
        available_height_px = area_height - header_margin_px
        box_gap_px = 15
        
        if layout_type == 'vertical':
            box_width_px = (area_width - (num_boxes - 1) * box_gap_px) / num_boxes
            box_height_px = available_height_px
            dx = box_width_px + box_gap_px
            dy = 0
            current_x = area_left
            current_y = current_top_px
        else:
            box_width_px = area_width
            box_height_px = (available_height_px - (num_boxes - 1) * box_gap_px) / num_boxes
            dx = 0
            dy = box_height_px + box_gap_px
            current_x = area_left
            current_y = current_top_px

        box_header_height_px = 35 
        
        for item in content_data:
            # ボックス描画
            header_shape = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, current_x, current_y, box_width_px, box_header_height_px)
            accent_color = item.get('accent_color_key', 'navy') # デフォルト色
            header_shape.Fill.ForeColor.RGB = self._to_ppt_rgb(accent_color)
            header_shape.Line.Visible = MSO_FALSE
            
            body_h_px = box_height_px - box_header_height_px
            body_shape = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, current_x, current_y + box_header_height_px, box_width_px, body_h_px)
            body_shape.Fill.ForeColor.RGB = self._to_ppt_rgb('gray_light')
            body_shape.Line.Visible = MSO_FALSE
            
            # テキスト配置
            current_x_r = current_x / self.slide_width
            current_y_r = current_y / self.slide_height
            box_w_r = box_width_px / self.slide_width
            header_h_r = box_header_height_px / self.slide_height
            body_h_r = body_h_px / self.slide_height
            
            title_text = item['title']
            title_color = item.get('title_color_key', 'white')
            self._add_text_box(slide, title_text, 
                               LayoutRatio(current_x_r + 0.01, current_y_r + 0.005, box_w_r - 0.02, header_h_r - 0.01),
                               bold=True, color_key=title_color, align=MSO_ALIGN_LEFT)
            
            body_text = item['body']
            max_font_size = self.config.layout.body_text_max_font_size
            auto_font_size = self._calculate_auto_font_size(body_text, max_font_size)
            
            self._add_text_box(slide, body_text,
                               LayoutRatio(current_x_r + 0.01, current_y_r + header_h_r + 0.01, 
                                           box_w_r - 0.02, body_h_r - 0.02,
                                           font_size=auto_font_size),
                               color_key='gray_dark')
            
            current_x += dx
            current_y += dy

    def generate(self, df: Any, cover_content: Dict, slides_structure: List[Dict], filename_prefix: str = "Presentation") -> None:
        try:
            print("Engine: 処理開始...")
            self._initialize_ppt()

            # 1. 表紙作成
            print("Engine: 表紙を作成中...")
            slide1 = self.prs.Slides(self.idx_cover)
            self._add_text_box(slide1, cover_content['main_title'], self.config.layout.cover_title, bold=True, align=MSO_ALIGN_CENTER)
            self._add_text_box(slide1, cover_content['sub_title'], self.config.layout.cover_sub, color_key='gray_dark', align=MSO_ALIGN_CENTER)
            self._add_text_box(slide1, cover_content.get('date', datetime.now().strftime("%Y年%m月%d日")), 
                               self.config.layout.cover_date, color_key='gray_dark', align=MSO_ALIGN_CENTER)

            template_slide = self.prs.Slides(self.idx_template_body)
            
            # 2. コンテンツスライド作成ループ
            for i, slide_def in enumerate(slides_structure):
                print(f"Engine: スライド {i+2} ('{slide_def['slide_title']}') を作成中...")
                new_slide = template_slide.Duplicate().Item(1)
                
                target_index = self.prs.Slides.Count - 1
                if target_index < 2: target_index = 2
                new_slide.MoveTo(target_index)

                self._add_text_box(new_slide, slide_def['slide_title'], self.config.layout.content_title, bold=True, align=MSO_ALIGN_RIGHT)

                layout_type = slide_def.get('layout_type', 'horizontal')
                if layout_type == 'vertical':
                    chart_area_ratio = self.config.layout.layout_vertical_chart
                else:
                    chart_area_ratio = self.config.layout.layout_horizontal_chart

                # グラフ生成
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
                        fig.savefig(img_path, dpi=150, bbox_inches='tight')
                        plt.close(fig)
                    
                    self._add_picture_fitted(new_slide, img_path, chart_area_ratio)
                
                content_data = slide_def.get('text_blocks') or slide_def.get('proposal_points')
                if content_data:
                    section_title = slide_def.get('proposal_section_title', "■ 主要な論点・ご提案")
                    self._add_text_content_boxes(new_slide, content_data, 
                                                 layout_type=layout_type, 
                                                 section_header_text=section_title)

            print("Engine: 仕上げ処理中...")
            template_slide.Delete()
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{filename_prefix}_{timestamp}.pptx"
            output_path = os.path.join(self.config.paths.output_dir, filename)
            
            print(f"[DEBUG] 最終ファイルを保存します: {output_path}")
            self.prs.SaveAs(output_path, PP_SAVE_AS_OPENXML_PRESENTATION)
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
                    try: self.ppt_app.WindowState = PP_WINDOW_NORMAL
                    except: pass
                    self.ppt_app.Quit()
                except: pass
            if os.path.exists(self.config.paths.temp_img_dir):
                 try: shutil.rmtree(self.config.paths.temp_img_dir)
                 except: pass