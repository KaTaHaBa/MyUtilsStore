import win32com.client
import os
from datetime import datetime
import shutil
import uuid
# --- Plotly ---
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- PowerPoint定数 ---
msoShapeRectangle = 1
msoTextOrientationHorizontal = 1
msoAlignCenter = 2
msoAlignLeft = 1
msoAlignRight = 3
msoTrue = -1
msoFalse = 0
ppLayoutBlank = 12
ppSaveAsOpenXMLPresentation = 11
ppWindowNormal = 1
ppWindowMinimized = 2
ppWindowMaximized = 3

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
        # HEXコードで定義 (PlotlyもPPT変換ロジックもこれに対応)
        self.palette = {
            'navy': '#002060', 
            'red': '#E60000', 
            'teal': '#008080',
            'gray_text': '#595959', 
            'gray_bg': '#F2F2F2',
            'white': '#FFFFFF', 
            'black': '#000000'
        }

class FontConfig:
    def __init__(self):
        # Plotly用のフォント名
        self.main_font = 'Meiryo' 

class LayoutElement:
    def __init__(self, left, top, width, height, font_size=None):
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.font_size = font_size

class LayoutConfig:
    def __init__(self):
        self.cover_title = LayoutElement(60, 200, 840, 60, font_size=32)
        self.cover_sub = LayoutElement(60, 270, 840, 40, font_size=20)
        self.cover_date = LayoutElement(60, 350, 840, 30, font_size=18)
        self.content_title = LayoutElement(40, 20, 880, 40, font_size=24)
        self.chart_area = LayoutElement(40, 100, 580, 400)
        self.proposal_area = LayoutElement(650, 100, 270, 400)

class SlideConfig:
    def __init__(self, template_path=None, output_dir=None):
        self.paths = PathConfig(template_path=template_path, output_dir=output_dir)
        self.colors = ColorConfig()
        self.fonts = FontConfig()
        self.layout = LayoutConfig()


# ==============================================================================
# B. グラフ描画戦略クラス (Plotly Version)
# ==============================================================================
class ChartStrategies:
    """
    Plotly を使用したグラフ描画ロジック
    """
    def __init__(self, config: SlideConfig):
        self.config = config

    def _get_color(self, color_key):
        """ConfigからHEXカラーコードを取得"""
        return self.config.colors.palette.get(color_key, '#000000')

    def _apply_common_layout(self, fig, chart_text):
        """共通のオシャレなレイアウト設定を適用"""
        fonts = self.config.fonts
        
        fig.update_layout(
            title=dict(
                text=f"<b>{chart_text.get('title', '')}</b>",
                font=dict(family=fonts.main_font, size=18, color=self._get_color('navy')),
                x=0.5, xanchor='center'
            ),
            legend=dict(
                orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
                font=dict(family=fonts.main_font)
            ),
            plot_bgcolor='rgba(0,0,0,0)', 
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family=fonts.main_font, color=self._get_color('gray_text')),
            margin=dict(l=20, r=20, t=60, b=60)
        )
        return fig

    def plot_combo_bar_line_2axis(self, df, mapping, chart_text):
        """
        【カテゴリ: COMBO_BAR_LINE_2AXIS】(Plotly版)
        """
        print(f"  [DEBUG] Plotlyで複合グラフを描画します。")
        
        # 2軸のサブプロット作成
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        x_col = mapping['x_col']

        # 1. 棒グラフ群 (左軸)
        for trace_def in mapping.get('bar_traces', []):
            fig.add_trace(
                go.Bar(
                    x=df[x_col], 
                    y=df[trace_def['col']], 
                    name=trace_def['name'],
                    marker_color=self._get_color(trace_def.get('color_key', 'navy')),
                    opacity=0.8
                ),
                secondary_y=False
            )

        # 2. 折れ線グラフ群 (右軸)
        for trace_def in mapping.get('line_traces', []):
            color_hex = self._get_color(trace_def.get('color_key', 'red'))
            marker_size = trace_def.get('marker_size', 8)
            line_width = trace_def.get('line_width', 2.5)
            
            fig.add_trace(
                go.Scatter(
                    x=df[x_col], 
                    y=df[trace_def['col']], 
                    name=trace_def['name'],
                    mode='lines+markers',
                    line=dict(color=color_hex, width=line_width),
                    marker=dict(size=marker_size, color=color_hex)
                ),
                secondary_y=True
            )

        # 共通レイアウト適用
        fig = self._apply_common_layout(fig, chart_text)

        # 軸ラベル設定
        fig.update_yaxes(title_text=chart_text.get('y1_label', ''), showgrid=True, gridcolor='lightgray', secondary_y=False)
        fig.update_yaxes(title_text=chart_text.get('y2_label', ''), showgrid=False, secondary_y=True)
        fig.update_xaxes(showgrid=False, title_text=chart_text.get('x_label', ''))

        return fig


# ==============================================================================
# C. スライド生成ロジッククラス (Core Engine Class)
# ==============================================================================
class PowerPointGeneratorEngine:
    def __init__(self, config: SlideConfig):
        self.config = config
        self.ppt_app = None
        self.prs = None
        self.stock_slide_index = 2
        self.strategies = ChartStrategies(config)
        
        if os.path.exists(self.config.paths.temp_img_dir):
             shutil.rmtree(self.config.paths.temp_img_dir)
        os.makedirs(self.config.paths.temp_img_dir, exist_ok=True)
        print(f"[DEBUG] 一時画像フォルダを準備しました: {self.config.paths.temp_img_dir}")
        os.makedirs(self.config.paths.output_dir, exist_ok=True)

    def _initialize_ppt(self):
        tpl_path = self.config.paths.template_file
        print(f"[DEBUG] テンプレートファイルを開きます: {tpl_path}")
        if not os.path.exists(tpl_path):
             raise FileNotFoundError(f"テンプレートが見つかりません: {tpl_path}\nパスを確認してください。")
        self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt_app.Visible = msoTrue
        try:
            self.ppt_app.WindowState = ppWindowMinimized
        except:
            pass
        self.prs = self.ppt_app.Presentations.Open(tpl_path, ReadOnly=msoTrue)

    def _to_ppt_rgb(self, color_key_or_tuple):
        """
        修正済み: HEX文字列('#RRGGBB')をPowerPoint用のRGB整数値に正しく変換します。
        """
        val = color_key_or_tuple
        
        # 1. キー名ならHEXコードを取得
        if isinstance(val, str) and not val.startswith('#'):
            val = self.config.colors.palette.get(val, '#000000')

        # 2. HEX文字列('#RRGGBB')ならRGBタプル(r, g, b)に変換
        if isinstance(val, str) and val.startswith('#'):
            hex_color = val.lstrip('#')
            if len(hex_color) == 6:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                val = (r, g, b)
            else:
                val = (0, 0, 0)
        
        # 3. RGBタプルをPPT用整数に変換 (R + G<<8 + B<<16)
        if not isinstance(val, tuple):
            val = (0, 0, 0)

        return val[0] + (val[1] << 8) + (val[2] << 16)

    def _add_text_box(self, slide, text, layout_elem, bold=False, color_key='navy', align=msoAlignLeft):
        shape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal,
                                        layout_elem.left, layout_elem.top,
                                        layout_elem.width, layout_elem.height)
        tf = shape.TextFrame
        tf.TextRange.Text = text
        if layout_elem.font_size: tf.TextRange.Font.Size = layout_elem.font_size
        tf.TextRange.Font.Bold = msoTrue if bold else msoFalse
        tf.TextRange.Font.Color.RGB = self._to_ppt_rgb(color_key)
        tf.TextRange.Font.Name = self.config.fonts.main_font
        tf.TextRange.ParagraphFormat.Alignment = align
        return shape

    def _add_picture(self, slide, image_path, layout_elem):
        print(f"  [DEBUG] PowerPointに画像を貼り付けます: {image_path}")
        if not os.path.exists(image_path):
            print(f"  [ERROR] 貼り付ける画像ファイルが見つかりません！: {image_path}")
            return
        slide.Shapes.AddPicture(image_path, msoFalse, msoTrue,
                                layout_elem.left, layout_elem.top,
                                layout_elem.width, layout_elem.height)

    def _add_proposal_points(self, slide, points_data, section_title_text="■ 主要な論点・ご提案"):
        layout = self.config.layout.proposal_area
        if not points_data: return
        self._add_text_box(slide, section_title_text, LayoutElement(layout.left, layout.top, layout.width, 30, 16), bold=True)
        margin = 15
        box_height = (layout.height - 40 - (len(points_data)-1)*margin) / len(points_data)
        current_top = layout.top + 40
        for point in points_data:
            box = slide.Shapes.AddShape(msoShapeRectangle, layout.left, current_top, layout.width, box_height)
            box.Fill.ForeColor.RGB = self._to_ppt_rgb('gray_bg')
            box.Line.Visible = msoFalse
            accent_color = point.get('accent_color_key', 'red')
            accent = slide.Shapes.AddShape(msoShapeRectangle, layout.left, current_top, layout.width, 30)
            accent.Fill.ForeColor.RGB = self._to_ppt_rgb(accent_color)
            accent.Line.Visible = msoFalse
            
            # 文字色指定の適用
            title_color = point.get('title_color_key', 'white')
            
            self._add_text_box(slide, point['title'], LayoutElement(layout.left+10, current_top, layout.width-20, 30, 14), bold=True, color_key=title_color, align=msoAlignCenter)
            body_box = self._add_text_box(slide, point['body'], LayoutElement(layout.left+15, current_top+35, layout.width-30, box_height-40, 11), color_key='gray_text')
            body_box.TextFrame.WordWrap = msoTrue
            current_top += box_height + margin

    def generate(self, df, cover_content, slides_structure, filename_prefix="Presentation"):
        try:
            print("Engine: 処理開始...")
            self._initialize_ppt()

            print("Engine: 表紙を作成中...")
            slide1 = self.prs.Slides(1)
            self._add_text_box(slide1, cover_content['main_title'], self.config.layout.cover_title, bold=True, align=msoAlignCenter)
            self._add_text_box(slide1, cover_content['sub_title'], self.config.layout.cover_sub, color_key='gray_text', align=msoAlignCenter)
            self._add_text_box(slide1, cover_content.get('date', datetime.now().strftime("%Y年%m月%d日")), 
                               self.config.layout.cover_date, color_key='gray_text', align=msoAlignCenter)

            stock_slide = self.prs.Slides(self.stock_slide_index)
            
            for i, slide_def in enumerate(slides_structure):
                print(f"Engine: スライド {i+2} ('{slide_def['slide_title']}') を作成中...")
                new_slide = stock_slide.Duplicate().Item(1)
                new_slide.MoveTo(self.prs.Slides.Count)
                self._add_text_box(new_slide, slide_def['slide_title'], self.config.layout.content_title, bold=True)

                category = slide_def['category']
                mapping = slide_def['data_mapping']
                chart_text = slide_def['chart_text']
                
                plot_method_name = f"plot_{category.lower()}"
                print(f"  [DEBUG] グラフ描画メソッドを探します: {plot_method_name}")
                plot_method = getattr(self.strategies, plot_method_name, None)
                
                if plot_method:
                    print(f"  [DEBUG] メソッドが見つかりました。グラフ描画を開始します。")
                    # PlotlyのFigureオブジェクトを取得
                    fig = plot_method(df, mapping, chart_text)
                    
                    img_filename = f"chart_{uuid.uuid4()}.png"
                    img_path = os.path.join(self.config.paths.temp_img_dir, img_filename)
                    
                    print(f"  [DEBUG] グラフ画像を保存しようとしています... パス: {img_path}")
                    # --- Plotlyでの保存 (Kaleido) ---
                    fig.write_image(img_path, width=800, height=500, scale=2)
                    # ------------------------------
                    
                    if os.path.exists(img_path):
                        print(f"  [DEBUG] ✅ 画像保存に成功しました！ ファイルサイズ: {os.path.getsize(img_path)} bytes")
                    else:
                        print(f"  [ERROR] ❌ 画像保存に失敗したようです。ファイルが存在しません。")

                    self._add_picture(new_slide, img_path, self.config.layout.chart_area)
                else:
                    print(f"Warning: Unknown chart category '{category}'. Skipping chart generation.")

                if 'proposal_points' in slide_def:
                    section_title = slide_def.get('proposal_section_title', "■ 主要な論点・ご提案")
                    self._add_proposal_points(new_slide, slide_def['proposal_points'], section_title_text=section_title)

            print("Engine: 仕上げ処理中...")
            self.prs.Slides(self.stock_slide_index).Delete()
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{filename_prefix}_{timestamp}.pptx"
            output_path = os.path.join(self.config.paths.output_dir, filename)
            print(f"[DEBUG] 最終ファイルを保存します: {output_path}")
            self.prs.SaveAs(output_path, ppSaveAsOpenXMLPresentation)
            print(f"Engine: 完了！ファイルが出力されました: {output_path}")
            
            print("[DEBUG] 保存を確実にするため、ファイルを閉じます...")
            self.prs.Close()
            self.prs = None

        except Exception as e:
            print(f"Engine Error: {e}")
            import traceback
            traceback.print_exc()
        
        finally:
            print("Engine: PowerPointを終了します...")
            if self.prs:
                try:
                    self.prs.Close()
                except: pass
            if self.ppt_app:
                try:
                    try: self.ppt_app.WindowState = ppWindowNormal
                    except: pass
                    self.ppt_app.Quit()
                except: pass
            
            if os.path.exists(self.config.paths.temp_img_dir):
                 print(f"[DEBUG] 一時画像フォルダを削除します: {self.config.paths.temp_img_dir}")
                 shutil.rmtree(self.config.paths.temp_img_dir)