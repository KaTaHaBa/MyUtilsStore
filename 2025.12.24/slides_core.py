
from __future__ import annotations

import os
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple, Union

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import pandas as pd
import seaborn as sns
import win32com.client
from PIL import Image

# --- PowerPoint constants ---
MSO_SHAPE_RECTANGLE = 1
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_ALIGN_LEFT = 1
MSO_ALIGN_CENTER = 2
MSO_ALIGN_RIGHT = 3
MSO_TRUE = -1
MSO_FALSE = 0
PP_SAVE_AS_OPENXML_PRESENTATION = 11
PP_SAVE_AS_PDF = 32
PP_WINDOW_NORMAL = 1
PP_WINDOW_MINIMIZED = 2


# ==============================
# A. Configuration
# ==============================

class PathConfig:
    """Resolve template and output paths used by the slide generator."""

    def __init__(
        self,
        template_path: Optional[str] = None,
        output_dir: Optional[str] = None,
        temp_dir_name: str = "temp_images_slide_gen",
    ) -> None:
        base_dir = os.getcwd()
        default_template = os.path.join(base_dir, "template", "template_16-9.pptx")
        fallback_template = os.path.join(base_dir, "template.pptx")

        if template_path:
            self.template_file = template_path
        elif os.path.exists(default_template):
            self.template_file = default_template
        else:
            self.template_file = fallback_template

        self.output_dir = output_dir if output_dir else os.path.join(base_dir, "output")
        self.temp_img_dir = os.path.join(base_dir, temp_dir_name)


class ColorConfig:
    """Define color palette by descriptive names."""

    def __init__(self) -> None:
        self.palette: Dict[str, str] = {
            "black": "#000000",
            "white": "#FFFFFF",
            "gray_dark": "#595959",
            "gray_medium": "#7F8C8D",
            "gray_light": "#F2F2F2",
            "navy": "#002060",
            "midnight_blue": "#1F3A5F",
            "teal": "#008080",
            "cadet_blue": "#5f9ea0",
            "sky_blue": "#a8d8ea",
            "ice_blue": "#e0f2f1",
            "red": "#E60000",
            "salmon": "#f38181",
            "pale_pink": "#fcbad3",
            "coral": "#ff6f69",
            "gold": "#C2A970",
            "cream": "#ffffd2",
            "mustard": "#ffcc5c",
            "spring_green": "#42e6a4",
            "mint": "#95e1d3",
            "lavender": "#aa96da",
            "dark_slate": "#2C3E50",
        }


class FontConfig:
    """Store default font and size settings."""

    def __init__(self) -> None:
        self.japanese_font: str = "Meiryo"
        self.english_font: str = "Segoe UI"
        plt.rcParams["font.family"] = self.japanese_font
        self.chart_title_size: int = 18
        self.chart_label_size: int = 12
        self.chart_tick_size: int = 11


class LayoutRatio:
    """Define a slide region by relative coordinates."""

    def __init__(
        self,
        left: float,
        top: float,
        width: float,
        height: float,
        font_size: Optional[int] = None,
    ) -> None:
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.font_size = font_size


class LayoutConfig:
    """Define layout ratios for cover and body slides."""

    def __init__(self, split_ratio: float = 0.5) -> None:
        self.split_ratio = split_ratio

        self.cover_title = LayoutRatio(0.1, 0.35, 0.8, 0.15, font_size=32)
        self.cover_sub = LayoutRatio(0.1, 0.52, 0.8, 0.10, font_size=20)
        self.cover_date = LayoutRatio(0.1, 0.85, 0.8, 0.05, font_size=18)
        self.content_title = LayoutRatio(0.40, 0.05, 0.55, 0.08, font_size=24)

        margin_outer = 0.05
        margin_inner = 0.02
        title_area_h = 0.15

        # Horizontal split
        h_chart_w = self.split_ratio - margin_outer - (margin_inner / 2)
        self.layout_horizontal_chart = LayoutRatio(margin_outer, title_area_h + 0.03, h_chart_w, 0.75)
        h_text_x = self.split_ratio + (margin_inner / 2)
        h_text_w = (1.0 - margin_outer) - h_text_x
        self.layout_horizontal_text = LayoutRatio(h_text_x, title_area_h + 0.03, h_text_w, 0.75)

        # Vertical split
        content_total_h = 1.0 - title_area_h - margin_outer
        v_chart_h = content_total_h * self.split_ratio
        self.layout_vertical_chart = LayoutRatio(margin_outer, title_area_h, 1.0 - (margin_outer * 2), v_chart_h)
        v_text_y = title_area_h + v_chart_h + margin_inner
        v_text_h = content_total_h * (1.0 - self.split_ratio) - margin_inner
        self.layout_vertical_text = LayoutRatio(margin_outer, v_text_y, 1.0 - (margin_outer * 2), v_text_h)

        self.body_text_max_font_size = 16


class SlideConfig:
    """Bundle slide configuration."""

    def __init__(
        self,
        template_path: Optional[str] = None,
        output_dir: Optional[str] = None,
        engine: str = "matplotlib",
        split_ratio: float = 0.5,
    ) -> None:
        self.paths = PathConfig(template_path=template_path, output_dir=output_dir)
        self.colors = ColorConfig()
        self.fonts = FontConfig()
        self.layout = LayoutConfig(split_ratio=split_ratio)
        self.engine = engine

# ==============================
# A2. Page models
# ==============================

@dataclass(frozen=True)
class SlidePageConfig:
    """Optional per-page configuration overrides."""

    colors: Optional[Dict[str, str]] = None
    fonts: Optional[Dict[str, Any]] = None
    layout: Optional[Dict[str, Any]] = None
    split_ratio: Optional[float] = None


@dataclass(frozen=True)
class SlidePage:
    """A single slide definition for UI-driven generation."""

    slide_title: str
    category: str
    data_mapping: Dict[str, Any]
    chart_text: Dict[str, Any]
    layout_type: str = "horizontal"
    data_source: Optional[str] = None
    data_frame: Optional[Any] = None
    data_columns: Optional[List[str]] = None
    text_blocks: Optional[List[Dict[str, Any]]] = None
    proposal_section_title: Optional[str] = None
    config: Optional[SlidePageConfig] = None


@dataclass(frozen=True)
class SlideCover:
    """Cover page content."""

    main_title: str
    sub_title: str
    date: str


@dataclass(frozen=True)
class SlideDeck:
    """Slide deck definition with cover and ordered pages."""

    cover: SlideCover
    pages: List[SlidePage]

# ==============================
# B. Data preparation helpers
# ==============================

@dataclass(frozen=True)
class SeriesSpec:
    """Define how a canonical key is exposed to charts."""

    canonical_key: str
    column_name: str
    color_key: str


_PERIOD_LABEL_RE = re.compile(r"(\d{4}-\d{2}-\d{2})-(\d{4}-\d{2}-\d{2})")


def _parse_period_label(label: str) -> Optional[Tuple[datetime, datetime]]:
    """Parse a period label in the form start-end."""
    match = _PERIOD_LABEL_RE.match(str(label))
    if not match:
        return None
    start_dt = datetime.strptime(match.group(1), "%Y-%m-%d")
    end_dt = datetime.strptime(match.group(2), "%Y-%m-%d")
    return start_dt, end_dt


def _format_period_label_fy(label: str) -> str:
    """Format a period label into FY{YYYY} using the period end year."""
    parsed = _parse_period_label(label)
    if not parsed:
        return str(label)
    end_dt = parsed[1]
    return f"FY{end_dt.year}"


def _get_period_columns(df: Any) -> List[str]:
    """Return period columns in the wide canonical tables."""
    if df is None or len(df) == 0:
        return []
    reserved = {"canonical_key", "label", "statement"}
    return [col for col in df.columns if col not in reserved]


def _select_latest_period(period_cols: Iterable[str]) -> Optional[str]:
    """Select the latest period label using the end date when possible."""
    parsed = []
    for col in period_cols:
        dates = _parse_period_label(col)
        if dates:
            parsed.append((col, dates[1]))
    if parsed:
        return sorted(parsed, key=lambda x: x[1])[-1][0]
    cols = list(period_cols)
    return sorted(cols)[-1] if cols else None


def build_trend_dataframe(df_wide: Any, series_specs: List[SeriesSpec]) -> Any:
    """Build a trend table with period labels as rows and metrics as columns."""
    if df_wide is None or len(df_wide) == 0:
        cols = ["period_label"] + [spec.column_name for spec in series_specs]
        return pd.DataFrame(columns=cols)
    period_cols = _get_period_columns(df_wide)
    data: Dict[str, List[Any]] = {"period_label": period_cols}

    for spec in series_specs:
        row = df_wide[df_wide["canonical_key"] == spec.canonical_key]
        if row.empty:
            data[spec.column_name] = [None] * len(period_cols)
            continue
        row = row.iloc[0]
        data[spec.column_name] = [row.get(col) for col in period_cols]

    return pd.DataFrame(data)


def build_snapshot_dataframe(
    df_wide: Any,
    series_specs: List[SeriesSpec],
    period_label: Optional[str] = None,
) -> Any:
    """Build a single-row snapshot table for a selected period."""
    if df_wide is None or len(df_wide) == 0:
        cols = ["period_label"] + [spec.column_name for spec in series_specs]
        return pd.DataFrame(columns=cols)
    period_cols = _get_period_columns(df_wide)
    target_period = period_label or _select_latest_period(period_cols)
    if not period_label and period_cols:
        def _has_numeric_value(row: pd.DataFrame, col: str) -> bool:
            if row.empty:
                return False
            val = pd.to_numeric(row.iloc[0].get(col), errors="coerce")
            return pd.notna(val)

        ordered_cols = sorted(
            period_cols,
            key=lambda col: _parse_period_label(col)[1] if _parse_period_label(col) else datetime.min,
            reverse=True,
        )
        preferred_row = (
            df_wide[df_wide["canonical_key"] == "TotalAssets"]
            if any(spec.column_name == "Total Assets" for spec in series_specs)
            else pd.DataFrame()
        )
        preferred_cols = [col for col in ordered_cols if _has_numeric_value(preferred_row, col)]
        if preferred_cols:
            target_period = preferred_cols[0]
        else:
            candidate = None
            for col in ordered_cols:
                has_value = False
                for spec in series_specs:
                    row = df_wide[df_wide["canonical_key"] == spec.canonical_key]
                    if _has_numeric_value(row, col):
                        has_value = True
                        break
                if has_value:
                    candidate = col
                    break
            if candidate:
                target_period = candidate
    row_data: Dict[str, Any] = {"period_label": target_period}

    for spec in series_specs:
        row = df_wide[df_wide["canonical_key"] == spec.canonical_key]
        if row.empty or not target_period:
            row_data[spec.column_name] = None
            continue
        row_data[spec.column_name] = row.iloc[0].get(target_period)

    return pd.DataFrame([row_data])


def build_slide_inputs_from_layout(
    df_pl: Any,
    df_bs: Any,
    df_cf: Any,
    layout: Dict[str, Any],
    company_name: str,
) -> Tuple[Dict[str, Any], Dict[str, Any], List[Dict[str, Any]]]:
    """Prepare data store, cover content, and slide definitions using a layout dict."""
    pl_series = [SeriesSpec(item["canonical_key"], item["column_name"], item.get("color_key", "navy")) for item in layout.get("pl_series", [])]
    cf_series = [SeriesSpec(item["canonical_key"], item["column_name"], item.get("color_key", "navy")) for item in layout.get("cf_series", [])]
    bs_series = [SeriesSpec(item["canonical_key"], item["column_name"], item.get("color_key", "navy")) for item in layout.get("bs_series", [])]

    pl_trend = build_trend_dataframe(df_pl, pl_series)
    cf_trend = build_trend_dataframe(df_cf, cf_series)
    bs_snapshot = build_snapshot_dataframe(df_bs, bs_series)

    data_store = {
        "pl_trend": pl_trend,
        "cf_trend": cf_trend,
        "bs_snapshot": bs_snapshot,
    }

    cover_content = {
        "main_title": f"{company_name} Financial Review",
        "sub_title": "Generated from XBRL",
        "date": datetime.now().strftime("%Y-%m-%d"),
    }

    pl_chart = layout.get("pl_chart", {})
    cf_chart = layout.get("cf_chart", {})
    bs_chart = layout.get("bs_chart", {})

    def _trace_list(series_defs: List[SeriesSpec], keys: List[str]) -> List[Dict[str, Any]]:
        color_map = {spec.column_name: spec.color_key for spec in series_defs}
        return [{"col": key, "name": key, "color_key": color_map.get(key, "navy")} for key in keys if key in color_map]

    slides_structure = [
        {
            "slide_title": pl_chart.get("slide_title", "PL"),
            "layout_type": "horizontal",
            "category": pl_chart.get("category", "combo_bar_line_2axis"),
            "data_source": "pl_trend",
            "data_mapping": {
                "x_col": pl_chart.get("x_col", "period_label"),
                "x_label_format": pl_chart.get("x_label_format", "fy"),
                "unit_scale": pl_chart.get("unit_scale", 1e9),
                "bar_traces": _trace_list(pl_series, pl_chart.get("bar_keys", [])),
                "line_traces": _trace_list(pl_series, pl_chart.get("line_keys", [])),
            },
            "chart_text": pl_chart.get("chart_text", {}),
        },
        {
            "slide_title": cf_chart.get("slide_title", "CF"),
            "layout_type": "horizontal",
            "category": cf_chart.get("category", "combo_bar_line_2axis"),
            "data_source": "cf_trend",
            "data_mapping": {
                "x_col": cf_chart.get("x_col", "period_label"),
                "x_label_format": cf_chart.get("x_label_format", "fy"),
                "unit_scale": cf_chart.get("unit_scale", 1e9),
                "bar_traces": _trace_list(cf_series, cf_chart.get("bar_keys", [])),
                "line_traces": _trace_list(cf_series, cf_chart.get("line_keys", [])),
            },
            "chart_text": cf_chart.get("chart_text", {}),
        },
        {
            "slide_title": bs_chart.get("slide_title", "BS"),
            "layout_type": "horizontal",
            "category": bs_chart.get("category", "balance_sheet"),
            "data_source": "bs_snapshot",
            "data_mapping": bs_chart,
            "chart_text": bs_chart.get("chart_text", {}),
        },
    ]

    return data_store, cover_content, slides_structure

# ==============================
# B2. Page helpers
# ==============================

def _update_layout_ratio(layout_ratio: LayoutRatio, data: Dict[str, Any]) -> None:
    """Update a LayoutRatio in-place from a dictionary."""
    if "left" in data:
        layout_ratio.left = float(data["left"])
    if "top" in data:
        layout_ratio.top = float(data["top"])
    if "width" in data:
        layout_ratio.width = float(data["width"])
    if "height" in data:
        layout_ratio.height = float(data["height"])
    if "font_size" in data:
        layout_ratio.font_size = data["font_size"]


def _apply_layout_overrides(layout: LayoutConfig, overrides: Dict[str, Any]) -> None:
    """Apply layout overrides for a single page."""
    if "content_title" in overrides:
        _update_layout_ratio(layout.content_title, overrides["content_title"])
    if "layout_horizontal_chart" in overrides:
        _update_layout_ratio(layout.layout_horizontal_chart, overrides["layout_horizontal_chart"])
    if "layout_horizontal_text" in overrides:
        _update_layout_ratio(layout.layout_horizontal_text, overrides["layout_horizontal_text"])
    if "layout_vertical_chart" in overrides:
        _update_layout_ratio(layout.layout_vertical_chart, overrides["layout_vertical_chart"])
    if "layout_vertical_text" in overrides:
        _update_layout_ratio(layout.layout_vertical_text, overrides["layout_vertical_text"])
    if "body_text_max_font_size" in overrides:
        layout.body_text_max_font_size = int(overrides["body_text_max_font_size"])


def _build_slide_config(base_config: SlideConfig, page_config: Optional[Union[SlidePageConfig, Dict[str, Any]]]) -> SlideConfig:
    """Build a per-page SlideConfig by applying overrides."""
    if page_config is None:
        return base_config

    if isinstance(page_config, SlidePageConfig):
        config_data = {
            "colors": page_config.colors,
            "fonts": page_config.fonts,
            "layout": page_config.layout,
            "split_ratio": page_config.split_ratio,
        }
    else:
        config_data = page_config

    split_ratio = config_data.get("split_ratio", base_config.layout.split_ratio)
    new_config = SlideConfig(
        template_path=base_config.paths.template_file,
        output_dir=base_config.paths.output_dir,
        engine=base_config.engine,
        split_ratio=split_ratio,
    )

    # Copy base palette and apply overrides
    new_config.colors.palette.update(base_config.colors.palette)
    if config_data.get("colors"):
        new_config.colors.palette.update(config_data["colors"])

    # Copy fonts and apply overrides
    new_config.fonts.japanese_font = base_config.fonts.japanese_font
    new_config.fonts.english_font = base_config.fonts.english_font
    new_config.fonts.chart_title_size = base_config.fonts.chart_title_size
    new_config.fonts.chart_label_size = base_config.fonts.chart_label_size
    new_config.fonts.chart_tick_size = base_config.fonts.chart_tick_size

    font_overrides = config_data.get("fonts") or {}
    if "japanese_font" in font_overrides:
        new_config.fonts.japanese_font = font_overrides["japanese_font"]
    if "english_font" in font_overrides:
        new_config.fonts.english_font = font_overrides["english_font"]
    if "chart_title_size" in font_overrides:
        new_config.fonts.chart_title_size = int(font_overrides["chart_title_size"])
    if "chart_label_size" in font_overrides:
        new_config.fonts.chart_label_size = int(font_overrides["chart_label_size"])
    if "chart_tick_size" in font_overrides:
        new_config.fonts.chart_tick_size = int(font_overrides["chart_tick_size"])

    layout_overrides = config_data.get("layout")
    if layout_overrides:
        _apply_layout_overrides(new_config.layout, layout_overrides)

    return new_config


def _resolve_slide_def(page: Union[SlidePage, Dict[str, Any]]) -> Dict[str, Any]:
    """Normalize slide definitions to a dictionary."""
    if isinstance(page, SlidePage):
        return {
            "slide_title": page.slide_title,
            "category": page.category,
            "data_mapping": page.data_mapping,
            "chart_text": page.chart_text,
            "layout_type": page.layout_type,
            "data_source": page.data_source,
            "data_frame": page.data_frame,
            "data_columns": page.data_columns,
            "text_blocks": page.text_blocks,
            "proposal_section_title": page.proposal_section_title,
            "config": page.config,
        }
    return page

# ==============================
# C. Chart strategy
# ==============================

class ChartStrategyBase:
    """Base class for chart rendering strategies."""

    def __init__(self, config: SlideConfig) -> None:
        self.config = config

    def _get_color(self, color_key: str) -> str:
        """Resolve color key to hex code."""
        return self.config.colors.palette.get(color_key, "#000000")

    def plot_combo_bar_line_2axis(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> Any:
        raise NotImplementedError

    def plot_balance_sheet(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> Any:
        raise NotImplementedError

    def plot_portfolio_timeseries(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> Any:
        raise NotImplementedError


def jpy_currency_formatter(x: float, pos: Any) -> str:
    """Format numbers for JPY charts with ASCII-friendly units."""
    if x == 0:
        return "0"
    abs_x = abs(x)
    if abs_x >= 1e12:
        return f"{x/1e12:,.1f}T"
    if abs_x >= 1e9:
        return f"{x/1e9:,.1f}B"
    if abs_x >= 1e6:
        return f"{x/1e6:,.0f}M"
    if abs_x >= 1e3:
        return f"{x/1e3:,.0f}K"
    return f"{x:,.0f}"


def scaled_number_formatter(x: float, pos: Any) -> str:
    """Format scaled values with a single decimal place."""
    return f"{x:,.1f}"


class MatplotlibStrategy(ChartStrategyBase):
    """Matplotlib-based rendering strategy."""

    def __init__(self, config: SlideConfig) -> None:
        super().__init__(config)
        sns.set_theme(style="white", rc={"axes.grid": False})
        plt.rcParams["font.family"] = config.fonts.japanese_font

    def _apply_common_style(self, fig: plt.Figure, ax: plt.Axes, chart_text: Dict[str, Any]) -> Tuple[plt.Figure, plt.Axes]:
        fonts = self.config.fonts
        ax.set_title(
            chart_text.get("title", ""),
            fontsize=fonts.chart_title_size,
            fontweight="bold",
            color=self._get_color("navy"),
            pad=20,
        )
        for axis in [ax.xaxis, ax.yaxis]:
            axis.label.set_color(self._get_color("gray_dark"))
            axis.label.set_fontsize(fonts.chart_label_size)
            axis.set_tick_params(colors=self._get_color("gray_dark"), labelsize=fonts.chart_tick_size)
        sns.despine(left=True, bottom=True)
        ax.yaxis.grid(True, color="#E0E0E0", linestyle="--", linewidth=0.5)
        return fig, ax

    def plot_combo_bar_line_2axis(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> plt.Figure:
        fig, ax1 = plt.subplots(figsize=(12, 7))
        x_col = mapping["x_col"]
        df_plot = df.copy()

        if mapping.get("x_label_format") == "fy":
            df_plot["_x_label"] = df_plot[x_col].map(_format_period_label_fy)
            x_col = "_x_label"

        unit_scale = float(mapping.get("unit_scale", 1.0) or 1.0)

        bar_cols = [t["col"] for t in mapping.get("bar_traces", [])]
        if bar_cols:
            if unit_scale != 1.0:
                for col in bar_cols:
                    if col in df_plot.columns:
                        df_plot[col] = df_plot[col] / unit_scale
            df_bar_melt = df_plot.melt(id_vars=[x_col], value_vars=bar_cols, var_name="Metric", value_name="Value")
            bar_palette = {t["col"]: self._get_color(t.get("color_key", "navy")) for t in mapping.get("bar_traces", [])}
            sns.barplot(data=df_bar_melt, x=x_col, y="Value", hue="Metric", palette=bar_palette, ax=ax1, alpha=0.85, dodge=True)

            handles, labels = ax1.get_legend_handles_labels()
            trace_name_map = {t["col"]: t["name"] for t in mapping.get("bar_traces", [])}
            new_labels = [trace_name_map.get(l, l) for l in labels]
            ax1.legend(
                handles,
                new_labels,
                loc="upper left",
                bbox_to_anchor=(0, -0.15),
                ncol=len(bar_cols),
                frameon=False,
                fontsize=self.config.fonts.chart_tick_size,
            )
            ax1.set_ylabel(chart_text.get("y1_label", ""), fontsize=self.config.fonts.chart_label_size)
            formatter = scaled_number_formatter if unit_scale != 1.0 else jpy_currency_formatter
            ax1.yaxis.set_major_formatter(ticker.FuncFormatter(formatter))

        line_traces = mapping.get("line_traces", [])
        if line_traces:
            ax2 = ax1.twinx()
            if unit_scale != 1.0:
                for trace_def in line_traces:
                    col = trace_def["col"]
                    if col in df_plot.columns:
                        df_plot[col] = df_plot[col] / unit_scale
            for trace_def in line_traces:
                color = self._get_color(trace_def.get("color_key", "red"))
                marker_size = trace_def.get("marker_size", 10)
                line_width = trace_def.get("line_width", 3.5)
                sns.lineplot(
                    data=df_plot,
                    x=x_col,
                    y=trace_def["col"],
                    ax=ax2,
                    color=color,
                    marker="o",
                    markersize=marker_size,
                    linewidth=line_width,
                    label=trace_def["name"],
                )
            ax2.set_ylabel(
                chart_text.get("y2_label", ""),
                fontsize=self.config.fonts.chart_label_size,
                color=self._get_color("gray_dark"),
            )
            ax2.tick_params(axis="y", colors=self._get_color("gray_dark"), labelsize=self.config.fonts.chart_tick_size)
            formatter = scaled_number_formatter if unit_scale != 1.0 else jpy_currency_formatter
            ax2.yaxis.set_major_formatter(ticker.FuncFormatter(formatter))
            sns.despine(ax=ax2, right=False, left=True, bottom=True)
            ax2.legend(
                loc="upper left",
                bbox_to_anchor=(0.5, -0.15),
                ncol=len(line_traces),
                frameon=False,
                fontsize=self.config.fonts.chart_tick_size,
            )
            ax2.grid(False)

        x_label_rotation = float(mapping.get("x_label_rotation", 0))
        tick_step = mapping.get("x_tick_step")
        if tick_step is None:
            tick_step = 2 if len(ax1.get_xticklabels()) > 12 else 1
        for idx, label in enumerate(ax1.get_xticklabels()):
            label.set_rotation(x_label_rotation)
            label.set_ha("right" if x_label_rotation else "center")
            if tick_step and idx % int(tick_step) != 0:
                label.set_visible(False)

        ax1.set_xlabel("")
        fig, ax1 = self._apply_common_style(fig, ax1, chart_text)
        plt.tight_layout()
        return fig

    def plot_portfolio_timeseries(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> plt.Figure:
        if len(df) == 0:
            raise ValueError("DataFrame is empty.")

        fig, ax = plt.subplots(figsize=(12, 7))
        df_plot = df.copy()
        x_col = mapping.get("x_col", "period_label")

        if mapping.get("x_label_format") == "fy":
            df_plot["_x_label"] = df_plot[x_col].map(_format_period_label_fy)
            x_col = "_x_label"

        unit_scale = float(mapping.get("unit_scale", 1.0) or 1.0)
        series_defs = mapping.get("series", [])
        if not series_defs:
            raise ValueError("Portfolio series definitions are empty.")

        for spec in series_defs:
            col = spec.get("col")
            if col in df_plot.columns and unit_scale != 1.0:
                df_plot[col] = df_plot[col] / unit_scale

        area_defs = [spec for spec in series_defs if spec.get("chart_type", "area") != "line"]
        line_defs = [spec for spec in series_defs if spec.get("chart_type", "area") == "line"]

        chart_style = mapping.get("chart_style", "area")
        x_labels = df_plot[x_col].tolist()

        if chart_style == "stacked_bar":
            x_vals = list(range(len(df_plot)))
            bar_width = float(mapping.get("bar_width", 0.6))
            bottom = [0] * len(df_plot)
            for spec in area_defs:
                col = spec.get("col")
                if not col or col not in df_plot.columns:
                    values = [0] * len(df_plot)
                else:
                    values = pd.to_numeric(df_plot[col], errors="coerce").fillna(0).tolist()
                color = self._get_color(spec.get("color_key", "navy"))
                ax.bar(x_vals, values, bottom=bottom, color=color, width=bar_width, alpha=0.85)
                bottom = [b + v for b, v in zip(bottom, values)]
        else:
            x_vals = df_plot[x_col]
            if area_defs:
                area_values = []
                for spec in area_defs:
                    col = spec.get("col")
                    if not col or col not in df_plot.columns:
                        area_values.append([0] * len(df_plot))
                        continue
                    area_values.append(pd.to_numeric(df_plot[col], errors="coerce").fillna(0).values)
                area_colors = [self._get_color(spec.get("color_key", "navy")) for spec in area_defs]
                ax.stackplot(x_vals, area_values, colors=area_colors, alpha=0.85)

        for spec in line_defs:
            col = spec.get("col")
            if not col or col not in df_plot.columns:
                continue
            color = self._get_color(spec.get("color_key", "red"))
            marker_size = spec.get("marker_size", 9)
            line_width = spec.get("line_width", 3.0)
            ax.plot(
                x_vals,
                df_plot[col],
                color=color,
                marker="o",
                markersize=marker_size,
                linewidth=line_width,
                label=spec.get("name", col),
            )

        formatter = scaled_number_formatter if unit_scale != 1.0 else jpy_currency_formatter
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(formatter))
        ax.set_ylabel(chart_text.get("y1_label", ""), fontsize=self.config.fonts.chart_label_size)

        legend_handles = []
        legend_labels = []
        for spec in area_defs:
            name = spec.get("name", spec["col"])
            legend_handles.append(mpatches.Patch(color=self._get_color(spec.get("color_key", "navy")), label=name))
            legend_labels.append(name)
        for spec in line_defs:
            name = spec.get("name", spec["col"])
            legend_handles.append(Line2D([0], [0], color=self._get_color(spec.get("color_key", "red")), marker="o", linewidth=2))
            legend_labels.append(name)
        if legend_handles:
            ax.legend(
                legend_handles,
                legend_labels,
                loc="upper left",
                bbox_to_anchor=(0, -0.15),
                ncol=max(1, len(legend_handles)),
                frameon=False,
                fontsize=self.config.fonts.chart_tick_size,
            )

        if chart_style == "stacked_bar":
            ax.set_xticks(x_vals)
            ax.set_xticklabels(x_labels)

        x_label_rotation = float(mapping.get("x_label_rotation", 0))
        tick_step = mapping.get("x_tick_step")
        if tick_step is None:
            tick_step = 2 if len(ax.get_xticklabels()) > 12 else 1
        for idx, label in enumerate(ax.get_xticklabels()):
            label.set_rotation(x_label_rotation)
            label.set_ha("right" if x_label_rotation else "center")
            if tick_step and idx % int(tick_step) != 0:
                label.set_visible(False)

        fig, ax = self._apply_common_style(fig, ax, chart_text)
        plt.tight_layout()
        return fig

    def plot_balance_sheet(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> plt.Figure:
        if len(df) == 0:
            raise ValueError("DataFrame is empty.")
        row = df.iloc[0]

        unit_scale = mapping.get("unit_scale", 1.0)
        total_assets_col = mapping.get("total_assets_col")
        total_assets_val = row.get(total_assets_col, 0)
        if pd.isna(total_assets_val):
            total_assets_val = 0
        total_assets_val = float(total_assets_val)
        auto_balance_assets = mapping.get("auto_balance_assets", False)
        auto_balance_liab_equity = mapping.get("auto_balance_liab_equity", False)
        other_assets_label = mapping.get("other_assets_label", "Other Assets")
        other_liab_equity_label = mapping.get("other_liab_equity_label", "Other Liab/Equity")

        def _build_stack_data(stack_def: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
            data = []
            for item in stack_def:
                val = row.get(item["col"], 0)
                val = pd.to_numeric(val, errors="coerce")
                if pd.isna(val):
                    val = 0
                data.append(
                    {
                        "label": item["name"],
                        "value": val / unit_scale,
                        "color": self._get_color(item.get("color_key", "gray_medium")),
                    }
                )
            return data

        def _select_stack_data(
            candidates: List[Tuple[str, List[Dict[str, Any]]]],
            total_target: float,
            preference: Optional[str],
            prefer_detail: bool,
        ) -> Tuple[str, List[Dict[str, Any]]]:
            balance_tolerance = float(mapping.get("balance_tolerance", 0.01))
            scored: List[Tuple[float, int, str, List[Dict[str, Any]]]] = []

            if preference in {"primary", "bank", "summary"}:
                preferred_stack = next((stack for name, stack in candidates if name == preference and stack), [])
                if preferred_stack:
                    preferred_data = _build_stack_data(preferred_stack)
                    preferred_total = sum(d["value"] for d in preferred_data)
                    if preferred_total > 0:
                        if total_target > 0:
                            gap_ratio = abs(total_target - preferred_total) / total_target
                            if gap_ratio <= balance_tolerance:
                                return preference, preferred_data
                        else:
                            return preference, preferred_data

            for name, stack_def in candidates:
                data = _build_stack_data(stack_def)
                total = sum(d["value"] for d in data)
                if total <= 0:
                    continue
                if total_target > 0:
                    gap_ratio = abs(total_target - total) / total_target
                else:
                    gap_ratio = 0
                nonzero_count = sum(1 for d in data if d["value"] > 0)
                scored.append((gap_ratio, nonzero_count, name, data))

            if not scored:
                return candidates[0][0], _build_stack_data(candidates[0][1])

            if prefer_detail:
                scored.sort(key=lambda x: (-x[1], x[0]))
                best = scored[0]
                if total_target > 0 and best[0] > balance_tolerance:
                    scored.sort(key=lambda x: (x[0], -x[1]))
                    return scored[0][2], scored[0][3]
                return best[2], best[3]

            within = [item for item in scored if item[0] <= balance_tolerance]
            if within:
                within.sort(key=lambda x: (-x[1], x[0]))
                return within[0][2], within[0][3]

            scored.sort(key=lambda x: (x[0], -x[1]))
            return scored[0][2], scored[0][3]

        def _apply_exclusive_groups(stack_def: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
            exclusive_groups = mapping.get("exclusive_groups", [])
            if not exclusive_groups:
                return stack_def
            drop_labels = set()
            for group in exclusive_groups:
                aggregate = group.get("aggregate")
                components = group.get("components", [])
                if not aggregate or not components:
                    continue
                has_component = False
                for name in components:
                    val = pd.to_numeric(row.get(name, 0), errors="coerce")
                    if pd.notna(val) and float(val) != 0:
                        has_component = True
                        break
                if has_component:
                    drop_labels.add(aggregate)
            if not drop_labels:
                return stack_def
            return [item for item in stack_def if item.get("name") not in drop_labels]

        def _collapse_equity_components(data_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
            if not mapping.get("collapse_equity_on_negative", True):
                return data_list
            equity_group = next(
                (group for group in mapping.get("exclusive_groups", []) if group.get("aggregate") == "Total Equity"),
                None,
            )
            if not equity_group:
                return data_list
            components = set(equity_group.get("components", []))
            if not components:
                return data_list
            equity_items = [d for d in data_list if d["label"] in components]
            if not equity_items:
                return data_list
            if not any(d["value"] < 0 for d in equity_items):
                return data_list
            total_equity_val = pd.to_numeric(row.get("Total Equity"), errors="coerce")
            if pd.isna(total_equity_val):
                return data_list
            filtered = [d for d in data_list if d["label"] not in components and d["label"] != "Total Equity"]
            filtered.append(
                {
                    "label": "Total Equity",
                    "value": float(total_equity_val) / unit_scale,
                    "color": self._get_color(mapping.get("equity_total_color_key", "sky_blue")),
                }
            )
            return filtered

        left_stack_def = _apply_exclusive_groups(mapping.get("left_stack", []))
        if auto_balance_assets and total_assets_col:
            filtered = [item for item in left_stack_def if item.get("col") != total_assets_col]
            if filtered:
                left_stack_def = filtered

        detail_requested = bool(mapping.get("detail_bars", False))
        is_bank = bool(mapping.get("is_bank", False))
        default_pref = "bank" if is_bank else "summary"
        if detail_requested:
            default_pref = "bank" if is_bank else "primary"

        stack_pref = mapping.get("stack_preference")
        left_pref = mapping.get("left_stack_preference") or stack_pref or default_pref
        right_pref = mapping.get("right_stack_preference") or stack_pref or default_pref
        if not is_bank and left_pref == "bank":
            left_pref = "primary" if detail_requested else "summary"
        if not is_bank and right_pref == "bank":
            right_pref = "primary" if detail_requested else "summary"

        left_candidates = [
            ("primary", left_stack_def),
            ("bank", (mapping.get("left_stack_for_bank") or []) if is_bank or left_pref == "bank" else []),
            ("summary", mapping.get("left_stack_summary") or []),
        ]
        right_candidates = [
            ("primary", _apply_exclusive_groups(mapping.get("right_stack") or [])),
            ("bank", (mapping.get("right_stack_for_bank") or []) if is_bank or right_pref == "bank" else []),
            ("summary", mapping.get("right_stack_summary") or []),
        ]

        selection_target = (total_assets_val / unit_scale) if total_assets_val else 0
        prefer_detail_left = bool(mapping.get("prefer_detail_left", auto_balance_assets))
        prefer_detail_right = bool(mapping.get("prefer_detail_right", auto_balance_liab_equity))
        if not detail_requested:
            prefer_detail_left = False
            prefer_detail_right = False

        left_stack_name, left_stack_data = _select_stack_data(
            left_candidates,
            selection_target,
            left_pref,
            prefer_detail_left or selection_target <= 0,
        )
        right_stack_name, right_stack_data = _select_stack_data(
            right_candidates,
            selection_target,
            right_pref,
            prefer_detail_right or selection_target <= 0,
        )

        right_stack_data = _collapse_equity_components(right_stack_data)

        left_total = sum(d["value"] for d in left_stack_data)
        right_total = sum(d["value"] for d in right_stack_data)

        balance_target = (total_assets_val / unit_scale) if total_assets_val > 0 else max(left_total, right_total)

        if balance_target > 0 and auto_balance_assets and left_total < balance_target:
            left_stack_data.append(
                {
                    "label": other_assets_label,
                    "value": balance_target - left_total,
                    "color": self._get_color(mapping.get("other_assets_color_key", "ice_blue")),
                }
            )

        if balance_target > 0 and auto_balance_liab_equity and right_total < balance_target:
            right_stack_data.append(
                {
                    "label": other_liab_equity_label,
                    "value": balance_target - right_total,
                    "color": self._get_color(mapping.get("other_liab_equity_color_key", "gray_light")),
                }
            )

        fig, ax = plt.subplots(figsize=(8, 8))

        def _stack_bars(x_pos: int, data_list: List[Dict[str, Any]]) -> None:
            bottom = 0
            for d in data_list:
                val = d["value"]
                if val > 0:
                    ax.bar(x_pos, val, bottom=bottom, color=d["color"], edgecolor="white", width=0.55)
                    bottom += val

        _stack_bars(0, left_stack_data)
        _stack_bars(1, right_stack_data)

        ax.set_xticks([0, 1])
        ax.set_xticklabels([mapping.get("left_label", "Assets"), mapping.get("right_label", "Liabilities & Equity")])
        ax.set_ylabel(chart_text.get("y1_label", ""))

        show_legend = bool(mapping.get("show_legend", False))
        if show_legend:
            legend_items: List[Dict[str, Any]] = []
            seen = set()
            legend_source = mapping.get("legend_source", "selected")
            if legend_source == "detail":
                detail_left = _apply_exclusive_groups(mapping.get("left_stack_for_bank") if is_bank else mapping.get("left_stack", []))
                detail_right = _apply_exclusive_groups(mapping.get("right_stack_for_bank") if is_bank else mapping.get("right_stack", []))
                source_data = _build_stack_data(detail_left) + _build_stack_data(detail_right)
            else:
                source_data = left_stack_data + right_stack_data

            for entry in source_data:
                if entry["value"] <= 0:
                    continue
                label = entry["label"]
                if label in seen:
                    continue
                seen.add(label)
                legend_items.append({"label": label, "color": entry["color"], "value": entry["value"]})
            legend_max = mapping.get("legend_max_items")
            if legend_max:
                legend_items = sorted(legend_items, key=lambda x: x["value"], reverse=True)[: int(legend_max)]
            if legend_items:
                handles = [mpatches.Patch(color=item["color"], label=item["label"]) for item in legend_items]
                ax.legend(
                    handles=handles,
                    loc="center left",
                    bbox_to_anchor=(1.02, 0.5),
                    frameon=False,
                    fontsize=self.config.fonts.chart_tick_size,
                )

        show_segment_labels = bool(mapping.get("show_segment_labels", True))
        show_summary_labels = bool(mapping.get("show_summary_labels", True))
        segment_label_min_ratio = float(mapping.get("segment_label_min_ratio", 0.08))
        segment_label_font_size = int(mapping.get("segment_label_font_size", max(self.config.fonts.chart_tick_size - 1, 8)))
        segment_label_color = self._get_color(mapping.get("segment_label_color_key", "gray_dark"))
        summary_label_max_length = int(mapping.get("summary_label_max_length", 30))

        def _annotate_segments(
            x_pos: int,
            data_list: List[Dict[str, Any]],
            max_length: int = 18,
            position: str = "inside",
        ) -> None:
            total = sum(d["value"] for d in data_list)
            if total <= 0:
                return
            bottom = 0
            for d in data_list:
                val = d["value"]
                if val <= 0:
                    continue
                ratio = val / total if total else 0
                label = d["label"]
                if ratio >= segment_label_min_ratio and len(label) <= max_length:
                    if position == "outside":
                        x_text = x_pos - 0.38 if x_pos == 0 else x_pos + 0.38
                        ha = "right" if x_pos == 0 else "left"
                    else:
                        x_text = x_pos
                        ha = "center"
                    ax.text(
                        x_text,
                        bottom + (val / 2),
                        label,
                        ha=ha,
                        va="center",
                        fontsize=segment_label_font_size,
                        color=segment_label_color,
                    )
                bottom += val

        def _group_spans(data_list: List[Dict[str, Any]], groups: List[Dict[str, Any]]) -> List[Tuple[float, float, str]]:
            if not groups:
                return []
            value_map = {entry["label"]: entry["value"] for entry in data_list}
            spans = []
            cursor = 0.0
            for group in groups:
                items = group.get("items", [])
                label = group.get("label", "")
                total = sum(value_map.get(item, 0) for item in items)
                if total <= 0:
                    continue
                spans.append((cursor, cursor + total, label))
                cursor += total
            return spans

        def _annotate_groups(x_pos: int, spans: List[Tuple[float, float, str]], offset: float) -> None:
            if not spans:
                return
            color = self._get_color(mapping.get("group_label_color_key", "gray_dark"))
            font_size = int(mapping.get("group_label_font_size", max(self.config.fonts.chart_tick_size - 1, 8)))
            for start, end, label in spans:
                if not label:
                    continue
                mid = (start + end) / 2
                ax.text(
                    x_pos + offset,
                    mid,
                    label,
                    ha="right" if offset < 0 else "left",
                    va="center",
                    fontsize=font_size,
                    color=color,
                )
                ax.hlines(end, x_pos - 0.28, x_pos + 0.28, color="#D0D0D0", linewidth=0.6)

        left_groups = mapping.get("left_groups", [])
        right_groups = mapping.get("right_groups", [])
        if left_stack_name == "bank":
            left_groups = mapping.get("left_groups_for_bank", left_groups)
        if right_stack_name == "bank":
            right_groups = mapping.get("right_groups_for_bank", right_groups)

        summary_label_position = mapping.get("summary_label_position", "inside")
        if show_segment_labels or (show_summary_labels and left_stack_name == "summary"):
            _annotate_segments(
                0,
                left_stack_data,
                summary_label_max_length if left_stack_name == "summary" else 18,
                summary_label_position if left_stack_name == "summary" else "inside",
            )
        if show_segment_labels or (show_summary_labels and right_stack_name == "summary"):
            _annotate_segments(
                1,
                right_stack_data,
                summary_label_max_length if right_stack_name == "summary" else 18,
                summary_label_position if right_stack_name == "summary" else "inside",
            )

        _annotate_groups(0, _group_spans(left_stack_data, left_groups), offset=-0.3)
        _annotate_groups(1, _group_spans(right_stack_data, right_groups), offset=0.3)

        fig, ax = self._apply_common_style(fig, ax, chart_text)
        ax.xaxis.grid(False)
        plt.tight_layout()
        return fig


class PlotlyStrategy(ChartStrategyBase):
    """Placeholder for Plotly rendering strategy."""

    def plot_combo_bar_line_2axis(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> Any:
        raise NotImplementedError("Plotly strategy is not implemented.")

    def plot_balance_sheet(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> Any:
        raise NotImplementedError("Plotly strategy is not implemented.")

    def plot_portfolio_timeseries(self, df: Any, mapping: Dict[str, Any], chart_text: Dict[str, Any]) -> Any:
        raise NotImplementedError("Plotly strategy is not implemented.")

# ==============================
# D. Slide generation engine
# ==============================

class PowerPointGeneratorEngine:
    """Generate slides using a PowerPoint template and chart images."""

    def __init__(self, config: SlideConfig) -> None:
        self.config = config
        self.ppt_app = None
        self.prs = None
        self.idx_cover = 1
        self.idx_template_body = 2
        self.idx_back_cover = 3
        self.strategy = MatplotlibStrategy(config) if self.config.engine == "matplotlib" else PlotlyStrategy(config)
        self.slide_width = 960
        self.slide_height = 540
        if os.path.exists(self.config.paths.temp_img_dir):
            shutil.rmtree(self.config.paths.temp_img_dir)
        os.makedirs(self.config.paths.temp_img_dir, exist_ok=True)
        os.makedirs(self.config.paths.output_dir, exist_ok=True)

    def _initialize_ppt(self) -> None:
        tpl_path = self.config.paths.template_file
        if not os.path.exists(tpl_path):
            raise FileNotFoundError(f"Template not found: {tpl_path}")
        self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt_app.Visible = MSO_TRUE
        try:
            self.ppt_app.WindowState = PP_WINDOW_MINIMIZED
        except Exception:
            pass
        self.prs = self.ppt_app.Presentations.Open(tpl_path, ReadOnly=MSO_TRUE)
        self.slide_width = self.prs.PageSetup.SlideWidth
        self.slide_height = self.prs.PageSetup.SlideHeight

    def _to_ppt_rgb(self, color_key_or_tuple: Union[str, Tuple[int, int, int]], config: Optional[SlideConfig] = None) -> int:
        val = color_key_or_tuple
        active_config = config or self.config
        if isinstance(val, str) and not val.startswith("#"):
            val = active_config.colors.palette.get(val, "#000000")
        if isinstance(val, str) and val.startswith("#"):
            hex_color = val.lstrip("#")
            if len(hex_color) == 6:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                val = (r, g, b)
            else:
                val = (0, 0, 0)
        if not isinstance(val, tuple):
            val = (0, 0, 0)
        return val[0] + (val[1] << 8) + (val[2] << 16)

    def _calc_rect(self, layout_ratio: LayoutRatio) -> Tuple[float, float, float, float]:
        left = layout_ratio.left * self.slide_width
        top = layout_ratio.top * self.slide_height
        width = layout_ratio.width * self.slide_width
        height = layout_ratio.height * self.slide_height
        return left, top, width, height

    def _add_text_box(
        self,
        slide: Any,
        text: str,
        layout_ratio: LayoutRatio,
        bold: bool = False,
        color_key: str = "navy",
        align: int = MSO_ALIGN_LEFT,
        config: Optional[SlideConfig] = None,
    ) -> Any:
        active_config = config or self.config
        left, top, width, height = self._calc_rect(layout_ratio)
        shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)
        tf = shape.TextFrame
        tf.TextRange.Text = text
        if layout_ratio.font_size:
            tf.TextRange.Font.Size = layout_ratio.font_size
        tf.TextRange.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        tf.TextRange.Font.Color.RGB = self._to_ppt_rgb(color_key, config=active_config)
        tf.TextRange.Font.Name = active_config.fonts.english_font
        tf.TextRange.Font.NameAscii = active_config.fonts.english_font
        tf.TextRange.Font.NameFarEast = active_config.fonts.japanese_font
        tf.TextRange.ParagraphFormat.Alignment = align
        return shape

    def _add_picture_fitted(self, slide: Any, image_path: str, layout_ratio: LayoutRatio) -> None:
        if not os.path.exists(image_path):
            return
        box_left, box_top, box_width, box_height = self._calc_rect(layout_ratio)
        with Image.open(image_path) as img:
            img_w, img_h = img.size
        scale = min(box_width / img_w, box_height / img_h)
        new_w, new_h = img_w * scale, img_h * scale
        final_left = box_left + (box_width - new_w) / 2
        final_top = box_top + (box_height - new_h) / 2
        slide.Shapes.AddPicture(image_path, MSO_FALSE, MSO_TRUE, final_left, final_top, new_w, new_h)

    def _calculate_auto_font_size(self, text: str, max_size: int) -> int:
        length = len(text)
        if length < 50:
            return max_size
        if length < 100:
            return max_size - 2
        if length < 200:
            return max_size - 4
        return max_size - 6

    def _add_text_content_boxes(
        self,
        slide: Any,
        content_data: List[Dict[str, Any]],
        layout_type: str = "horizontal",
        section_header_text: str = "Key Findings",
        config: Optional[SlideConfig] = None,
    ) -> None:
        if not content_data:
            return
        num_boxes = len(content_data)
        if num_boxes == 0:
            return

        active_config = config or self.config
        layout_ratio = (
            active_config.layout.layout_vertical_text if layout_type == "vertical" else active_config.layout.layout_horizontal_text
        )

        area_left, area_top, area_width, area_height = self._calc_rect(layout_ratio)

        # Section header
        title_h_ratio = 0.05
        self._add_text_box(
            slide,
            section_header_text,
            LayoutRatio(layout_ratio.left, layout_ratio.top, layout_ratio.width, title_h_ratio),
            bold=True,
            config=active_config,
        )

        header_margin_px = 40
        current_top_px = area_top + header_margin_px
        available_height_px = area_height - header_margin_px
        box_gap_px = 15

        if layout_type == "vertical":
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
            if isinstance(item, str):
                item = {"body": item}
            header_shape = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, current_x, current_y, box_width_px, box_header_height_px)
            accent_color = item.get("accent_color_key", "navy")
            header_shape.Fill.ForeColor.RGB = self._to_ppt_rgb(accent_color, config=active_config)
            header_shape.Line.Visible = MSO_FALSE

            body_h_px = box_height_px - box_header_height_px
            body_shape = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, current_x, current_y + box_header_height_px, box_width_px, body_h_px)
            body_shape.Fill.ForeColor.RGB = self._to_ppt_rgb("gray_light", config=active_config)
            body_shape.Line.Visible = MSO_FALSE

            current_x_r = current_x / self.slide_width
            current_y_r = current_y / self.slide_height
            box_w_r = box_width_px / self.slide_width
            header_h_r = box_header_height_px / self.slide_height
            body_h_r = body_h_px / self.slide_height

            title_text = item.get("title") or item.get("header") or ""
            title_color = item.get("title_color_key", "white")
            self._add_text_box(
                slide,
                title_text,
                LayoutRatio(current_x_r + 0.01, current_y_r + 0.005, box_w_r - 0.02, header_h_r - 0.01),
                bold=True,
                color_key=title_color,
                align=MSO_ALIGN_LEFT,
                config=active_config,
            )

            body_text = item.get("body") or item.get("text") or item.get("value") or ""
            if not title_text and not body_text:
                current_x += dx
                current_y += dy
                continue
            max_font_size = self.config.layout.body_text_max_font_size
            if config is not None:
                max_font_size = config.layout.body_text_max_font_size
            auto_font_size = self._calculate_auto_font_size(body_text, max_font_size)

            self._add_text_box(
                slide,
                body_text,
                LayoutRatio(
                    current_x_r + 0.01,
                    current_y_r + header_h_r + 0.01,
                    box_w_r - 0.02,
                    body_h_r - 0.02,
                    font_size=auto_font_size,
                ),
                color_key="gray_dark",
                config=active_config,
            )

            current_x += dx
            current_y += dy

    def generate(
        self,
        data_store: Optional[Union[Dict[str, Any], Any]] = None,
        cover_content: Optional[Dict[str, Any]] = None,
        slides_structure: Optional[List[Union[SlidePage, Dict[str, Any]]]] = None,
        filename_prefix: str = "Presentation",
        deck: Optional[SlideDeck] = None,
    ) -> None:
        try:
            if deck is not None:
                cover_content = {
                    "main_title": deck.cover.main_title,
                    "sub_title": deck.cover.sub_title,
                    "date": deck.cover.date,
                }
                slides_structure = deck.pages

            if cover_content is None or slides_structure is None:
                raise ValueError("cover_content and slides_structure are required.")

            if data_store is None:
                data_store = {}

            self._initialize_ppt()

            # Cover slide
            slide1 = self.prs.Slides(self.idx_cover)
            self._add_text_box(slide1, cover_content["main_title"], self.config.layout.cover_title, bold=True, align=MSO_ALIGN_CENTER)
            self._add_text_box(
                slide1,
                cover_content["sub_title"],
                self.config.layout.cover_sub,
                color_key="gray_dark",
                align=MSO_ALIGN_CENTER,
            )
            self._add_text_box(
                slide1,
                cover_content.get("date", datetime.now().strftime("%Y-%m-%d")),
                self.config.layout.cover_date,
                color_key="gray_dark",
                align=MSO_ALIGN_CENTER,
            )

            template_slide = self.prs.Slides(self.idx_template_body)

            # Content slides
            for i, slide_def_raw in enumerate(slides_structure):
                slide_def = _resolve_slide_def(slide_def_raw)
                page_config = _build_slide_config(self.config, slide_def.get("config"))
                page_strategy = MatplotlibStrategy(page_config) if page_config.engine == "matplotlib" else PlotlyStrategy(page_config)
                new_slide = template_slide.Duplicate().Item(1)
                target_index = self.prs.Slides.Count - 1
                if target_index < 2:
                    target_index = 2
                new_slide.MoveTo(target_index)

                self._add_text_box(
                    new_slide,
                    slide_def["slide_title"],
                    page_config.layout.content_title,
                    bold=True,
                    align=MSO_ALIGN_RIGHT,
                    config=page_config,
                )

                layout_type = slide_def.get("layout_type", "horizontal")
                chart_area_ratio = (
                    page_config.layout.layout_vertical_chart if layout_type == "vertical" else page_config.layout.layout_horizontal_chart
                )

                data_frame = slide_def.get("data_frame")
                data_source = slide_def.get("data_source")
                if data_frame is not None:
                    df = data_frame
                elif isinstance(data_store, dict):
                    df = data_store.get(data_source, [])
                else:
                    df = data_store

                category = slide_def["category"]
                mapping = slide_def["data_mapping"]
                chart_text = slide_def["chart_text"]

                plot_method_name = f"plot_{category.lower()}"
                plot_method = getattr(page_strategy, plot_method_name, None)

                if plot_method:
                    fig = plot_method(df, mapping, chart_text)
                    img_filename = f"chart_{uuid.uuid4()}.png"
                    img_path = os.path.join(self.config.paths.temp_img_dir, img_filename)

                    fig.savefig(img_path, dpi=150, bbox_inches="tight")
                    plt.close(fig)

                    self._add_picture_fitted(new_slide, img_path, chart_area_ratio)

                content_data = slide_def.get("text_blocks") or slide_def.get("proposal_points")
                if content_data:
                    section_title = slide_def.get("proposal_section_title", "Key Findings")
                    self._add_text_content_boxes(
                        new_slide,
                        content_data,
                        layout_type=layout_type,
                        section_header_text=section_title,
                        config=page_config,
                    )

            template_slide.Delete()

            filename = f"{filename_prefix}.pptx"
            output_path = os.path.join(self.config.paths.output_dir, filename)

            self.prs.SaveAs(output_path, PP_SAVE_AS_OPENXML_PRESENTATION)
            pdf_filename = f"{filename_prefix}.pdf"
            pdf_output_path = os.path.join(self.config.paths.output_dir, pdf_filename)
            try:
                self.prs.SaveAs(pdf_output_path, PP_SAVE_AS_PDF)
            except Exception:
                pass
            try:
                self.prs.Close()
            except Exception:
                pass
            self.prs = None

        except Exception as exc:
            print(f"Engine Error: {exc}")
            import traceback

            traceback.print_exc()

        finally:
            if self.prs:
                try:
                    self.prs.Close()
                except Exception:
                    pass
            if self.ppt_app:
                try:
                    try:
                        self.ppt_app.WindowState = PP_WINDOW_NORMAL
                    except Exception:
                        pass
                    self.ppt_app.Quit()
                except Exception:
                    pass
            if os.path.exists(self.config.paths.temp_img_dir):
                try:
                    shutil.rmtree(self.config.paths.temp_img_dir)
                except Exception:
                    pass

