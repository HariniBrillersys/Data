"""
Chart Builder - Generate charts/graphs as PNG images using matplotlib.

Charts use the active template's theme colors for visual consistency.
The generated PNG files can be placed into slide picture placeholders
via the 'image' field in slide content dicts.

Uses matplotlib (core dependency) for chart rendering.

Supported chart types:
  - bar: Vertical bar chart
  - horizontal_bar: Horizontal bar chart
  - stacked_bar: Stacked vertical bars
  - line: Line chart
  - pie: Pie chart
  - donut: Donut chart
  - scatter: Scatter plot
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path
from typing import Optional

import matplotlib

matplotlib.use("Agg")  # Non-interactive backend
import matplotlib.pyplot as plt
import numpy as np

from .theme_colors import ThemeColors

log = logging.getLogger(__name__)

_MAX_LABELS = 100
_MAX_SERIES = 20


def _setup_style(theme: ThemeColors, font_family: str = "Calibri") -> None:
    """Configure matplotlib style to match the template theme."""
    plt.rcParams.update(
        {
            "font.family": "sans-serif",
            "font.sans-serif": [font_family, "DejaVu Sans", "Arial", "Helvetica", "Calibri"],
            "font.size": 12,
            "axes.facecolor": "white",
            "figure.facecolor": "white",
            "axes.edgecolor": "#CCCCCC",
            "axes.grid": True,
            "grid.color": "#EEEEEE",
            "grid.linewidth": 0.5,
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )


def _hex_to_mpl(hex_color: str) -> str:
    """Ensure hex color is in matplotlib-compatible format."""
    return hex_color if hex_color.startswith("#") else f"#{hex_color}"


def _build_color_cycle(theme: ThemeColors) -> list[str]:
    """Build a matplotlib-compatible color cycle from theme colors."""
    # Use accents, then fall back to custom colors for more variety
    colors = []
    for c in theme.accent_cycle():
        mpl_c = _hex_to_mpl(c)
        if mpl_c not in colors:
            colors.append(mpl_c)

    # Add custom colors for variety if we have few unique accents
    if len(colors) < 6:
        for name, hex_val in theme.custom_colors.items():
            mpl_c = _hex_to_mpl(hex_val)
            if mpl_c not in colors:
                colors.append(mpl_c)
            if len(colors) >= 10:
                break

    return colors or ["#0F69AF", "#96D7D2", "#FFDCB9", "#503291", "#149B5F", "#E61E50"]


def _normalize_data(data: dict | list[dict]) -> tuple[list[str], list[dict]]:
    """Normalize data input to a consistent format.

    Returns:
        Tuple of (labels, series_list).
        Each series is {"name": str, "values": list[float]}.
    """
    if isinstance(data, dict):
        # Single series: {"Q1": 10, "Q2": 15, ...}
        if len(data) > _MAX_LABELS:
            raise ValueError(f"Too many data points ({len(data)}). Maximum is {_MAX_LABELS} labels per chart.")
        labels = list(data.keys())
        values = [float(v) for v in data.values()]
        return labels, [{"name": "", "values": values}]

    elif isinstance(data, list):
        if not data:
            return [], []

        if len(data) > _MAX_SERIES:
            raise ValueError(f"Too many data series ({len(data)}). Maximum is {_MAX_SERIES} series per chart.")

        # Multi-series: [{"name": "Revenue", "values": {"Q1": 10, ...}}, ...]
        # Preserve insertion order from the first series that has dict values,
        # then append any extra keys from subsequent series.
        labels: list[str] = []
        seen: set[str] = set()
        for series in data:
            if isinstance(series.get("values"), dict):
                for key in series["values"]:
                    if key not in seen:
                        labels.append(key)
                        seen.add(key)

        if len(labels) > _MAX_LABELS:
            raise ValueError(f"Too many data points ({len(labels)}). Maximum is {_MAX_LABELS} labels per chart.")

        series_list = []
        for series in data:
            name = series.get("name", "")
            vals = series.get("values", {})
            if isinstance(vals, dict):
                values = [float(vals.get(lbl, 0)) for lbl in labels]
            elif isinstance(vals, list):
                values = [float(v) for v in vals]
                if not labels:
                    labels = [str(i + 1) for i in range(len(values))]
            else:
                continue
            series_list.append({"name": name, "values": values})

        return labels, series_list

    return [], []


class ChartBuilder:
    """Generate chart images from data using matplotlib."""

    def __init__(self, images_dir: str | Path = "outputs/images", font_family: str = "Calibri"):
        self.images_dir = Path(images_dir)
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.font_family = font_family
        self._lock = __import__("threading").Lock()
        self._warmup()

    @staticmethod
    def _warmup() -> None:
        """Pre-warm matplotlib to avoid first-call delays.

        The first render triggers font cache building and backend
        initialization which can take 10-60s. Paying that cost at
        startup keeps tool-call latency predictable.
        """
        try:
            fig, ax = plt.subplots(figsize=(1, 1))
            ax.text(0.5, 0.5, "warmup", fontsize=8)
            fig.savefig(__import__("io").BytesIO(), format="png")
            plt.close(fig)
            log.debug("matplotlib warmup complete")
        except Exception as exc:
            log.warning("matplotlib warmup failed: %s", exc)

    def generate(
        self,
        chart_type: str,
        data: dict | list[dict],
        title: str = "",
        output_name: str = "",
        xlabel: str = "",
        ylabel: str = "",
        legend: bool = True,
        theme_colors: Optional[ThemeColors] = None,
        figsize: tuple[float, float] = (10, 6),
    ) -> dict:
        """Generate a chart and save as PNG.

        Args:
            chart_type: 'bar', 'horizontal_bar', 'stacked_bar', 'line', 'pie', 'donut', 'scatter'.
            data: Single series dict or multi-series list.
            title: Chart title.
            output_name: Output filename. Auto-generated if empty.
            xlabel, ylabel: Axis labels.
            legend: Show legend for multi-series.
            theme_colors: Template theme colors for styling.
            figsize: Figure size in inches (width, height).

        Returns:
            Dict with success, path, chart_type, file_size_bytes.
        """
        if not output_name:
            output_name = f"chart_{uuid.uuid4().hex[:8]}.png"
        if not output_name.lower().endswith(".png"):
            output_name += ".png"

        output_path = self.images_dir / output_name
        theme = theme_colors or ThemeColors.uptimize_defaults()

        with self._lock:
            _setup_style(theme, self.font_family)
            colors = _build_color_cycle(theme)

            try:
                labels, series_list = _normalize_data(data)
                if not labels and chart_type not in ("pie", "donut"):
                    return {"success": False, "error": "No data labels found in the provided data."}
                if not series_list:
                    return {"success": False, "error": "No data series found in the provided data."}
                chart_func = {
                    "bar": self._bar,
                    "horizontal_bar": self._horizontal_bar,
                    "stacked_bar": self._stacked_bar,
                    "line": self._line,
                    "pie": self._pie,
                    "donut": self._donut,
                    "scatter": self._scatter,
                }.get(chart_type.lower())

                if chart_func is None:
                    return {
                        "success": False,
                        "error": f"Unknown chart type '{chart_type}'. "
                        f"Available: bar, horizontal_bar, stacked_bar, line, pie, donut, scatter.",
                    }

                fig, ax = plt.subplots(figsize=figsize)

                chart_func(ax, labels, series_list, colors)

                if title:
                    ax.set_title(title, fontsize=14, fontweight="bold", pad=15)
                if xlabel and chart_type not in ("pie", "donut"):
                    ax.set_xlabel(xlabel)
                if ylabel and chart_type not in ("pie", "donut"):
                    ax.set_ylabel(ylabel)

                if legend and len(series_list) > 1 and chart_type not in ("pie", "donut"):
                    ax.legend(frameon=False)

                plt.tight_layout()
                fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
                plt.close(fig)

                return {
                    "success": True,
                    "path": str(output_path.resolve()),
                    "chart_type": chart_type,
                    "file_size_bytes": output_path.stat().st_size,
                    "image_mode": "fit",
                }

            except Exception as e:
                plt.close("all")
                return {"success": False, "error": f"Chart rendering failed: {e}"}

    # --- Chart type renderers ---

    def _bar(self, ax, labels, series_list, colors):
        x = np.arange(len(labels))
        n = len(series_list)
        width = 0.8 / n

        for i, series in enumerate(series_list):
            offset = (i - n / 2 + 0.5) * width
            ax.bar(x + offset, series["values"], width, label=series["name"], color=colors[i % len(colors)])

        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right" if len(labels) > 6 else "center")

    def _horizontal_bar(self, ax, labels, series_list, colors):
        y = np.arange(len(labels))
        n = len(series_list)
        height = 0.8 / n

        for i, series in enumerate(series_list):
            offset = (i - n / 2 + 0.5) * height
            ax.barh(y + offset, series["values"], height, label=series["name"], color=colors[i % len(colors)])

        ax.set_yticks(y)
        ax.set_yticklabels(labels)

    def _stacked_bar(self, ax, labels, series_list, colors):
        x = np.arange(len(labels))
        bottom = np.zeros(len(labels))

        for i, series in enumerate(series_list):
            values = np.array(series["values"])
            ax.bar(x, values, 0.6, bottom=bottom, label=series["name"], color=colors[i % len(colors)])
            bottom += values

        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right" if len(labels) > 6 else "center")

    def _line(self, ax, labels, series_list, colors):
        for i, series in enumerate(series_list):
            ax.plot(
                labels,
                series["values"],
                marker="o",
                linewidth=2,
                markersize=6,
                label=series["name"],
                color=colors[i % len(colors)],
            )

        if len(labels) > 6:
            plt.xticks(rotation=45, ha="right")

    def _pie(self, ax, labels, series_list, colors):
        values = series_list[0]["values"]
        wedges, texts, autotexts = ax.pie(
            values,
            labels=labels,
            colors=colors[: len(values)],
            autopct="%1.1f%%",
            startangle=90,
        )
        for text in autotexts:
            text.set_fontsize(10)
            text.set_fontweight("bold")
        ax.axis("equal")

    def _donut(self, ax, labels, series_list, colors):
        values = series_list[0]["values"]
        wedges, texts, autotexts = ax.pie(
            values,
            labels=labels,
            colors=colors[: len(values)],
            autopct="%1.1f%%",
            startangle=90,
            pctdistance=0.8,
            wedgeprops={"width": 0.4},
        )
        for text in autotexts:
            text.set_fontsize(10)
            text.set_fontweight("bold")
        ax.axis("equal")

    def _scatter(self, ax, labels, series_list, colors):
        # For scatter, labels are X values (try to parse as numbers)
        try:
            x_values = [float(lbl) for lbl in labels]
        except (ValueError, TypeError):
            x_values = list(range(len(labels)))

        for i, series in enumerate(series_list):
            ax.scatter(
                x_values,
                series["values"],
                s=60,
                label=series["name"],
                color=colors[i % len(colors)],
                alpha=0.8,
            )
