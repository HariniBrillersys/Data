"""
Theme Color Extraction from PowerPoint templates.

Extracts the color scheme from a template's theme XML (ppt/theme/theme1.xml)
so that shapes, charts, and other generated elements can use on-brand colors.

PowerPoint themes define 12 standard color slots:
  dk1, dk2, lt1, lt2, accent1-accent6, hlink, folHlink

Plus optional custom colors (<a:custClrLst>) for extended brand palettes.
"""

from __future__ import annotations

import logging
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from lxml import etree

log = logging.getLogger(__name__)

_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS = {"a": _DRAWINGML_NS}

# Standard theme color slot names in XML order
_SLOT_NAMES = [
    "dk1",
    "dk2",
    "lt1",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
]


def _extract_color(element) -> Optional[str]:
    """Extract hex color string from a theme color element.

    Handles both <a:srgbClr val="0F69AF"/> and <a:sysClr lastClr="000000"/>.
    """
    srgb = element.find("a:srgbClr", _NS)
    if srgb is not None:
        val = srgb.get("val")
        if val:
            return f"#{val.upper()}"

    sys_clr = element.find("a:sysClr", _NS)
    if sys_clr is not None:
        last = sys_clr.get("lastClr")
        if last:
            return f"#{last.upper()}"

    return None


@dataclass
class ThemeColors:
    """Color palette extracted from a PowerPoint template's theme.

    All colors are stored as hex strings (e.g. '#0F69AF').
    Font families are extracted from theme font scheme.
    """

    dk1: str = "#000000"
    dk2: str = "#0F69AF"
    lt1: str = "#FFFFFF"
    lt2: str = "#96D7D2"
    accent1: str = "#96D7D2"
    accent2: str = "#0F69AF"
    accent3: str = "#FFDCB9"
    accent4: str = "#96D7D2"
    accent5: str = "#0F69AF"
    accent6: str = "#FFDCB9"
    hlink: str = "#0F69AF"
    folHlink: str = "#0F69AF"
    custom_colors: dict[str, str] = field(default_factory=dict)
    major_font: str = "Calibri"  # For headings
    minor_font: str = "Calibri"  # For body text

    @classmethod
    def uptimize_defaults(cls) -> ThemeColors:
        """Return the Uptimize Master template's default color palette."""
        return cls(
            dk1="#000000",
            dk2="#0F69AF",
            lt1="#FFFFFF",
            lt2="#96D7D2",
            accent1="#96D7D2",
            accent2="#0F69AF",
            accent3="#FFDCB9",
            accent4="#96D7D2",
            accent5="#0F69AF",
            accent6="#FFDCB9",
            hlink="#0F69AF",
            folHlink="#0F69AF",
            custom_colors={
                "Rich Purple": "#503291",
                "Rich Blue": "#0F69AF",
                "Rich Green": "#149B5F",
                "Rich Red": "#E61E50",
                "Vibrant Magenta": "#EB3C96",
                "Vibrant Cyan": "#2DBECD",
                "Vibrant Green": "#A5CD50",
                "Vibrant Yellow": "#FFC832",
                "Sensitive Pink": "#E1C3CD",
                "Sensitive Blue": "#96D7D2",
                "Sensitive Green": "#B4DC96",
                "Sensitive Yellow": "#FFDCB9",
            },
            major_font="Verdana",
            minor_font="Verdana",
        )

    @classmethod
    def from_template(cls, template_path: str | Path) -> ThemeColors:
        """Extract theme colors from a PowerPoint template file (.potx or .pptx).

        Falls back to Uptimize defaults if extraction fails.
        """
        template_path = Path(template_path)

        try:
            with zipfile.ZipFile(template_path, "r") as zf:
                # Find the primary theme file (theme1.xml is the slide master theme)
                theme_file = None
                theme_candidates = sorted(
                    [n for n in zf.namelist() if n.startswith("ppt/theme/theme") and n.endswith(".xml")]
                )
                # Prefer theme1.xml (slide master theme, has custom colors)
                for candidate in theme_candidates:
                    if candidate == "ppt/theme/theme1.xml":
                        theme_file = candidate
                        break
                if theme_file is None and theme_candidates:
                    theme_file = theme_candidates[0]

                if theme_file is None:
                    log.debug("No theme file found in %s, using defaults", template_path.name)
                    return cls.uptimize_defaults()

                theme_xml = zf.read(theme_file)
                return cls._parse_theme_xml(theme_xml)

        except Exception as e:
            log.debug("Failed to extract theme from %s: %s", template_path.name, e)
            return cls.uptimize_defaults()

    @classmethod
    def _parse_theme_xml(cls, xml_bytes: bytes) -> ThemeColors:
        """Parse theme XML and extract color scheme."""
        root = etree.fromstring(xml_bytes)

        # Find <a:clrScheme>
        clr_scheme = root.find(".//a:clrScheme", _NS)
        if clr_scheme is None:
            return cls.uptimize_defaults()

        colors = {}
        for slot_name in _SLOT_NAMES:
            slot_elem = clr_scheme.find(f"a:{slot_name}", _NS)
            if slot_elem is not None:
                color = _extract_color(slot_elem)
                if color:
                    colors[slot_name] = color

        # Parse custom colors
        custom = {}
        cust_clr_lst = root.find(".//a:custClrLst", _NS)
        if cust_clr_lst is not None:
            for cust_clr in cust_clr_lst.findall("a:custClr", _NS):
                name = cust_clr.get("name", "").strip()
                srgb = cust_clr.find("a:srgbClr", _NS)
                if srgb is not None and name:
                    val = srgb.get("val")
                    if val:
                        custom[name] = f"#{val.upper()}"

        # Parse font families
        major_font = "Calibri"
        minor_font = "Calibri"
        font_scheme = root.find(".//a:fontScheme", _NS)
        if font_scheme is not None:
            # Major font (headings)
            major_font_elem = font_scheme.find("a:majorFont", _NS)
            if major_font_elem is not None:
                latin = major_font_elem.find("a:latin", _NS)
                if latin is not None:
                    typeface = latin.get("typeface")
                    if typeface:
                        major_font = typeface

            # Minor font (body text)
            minor_font_elem = font_scheme.find("a:minorFont", _NS)
            if minor_font_elem is not None:
                latin = minor_font_elem.find("a:latin", _NS)
                if latin is not None:
                    typeface = latin.get("typeface")
                    if typeface:
                        minor_font = typeface

        return cls(
            dk1=colors.get("dk1", "#000000"),
            dk2=colors.get("dk2", "#0F69AF"),
            lt1=colors.get("lt1", "#FFFFFF"),
            lt2=colors.get("lt2", "#96D7D2"),
            accent1=colors.get("accent1", "#96D7D2"),
            accent2=colors.get("accent2", "#0F69AF"),
            accent3=colors.get("accent3", "#FFDCB9"),
            accent4=colors.get("accent4", "#96D7D2"),
            accent5=colors.get("accent5", "#0F69AF"),
            accent6=colors.get("accent6", "#FFDCB9"),
            hlink=colors.get("hlink", "#0F69AF"),
            folHlink=colors.get("folHlink", "#0F69AF"),
            custom_colors=custom,
            major_font=major_font,
            minor_font=minor_font,
        )

    def accent_cycle(self) -> list[str]:
        """Return accent colors as a list for chart/shape color cycling."""
        return [self.accent1, self.accent2, self.accent3, self.accent4, self.accent5, self.accent6]

    def resolve_color(self, color_ref: str) -> str:
        """Resolve a color reference to a hex string.

        Accepts:
          - Theme slot names: 'accent1', 'accent2', ..., 'dark', 'light'
          - Custom color names: 'Rich Purple', 'Vibrant Cyan', ...
          - Direct hex: '#0F69AF'
        """
        # Direct hex
        if color_ref.startswith("#"):
            return color_ref.upper()

        # Theme slot aliases
        aliases = {
            "dark": "dk1",
            "light": "lt1",
            "dark2": "dk2",
            "light2": "lt2",
        }
        slot = aliases.get(color_ref.lower(), color_ref.lower())

        # Check standard slots
        if hasattr(self, slot):
            val = getattr(self, slot)
            if isinstance(val, str) and val.startswith("#"):
                return val

        # Check custom colors
        if color_ref in self.custom_colors:
            return self.custom_colors[color_ref]

        # Case-insensitive custom color search
        for name, hex_val in self.custom_colors.items():
            if name.lower() == color_ref.lower():
                return hex_val

        # Default to accent2 (primary brand color)
        return self.accent2
