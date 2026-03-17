"""
PowerPoint MCP Server - FastMCP implementation.

Exposes tools for creating executive-quality presentations from templates
via the Model Context Protocol (MCP).

Configuration (env vars):
    PPTX_TEMPLATES_DIR  - Path to templates folder (default: ./templates)
    PPTX_OUTPUTS_DIR    - Path to outputs folder (default: ./outputs)
    PPTX_TRANSPORT      - Transport: "stdio" or "sse" (default: stdio)
    PPTX_HOST           - SSE host (default: 127.0.0.1)
    PPTX_PORT           - SSE port (default: 8000)
    PPTX_LOG_LEVEL      - Logging level: DEBUG, INFO, WARNING (default: INFO)
    UPTIMIZE_OPENAI_API_KEY - API key for UPTIMIZE image generation
    UPTIMIZE_ENV        - UPTIMIZE environment: "dev" or "p" (default: dev)
"""

from __future__ import annotations

import base64
import logging
import os
import sys
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from .chart_builder import ChartBuilder
from .composer import PresentationComposer
from .template_engine import TemplateEngine

try:
    from .image_generator import ImageGenerator

    _HAS_IMAGE_GEN = True
except ImportError:
    _HAS_IMAGE_GEN = False
    ImageGenerator = None  # type: ignore[assignment, misc]

log = logging.getLogger("pptx_mcp")

# ---------------------------------------------------------------------------
# Configuration via environment variables
# ---------------------------------------------------------------------------

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent

_TEMPLATES_DIR = Path(os.environ.get("PPTX_TEMPLATES_DIR", _PROJECT_ROOT / "templates"))
_OUTPUTS_DIR = Path(os.environ.get("PPTX_OUTPUTS_DIR", _PROJECT_ROOT / "outputs"))
_IMAGES_DIR = _OUTPUTS_DIR / "images"
_DEFAULT_TEMPLATE = os.environ.get("PPTX_DEFAULT_TEMPLATE", "")
_TRANSPORT = os.environ.get("PPTX_TRANSPORT", "stdio")
_HOST = os.environ.get("PPTX_HOST", "127.0.0.1")
_PORT = int(os.environ.get("PPTX_PORT", "8000"))

# UPTIMIZE image generation
_UPTIMIZE_API_KEY = os.environ.get("UPTIMIZE_OPENAI_API_KEY", "")
_UPTIMIZE_ENV = os.environ.get("UPTIMIZE_ENV", "dev")
_UPTIMIZE_BASE_URL = f"https://api.nlp.{_UPTIMIZE_ENV}.uptimize.merckgroup.com/openai/v1"

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

_LOG_LEVEL = os.environ.get("PPTX_LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, _LOG_LEVEL, logging.INFO),
    format="%(name)s %(levelname)s: %(message)s",
    stream=sys.stderr,
)

# ---------------------------------------------------------------------------
# Server setup
# ---------------------------------------------------------------------------

_TEMPLATE_INSTRUCTION = (
    f"Default template: {_DEFAULT_TEMPLATE}. Use this unless the user asks for a different one."
    if _DEFAULT_TEMPLATE
    else "No default template configured. Call list_templates first and ask the user which to use."
)
_IMAGE_INSTRUCTION = (
    f"Image generation available via generate_image. Images saved to {_IMAGES_DIR}."
    if _UPTIMIZE_API_KEY
    else "Image generation NOT available. Requires: pip install pptx-mcp[images] and UPTIMIZE_OPENAI_API_KEY env var."
)

# NOTE: _instructions is an f-string. Any literal curly braces (e.g. JSON
# examples) MUST be doubled ({{ }}) to avoid ValueError at import time.
_instructions = f"""You are a PowerPoint presentation generator that creates
slide decks from corporate templates.

Any organization can provide their own .potx or .pptx template files.
The server auto-detects all layouts and placeholder roles in each template.

TEMPLATE SELECTION:
{_TEMPLATE_INSTRUCTION}

IMAGE GENERATION:
{_IMAGE_INSTRUCTION}

CHART GENERATION:
Charts can be generated via the generate_chart tool.
Returns a PNG file path usable as the 'image' field in any slide dict.
Charts default to image_mode: "fit" — no cropping, full content visible.
Agents can override with image_mode: "fill" if cropping is desired (rare).

IMAGE PLACEMENT MODES:
  - image_mode: "fill" (default) — image is cropped to fill a picture
    placeholder, integrating with the template's mask/shape design.
    Use this for decorative and AI-generated images, especially on
    title and intro slides where the image should blend with the
    template's visual design rather than appear as a standalone rectangle.
    generate_image returns image_mode: "fill" — use it unless the
    generated image contains data that cannot be cropped.
  - image_mode: "fit" — image is placed as a freestanding shape, scaled
    proportionally with NO cropping. Use this for charts, diagrams, data
    images, screenshots, and any visual where the full content must remain
    visible. generate_chart returns image_mode: "fit" — always use it.
    When the image field is a list, images are arranged in a clean grid.
  - image_mode: "collage" — multiple images arranged with overlapping,
    staggered positions and slight size variation. Creates a dynamic
    visual feel. Best with 2-9 images on a title-only or blank layout.

CHOOSING IMAGE MODE (use the hint returned by generate tools):
  - generate_image → image_mode: "fill" (decorative, can be masked/cropped)
  - generate_chart → image_mode: "fit" (data content, must not be cropped)
  - User-provided images: use "fill" for photos/decorative, "fit" for
    screenshots/diagrams/data visualizations
  - Title/intro slides: always prefer "fill" for generated images — the
    template's picture placeholder masks create polished, branded visuals

MULTI-IMAGE:
The image field can be a list of paths or dicts for multi-image slides:
  - ["path1.png", "path2.png"]  (auto z-order from list position)
  - [{{"path": "bg.png", "z_order": 0}}, {{"path": "hero.png", "z_order": 1}}]
List order determines layering: last item is on top (most visible).
Use z_order in dicts to override. Max 9 images per slide.
The layout classifier prefers title-only or blank layouts for multi-image.

SPEAKER NOTES:
Each slide can include a 'notes' field (string or list of strings) with
talking points for the presenter. Notes appear in Presenter View only,
not on the projected slide. Use notes to keep slides clean and lean —
put storytelling details, context, and talking points in notes rather
than overloading slides with text. Keep notes concise — a few key
talking points or reminders, not full scripts. Think bullet points
the presenter can glance at, not paragraphs to read aloud.

SHAPE ANNOTATIONS:
Slides can include decorative shapes (arrows, callouts, badges, etc.) via
the 'shapes' key in each slide dict. Shapes use the template's color scheme.

SLIDE DESIGN PRINCIPLES:
- Keep slides lean: aim for 6 bullets max, ~6 words each (the 6×6 guideline)
- Use speaker notes for supporting detail — projected slides should be visual, not documents
- Check placeholder capacity via get_template_layouts: each placeholder reports
  max_comfortable_words and max_bullet_items
- If create_presentation returns density warnings, revise content or split across slides
- Content-to-visual ratio: if a slide is all text, add an image or split into two slides

WORKFLOW:
1. Call list_templates to see available templates
2. If the user hasn't specified a template, ask them to pick one
3. Call get_template_layouts to inspect the chosen template's layouts
4. Optionally call generate_image or generate_chart to create visuals
5. Call create_presentation with structured slide content (including optional shapes)
6. Optionally call download_presentation to get the file

Each slide dict supports: title, subtitle, body, content, content_left,
content_right, content_1/2/3, notice, image, image_mode, notes, shapes.

SHAPES:
Each entry in 'shapes' is a dict with 'type' and additional params:
  - arrow: direction (right/left/up/down), label, color, position
  - callout: text, color, position
  - badge: text, color, position (top-right, bottom-left, etc.)
  - highlight: target (placeholder role), color, opacity
  - connector: start, end, style (solid/dashed)
  - process_arrow: steps (list of labels), color
Colors can be theme names (accent1-accent6, dark, light) or hex (#0F69AF)."""

mcp = FastMCP(
    "PowerPoint Presentation Generator",
    instructions=_instructions,
)

# Initialize engine and composer
log.info("Initializing PowerPoint MCP Server...")
log.info("  Templates dir: %s", _TEMPLATES_DIR)
log.info("  Outputs dir:   %s", _OUTPUTS_DIR)
if _DEFAULT_TEMPLATE:
    log.info("  Default template: %s", _DEFAULT_TEMPLATE)

_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
_OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
_IMAGES_DIR.mkdir(parents=True, exist_ok=True)

engine = TemplateEngine(_TEMPLATES_DIR)
composer = PresentationComposer(engine, _OUTPUTS_DIR)

# Initialize image generator (optional — requires openai package + API key)
image_gen: ImageGenerator | None = None  # type: ignore[assignment]
if not _HAS_IMAGE_GEN:
    log.info("  Image generation: disabled (pip install pptx-mcp[images])")
elif not _UPTIMIZE_API_KEY:
    log.info("  Image generation: disabled (set UPTIMIZE_OPENAI_API_KEY to enable)")
else:
    try:
        image_gen = ImageGenerator(
            api_key=_UPTIMIZE_API_KEY,
            base_url=_UPTIMIZE_BASE_URL,
            images_dir=_IMAGES_DIR,
        )
        log.info("  Image generation: enabled (env=%s)", _UPTIMIZE_ENV)
    except Exception as e:
        log.warning("  Image generation: failed to initialize (%s)", e)

# Chart builder — initialized lazily after first template scan
_chart_builder: ChartBuilder | None = None


def _ensure_chart_builder(template_name: str = "") -> ChartBuilder:
    """Initialize chart builder lazily, using template font if available."""
    global _chart_builder
    if _chart_builder is not None:
        return _chart_builder

    default_font = "Calibri"
    tn = template_name or _DEFAULT_TEMPLATE
    if tn:
        info = engine.get_template(tn)
        if info:
            default_font = info.theme_colors.minor_font
    elif not tn:
        # Try first available template
        available = engine.list_available()
        if available:
            info = engine.get_template(available[0]["name"])
            if info:
                default_font = info.theme_colors.minor_font

    _chart_builder = ChartBuilder(images_dir=_IMAGES_DIR, font_family=default_font)
    log.info("  Chart generation: enabled (font: %s)", default_font)
    return _chart_builder


log.info("  Transport: %s", _TRANSPORT)


# ---------------------------------------------------------------------------
# MCP Tools
# ---------------------------------------------------------------------------


@mcp.tool()
def create_presentation(
    slides: list[dict],
    output_name: str = "presentation.pptx",
    template_name: str = "",
) -> dict:
    """
    Create a PowerPoint presentation from structured slide content.

    Each slide in the list is a dict that can contain:
      - layout (str, optional): Preferred layout name (e.g. "Title 01", "Content 01").
        If omitted, the best layout is auto-selected based on content.
      - title (str): Slide title text.
      - subtitle (str): Subtitle (for title/intro slides).
      - body (str): Sub-headline or category text (small text above main content).
      - content (str | list[str]): Main content. String for paragraph, list for bullets.
      - content_left (str | list[str]): Left column (for two-column layouts).
      - content_right (str | list[str]): Right column (for two-column layouts).
      - content_1, content_2, content_3 (str | list[str]): For three-column layouts.
      - notice (str | list[str]): Sidebar/notice panel content.
      - image (str | list): Path to an image file, or a list of paths/dicts
        for multi-image placement (max 9). Each dict in a list can have
        "path" (required) and "z_order" (int, 0=back, higher=front).
        List order = z-order by default (last item on top).
      - image_mode (str, optional): How images are placed on the slide.
        "fill" (default) crops the image to fill a picture placeholder.
        "fit" places image(s) as freestanding shape(s) scaled proportionally
        with no cropping. Multiple images are arranged in a clean grid.
        "collage" arranges multiple images with overlap and stagger.
        Charts default to image_mode: "fit" (not just recommended).
      - notes (str | list[str]): Speaker notes shown in Presenter View.
        String for a single note, list for multiple talking points.
        Keep notes brief — a few key reminders, not full scripts.
      - shapes (list[dict]): Decorative shapes to overlay on the slide.
        Each shape dict requires 'type' plus type-specific params.
        Types: arrow, callout, badge, highlight, connector, process_arrow.
        Colors: theme names (accent1-accent6, dark, light) or hex (#0F69AF).

    Args:
        slides: List of slide content dicts.
        output_name: Output filename (saved to outputs/ directory).
        template_name: Template to use. Call list_templates to see options.
            Use get_template_layouts to see placeholder capacity guidance
            (max_comfortable_words, max_bullet_items) for each placeholder.

    Returns:
        Dict with success, output_path, file_size, num_slides, and per-slide details.
    """
    effective_template = template_name or _DEFAULT_TEMPLATE
    if not effective_template:
        available = engine.list_available()
        if len(available) == 1:
            effective_template = available[0]["name"]
        else:
            return {
                "success": False,
                "error": (
                    "No template specified and no default configured. "
                    f"Available templates: {[t['name'] for t in available]}. "
                    "Please ask the user which template to use."
                ),
            }

    return composer.create_presentation(
        slides=slides,
        template_name=effective_template,
        output_name=output_name,
    )


@mcp.tool()
def list_templates() -> dict:
    """
    List all available PowerPoint templates and their key information.

    Returns instantly with template names and files. Layout details are
    included for templates that have already been analyzed. Use
    get_template_layouts to trigger a full analysis of a specific template.

    Returns:
        Dict with template names and available layout names (if analyzed).
    """
    available = engine.list_available()
    result = {"success": True, "templates_directory": str(_TEMPLATES_DIR), "templates": []}
    for tmpl in available:
        entry: dict[str, Any] = {"name": tmpl["name"], "file": tmpl["file"]}
        # Include layout details if already analyzed (from memory or disk cache)
        info = engine.get_template(tmpl["name"])
        if info:
            entry["author"] = info.author
            entry["total_layouts"] = len(info.layouts)
            entry["layout_names"] = info.layout_names()
        result["templates"].append(entry)
    return result


@mcp.tool()
def get_template_layouts(template_name: str) -> dict:
    """
    Get detailed layout information for a template.

    Args:
        template_name: Name of the template to inspect.

    Returns:
        Dict with layouts, each showing name and available placeholder roles.
    """
    info = engine.get_template(template_name)
    if info is None:
        available = [t["name"] for t in engine.list_available()]
        return {
            "success": False,
            "error": f"Template '{template_name}' not found. Available: {available}",
        }

    layouts = []
    for layout in info.layouts:
        fillable = layout.get_fillable_placeholders()
        layouts.append(
            {
                "name": layout.name,
                "design_intent": layout.design_intent,
                "accepts": [p.role.value for p in fillable],
                "has_title": layout.has_title,
                "has_content": layout.has_content,
                "has_picture": layout.has_picture,
                "content_areas": layout.content_count,
                "placeholders": [
                    {
                        "idx": p.idx,
                        "role": p.role.value,
                        "name": p.name,
                        "hint_text": p.hint_text,
                        "default_font_size": p.default_font_size,
                        "font_family": p.font_family,
                        "text_alignment": p.text_alignment,
                        "is_primary": p.is_primary,
                        "semantic_role": p.semantic_role_hint,
                        "formatting_advice": p.formatting_recommendation,
                        "visual_priority": p.visual_priority,
                        "has_crop_geometry": p.has_crop_geometry,
                        "max_comfortable_words": p.max_comfortable_words,
                        "max_comfortable_lines": p.max_comfortable_lines,
                        "max_bullet_items": p.max_bullet_items,
                    }
                    for p in fillable
                ],
            }
        )

    return {
        "success": True,
        "template": template_name,
        "total_layouts": len(layouts),
        "layouts": layouts,
    }


@mcp.tool()
def register_template(template_path: str, template_display_name: str = "") -> dict:
    """
    Register a new template from a file path on disk.

    Args:
        template_path: Absolute path to the .potx or .pptx file.
        template_display_name: Optional display name. Defaults to filename stem.

    Returns:
        Dict with the registered template info including all discovered layouts.
    """
    try:
        name = template_display_name if template_display_name else None
        info = engine.register_template(template_path, name)
        return {
            "success": True,
            "template_name": info.name,
            "file": info.path.name,
            "total_layouts": len(info.layouts),
            "layout_names": info.layout_names(),
            "message": f"Template '{info.name}' registered with {len(info.layouts)} layouts.",
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def generate_image(
    prompt: str,
    output_name: str = "",
    size: str = "1024x1024",
    quality: str = "standard",
) -> dict:
    """
    Generate an image using UPTIMIZE gpt-image-1-mini-gs model.

    The returned file path can be passed as the 'image' field in a slide dict.

    CRITICAL: Generate ONE image at a time. Wait for completion before requesting
    the next image. Image generation takes 15-60 seconds per call. Even calling
    two in parallel will cause timeouts and failures due to API rate limits.

    You MUST generate images sequentially:
    1. Call generate_image for first image
    2. Wait for response
    3. Call generate_image for second image
    4. Wait for response
    5. Continue one at a time

    DO NOT use parallel tool calls for image generation.

    Args:
        prompt: Detailed description of the image to generate.
        output_name: Filename (e.g. 'hero-banner.png'). Auto-generated if empty.
        size: '1024x1024', '1536x1024' (landscape), '1024x1536' (portrait), or 'auto'.
        quality: 'low', 'medium', 'high', 'auto', 'standard', or 'hd'.

    Returns:
        Dict with success, path, prompt, size, file_size_bytes, and
        image_mode ("fill" — use this when placing the image in a slide
        so it integrates with the template's picture placeholder masks,
        especially on title/intro slides. Override with "fit" only if the
        generated image contains data content that must not be cropped).
    """
    if image_gen is None:
        if not _HAS_IMAGE_GEN:
            return {
                "success": False,
                "error": "Image generation requires the openai package. Install: pip install pptx-mcp[images]",
            }
        return {
            "success": False,
            "error": "Image generation is not configured. Set UPTIMIZE_OPENAI_API_KEY to enable.",
        }

    try:
        return image_gen.generate(
            prompt=prompt,
            output_name=output_name,
            size=size,
            quality=quality,
        )
    except Exception as e:
        return {"success": False, "error": f"Image generation failed: {e}"}


@mcp.tool()
def generate_chart(
    chart_type: str,
    data: dict | list[dict],
    title: str = "",
    output_name: str = "",
    xlabel: str = "",
    ylabel: str = "",
    legend: bool = True,
    template_name: str = "",
) -> dict:
    """
    Generate a chart as a PNG image using matplotlib.

    The returned file path can be passed as the 'image' field in a slide dict.
    Charts use the template's color scheme for visual consistency.

    NOTE: When generating multiple charts, call this tool sequentially
    (one at a time) rather than in parallel to avoid resource conflicts.

    Args:
        chart_type: 'bar', 'horizontal_bar', 'stacked_bar', 'line', 'pie', 'donut', 'scatter'.
        data: Single series dict {"Q1": 10, "Q2": 15} or multi-series list
            [{"name": "Revenue", "values": {"Q1": 10}}, ...].
        title: Chart title (optional).
        output_name: Filename (e.g. 'chart.png'). Auto-generated if empty.
        xlabel: X-axis label (optional).
        ylabel: Y-axis label (optional).
        legend: Show legend for multi-series (default: True).
        template_name: Template for color extraction (default: first available).

    Returns:
        Dict with success, path, chart_type, and file_size_bytes.
    """
    try:
        tn = template_name or _DEFAULT_TEMPLATE
        if not tn:
            available = engine.list_available()
            if available:
                tn = available[0]["name"]

        chart_builder = _ensure_chart_builder(tn)

        theme_colors = None
        if tn:
            info = engine.get_template(tn)
            if info:
                theme_colors = info.theme_colors

        return chart_builder.generate(
            chart_type=chart_type,
            data=data,
            title=title,
            output_name=output_name,
            xlabel=xlabel,
            ylabel=ylabel,
            legend=legend,
            theme_colors=theme_colors,
        )
    except Exception as e:
        return {"success": False, "error": f"Chart generation failed: {e}"}


@mcp.tool()
def list_generated_presentations() -> dict:
    """
    List all previously generated presentations in the outputs folder.

    Returns:
        Dict with list of files including filename, path, and size.
    """
    return composer.list_outputs()


@mcp.tool()
def download_presentation(filename: str) -> dict:
    """
    Get a generated presentation as base64-encoded content for download.

    Args:
        filename: Name of the file in the outputs directory.

    Returns:
        Dict with base64 content, file size, and MIME type.
    """
    file_path = _OUTPUTS_DIR / filename
    if not file_path.exists():
        return {"success": False, "error": f"File not found: {filename}"}

    with open(file_path, "rb") as f:
        data = f.read()

    return {
        "success": True,
        "filename": filename,
        "file_size": len(data),
        "content_base64": base64.b64encode(data).decode("utf-8"),
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main():
    """Run the MCP server."""
    transport = _TRANSPORT.lower()
    if transport == "sse":
        log.info("Starting PowerPoint MCP Server (SSE on %s:%s)...", _HOST, _PORT)
        mcp.run(transport="sse", host=_HOST, port=_PORT)
    else:
        log.info("Starting PowerPoint MCP Server (stdio)...")
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
