import json
import logging
import sys
from pathlib import Path

# Add src to the path
sys.path.append(str(Path(__file__).parent / "src"))

from pptx_mcp.composer import PresentationComposer
from pptx_mcp.template_engine import TemplateEngine

log = logging.getLogger(__name__)

def generate_ppt_from_json(json_data: str | dict, output_filename: str = "presentation.pptx", templates_dir: str = "templates", outputs_dir: str = "outputs") -> str:
    """
    Generates a PowerPoint presentation from a strictly structured JSON response.

    Args:
        json_data: The JSON string or dictionary output by the LLM.
        output_filename: The name of the resulting PPTX file.
        templates_dir: Path to the PowerPoint templates directory.
        outputs_dir: Path where the generated PPTX will be saved.

    Returns:
        The absolute path to the generated presentation.
    """
    if isinstance(json_data, str):
        try:
            data = json.loads(json_data)
        except json.JSONDecodeError as e:
            log.error(f"Failed to parse JSON data: {e}")
            raise ValueError(f"Invalid JSON data provided: {e}")
    else:
        data = json_data

    presentation_data = data.get("presentation", {})
    if not presentation_data:
         raise ValueError("JSON data must contain a 'presentation' key at the root.")

    title = presentation_data.get("title", "Data Analysis Report")
    slides_data = presentation_data.get("slides", [])

    if not slides_data:
        log.warning("No slides found in the JSON data.")

    # Initialize the template engine and presentation composer
    engine = TemplateEngine(Path(templates_dir))
    composer = PresentationComposer(engine, Path(outputs_dir))

    # Construct the slides list for the composer
    formatted_slides = []

    # 1. Add a Title Slide
    formatted_slides.append({
        "layout": "Title Slide", # Try to use standard layout names. The Composer will fallback if needed.
        "title": title,
        "subtitle": "Generated via LangGraph Data Agent"
    })

    # 2. Add content slides
    for slide in slides_data:
        slide_title = slide.get("title", "Insight")
        content = slide.get("content", [])
        notes = slide.get("speaker_notes", "")

        formatted_slides.append({
            "title": slide_title,
            "content": content,
            "notes": notes
        })

    # Generate the presentation
    # The default template logic will pick the first available one if not specified
    available = engine.list_available()
    template_name = available[0]["name"] if available else ""

    result = composer.create_presentation(
        slides=formatted_slides,
        output_name=output_filename,
        template_name=template_name
    )

    if result.get("success"):
        return result.get("output_path")
    else:
        error_msg = result.get("error", "Unknown error during presentation generation.")
        log.error(f"Presentation generation failed: {error_msg}")
        raise RuntimeError(f"Failed to generate presentation: {error_msg}")
