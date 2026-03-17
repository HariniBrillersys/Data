"""
Presentation Composer - Orchestrate multi-slide presentation creation.

Takes structured content (list of slides with content) and produces a
complete, professional presentation using the template system.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from .layout_classifier import LayoutClassifier
from .overflow import validate_presentation
from .shape_builder import ShapeAnnotator
from .slide_builder import SlideBuilder
from .template_engine import TemplateEngine

log = logging.getLogger(__name__)


class PresentationComposer:
    """
    High-level orchestrator for creating presentations.

    Usage:
        engine = TemplateEngine("templates")
        engine.scan()
        composer = PresentationComposer(engine)

        result = composer.create_presentation(
            template_name="Uptimize Master",
            slides=[
                {"layout": "Title 01", "title": "My Deck", "subtitle": "Q4 2025"},
                {"title": "Key Findings", "content": ["Point 1", "Point 2"]},
                {"title": "Details", "content_left": "Left", "content_right": "Right"},
            ],
            output_path="outputs/my_deck.pptx",
        )
    """

    def __init__(self, engine: TemplateEngine, outputs_dir: str | Path = "outputs"):
        self.engine = engine
        self.outputs_dir = Path(outputs_dir)
        self.outputs_dir.mkdir(parents=True, exist_ok=True)

    def create_presentation(
        self,
        slides: list[dict],
        template_name: str = "Uptimize Master",
        output_path: Optional[str | Path] = None,
        output_name: Optional[str] = None,
    ) -> dict:
        """
        Create a complete presentation from a list of slide definitions.

        Args:
            slides: List of slide content dicts. Each dict can contain:
                - layout (str, optional): Preferred layout name
                - title (str, optional): Slide title
                - subtitle (str, optional): Subtitle (for title slides)
                - body (str, optional): Sub-headline text
                - content (str | list, optional): Main content
                - content_left, content_right: For two-column layouts
                - content_1, content_2, content_3: For three-column layouts
                - notice (str | list, optional): Sidebar content
                - image (str | list, optional): Path to image file, or a list
                    of paths/dicts for multi-image (max 9 per slide).
                - image_mode (str, optional): "fill" (default, crop to placeholder),
                    "fit" (no crop grid layout), or "collage" (overlapping).
                    Use "fit" for charts/diagrams, "collage" for visual spreads.
                - notes (str | list[str], optional): Speaker notes for presenter
                    view. String for a single note or list for bullet points.
                - shapes (list[dict], optional): Decorative shape annotations
            template_name: Name of template to use.
            output_path: Full output path. If not given, uses output_name.
            output_name: Output filename (placed in outputs_dir).

        Returns:
            Dict with success status, path, and details.
        """
        try:
            # Resolve output path
            if output_path:
                out = Path(output_path)
            elif output_name:
                out = self.outputs_dir / output_name
            else:
                out = self.outputs_dir / "presentation.pptx"

            # Get template info
            template_info = self.engine.get_template(template_name)
            if template_info is None:
                return {
                    "success": False,
                    "error": f"Template '{template_name}' not found. Available: {self.engine.list_templates()}",
                }

            # Open a fresh presentation from template
            prs = self.engine.open_presentation(template_name)
            builder = SlideBuilder(prs, template_info)
            classifier = LayoutClassifier(template_info)
            annotator = ShapeAnnotator(
                theme_colors=template_info.theme_colors,
                slide_width=template_info.slide_width,
                slide_height=template_info.slide_height,
            )

            # Process each slide
            slide_results = []
            for i, slide_def in enumerate(slides):
                result = self._add_slide(builder, classifier, annotator, slide_def, i + 1)
                slide_results.append(result)

            # Post-generation validation: detect and fix text overflow
            overflow_issues = validate_presentation(prs)
            if overflow_issues:
                for issue_group in overflow_issues:
                    sn = issue_group["slide_number"]
                    for issue in issue_group["issues"]:
                        log.debug(
                            "  [overflow fix] Slide %d, '%s': ratio=%sx, %s",
                            sn,
                            issue["placeholder_name"],
                            issue["overflow_ratio"],
                            issue["action_taken"],
                        )

            # Save
            saved_path = builder.save(out)

            return {
                "success": True,
                "output_path": str(saved_path),
                "filename": saved_path.name,
                "file_size": saved_path.stat().st_size,
                "num_slides": len(slides),
                "template_used": template_name,
                "slides": slide_results,
                "overflow_fixes": overflow_issues if overflow_issues else None,
                "message": f"Presentation created with {len(slides)} slides using '{template_name}' template.",
            }

        except Exception as e:
            return {
                "success": False,
                "error": str(e),
            }

    def _analyze_content_density(self, content: dict, layout) -> list[dict]:
        """Analyze content density before filling placeholders.

        Args:
            content: Content dict with fields like title, content, content_left, etc.
            layout: LayoutInfo object with placeholder metadata

        Returns:
            List of warning dicts for overloaded content
        """
        from .template_engine import PlaceholderRole

        warnings = []

        # Map content field names to placeholder roles
        field_to_role = {
            "title": PlaceholderRole.TITLE,
            "subtitle": PlaceholderRole.SUBTITLE,
            "body": PlaceholderRole.BODY,
            "content": PlaceholderRole.CONTENT,
            "content_left": PlaceholderRole.CONTENT_LEFT,
            "content_right": PlaceholderRole.CONTENT_RIGHT,
            "content_1": PlaceholderRole.CONTENT_1,
            "content_2": PlaceholderRole.CONTENT_2,
            "content_3": PlaceholderRole.CONTENT_3,
            "notice": PlaceholderRole.NOTICE,
        }

        for field_name, role in field_to_role.items():
            if field_name not in content:
                continue

            field_content = content[field_name]
            if not field_content:
                continue

            # Find placeholder for this role
            placeholder = layout.get_by_role(role)
            if not placeholder or placeholder.max_comfortable_words == 0:
                continue

            # Count content volume
            if isinstance(field_content, list):
                # Bullet list
                bullet_count = len(field_content)
                total_words = sum(len(str(item).split()) for item in field_content)

                # Check bullet count
                if placeholder.max_bullet_items > 0:
                    bullet_ratio = bullet_count / placeholder.max_bullet_items
                    if bullet_ratio > 1.0:
                        severity = "severe" if bullet_ratio > 1.5 else "mild"
                        warnings.append(
                            {
                                "type": "content_overflow" if severity == "severe" else "dense_content",
                                "placeholder": field_name,
                                "bullets": bullet_count,
                                "comfortable_max_bullets": placeholder.max_bullet_items,
                                "bullet_ratio": round(bullet_ratio, 2),
                                "words": total_words,
                                "comfortable_max": placeholder.max_comfortable_words,
                                "ratio": round(total_words / placeholder.max_comfortable_words, 2),
                                "needs_reword": bullet_ratio > 1.5,
                                "suggestion": (
                                    "Content significantly exceeds placeholder capacity. "
                                    "Move detail to speaker notes or split across slides."
                                    if severity == "severe"
                                    else "Consider moving supporting detail to speaker notes"
                                ),
                            }
                        )
                        continue

                # Check word count
                word_count = total_words
            else:
                # String content
                word_count = len(str(field_content).split())

            # Check if content fits comfortably
            ratio = word_count / placeholder.max_comfortable_words
            if ratio <= 1.0:
                continue  # Fits comfortably

            # Generate warning
            severity = "severe" if ratio > 1.5 else "mild"
            warnings.append(
                {
                    "type": "content_overflow" if severity == "severe" else "dense_content",
                    "placeholder": field_name,
                    "words": word_count,
                    "comfortable_max": placeholder.max_comfortable_words,
                    "ratio": round(ratio, 2),
                    "needs_reword": ratio > 1.5,
                    "suggestion": (
                        "Content significantly exceeds placeholder capacity. "
                        "Move detail to speaker notes or split across slides."
                        if severity == "severe"
                        else "Consider moving supporting detail to speaker notes"
                    ),
                }
            )

        return warnings

    def _add_slide(
        self,
        builder: SlideBuilder,
        classifier: LayoutClassifier,
        annotator: ShapeAnnotator,
        slide_def: dict,
        slide_num: int,
    ) -> dict:
        """Add a single slide to the presentation."""
        # Extract layout preference and shapes without mutating the original dict
        preferred_layout = slide_def.get("layout")
        shapes_defs = slide_def.get("shapes")

        # Build content dict (everything except 'layout' and 'shapes')
        content = {k: v for k, v in slide_def.items() if v is not None and k not in ("layout", "shapes")}

        # Auto-select layout and classify content
        layout, spec = classifier.auto_select(content, preferred_layout)

        log.debug(
            "  Slide %d: intent=%s -> layout='%s'",
            slide_num,
            spec.intent.value,
            layout.name,
        )

        # Pre-flight content analysis
        warnings = self._analyze_content_density(content, layout)

        # Add slide and fill content
        proxy = builder.add_slide(layout.name)
        fill_results = proxy.fill(spec.to_content_dict())

        # Add shape annotations if provided
        shapes_added = []
        if shapes_defs and proxy._slide:
            shapes_added = annotator.annotate(proxy._slide, shapes_defs)

        result = {
            "slide_number": slide_num,
            "layout_used": layout.name,
            "intent": spec.intent.value,
            "fields_filled": fill_results,
            "shapes_added": shapes_added if shapes_added else None,
        }

        # Add warnings if any
        if warnings:
            result["warnings"] = warnings

        return result

    def list_outputs(self) -> dict:
        """List all generated presentations in the outputs directory."""
        files = []
        for path in sorted(self.outputs_dir.glob("*.pptx")):
            if path.name.startswith("~$"):
                continue
            files.append(
                {
                    "filename": path.name,
                    "path": str(path),
                    "size_bytes": path.stat().st_size,
                }
            )
        return {
            "success": True,
            "outputs_directory": str(self.outputs_dir),
            "total_files": len(files),
            "files": files,
        }
