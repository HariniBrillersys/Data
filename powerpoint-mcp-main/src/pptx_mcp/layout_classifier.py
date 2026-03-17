"""
Layout Classifier - Intelligently select the best layout for given content.

Maps content intent to the most appropriate template layout, considering
what placeholders the content needs and what each layout provides.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Optional

from .template_engine import LayoutInfo, PlaceholderRole, TemplateInfo


class SlideIntent(str, Enum):
    """What kind of slide the user wants to create."""

    TITLE_SLIDE = "title_slide"
    SECTION_DIVIDER = "section_divider"
    SINGLE_CONTENT = "single_content"
    TWO_COLUMN = "two_column"
    THREE_COLUMN = "three_column"
    CONTENT_WITH_IMAGE = "content_with_image"
    FULL_IMAGE = "full_image"
    DATA_IMAGE = "data_image"  # Chart/diagram that must not be cropped
    KEY_MESSAGE = "key_message"
    TABLE_OF_CONTENTS = "table_of_contents"
    CONTENT_WITH_NOTICE = "content_with_notice"
    BLANK = "blank"


@dataclass
class SlideSpec:
    """Specification for what a slide should contain."""

    intent: SlideIntent
    title: Optional[str] = None
    subtitle: Optional[str] = None
    body: Optional[str] = None
    content: Optional[str | list] = None
    content_left: Optional[str | list] = None
    content_right: Optional[str | list] = None
    content_1: Optional[str | list] = None
    content_2: Optional[str | list] = None
    content_3: Optional[str | list] = None
    notice: Optional[str | list] = None
    image: Optional[str] = None
    image_mode: Optional[str] = None  # "fill" (default) or "fit" (no crop)
    notes: Optional[str | list] = None  # Speaker notes for presenter view

    def to_content_dict(self) -> dict:
        """Convert to the content dict expected by SlideProxy.fill()."""
        d = {}
        for key in [
            "title",
            "subtitle",
            "body",
            "content",
            "content_left",
            "content_right",
            "content_1",
            "content_2",
            "content_3",
            "notice",
            "image",
            "image_mode",
            "notes",
        ]:
            val = getattr(self, key, None)
            if val is not None:
                d[key] = val
        return d


class LayoutClassifier:
    """
    Select the best layout from a template for a given content specification.

    Uses a scoring system that considers:
    1. Required placeholders (content needs vs. layout provides)
    2. Layout name heuristics (keywords like "Title", "Divider", etc.)
    3. Placeholder count match (prefer layouts that match content complexity)
    """

    INTENT_PATTERNS: dict[SlideIntent, list[str]] = {
        SlideIntent.TITLE_SLIDE: ["title 0", "title 01", "title 02", "title 03", "intro"],
        SlideIntent.SECTION_DIVIDER: ["divider", "section"],
        SlideIntent.SINGLE_CONTENT: [
            "content 0",
            "content 01",
            "content 02",
            "content 03",
            "content 04",
            "content 05",
            "content 06",
            "content white",
            "one content",
        ],
        SlideIntent.TWO_COLUMN: ["two content"],
        SlideIntent.THREE_COLUMN: ["three content"],
        SlideIntent.CONTENT_WITH_IMAGE: ["content with image", "content with visual"],
        SlideIntent.FULL_IMAGE: ["big picture"],
        SlideIntent.DATA_IMAGE: ["only title", "blank", "content 0"],
        SlideIntent.KEY_MESSAGE: ["key message"],
        SlideIntent.TABLE_OF_CONTENTS: ["table of content", "agenda", "toc"],
        SlideIntent.CONTENT_WITH_NOTICE: ["content and notice"],
        SlideIntent.BLANK: ["blank", "only title"],
    }

    def __init__(self, template_info: TemplateInfo):
        self.template = template_info

    def select_layout(
        self,
        spec: SlideSpec,
        preferred_layout: Optional[str] = None,
    ) -> LayoutInfo:
        """Select the best layout for a given slide specification."""
        if preferred_layout:
            layout = self.template.get_layout(preferred_layout)
            if layout:
                return layout
            layout = self.template.find_layout(preferred_layout)
            if layout:
                return layout

        scored = []
        for layout in self.template.layouts:
            score = self._score_layout(layout, spec)
            if score > 0:
                scored.append((score, layout))

        if not scored:
            for layout in self.template.layouts:
                if layout.has_content or layout.has_title:
                    return layout
            raise ValueError("No suitable layout found in template")

        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1]

    def classify_content(self, content: dict) -> SlideSpec:
        """Classify raw user content into a SlideSpec with intent."""
        has_image = "image" in content
        has_content = "content" in content
        has_left = "content_left" in content
        has_right = "content_right" in content
        has_c2 = "content_2" in content
        has_c3 = "content_3" in content
        has_notice = "notice" in content
        has_subtitle = "subtitle" in content
        image_mode = content.get("image_mode", "fill")
        is_fit = image_mode in ("fit", "collage")

        if has_c3 or has_c2:
            intent = SlideIntent.THREE_COLUMN if has_c3 else SlideIntent.TWO_COLUMN
        elif has_left and has_right:
            intent = SlideIntent.TWO_COLUMN
        elif has_subtitle and not has_content:
            # Title/intro slide — subtitle signals a title slide even when an
            # image is present (the image is decorative and should fill the
            # picture placeholder mask).
            intent = SlideIntent.TITLE_SLIDE
        elif has_image and is_fit and not has_content:
            # Data image (chart/diagram) without text — needs full uncropped display
            intent = SlideIntent.DATA_IMAGE
        elif has_image and is_fit and has_content:
            # Data image with accompanying text — use content layout, image placed fit
            intent = SlideIntent.DATA_IMAGE
        elif has_image and has_content:
            intent = SlideIntent.CONTENT_WITH_IMAGE
        elif has_image and not has_content:
            intent = SlideIntent.FULL_IMAGE
        elif has_notice:
            intent = SlideIntent.CONTENT_WITH_NOTICE
        elif has_content:
            intent = SlideIntent.SINGLE_CONTENT
        else:
            intent = SlideIntent.TITLE_SLIDE if has_subtitle else SlideIntent.SECTION_DIVIDER

        spec = SlideSpec(intent=intent)
        for key in [
            "title",
            "subtitle",
            "body",
            "content",
            "content_left",
            "content_right",
            "content_1",
            "content_2",
            "content_3",
            "notice",
            "image",
            "image_mode",
            "notes",
        ]:
            if key in content:
                setattr(spec, key, content[key])

        return spec

    def auto_select(
        self,
        content: dict,
        preferred_layout: Optional[str] = None,
    ) -> tuple[LayoutInfo, SlideSpec]:
        """Full auto: classify content, then select the best layout."""
        spec = self.classify_content(content)
        layout = self.select_layout(spec, preferred_layout)
        return layout, spec

    def _score_layout(self, layout: LayoutInfo, spec: SlideSpec) -> int:
        """Score how well a layout matches a slide specification."""
        score = 0
        name_lower = layout.name.lower()

        patterns = self.INTENT_PATTERNS.get(spec.intent, [])
        for pattern in patterns:
            if pattern in name_lower:
                score += 50
                break

        if spec.title and layout.has_title:
            score += 10
        if spec.content and layout.has_content:
            score += 20
        if spec.image and layout.has_picture:
            # Check if picture placeholders have crop geometry
            pic_phs = [p for p in layout.placeholders if p.role == PlaceholderRole.PICTURE]
            has_crop = any(p.has_crop_geometry for p in pic_phs)

            # For DATA_IMAGE intent, picture placeholders are undesirable
            # because they crop images. Prefer layouts without them.
            if spec.intent == SlideIntent.DATA_IMAGE:
                score -= 25  # Base penalty for any picture placeholder
                if has_crop:
                    score -= 30  # Additional penalty for crop geometry (total -55)
            else:
                # For decorative images (CONTENT_WITH_IMAGE, FULL_IMAGE, TITLE_SLIDE)
                # Picture placeholders are desired — crop geometry (masks) are
                # part of the template design and create branded visuals.
                score += 30  # Base bonus for picture placeholder

        if spec.intent == SlideIntent.TWO_COLUMN:
            if layout.content_count == 2:
                score += 40
            elif layout.content_count == 1:
                score -= 20
        elif spec.intent == SlideIntent.THREE_COLUMN:
            if layout.content_count == 3:
                score += 40
            elif layout.content_count != 3:
                score -= 20

        # DATA_IMAGE: prefer simple layouts with space for a freestanding image
        if spec.intent == SlideIntent.DATA_IMAGE:
            if not layout.has_picture:
                score += 15  # bonus for no picture placeholder
            if layout.has_title and layout.content_count == 0 and not layout.has_picture:
                score += 20  # ideal: title-only layout (maximum image space)
            if spec.content and layout.has_content:
                score += 10  # has text content too, so content area is useful

        if spec.intent == SlideIntent.TITLE_SLIDE:
            if "content" in name_lower and "title" not in name_lower:
                score -= 30
            # When a title slide has an image, strongly prefer layouts with a
            # picture placeholder so the image fills the template's mask shape.
            if spec.image and layout.has_picture:
                score += 25

        if spec.intent == SlideIntent.SECTION_DIVIDER:
            if layout.content_count > 0 and "divider" not in name_lower:
                score -= 20

        if "01" in layout.name:
            score += 3
        elif "02" in layout.name:
            score += 2

        # Design intent matching (use layout.design_intent property)
        intent_to_design = {
            SlideIntent.TITLE_SLIDE: "title_heavy",
            SlideIntent.SECTION_DIVIDER: "divider",
            SlideIntent.SINGLE_CONTENT: "content_heavy",
            SlideIntent.CONTENT_WITH_IMAGE: "visual_heavy",
            SlideIntent.FULL_IMAGE: "visual_heavy",
            SlideIntent.DATA_IMAGE: "balanced",
        }
        expected_design = intent_to_design.get(spec.intent, "balanced")
        if layout.design_intent == expected_design:
            score += 20  # Significant bonus for matching design intent

        # Primary placeholder bonus: boost scores when is_primary matches content intent
        if spec.title and layout.has_title:
            title_phs = [p for p in layout.placeholders if p.role == PlaceholderRole.TITLE]
            if any(p.is_primary for p in title_phs):
                score += 15  # Primary title = strong match for title content

        if spec.content:
            content_phs = [p for p in layout.placeholders if p.role in (PlaceholderRole.CONTENT, PlaceholderRole.BODY)]
            if any(p.is_primary for p in content_phs):
                score += 10  # Primary content area = good match for content-heavy slides

        # Semantic role matching: prefer layouts where semantic hints align with intent
        for ph in layout.placeholders:
            if spec.intent == SlideIntent.TITLE_SLIDE and "emphasis" in ph.semantic_role_hint:
                score += 5  # Emphasis placeholders good for title slides
            if spec.intent == SlideIntent.SECTION_DIVIDER and "chapter" in ph.semantic_role_hint:
                score += 5  # Chapter placeholders good for dividers

        # Visual priority bonus: prefer layouts with prominent placeholders for key content
        if spec.content:
            content_phs = [p for p in layout.placeholders if p.role == PlaceholderRole.CONTENT]
            if content_phs:
                avg_priority = sum(p.visual_priority for p in content_phs) / len(content_phs)
                if avg_priority >= 70:
                    score += 10  # Bonus for visually prominent content placeholders

        return score
