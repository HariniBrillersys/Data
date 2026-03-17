# Presentation Quality Overhaul — Design

**Date:** 2026-02-25
**Status:** Approved
**Scope:** Image cropping protection, content density management, formatting bug fixes

## Context

After Phase 1 (Placeholder Intelligence) delivered template introspection and metadata, a review of generated presentation quality revealed three categories of issues:

1. **Charts and data images get cropped** by template picture placeholders that have masking geometry (rounded shapes, custom crop regions). The system hints `recommended_image_mode: "fit"` but agents can ignore it.
2. **No content density management** — slides accept unlimited text, shrink it to unreadable font sizes via reactive overflow handling. No pre-flight capacity analysis, no guidance to agents about slide design principles.
3. **Formatting bugs** degrade output quality: template formatting lost when setting text, chart fonts hardcoded to Verdana, broken hint text extraction, poor shape text contrast.

## Decisions

- **Phases 1-2 and 1-3 from original roadmap (Multi-User Architecture, Template Upload) are moved to backlog.** This quality overhaul takes priority.
- **Liquid Design System icon integration is deferred** to a follow-up phase. Quality foundation first.
- **Advisory over restrictive**: the system provides information and warnings but never truncates or refuses content. Presentations always get generated.

---

## Section 1: Image Cropping Protection

### Problem

Template picture placeholders can have non-rectangular geometry (custom shapes, crop regions, large rounded corners) that acts as a mask. When a chart or graph is placed via `insert_picture()`, PowerPoint crops the image to the placeholder shape at render time. The data is lost visually.

The current `generate_chart` tool returns `recommended_image_mode: "fit"` as a hint, but nothing enforces this. The layout classifier penalizes picture placeholders for `DATA_IMAGE` intent (-25 points) but has no awareness of which specific placeholders have masking geometry.

### Design

**Placeholder geometry detection:**
- During template scanning, inspect each picture placeholder's XML for:
  - `<a:custGeom>` — custom (non-rectangular) geometry
  - `<a:srcRect>` / `<a:fillRect>` — pre-defined crop regions
  - Rounded corners with radii that significantly clip content
- Add `has_crop_geometry: bool` to `PlaceholderInfo`

**Layout classifier enhancement:**
- Placeholders with `has_crop_geometry=True` receive an additional scoring penalty beyond the existing -25 for picture placeholders
- Strongly prefers layouts with safe rectangular picture areas or title-only layouts for data images

**Automatic mode switching (safety net):**
- In `slide_builder.py`, when `image_mode="fill"` is requested but the target placeholder has crop geometry, automatically switch to `"fit"` mode and log a warning
- This catches cases where agents ignore hints

**Chart builder enforcement:**
- `generate_chart` returns `image_mode: "fit"` as the **default** (not just a recommendation)
- Agents can override with explicit `image_mode: "fill"` if they want cropping (opt-in)

---

## Section 2: Content Density Management

### Problem

No guardrails on text volume. A 500-word paragraph gets shrunk to unreadable 10pt font. No bullet count limits. No pre-flight analysis. The overflow system is purely reactive — detects overflow after content is placed and can only shrink text or flag `needs_reword`.

### Design

**Placeholder capacity calculation:**
- Each placeholder gets computed capacity metadata based on actual dimensions and font size:
  - `max_comfortable_words`: words that fit at readable font size
  - `max_comfortable_lines`: lines that fit without shrinking
  - `max_bullet_items`: recommended max bullet count (typically 5-7)
- Exposed in `PlaceholderInfo` and available via `get_template_layouts`

**Pre-flight content analysis:**
- Before filling each placeholder, composer measures incoming content against comfortable capacity
- Three outcomes:
  - **Fits comfortably** — proceed normally
  - **Mild overflow (1.0x-1.5x)** — proceed, return warning: `{"warning": "dense_content", "placeholder": "content", "words": 150, "comfortable_max": 100, "suggestion": "Consider moving supporting detail to speaker notes"}`
  - **Severe overflow (>1.5x)** — proceed, return stronger warning with `needs_reword: true`

**Agent instruction guidance:**
- Add slide design principles to tool instruction string:
  - 6x6 guideline (max ~6 bullets, ~6 words each)
  - Content-to-visual ratio guidance
  - Speaker notes as primary relief valve for detail
- Teaching instructions, not enforcement

**Capacity hints in response:**
- `create_presentation` per-slide results include capacity data
- Agents can iteratively improve overloaded slides

### Key Principle

Advisory, not restrictive. Never truncates or refuses content. Provides information for informed decisions. The presentation always gets generated.

---

## Section 3: Formatting Bug Fixes

### Problem

Several concrete bugs degrade output quality.

### Fixes

1. **`_set_text()` formatting preservation** — Use `p.add_run()` instead of `p.text = text` when `p.runs` is empty. Preserves template run-level formatting (font, size, color, bold).

2. **Chart font from template** — Extract `<a:majorFont>` / `<a:minorFont>` from theme XML. Pass to chart builder instead of hardcoded Verdana.

3. **`_extract_hint_text()` fix** — Change `first_p.text` to `''.join(first_p.itertext())` to read text inside `<a:r>` run elements correctly.

4. **Consolidate font size resolution** — Merge duplicated logic in `overflow.py` and `template_engine.py` into shared utility. Eliminates divergent fallback defaults.

5. **Shape text contrast** — Calculate fill color luminance: `(0.299*R + 0.587*G + 0.114*B)`. Dark text when luminance > 0.5, white text when <= 0.5.

6. **Stale `to_catalog()`** — Update to include Phase 1 metadata fields or remove if unused.

---

## Out of Scope

- Liquid Design System icon integration (deferred to follow-up phase)
- Multi-user architecture (moved to backlog)
- Template upload workflow (moved to backlog)
- Shape collision detection with placeholders (future enhancement)
- Vector chart output / SVG/EMF rendering (future enhancement)
- Additional chart types (waterfall, area, radar)
