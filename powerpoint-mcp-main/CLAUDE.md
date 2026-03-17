# CLAUDE.md

Guidance for AI coding assistants working on this repository.

## What This Is

MCP server that generates PowerPoint presentations from corporate templates. Users provide content + structure; the server selects appropriate template layouts and fills their pre-designed placeholders. Shapes, charts, and images can be layered on top for visual enhancement.

## Architecture

```
src/pptx_mcp/
  __init__.py
  server.py            # FastMCP entry point (tools, stdio + SSE transport)
  template_engine.py   # Load .potx/.pptx templates, introspect layouts/placeholders
  slide_builder.py     # Fill existing placeholders by idx (NEVER create new text shapes)
  layout_classifier.py # Map content intent -> best layout selection
  composer.py          # Orchestrate multi-slide creation
  overflow.py          # Post-generation text overflow detection & auto-fit
  shape_builder.py     # Decorative shape annotations (arrows, callouts, badges, etc.)
  chart_builder.py     # Chart/graph generation via matplotlib
  theme_colors.py      # Extract color scheme from template theme XML
  image_generator.py   # AI image generation via UPTIMIZE API (optional)
```

### Core Design Rules

1. **ALL text content goes into EXISTING template placeholders identified by their idx.** Never call `slide.shapes.add_textbox()` or create new text shapes. This preserves template fonts, colors, positions, and design integrity.

2. **Shapes may be added on top for visual enhancement** via `shape_builder.py`. These are decorative elements (arrows, callouts, badges) that complement the placeholder content -- they must not replace or overlap placeholder text. Shapes use the template's theme colors for brand consistency.

3. **Charts are rendered as PNG images** via `chart_builder.py` and placed into picture placeholders. They are not native PowerPoint chart objects.

4. **Images support three placement modes** via `image_mode`:
   - `"fill"` (default): Inserted into a PICTURE placeholder, cropped to fill. Best for storytelling/decorative images.
   - `"fit"`: Placed as a freestanding picture shape via `slide.shapes.add_picture()`, scaled proportionally with NO cropping. Best for charts, diagrams, and data images. The layout classifier uses a dedicated `DATA_IMAGE` intent for this mode, preferring title-only or blank layouts to maximise image space. Multiple images are arranged in a clean grid.
   - `"collage"`: Multiple images arranged with overlapping, staggered positions and slight size variation. Creates a dynamic visual feel. The `image` field accepts a list of paths or dicts (`{"path": str, "z_order": int}`) for multi-image placement (max 9). List order = z-order (last = front).

5. **Speaker notes** can be added to any slide via the `notes` field (string or list of strings). Notes appear only in Presenter View. Use them for brief talking points and key reminders to keep slides lean. Notes should be concise -- a presenter glances at them, not reads them verbatim.

### Placeholder Role System

Every placeholder in a layout is classified into a `PlaceholderRole`:
- `title`, `subtitle`, `body` (sub-headline), `content` (main area)
- `content_left`, `content_right` (two-column), `content_1/2/3` (three-column)
- `picture`, `notice` (sidebar), `footer`, `slide_number`

The `LayoutClassifier` maps user content intent to the best available layout using a scoring system based on name patterns and placeholder availability.

### Theme Color System

Colors are extracted from the template's `ppt/theme/theme1.xml` at load time. The `ThemeColors` dataclass provides:
- 12 standard slots: dk1, dk2, lt1, lt2, accent1-6, hlink, folHlink
- Custom brand colors from `<a:custClrLst>`
- `resolve_color()` method that accepts theme names, custom names, or hex values
- `accent_cycle()` for chart and shape color sequences

### Overflow Detection

After every presentation generation, `overflow.py` estimates whether text fits in each placeholder. Title/subtitle placeholders use a stricter threshold (0.85) to prevent wrapping. If overflow is detected:
1. Enables `normAutofit` (PowerPoint's shrink-to-fit)
2. For severe overflow (>1.5x), also explicitly reduces font size
3. Reads actual margins from the placeholder's `bodyPr` XML attributes

### Configuration

All config is via environment variables:
- `PPTX_TEMPLATES_DIR` - templates folder path
- `PPTX_OUTPUTS_DIR` - outputs folder path
- `PPTX_DEFAULT_TEMPLATE` - default template name (empty = agent asks user)
- `PPTX_TRANSPORT` - `stdio` or `sse`
- `PPTX_HOST` / `PPTX_PORT` - SSE bind settings
- `PPTX_LOG_LEVEL` - DEBUG, INFO, WARNING (default: INFO)

## Development Commands

```bash
# Install (with all dev dependencies including openai for image generation)
pip install -e ".[dev]"

# Run MCP server
pptx-mcp              # entry point
python -m pptx_mcp    # module

# Run tests
pytest tests/ -v

# Lint
ruff check src/
```

## Templates

Located in `templates/`. Accepts `.potx` and `.pptx` files. Bundled template:
- **Uptimize Master** - 35 layouts, Uptimize/Merck branding

## Key Technical Details

- PPTX/POTX files are ZIP archives containing XML (no PowerPoint installation needed)
- `.potx` files are converted to `.pptx` by rewriting `[Content_Types].xml` in the ZIP
- python-pptx handles text, shapes, tables, images, basic formatting
- python-pptx does NOT handle: complex animations, think-cell objects, macros
- All template slides are removed before generation; only layouts are preserved
- Shapes are generated via python-pptx's shape API (MSO_SHAPE) -- pure XML, no PowerPoint needed
- Charts use matplotlib (core dependency, always available)
- Title/subtitle placeholders get proactive `normAutofit` to prevent text wrapping

See [CONTRIBUTING.md](CONTRIBUTING.md) for design principles, scope guidelines, and git workflow.

## Testing

```bash
pytest tests/ -v
```

Test presentations are saved to `outputs/`. Open them in PowerPoint to verify:
1. Text appears in the correct placeholder positions
2. Fonts and colors match the template design
3. Bullet points are properly formatted
4. Layout selection matches content intent
5. Shapes render correctly with theme colors
6. Long titles shrink-to-fit rather than wrapping
7. Speaker notes appear in Presenter View
8. Images with `image_mode: "fit"` display without cropping
