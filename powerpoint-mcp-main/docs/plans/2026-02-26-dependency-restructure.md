# Dependency Restructure Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Make charts a core dependency (always available) and images an optional extra (`pip install pptx-mcp[images]`).

**Architecture:** Swap the gating pattern: remove matplotlib's import guard and optional extra, add an import guard and optional extra for openai. Update server initialization, tool runtime guards, instructions, docs, and tests.

**Tech Stack:** Python packaging (pyproject.toml), conditional imports, pytest

---

### Task 1: Update pyproject.toml dependencies

**Files:**
- Modify: `pyproject.toml`

**Step 1: Edit pyproject.toml**

Move `matplotlib>=3.8.0` to core dependencies, move `openai>=1.0.0` to a new `[images]` extra, remove the `[charts]` extra, and update `[dev]` to include `openai` instead of `matplotlib` (since matplotlib is now core).

```toml
[project]
dependencies = [
    "python-pptx>=1.0.0",
    "mcp[cli]>=1.0.0",
    "lxml>=4.9.0",
    "matplotlib>=3.8.0",
]

[project.optional-dependencies]
images = ["openai>=1.0.0"]
dev = [
    "pytest>=7.0",
    "ruff>=0.4.0",
    "openai>=1.0.0",
    "pre-commit>=3.5.0",
]
```

**Step 2: Reinstall the package**

Run: `pip install -e ".[dev]"`
Expected: Installs successfully, matplotlib in core, openai via dev extra.

**Step 3: Commit**

```bash
git add pyproject.toml
git commit -m "refactor: move matplotlib to core deps, openai to optional [images] extra"
```

---

### Task 2: Remove chart_builder.py import guard

**Files:**
- Modify: `src/pptx_mcp/chart_builder.py`

**Step 1: Replace the guarded import with direct imports**

Change lines 28-38 from:

```python
try:
    import matplotlib

    matplotlib.use("Agg")  # Non-interactive backend
    import matplotlib.pyplot as plt
    import numpy as np

    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False
    np = None  # type: ignore[assignment]
```

To:

```python
import matplotlib

matplotlib.use("Agg")  # Non-interactive backend
import matplotlib.pyplot as plt
import numpy as np
```

**Step 2: Remove the HAS_MATPLOTLIB check in __init__**

In `ChartBuilder.__init__` (line 152-154), remove:

```python
if not HAS_MATPLOTLIB:
    raise ImportError("matplotlib is required for chart generation")
```

**Step 3: Update module docstring**

Change lines 8-9 from:

```
Requires matplotlib (optional dependency):
    pip install pptx-mcp[charts]
```

To:

```
Uses matplotlib (core dependency) for chart rendering.
```

**Step 4: Run tests to verify charts still work**

Run: `pytest tests/test_smoke.py::TestChartGeneration -v`
Expected: All chart tests PASS (the `@pytest.mark.skipif` will never trigger since matplotlib is always available, but that gets cleaned up in Task 5).

**Step 5: Commit**

```bash
git add src/pptx_mcp/chart_builder.py
git commit -m "refactor: remove matplotlib import guard (now core dependency)"
```

---

### Task 3: Add import guard to image_generator.py

**Files:**
- Modify: `src/pptx_mcp/image_generator.py`

**Step 1: Add guarded import**

Change line 16 from:

```python
from openai import OpenAI
```

To:

```python
try:
    from openai import OpenAI

    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False
    OpenAI = None  # type: ignore[assignment, misc]
```

**Step 2: Add guard in __init__**

At the start of `ImageGenerator.__init__` (line 26), before the existing `if not api_key:` check, add:

```python
if not HAS_OPENAI:
    raise ImportError(
        "Image generation requires the openai package. "
        "Install with: pip install pptx-mcp[images]"
    )
```

**Step 3: Update module docstring**

Change the docstring (lines 1-7) to:

```python
"""
Image generation via UPTIMIZE OpenAI API.

Uses the gpt-image-1-mini-gs model through the Merck UPTIMIZE proxy.
Tries the Images API first, then falls back to the Responses API if
the endpoint is not available.

Requires the openai package (optional dependency):
    pip install pptx-mcp[images]
"""
```

**Step 4: Verify import works without openai**

This can't easily be tested in the same process, but the pattern is proven (it's exactly what charts used to do). The server.py changes in Task 4 will complete the integration.

**Step 5: Commit**

```bash
git add src/pptx_mcp/image_generator.py
git commit -m "refactor: add openai import guard (now optional dependency)"
```

---

### Task 4: Update server.py initialization and instructions

**Files:**
- Modify: `src/pptx_mcp/server.py`

**Step 1: Guard the ImageGenerator import**

Change lines 29-30 from:

```python
from .composer import PresentationComposer
from .image_generator import ImageGenerator
from .template_engine import TemplateEngine
```

To:

```python
from .composer import PresentationComposer
from .template_engine import TemplateEngine

try:
    from .image_generator import ImageGenerator

    _HAS_IMAGE_GEN = True
except ImportError:
    _HAS_IMAGE_GEN = False
    ImageGenerator = None  # type: ignore[assignment, misc]
```

**Step 2: Update image initialization block**

Change lines 195-208 from:

```python
# Initialize image generator (optional)
image_gen: ImageGenerator | None = None
if _UPTIMIZE_API_KEY:
    try:
        image_gen = ImageGenerator(
            api_key=_UPTIMIZE_API_KEY,
            base_url=_UPTIMIZE_BASE_URL,
            images_dir=_IMAGES_DIR,
        )
        log.info("  Image generation: enabled (env=%s)", _UPTIMIZE_ENV)
    except Exception as e:
        log.warning("  Image generation: failed to initialize (%s)", e)
else:
    log.info("  Image generation: disabled (set UPTIMIZE_OPENAI_API_KEY to enable)")
```

To:

```python
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
```

**Step 3: Update chart initialization — remove try/except**

Change lines 210-225 from:

```python
# Initialize chart builder (optional)
try:
    from .chart_builder import ChartBuilder

    # Use default template's minor font if available
    default_font = "Calibri"
    if template_names:
        default_template = engine.get_template(template_names[0])
        if default_template:
            default_font = default_template.theme_colors.minor_font

    _chart_builder = ChartBuilder(images_dir=_IMAGES_DIR, font_family=default_font)
    log.info("  Chart generation: enabled (font: %s)", default_font)
except ImportError:
    _chart_builder = None
    log.info("  Chart generation: disabled (pip install pptx-mcp[charts])")
```

To:

```python
# Initialize chart builder (core dependency — always available)
from .chart_builder import ChartBuilder

default_font = "Calibri"
if template_names:
    default_template = engine.get_template(template_names[0])
    if default_template:
        default_font = default_template.theme_colors.minor_font

_chart_builder = ChartBuilder(images_dir=_IMAGES_DIR, font_family=default_font)
log.info("  Chart generation: enabled (font: %s)", default_font)
```

**Step 4: Update _IMAGE_INSTRUCTION**

Change lines 74-78 from:

```python
_IMAGE_INSTRUCTION = (
    f"Image generation available via generate_image. Images saved to {_IMAGES_DIR}."
    if _UPTIMIZE_API_KEY
    else "Image generation NOT configured. Set UPTIMIZE_OPENAI_API_KEY to enable."
)
```

To:

```python
_IMAGE_INSTRUCTION = (
    f"Image generation available via generate_image. Images saved to {_IMAGES_DIR}."
    if _UPTIMIZE_API_KEY
    else "Image generation NOT available. Requires: pip install pptx-mcp[images] and UPTIMIZE_OPENAI_API_KEY env var."
)
```

**Step 5: Update chart instruction in _instructions string**

Change lines 94-98 from:

```python
CHART GENERATION:
Charts can be generated via the generate_chart tool (requires matplotlib).
Returns a PNG file path usable as the 'image' field in any slide dict.
Charts default to image_mode: "fit" — no cropping, full content visible.
Agents can override with image_mode: "fill" if cropping is desired (rare).
```

To:

```python
CHART GENERATION:
Charts can be generated via the generate_chart tool.
Returns a PNG file path usable as the 'image' field in any slide dict.
Charts default to image_mode: "fit" — no cropping, full content visible.
Agents can override with image_mode: "fill" if cropping is desired (rare).
```

**Step 6: Update generate_image runtime error for better messaging**

Change lines 460-464 from:

```python
if image_gen is None:
    return {
        "success": False,
        "error": "Image generation is not configured. Set UPTIMIZE_OPENAI_API_KEY to enable.",
    }
```

To:

```python
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
```

**Step 7: Remove generate_chart runtime guard**

Change lines 511-515 from:

```python
if _chart_builder is None:
    return {
        "success": False,
        "error": "Chart generation requires matplotlib. Install: pip install pptx-mcp[charts]",
    }
```

To: (remove entirely — `_chart_builder` is always initialized)

**Step 8: Run tests**

Run: `pytest tests/test_server.py -v`
Expected: All pass. Some test assertions about instruction wording may need updating (handled in Task 5).

**Step 9: Commit**

```bash
git add src/pptx_mcp/server.py
git commit -m "refactor: charts always initialized, images guarded by package + env var"
```

---

### Task 5: Update tests

**Files:**
- Modify: `tests/test_smoke.py`
- Modify: `tests/test_server.py`

**Step 1: Remove skipif from chart tests in test_smoke.py**

Change lines 35-41 from:

```python
# Check if matplotlib is available
try:
    from pptx_mcp.chart_builder import ChartBuilder

    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False
```

To:

```python
from pptx_mcp.chart_builder import ChartBuilder
```

Change line 350 from:

```python
@pytest.mark.skipif(not HAS_MATPLOTLIB, reason="matplotlib not installed")
class TestChartGeneration:
```

To:

```python
class TestChartGeneration:
```

**Step 2: Fix the pre-existing test key mismatch**

Change lines 423-433 — the test checks `recommended_image_mode` but the code returns `image_mode`:

```python
def test_chart_returns_recommended_image_mode(self, output_dir):
    """Verify chart generation returns recommended_image_mode: fit."""
    builder = ChartBuilder(images_dir=output_dir)

    result = builder.generate(
        chart_type="bar",
        data={"A": 1, "B": 2},
    )

    assert result["success"] is True
    assert result["recommended_image_mode"] == "fit"
```

To:

```python
def test_chart_returns_image_mode_fit(self, output_dir):
    """Verify chart generation returns image_mode: fit."""
    builder = ChartBuilder(images_dir=output_dir)

    result = builder.generate(
        chart_type="bar",
        data={"A": 1, "B": 2},
    )

    assert result["success"] is True
    assert result["image_mode"] == "fit"
```

**Step 3: Update test_server.py image instruction assertions**

Change lines 120-127 — the disabled message now mentions pip install:

```python
def test_image_gen_disabled(self, isolated_server_env):
    """Without API key, instructions warn about disabled generation."""
    import pptx_mcp.server as server

    importlib.reload(server)

    assert "Image generation NOT configured" in server._instructions
    assert "Set UPTIMIZE_OPENAI_API_KEY" in server._instructions
```

To:

```python
def test_image_gen_disabled(self, isolated_server_env):
    """Without API key or package, instructions warn about disabled generation."""
    import pptx_mcp.server as server

    importlib.reload(server)

    assert "Image generation NOT available" in server._instructions
    assert "pptx-mcp[images]" in server._instructions
```

**Step 4: Run all tests**

Run: `pytest tests/ -v`
Expected: All pass.

**Step 5: Lint**

Run: `ruff check src/`
Expected: Clean.

**Step 6: Commit**

```bash
git add tests/test_smoke.py tests/test_server.py
git commit -m "test: update tests for dependency restructure (charts core, images optional)"
```

---

### Task 6: Update documentation

**Files:**
- Modify: `CLAUDE.md`
- Modify: `CONTRIBUTING.md`
- Modify: `README.md`

**Step 1: Update CLAUDE.md**

In the Architecture section, change:
- `chart_builder.py     # Chart/graph generation via matplotlib (optional)` → `chart_builder.py     # Chart/graph generation via matplotlib`
- `image_generator.py   # AI image generation via UPTIMIZE API` → `image_generator.py   # AI image generation via UPTIMIZE API (optional)`

In Key Technical Details, change:
- `Charts require matplotlib (optional [charts] dependency)` → `Charts use matplotlib (core dependency, always available)`

In Development Commands, change:
- `# Install (with all dev dependencies including matplotlib)` → `# Install (with all dev dependencies including openai for image generation)`

**Step 2: Update CONTRIBUTING.md**

In the Scope section, change:
- `Optional dependencies gated behind extras (like matplotlib for charts) are acceptable.` → `Optional dependencies gated behind extras (like openai for AI image generation) are acceptable.`

In Development Commands, change:
- `# Install (with all dev dependencies including matplotlib)` → `# Install (with all dev dependencies including openai for image generation)`

**Step 3: Update README.md**

In the Installation Options table, change:
- `| **With charts** | pip install "pptx-mcp[charts] @ ..." |` → `| **With images** | pip install "pptx-mcp[images] @ ..." |`

In the MCP Tools table, change:
- `generate_chart` description: remove `(requires [charts])`
- `generate_image` description: add `(requires [images])`

In the Chart Generation section, remove the trailing install instruction:
- Remove: `Charts automatically use the template's color scheme. Install the charts extra: pip install "pptx-mcp[charts] @ ..."`
- Replace with: `Charts are a core feature and always available. They automatically use the template's color scheme.`

**Step 4: Commit**

```bash
git add CLAUDE.md CONTRIBUTING.md README.md
git commit -m "docs: update for dependency restructure (charts core, images optional)"
```

---

### Task 7: Final verification

**Step 1: Run full test suite**

Run: `pytest tests/ -v`
Expected: All pass.

**Step 2: Run linter**

Run: `ruff check src/`
Expected: Clean.

**Step 3: Verify server starts**

Run: `python -c "from pptx_mcp import server; print('OK')"` 
Expected: Prints `OK` with log lines showing chart generation enabled and image generation status.

**Step 4: Commit (if any final fixes needed)**

Only if steps 1-3 revealed issues.
