# Contributing

Guidelines for contributing to the PowerPoint MCP server.

## Design Principles

### Scope

This MCP focuses on generating and working with PowerPoint presentations from corporate templates. Features that naturally extend this capability -- reading slides, charts, images, speaker notes, overflow handling -- belong here. Features that are conceptually unrelated (e.g., rendering source code screenshots, converting to PDF) are better served by their own dedicated MCP servers.

When adding new functionality, prefer extending existing tool parameters over adding new MCP tools. For example, adding a key to a slide dict is simpler and more composable than introducing a new tool. New tools are fine when they represent a genuinely distinct operation (e.g., `generate_chart` is separate from `create_presentation`).

Avoid adding dependencies unless they provide clear value. Prefer what's already available transitively (e.g., Pillow via python-pptx). Optional dependencies gated behind extras (like openai for AI image generation) are acceptable.

### Presentation Quality

- Slides should be lean. Avoid overloading them with text.
- Use speaker notes (`notes` field) for brief talking points and key reminders -- keep the projected slide clean and visual. Notes should be concise; a presenter glances at them, not reads them verbatim.
- Storytelling/decorative images can be cropped to fill placeholders (`image_mode: "fill"`).
- Data images (charts, diagrams, screenshots) must never be cropped -- use `image_mode: "fit"` to preserve full content.

### API Design

- Prefer hints over enforcement. For example, `generate_chart` returns `recommended_image_mode: "fit"` as a suggestion; the calling agent makes the final decision.
- New parameters must default to backward-compatible values. Existing tool calls must continue to work without modification (e.g., `image_mode` defaults to `"fill"`, `notes` is optional).
- All text goes into existing template placeholders. Only images (fit mode) and decorative shapes may be added as freestanding elements.

### Git Workflow

- The `main` branch is protected. Always create a feature branch before making changes.
- One branch per feature set / release.
- Run `pytest tests/ -v` and `ruff check src/` before considering work complete.

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

## Testing

### Test Organisation

- **`test_server.py`** -- Server module import and configuration tests
- **`test_smoke.py`** -- End-to-end functionality tests (presentations, layouts, charts)
- **`conftest.py`** -- Shared fixtures and test utilities

### Running Tests

```bash
# All tests
pytest

# Specific file
pytest tests/test_server.py

# Specific class or test
pytest tests/test_server.py::TestServerImport
pytest tests/test_server.py::TestServerImport::test_module_imports
```

### Environment Variable Testing

Tests that verify environment variable behaviour use the `isolated_server_env`
fixture from `conftest.py`. It clears all `PPTX_*` and `UPTIMIZE_*` env vars
and forces a fresh import of the server module for full isolation.

```python
def test_custom_setting(isolated_server_env, monkeypatch):
    monkeypatch.setenv("PPTX_TRANSPORT", "sse")
    import pptx_mcp.server as server
    importlib.reload(server)
    assert server._TRANSPORT == "sse"
```

**Important:** Never commit actual API keys or tokens. Tests use mock values.
