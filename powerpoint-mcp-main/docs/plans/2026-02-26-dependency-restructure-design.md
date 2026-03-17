# Dependency Restructure — Design

**Date:** 2026-02-26
**Status:** Approved
**Scope:** Swap chart/image dependency positions; defer Phases 3-6 to backlog

## Context

The project has completed Phases 1-2 (Placeholder Intelligence, Presentation Quality Overhaul). The remaining phases (3: Multi-User Architecture, 4: Template Upload, 5: File Download, 6: HTTP Transport & AWS Deployment) are deferred indefinitely to a backlog — there are no users yet and these features are premature.

Two dependency issues need fixing:

1. **`openai>=1.0.0` is a core dependency** but image generation requires an API key tied to the Merck UPTIMIZE proxy. Individual users cannot share their API keys, and there is no server-side key provisioning yet. The `openai` package and its transitive dependencies (~20MB) are installed for every user even if they never use image generation.

2. **`matplotlib>=3.8.0` is an optional dependency** (`pip install pptx-mcp[charts]`) but chart generation is self-contained — it runs locally with no API keys, no network calls, no credentials. It should just work out of the box.

## Decision

Swap their positions:

| Feature | Current | New |
|---------|---------|-----|
| Charts (matplotlib) | Optional `[charts]` extra | **Core dependency** |
| Images (openai) | Core dependency | **Optional `[images]` extra** |

**Principle:** If it works offline with no credentials, it's core. If it needs external services and API keys, it's optional.

---

## Section 1: Charts → Core Dependency

### Changes

**`pyproject.toml`:**
- Move `matplotlib>=3.8.0` from `[project.optional-dependencies].charts` to `dependencies`
- Remove the `charts` extra entirely (no longer needed)
- Remove `matplotlib>=3.8.0` from the `dev` extra (it's already in core)

**`chart_builder.py`:**
- Remove the `try/except ImportError` guard around `import matplotlib`
- Remove `HAS_MATPLOTLIB` flag
- Remove the `ImportError` check in `ChartBuilder.__init__`
- Import `matplotlib`, `matplotlib.pyplot`, and `numpy` directly at module level

**`server.py`:**
- Remove the `try/except ImportError` block around `ChartBuilder` initialization
- Initialize `_chart_builder` directly (it will always succeed)
- Remove the runtime `None` check in `generate_chart` tool
- Update the chart instruction text in `_instructions` (remove "requires matplotlib" language)

**`CLAUDE.md` / `CONTRIBUTING.md`:**
- Update install commands (remove `[charts]` references)
- Update architecture notes about chart dependency

**Tests:**
- No new tests needed — charts already have test coverage in `test_smoke.py`
- Update any test comments referencing `[charts]` extra

---

## Section 2: Images → Optional Dependency

### Changes

**`pyproject.toml`:**
- Remove `openai>=1.0.0` from core `dependencies`
- Add new optional extra: `images = ["openai>=1.0.0"]`
- Add `openai>=1.0.0` to the `dev` extra (so dev installs still get it)

**`image_generator.py`:**
- Add a `try/except ImportError` guard around `from openai import OpenAI`
- Add `HAS_OPENAI = True/False` flag (mirrors old chart pattern)
- Check `HAS_OPENAI` in `ImageGenerator.__init__` and raise `ImportError` if missing

**`server.py`:**
- Change the `image_generator` import to a guarded `try/except ImportError`
- Initialization becomes: try import → check API key → instantiate (three gates)
- Update the runtime error message in `generate_image` to distinguish between:
  - Package not installed: `"Image generation requires the openai package. Install: pip install pptx-mcp[images]"`
  - Package installed but no API key: `"Image generation is not configured. Set UPTIMIZE_OPENAI_API_KEY to enable."`
- Update `_IMAGE_INSTRUCTION` in the instructions string to reflect the new gating

**Tests:**
- Update `test_image_gen_disabled` to verify the new error message wording
- Update `test_image_gen_enabled` to verify the enabled instruction wording
- Add a test verifying the instruction mentions `[images]` install when openai is unavailable (can be done via the existing `isolated_server_env` fixture)

---

## Section 3: Roadmap Cleanup

### Changes

- Phases 3-6 (Multi-User Architecture, Template Upload, File Download, HTTP Transport) remain documented in the roadmap as "Backlog" status
- No planning or implementation work is expected for these phases
- The roadmap's progress table reflects the current state: 2/6 phases complete, 4 deferred

### What This Means

The v1.0 milestone scope is effectively: Phase 1 (Placeholder Intelligence) + Phase 2 (Presentation Quality) + this dependency restructure. Phases 3-6 are deferred to a future milestone when multi-user support becomes relevant.

---

## Out of Scope

- Server-side API key provisioning (future — when multi-user support is needed)
- Conditional tool registration (tools remain always-visible, guard at runtime)
- Any changes to the `ImageGenerator` class internals (API logic, dual-endpoint strategy)
- Changes to how the `generate_image` or `generate_chart` tool parameters or return values work
