"""Shared test fixtures."""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image as PILImage

# Project root (one level up from tests/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"


@pytest.fixture
def templates_dir():
    """Path to the templates directory."""
    return TEMPLATES_DIR


@pytest.fixture
def output_dir(tmp_path):
    """Temporary output directory for test presentations."""
    return tmp_path


@pytest.fixture
def test_image(tmp_path) -> Path:
    """Create a small test PNG image (200x100, landscape)."""
    img_path = tmp_path / "test_image.png"
    img = PILImage.new("RGB", (200, 100), color=(0, 111, 175))
    img.save(img_path)
    return img_path


@pytest.fixture
def test_image_wide(tmp_path) -> Path:
    """Create a wide 16:9 test PNG image (1920x1080)."""
    img_path = tmp_path / "test_chart.png"
    img = PILImage.new("RGB", (1920, 1080), color=(150, 215, 210))
    img.save(img_path)
    return img_path


@pytest.fixture
def test_images_multi(tmp_path) -> list[Path]:
    """Create 4 small test PNG images of different colors."""
    colors = [
        (0, 111, 175),  # blue
        (150, 215, 210),  # teal
        (255, 220, 185),  # peach
        (80, 50, 145),  # purple
    ]
    paths = []
    for i, color in enumerate(colors):
        img_path = tmp_path / f"multi_{i}.png"
        # Vary dimensions slightly to test proportional scaling
        w = 400 + i * 50
        h = 300 - i * 20
        img = PILImage.new("RGB", (w, h), color=color)
        img.save(img_path)
        paths.append(img_path)
    return paths


# ---------------------------------------------------------------------------
# Server module isolation
# ---------------------------------------------------------------------------

# All env vars read by server.py at module level
_SERVER_ENV_VARS = [
    "PPTX_TEMPLATES_DIR",
    "PPTX_OUTPUTS_DIR",
    "PPTX_DEFAULT_TEMPLATE",
    "PPTX_TRANSPORT",
    "PPTX_HOST",
    "PPTX_PORT",
    "PPTX_LOG_LEVEL",
    "UPTIMIZE_OPENAI_API_KEY",
    "UPTIMIZE_ENV",
]


@pytest.fixture
def isolated_server_env(monkeypatch):
    """Clear all PPTX_*/UPTIMIZE_* env vars and force a fresh server import.

    Use this fixture whenever a test needs to verify server configuration
    behaviour in a clean environment.
    """
    import sys

    for var in _SERVER_ENV_VARS:
        monkeypatch.delenv(var, raising=False)

    # Drop cached module so the next import re-evaluates module-level code
    if "pptx_mcp.server" in sys.modules:
        del sys.modules["pptx_mcp.server"]

    yield

    if "pptx_mcp.server" in sys.modules:
        del sys.modules["pptx_mcp.server"]
