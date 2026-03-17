"""
Tests for the MCP server module and entry point.

Verifies that server.py can be imported (catches module-level errors like
broken f-strings), that environment variables are handled correctly, and
that sensitive values are never leaked into user-facing instructions.
"""

from __future__ import annotations

import importlib


class TestServerImport:
    """Verify the server module loads without errors."""

    def test_module_imports(self):
        """Import server.py — catches f-string and syntax errors."""
        from pptx_mcp import server

        assert hasattr(server, "mcp")
        assert hasattr(server, "main")

    def test_instructions_valid(self):
        """The _instructions f-string should contain key documentation sections."""
        from pptx_mcp import server

        assert isinstance(server._instructions, str)
        assert "TEMPLATE SELECTION:" in server._instructions
        assert "IMAGE PLACEMENT MODES:" in server._instructions
        assert "MULTI-IMAGE:" in server._instructions
        # Dict examples should render with literal braces, not be swallowed
        assert '{"path":' in server._instructions


class TestServerEnvVars:
    """Verify environment variable handling and defaults."""

    def test_defaults(self, isolated_server_env):
        """Defaults are used when no env vars are set."""
        import pptx_mcp.server as server

        importlib.reload(server)

        assert server._TRANSPORT == "stdio"
        assert server._HOST == "127.0.0.1"
        assert server._PORT == 8000
        assert server._DEFAULT_TEMPLATE == ""
        assert server._UPTIMIZE_ENV == "dev"
        assert server._LOG_LEVEL == "INFO"

    def test_custom_transport(self, isolated_server_env, monkeypatch):
        """PPTX_TRANSPORT overrides the default."""
        monkeypatch.setenv("PPTX_TRANSPORT", "sse")

        import pptx_mcp.server as server

        importlib.reload(server)

        assert server._TRANSPORT == "sse"

    def test_custom_port(self, isolated_server_env, monkeypatch):
        """PPTX_PORT is converted to int."""
        monkeypatch.setenv("PPTX_PORT", "9000")

        import pptx_mcp.server as server

        importlib.reload(server)

        assert server._PORT == 9000
        assert isinstance(server._PORT, int)

    def test_custom_template(self, isolated_server_env, monkeypatch):
        """PPTX_DEFAULT_TEMPLATE appears in instructions."""
        monkeypatch.setenv("PPTX_DEFAULT_TEMPLATE", "Corporate Template")

        import pptx_mcp.server as server

        importlib.reload(server)

        assert server._DEFAULT_TEMPLATE == "Corporate Template"
        assert "Corporate Template" in server._instructions

    def test_api_key_not_leaked(self, isolated_server_env, monkeypatch):
        """API key must never appear in user-facing instructions."""
        monkeypatch.setenv("UPTIMIZE_OPENAI_API_KEY", "sk-secret-12345")

        import pptx_mcp.server as server

        importlib.reload(server)

        assert server._UPTIMIZE_API_KEY == "sk-secret-12345"
        assert "sk-secret" not in server._instructions

    def test_log_level_uppercase(self, isolated_server_env, monkeypatch):
        """PPTX_LOG_LEVEL is normalised to uppercase."""
        monkeypatch.setenv("PPTX_LOG_LEVEL", "debug")

        import pptx_mcp.server as server

        importlib.reload(server)

        assert server._LOG_LEVEL == "DEBUG"


class TestServerConfig:
    """Verify conditional instruction generation."""

    def test_image_gen_enabled(self, isolated_server_env, monkeypatch):
        """With API key, instructions mention image generation."""
        monkeypatch.setenv("UPTIMIZE_OPENAI_API_KEY", "test-key")

        import pptx_mcp.server as server

        importlib.reload(server)

        assert "Image generation available" in server._instructions
        assert "generate_image" in server._instructions

    def test_image_gen_disabled(self, isolated_server_env):
        """Without API key or package, instructions warn about disabled generation."""
        import pptx_mcp.server as server

        importlib.reload(server)

        assert "Image generation NOT available" in server._instructions
        assert "pptx-mcp[images]" in server._instructions
