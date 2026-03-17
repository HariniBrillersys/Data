"""
Image generation via UPTIMIZE OpenAI API.

Uses the gpt-image-1-mini-gs model through the Merck UPTIMIZE proxy.
Tries the Images API first, then falls back to the Responses API if
the endpoint is not available.

Requires the openai package (optional dependency):
    pip install pptx-mcp[images]
"""

from __future__ import annotations

import base64
import logging
import uuid
from pathlib import Path

try:
    from openai import OpenAI

    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False
    OpenAI = None  # type: ignore[assignment, misc]

log = logging.getLogger(__name__)


class ImageGenerator:
    """Generate images via the UPTIMIZE OpenAI proxy."""

    MODEL = "gpt-image-1-mini-gs"

    def __init__(self, api_key: str, base_url: str, images_dir: Path) -> None:
        if not HAS_OPENAI:
            raise ImportError(
                "Image generation requires the openai package. Install with: pip install pptx-mcp[images]"
            )
        if not api_key:
            raise ValueError("UPTIMIZE_OPENAI_API_KEY is not set.")
        self.client = OpenAI(
            api_key=api_key,
            base_url=base_url,
            default_query={"api-version": "preview"},
        )
        self.images_dir = images_dir
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self._use_responses_api: bool | None = None
        self._lock = __import__("threading").Lock()

    def generate(
        self,
        prompt: str,
        output_name: str = "",
        size: str = "1024x1024",
        quality: str = "standard",
    ) -> dict:
        """Generate an image and save it to disk.

        Thread-safe: Uses a lock to ensure only one image is generated at a time,
        preventing API rate limit issues and timeout errors.
        """
        if not output_name:
            output_name = f"image_{uuid.uuid4().hex[:8]}.png"
        if not output_name.lower().endswith(".png"):
            output_name += ".png"

        output_path = self.images_dir / output_name

        # Enforce sequential generation with a lock
        with self._lock:
            if self._use_responses_api is None:
                return self._try_both(prompt, output_path, size, quality)
            elif self._use_responses_api:
                return self._generate_via_responses(prompt, output_path, size, quality)
            else:
                return self._generate_via_images(prompt, output_path, size, quality)

    def _try_both(self, prompt: str, output_path: Path, size: str, quality: str) -> dict:
        """Try Images API first, fall back to Responses API."""
        try:
            log.debug("Trying Images API (images.generate)...")
            result = self._generate_via_images(prompt, output_path, size, quality)
            if result["success"]:
                self._use_responses_api = False
                log.debug("Images API works. Caching for future calls.")
                return result
        except Exception as e:
            log.debug("Images API failed: %s", e)

        try:
            log.debug("Trying Responses API (responses.create)...")
            result = self._generate_via_responses(prompt, output_path, size, quality)
            if result["success"]:
                self._use_responses_api = True
                log.debug("Responses API works. Caching for future calls.")
                return result
        except Exception as e:
            log.warning("Responses API also failed: %s", e)

        return {
            "success": False,
            "error": (
                "Both Images API and Responses API failed. "
                "Check that UPTIMIZE_OPENAI_API_KEY is valid and the "
                "gpt-image-1-mini-gs model is available."
            ),
        }

    _QUALITY_MAP = {
        "standard": "medium",
        "hd": "high",
        "low": "low",
        "medium": "medium",
        "high": "high",
        "auto": "auto",
    }

    _SIZE_MAP = {
        "1024x1024": "1024x1024",
        "1792x1024": "1536x1024",
        "1024x1792": "1024x1536",
        "1536x1024": "1536x1024",
        "1024x1536": "1024x1536",
        "auto": "auto",
        "landscape": "1536x1024",
        "portrait": "1024x1536",
        "square": "1024x1024",
    }

    def _generate_via_images(self, prompt: str, output_path: Path, size: str, quality: str) -> dict:
        """Use the standard OpenAI Images API endpoint."""
        azure_quality = self._QUALITY_MAP.get(quality, "medium")
        azure_size = self._SIZE_MAP.get(size, "1024x1024")

        response = self.client.images.generate(
            model=self.MODEL,
            prompt=prompt,
            size=azure_size,
            quality=azure_quality,
            n=1,
        )

        image_data = response.data[0]

        if hasattr(image_data, "b64_json") and image_data.b64_json:
            img_bytes = base64.b64decode(image_data.b64_json)
        elif hasattr(image_data, "url") and image_data.url:
            import urllib.request

            with urllib.request.urlopen(image_data.url) as resp:
                img_bytes = resp.read()
        else:
            return {
                "success": False,
                "error": "No image data in response (no b64_json or url field).",
            }

        output_path.write_bytes(img_bytes)

        return {
            "success": True,
            "path": str(output_path.resolve()),
            "prompt": prompt,
            "size": size,
            "file_size_bytes": len(img_bytes),
            "method": "images.generate",
            "image_mode": "fill",
        }

    def _generate_via_responses(self, prompt: str, output_path: Path, size: str, quality: str) -> dict:
        """Use the UPTIMIZE Responses API endpoint."""
        response = self.client.responses.create(
            model=self.MODEL,
            input=prompt,
        )

        img_bytes = None

        if hasattr(response, "output") and response.output:
            for item in response.output:
                if hasattr(item, "type") and item.type == "image":
                    if hasattr(item, "image") and hasattr(item.image, "b64_json"):
                        img_bytes = base64.b64decode(item.image.b64_json)
                        break
                if hasattr(item, "content"):
                    for content_block in item.content:
                        if hasattr(content_block, "type") and content_block.type == "image":
                            if hasattr(content_block, "image_url"):
                                url = content_block.image_url
                                if hasattr(url, "url") and url.url.startswith("data:"):
                                    b64_part = url.url.split(",", 1)[1]
                                    img_bytes = base64.b64decode(b64_part)
                                    break
                        if hasattr(content_block, "type") and content_block.type == "output_image":
                            if hasattr(content_block, "image_base64"):
                                img_bytes = base64.b64decode(content_block.image_base64)
                                break

        if img_bytes is None and hasattr(response, "output_text") and response.output_text:
            text = response.output_text.strip()
            try:
                img_bytes = base64.b64decode(text)
                if not img_bytes[:4] == b"\x89PNG":
                    img_bytes = None
            except Exception:
                img_bytes = None

        if img_bytes is None:
            return {
                "success": False,
                "error": "Responses API returned no extractable image data.",
            }

        output_path.write_bytes(img_bytes)

        return {
            "success": True,
            "path": str(output_path.resolve()),
            "prompt": prompt,
            "size": size,
            "file_size_bytes": len(img_bytes),
            "method": "responses.create",
            "image_mode": "fill",
        }
