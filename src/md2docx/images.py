"""
Image utilities for handling logos and images in Word documents.

Supports downloading from URLs and converting SVG to PNG.
"""

import os
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional
from urllib.parse import urlparse
import tempfile

# Environment variable that provides a comma-separated host allowlist
ALLOWED_HOSTS_ENV = "MD2DOCX_ALLOWED_IMAGE_HOSTS"

# Allowlist provider hook so alternative providers can be plugged in later
_allowed_hosts_provider: Callable[[], set[str]] = lambda: {
    host.strip()
    for host in os.getenv(ALLOWED_HOSTS_ENV, "").split(",")
    if host.strip()
}


def load_image(path_or_url: str) -> BytesIO:
    """
    Load an image from a file path or URL.

    Automatically handles:
    - Local file paths
    - HTTP/HTTPS URLs
    - SVG to PNG conversion

    Args:
        path_or_url: File path or URL to the image.

    Returns:
        BytesIO stream containing the image data (PNG format if SVG).

    Raises:
        ValueError: If the image cannot be loaded.
    """
    if not path_or_url:
        raise ValueError("Empty path or URL provided")

    # Check if it's a URL
    if path_or_url.startswith(("http://", "https://")):
        if not _is_host_allowed(path_or_url):
            raise ValueError("Remote image host is not allowed")
        return _download_image(path_or_url)
    else:
        return _load_local_image(path_or_url)


def set_allowed_hosts_provider(provider: Callable[[], set[str]]) -> None:
    """Register a custom provider for allowed remote hosts."""
    global _allowed_hosts_provider
    _allowed_hosts_provider = provider


def _is_host_allowed(url: str) -> bool:
    """Check whether the URL host is in the configured allowlist."""
    host = urlparse(url).hostname
    if not host:
        return False
    allowed_hosts = _allowed_hosts_provider()
    return bool(allowed_hosts) and host in allowed_hosts


def _download_image(url: str) -> BytesIO:
    """Download image from URL."""
    try:
        import requests
    except ImportError:
        raise ImportError("requests library is required for downloading images. Install with: pip install requests")

    try:
        # Use a proper User-Agent to avoid being blocked
        headers = {
            "User-Agent": "Mozilla/5.0 (compatible; MD2DOCX/1.0)"
        }
        response = requests.get(url, timeout=30, headers=headers)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        image_data = response.content

        # Check if SVG
        if "svg" in content_type.lower() or url.lower().endswith(".svg"):
            return _convert_svg_to_png(image_data)

        return BytesIO(image_data)

    except requests.RequestException as e:
        raise ValueError(f"Failed to download image from {url}: {e}")


def _load_local_image(path: str) -> BytesIO:
    """Load image from local file."""
    file_path = Path(path)

    if not file_path.exists():
        raise ValueError(f"Image file not found: {path}")

    # Check if SVG
    if file_path.suffix.lower() == ".svg":
        with open(file_path, "rb") as f:
            svg_data = f.read()
        return _convert_svg_to_png(svg_data)

    # Load regular image
    with open(file_path, "rb") as f:
        return BytesIO(f.read())


def _convert_svg_to_png(svg_data: bytes) -> BytesIO:
    """
    Convert SVG data to PNG format.

    Args:
        svg_data: SVG file content as bytes.

    Returns:
        BytesIO stream containing PNG image data.
    """
    try:
        import cairosvg
    except ImportError:
        raise ImportError(
            "cairosvg library is required for SVG conversion. "
            "Install with: pip install cairosvg"
        )

    try:
        png_data = cairosvg.svg2png(bytestring=svg_data)
        return BytesIO(png_data)
    except Exception as e:
        raise ValueError(f"Failed to convert SVG to PNG: {e}")


def get_image_dimensions(image_stream: BytesIO) -> tuple[int, int]:
    """
    Get the dimensions of an image.

    Args:
        image_stream: BytesIO stream containing image data.

    Returns:
        Tuple of (width, height) in pixels.
    """
    try:
        from PIL import Image
    except ImportError:
        raise ImportError(
            "Pillow library is required for image processing. "
            "Install with: pip install Pillow"
        )

    # Save position and reset
    pos = image_stream.tell()
    image_stream.seek(0)

    try:
        with Image.open(image_stream) as img:
            width, height = img.size
    finally:
        image_stream.seek(pos)

    return width, height


def is_supported_format(path_or_url: str) -> bool:
    """
    Check if the image format is supported.

    Supported formats: PNG, JPEG, GIF, TIFF, BMP, SVG (converted to PNG)

    Args:
        path_or_url: File path or URL to check.

    Returns:
        True if the format is supported, False otherwise.
    """
    lower_path = path_or_url.lower()
    supported_extensions = [".png", ".jpg", ".jpeg", ".gif", ".tiff", ".tif", ".bmp", ".svg"]
    return any(lower_path.endswith(ext) for ext in supported_extensions)
