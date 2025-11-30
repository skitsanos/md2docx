"""
Markdown to Word Document Converter

A modular library for converting Markdown content to Word documents
with support for custom branding (fonts, colors, headers, footers).
"""

__version__ = "0.1.0"

from .parser import MarkdownParser
from .generator import WordGenerator
from .branding import BrandingConfig

__all__ = ["MarkdownParser", "WordGenerator", "BrandingConfig"]
