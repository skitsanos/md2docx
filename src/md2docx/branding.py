"""
Branding configuration for Word document generation.

Defines fonts, colors, headers, footers, and other styling options.
"""

from dataclasses import dataclass, field
from typing import Optional
from docx.shared import Pt, RGBColor, Inches


@dataclass
class FontConfig:
    """Font configuration for document elements."""
    name: str = "Calibri"
    size: Pt = field(default_factory=lambda: Pt(11))
    color: Optional[RGBColor] = None
    bold: bool = False
    italic: bool = False


@dataclass
class HeadingConfig:
    """Configuration for heading levels."""
    font_name: str = "Calibri"
    font_size: Pt = field(default_factory=lambda: Pt(16))
    color: Optional[RGBColor] = None
    bold: bool = True
    italic: bool = False
    space_before: Pt = field(default_factory=lambda: Pt(12))
    space_after: Pt = field(default_factory=lambda: Pt(6))


@dataclass
class HeaderFooterConfig:
    """Configuration for page headers and footers."""
    text: str = ""  # Center text (or single text for simple header)
    left_text: str = ""  # Left-aligned text
    right_text: str = ""  # Right-aligned text
    font_name: str = "Calibri"
    font_size: Pt = field(default_factory=lambda: Pt(9))
    color: Optional[RGBColor] = None
    include_page_number: bool = False
    page_number_position: str = "right"  # "left", "center", or "right"
    logo_path: str = ""  # Path or URL to logo image (PNG, JPEG, GIF, TIFF, BMP)
    logo_position: str = "left"  # "left", "center", or "right"
    logo_width: Optional[float] = None  # Logo width in inches (None = auto)


@dataclass
class BrandingConfig:
    """
    Complete branding configuration for Word document generation.

    This class holds all styling and branding options for the generated
    Word document, including fonts, colors, headers, footers, and margins.
    """

    # Document metadata
    title: str = ""
    author: str = ""
    company: str = ""

    # Page setup
    page_width: Inches = field(default_factory=lambda: Inches(8.5))
    page_height: Inches = field(default_factory=lambda: Inches(11))
    margin_top: Inches = field(default_factory=lambda: Inches(1))
    margin_bottom: Inches = field(default_factory=lambda: Inches(1))
    margin_left: Inches = field(default_factory=lambda: Inches(1))
    margin_right: Inches = field(default_factory=lambda: Inches(1))

    # Body text
    body_font: FontConfig = field(default_factory=FontConfig)

    # Headings (h1 through h6)
    heading1: HeadingConfig = field(default_factory=lambda: HeadingConfig(
        font_size=Pt(24), bold=True, space_before=Pt(18), space_after=Pt(12)
    ))
    heading2: HeadingConfig = field(default_factory=lambda: HeadingConfig(
        font_size=Pt(20), bold=True, space_before=Pt(16), space_after=Pt(10)
    ))
    heading3: HeadingConfig = field(default_factory=lambda: HeadingConfig(
        font_size=Pt(16), bold=True, space_before=Pt(14), space_after=Pt(8)
    ))
    heading4: HeadingConfig = field(default_factory=lambda: HeadingConfig(
        font_size=Pt(14), bold=True, space_before=Pt(12), space_after=Pt(6)
    ))
    heading5: HeadingConfig = field(default_factory=lambda: HeadingConfig(
        font_size=Pt(12), bold=True, space_before=Pt(10), space_after=Pt(4)
    ))
    heading6: HeadingConfig = field(default_factory=lambda: HeadingConfig(
        font_size=Pt(11), bold=True, italic=True, space_before=Pt(8), space_after=Pt(4)
    ))

    # Code blocks
    code_font: FontConfig = field(default_factory=lambda: FontConfig(
        name="Courier New", size=Pt(10)
    ))
    code_background_color: Optional[RGBColor] = field(default_factory=lambda: RGBColor(245, 245, 245))

    # Links
    link_color: RGBColor = field(default_factory=lambda: RGBColor(0, 0, 255))
    link_underline: bool = True

    # Lists
    list_indent: Inches = field(default_factory=lambda: Inches(0.5))

    # Headers and footers
    header: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)
    footer: HeaderFooterConfig = field(default_factory=lambda: HeaderFooterConfig(
        include_page_number=True
    ))

    def get_heading_config(self, level: int) -> HeadingConfig:
        """Get heading configuration for a specific level (1-6)."""
        heading_map = {
            1: self.heading1,
            2: self.heading2,
            3: self.heading3,
            4: self.heading4,
            5: self.heading5,
            6: self.heading6,
        }
        return heading_map.get(level, self.heading6)

    @classmethod
    def from_dict(cls, data: dict) -> "BrandingConfig":
        """Create a BrandingConfig from a dictionary."""
        config = cls()

        # Update simple fields
        for key in ["title", "author", "company"]:
            if key in data:
                setattr(config, key, data[key])

        # Update page setup
        if "page" in data:
            page = data["page"]
            if "width" in page:
                config.page_width = Inches(page["width"])
            if "height" in page:
                config.page_height = Inches(page["height"])
            if "margin_top" in page:
                config.margin_top = Inches(page["margin_top"])
            if "margin_bottom" in page:
                config.margin_bottom = Inches(page["margin_bottom"])
            if "margin_left" in page:
                config.margin_left = Inches(page["margin_left"])
            if "margin_right" in page:
                config.margin_right = Inches(page["margin_right"])

        # Update body font
        if "body_font" in data:
            config.body_font = cls._parse_font_config(data["body_font"])

        # Update headings
        for i in range(1, 7):
            key = f"heading{i}"
            if key in data:
                setattr(config, key, cls._parse_heading_config(data[key]))

        # Update code font
        if "code_font" in data:
            config.code_font = cls._parse_font_config(data["code_font"])

        if "code_background_color" in data:
            config.code_background_color = cls._parse_color(data["code_background_color"])

        # Update link settings
        if "link_color" in data:
            config.link_color = cls._parse_color(data["link_color"])
        if "link_underline" in data:
            config.link_underline = data["link_underline"]

        # Update header/footer
        if "header" in data:
            config.header = cls._parse_header_footer_config(data["header"])
        if "footer" in data:
            config.footer = cls._parse_header_footer_config(data["footer"])

        return config

    @staticmethod
    def _parse_color(color_data) -> RGBColor:
        """Parse color from various formats."""
        if isinstance(color_data, (list, tuple)) and len(color_data) == 3:
            return RGBColor(color_data[0], color_data[1], color_data[2])
        elif isinstance(color_data, str):
            # Handle hex color like "#FF0000" or "FF0000"
            hex_color = color_data.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            return RGBColor(r, g, b)
        return color_data

    @staticmethod
    def _parse_font_config(data: dict) -> FontConfig:
        """Parse font configuration from dictionary."""
        config = FontConfig()
        if "name" in data:
            config.name = data["name"]
        if "size" in data:
            config.size = Pt(data["size"])
        if "color" in data:
            config.color = BrandingConfig._parse_color(data["color"])
        if "bold" in data:
            config.bold = data["bold"]
        if "italic" in data:
            config.italic = data["italic"]
        return config

    @staticmethod
    def _parse_heading_config(data: dict) -> HeadingConfig:
        """Parse heading configuration from dictionary."""
        config = HeadingConfig()
        if "font_name" in data:
            config.font_name = data["font_name"]
        if "font_size" in data:
            config.font_size = Pt(data["font_size"])
        if "color" in data:
            config.color = BrandingConfig._parse_color(data["color"])
        if "bold" in data:
            config.bold = data["bold"]
        if "space_before" in data:
            config.space_before = Pt(data["space_before"])
        if "space_after" in data:
            config.space_after = Pt(data["space_after"])
        return config

    @staticmethod
    def _parse_header_footer_config(data: dict) -> HeaderFooterConfig:
        """Parse header/footer configuration from dictionary."""
        config = HeaderFooterConfig()
        if "text" in data:
            config.text = data["text"]
        if "left_text" in data:
            config.left_text = data["left_text"]
        if "right_text" in data:
            config.right_text = data["right_text"]
        if "font_name" in data:
            config.font_name = data["font_name"]
        if "font_size" in data:
            config.font_size = Pt(data["font_size"])
        if "color" in data:
            config.color = BrandingConfig._parse_color(data["color"])
        if "include_page_number" in data:
            config.include_page_number = data["include_page_number"]
        if "page_number_position" in data:
            config.page_number_position = data["page_number_position"]
        if "logo_path" in data:
            config.logo_path = data["logo_path"]
        if "logo_position" in data:
            config.logo_position = data["logo_position"]
        if "logo_width" in data:
            config.logo_width = data["logo_width"]
        return config
