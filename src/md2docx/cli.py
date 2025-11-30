#!/usr/bin/env python3
"""
Command-line interface for Markdown to Word document conversion.

This module provides a CLI tool to convert Markdown files to Word documents
with optional custom branding configuration.
"""

import argparse
import json
import sys
from pathlib import Path

from .parser import MarkdownParser
from .generator import WordGenerator
from .branding import BrandingConfig


def main():
    """Main entry point for the CLI tool."""
    parser = argparse.ArgumentParser(
        description="Convert Markdown files to Word documents with custom branding.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic conversion
  md2docx input.md output.docx

  # With custom branding config
  md2docx input.md output.docx --config branding.json

  # With inline branding options
  md2docx input.md output.docx --title "My Document" --author "John Doe"

  # Parse and show AST only
  md2docx input.md --ast

  # Generate sample branding config
  md2docx --generate-config branding.json
        """,
    )

    parser.add_argument(
        "input",
        nargs="?",
        help="Input Markdown file path",
    )

    parser.add_argument(
        "output",
        nargs="?",
        help="Output Word document path (.docx)",
    )

    parser.add_argument(
        "--config",
        "-c",
        type=str,
        help="Path to JSON branding configuration file",
    )

    parser.add_argument(
        "--title",
        type=str,
        help="Document title",
    )

    parser.add_argument(
        "--author",
        type=str,
        help="Document author",
    )

    parser.add_argument(
        "--company",
        type=str,
        help="Company name",
    )

    parser.add_argument(
        "--header",
        type=str,
        help="Header text for all pages",
    )

    parser.add_argument(
        "--footer",
        type=str,
        help="Footer text for all pages",
    )

    parser.add_argument(
        "--no-page-numbers",
        action="store_true",
        help="Disable page numbers in footer",
    )

    parser.add_argument(
        "--font",
        type=str,
        help="Body font name (e.g., 'Arial', 'Times New Roman')",
    )

    parser.add_argument(
        "--font-size",
        type=int,
        help="Body font size in points",
    )

    parser.add_argument(
        "--ast",
        action="store_true",
        help="Parse and display the AST only (no Word generation)",
    )

    parser.add_argument(
        "--generate-config",
        type=str,
        metavar="FILE",
        help="Generate a sample branding configuration file",
    )

    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose output",
    )

    args = parser.parse_args()

    # Handle config generation
    if args.generate_config:
        generate_sample_config(args.generate_config)
        print(f"Sample configuration file generated: {args.generate_config}")
        return 0

    # Validate input arguments
    if not args.input:
        parser.error("Input file is required (unless using --generate-config)")

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        return 1

    if not input_path.suffix.lower() in (".md", ".markdown", ".txt"):
        print(
            f"Warning: Input file may not be Markdown: {input_path.suffix}",
            file=sys.stderr,
        )

    # Parse the Markdown file
    if args.verbose:
        print(f"Parsing Markdown file: {input_path}")

    try:
        md_parser = MarkdownParser()
        ast = md_parser.parse_file(str(input_path))
    except Exception as e:
        print(f"Error parsing Markdown: {e}", file=sys.stderr)
        return 1

    # If AST-only mode, display and exit
    if args.ast:
        print("Abstract Syntax Tree:")
        print("=" * 50)
        print(MarkdownParser.pretty_print_ast(ast))
        return 0

    # Validate output argument
    if not args.output:
        parser.error("Output file is required (unless using --ast)")

    output_path = Path(args.output)
    if not output_path.suffix.lower() == ".docx":
        print(
            f"Warning: Output file should have .docx extension, got: {output_path.suffix}",
            file=sys.stderr,
        )

    # Load branding configuration
    branding = load_branding_config(args)

    if args.verbose:
        print(f"Using branding configuration:")
        print(f"  Title: {branding.title or '(not set)'}")
        print(f"  Author: {branding.author or '(not set)'}")
        print(f"  Company: {branding.company or '(not set)'}")
        print(f"  Body Font: {branding.body_font.name} {branding.body_font.size}")

    # Generate Word document
    if args.verbose:
        print(f"Generating Word document...")

    try:
        generator = WordGenerator(branding)
        document = generator.generate(ast)
        generator.save(document, output_path)
    except Exception as e:
        print(f"Error generating Word document: {e}", file=sys.stderr)
        return 1

    if args.verbose:
        print(f"Word document saved: {output_path}")
    else:
        print(f"Created: {output_path}")

    return 0


def load_branding_config(args) -> BrandingConfig:
    """Load branding configuration from file and/or command-line arguments."""
    # Start with default config or load from file
    if args.config:
        config_path = Path(args.config)
        if not config_path.exists():
            print(
                f"Warning: Config file not found: {config_path}, using defaults",
                file=sys.stderr,
            )
            branding = BrandingConfig()
        else:
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config_data = json.load(f)
                branding = BrandingConfig.from_dict(config_data)
            except json.JSONDecodeError as e:
                print(
                    f"Warning: Invalid JSON in config file: {e}, using defaults",
                    file=sys.stderr,
                )
                branding = BrandingConfig()
    else:
        branding = BrandingConfig()

    # Override with command-line arguments
    if args.title:
        branding.title = args.title
    if args.author:
        branding.author = args.author
    if args.company:
        branding.company = args.company
    if args.header:
        branding.header.text = args.header
    if args.footer:
        branding.footer.text = args.footer
    if args.no_page_numbers:
        branding.footer.include_page_number = False
    if args.font:
        branding.body_font.name = args.font
    if args.font_size:
        from docx.shared import Pt

        branding.body_font.size = Pt(args.font_size)

    return branding


def generate_sample_config(output_path: str) -> None:
    """Generate a sample branding configuration JSON file."""
    sample_config = {
        "title": "My Document",
        "author": "Author Name",
        "company": "Company Name",
        "page": {
            "width": 8.5,
            "height": 11,
            "margin_top": 1,
            "margin_bottom": 1,
            "margin_left": 1,
            "margin_right": 1,
        },
        "body_font": {
            "name": "Calibri",
            "size": 11,
            "color": "#000000",
        },
        "heading1": {
            "font_name": "Calibri",
            "font_size": 24,
            "color": "#2F5496",
            "bold": True,
            "space_before": 18,
            "space_after": 12,
        },
        "heading2": {
            "font_name": "Calibri",
            "font_size": 20,
            "color": "#2F5496",
            "bold": True,
            "space_before": 16,
            "space_after": 10,
        },
        "heading3": {
            "font_name": "Calibri",
            "font_size": 16,
            "color": "#2F5496",
            "bold": True,
            "space_before": 14,
            "space_after": 8,
        },
        "code_font": {
            "name": "Courier New",
            "size": 10,
        },
        "code_background_color": "#F5F5F5",
        "link_color": "#0563C1",
        "link_underline": True,
        "header": {
            "left_text": "Company Name",
            "text": "Document Title",
            "right_text": "",
            "font_name": "Calibri",
            "font_size": 9,
            "color": "#808080",
        },
        "footer": {
            "left_text": "Confidential",
            "text": "",
            "right_text": "",
            "font_name": "Calibri",
            "font_size": 9,
            "color": "#808080",
            "include_page_number": True,
            "page_number_position": "right",
        },
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(sample_config, f, indent=2)


if __name__ == "__main__":
    sys.exit(main())
