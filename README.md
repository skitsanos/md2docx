# MD2DOCX

Convert Markdown content to Word documents with custom branding support.

## Features

- **Markdown to Word conversion** using mistune AST and python-docx
- **Custom branding**: fonts, colors, headers, footers, page margins
- **CLI tool** for command-line conversions
- **REST API** with FastAPI for programmatic access
- **Full Markdown support**: headings, lists, tables, code blocks, links, and more

## Installation

```bash
# Create virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install dependencies
pip install -e .
```

Or using Task:

```bash
task install
```

## Quick Start

### CLI Usage

```bash
# Basic conversion
md2docx input.md output.docx

# With branding
md2docx input.md output.docx --title "My Document" --author "Author Name"

# Using configuration file
md2docx input.md output.docx --config branding.json

# Generate sample branding config
md2docx --generate-config branding.json
```

### REST API

```bash
# Start server
PYTHONPATH=src uvicorn md2docx.api:app --reload --port 8000

# Convert via API
curl -X POST http://localhost:8000/convert \
  -H "Content-Type: application/json" \
  -d '{
    "markdown": "# Hello World\n\nThis is **bold** text.",
    "branding": {
      "title": "My Document",
      "author": "API User"
    }
  }' -o document.docx
```

### Remote Logo Allowlist

Remote logos/images are only fetched from hosts listed in the `MD2DOCX_ALLOWED_IMAGE_HOSTS` environment variable (comma-separated). Example:

```bash
export MD2DOCX_ALLOWED_IMAGE_HOSTS=cdn.example.com,assets.example.org
```

Requests with logo URLs outside this allowlist will be rejected.

## Supported Markdown Elements

- Headings (h1-h6)
- Bold, italic, strikethrough
- Ordered and unordered lists (with nesting)
- Tables with alignment (left, center, right)
- Code blocks and inline code
- Block quotes
- Hyperlinks
- Horizontal rules
- Special characters and emoji

## Branding Configuration

Customize the output with:

- Document metadata (title, author, company)
- Page size and margins
- Body font (name, size, color)
- Heading styles (font, size, color, spacing)
- Code block styling
- Link appearance
- Multi-zone headers and footers (left/center/right)
- Automatic page numbering

Example configuration:

```json
{
  "title": "Project Report",
  "author": "Development Team",
  "header": {
    "left_text": "Company Name",
    "text": "Report Title",
    "right_text": "2024"
  },
  "footer": {
    "left_text": "Confidential",
    "include_page_number": true,
    "page_number_position": "right"
  },
  "heading1": {
    "color": "#2F5496"
  }
}
```

## Documentation

- [CLI Documentation](docs/CLI.md)
- [Branding Configuration](docs/BRANDING.md)
- [REST API Reference](docs/API.md)

## Project Structure

```
src/md2docx/
├── __init__.py      # Package initialization
├── parser.py        # Markdown to AST parser (mistune)
├── generator.py     # AST to Word generator (python-docx)
├── branding.py      # Branding configuration
├── cli.py           # Command-line interface
├── api.py           # REST API (FastAPI)
└── __main__.py      # API server entry point
```

## Docker Support

```bash
# Build and run
task docker-build
task docker-run

# Or with docker-compose
task up
```

## Requirements

- Python 3.10+
- python-docx
- mistune
- FastAPI
- uvicorn
