# MD2DOCX CLI Documentation

Command-line tool for converting Markdown files to Word documents with custom branding support.

## Installation

```bash
# Create virtual environment
python3 -m venv .venv
source .venv/bin/activate  # Linux/Mac
# or .venv\Scripts\activate on Windows

# Install dependencies
pip install -e .

# Ensure python-docx deps are met; use Python 3.10+
```

## Basic Usage

```bash
# Convert Markdown to Word
md2docx input.md output.docx

# With verbose output
md2docx input.md output.docx --verbose
```

## Branding Options

### Inline Options

```bash
md2docx input.md output.docx \
  --title "My Document" \
  --author "John Doe" \
  --company "ACME Corp" \
  --header "Header Text" \
  --footer "Footer Text" \
  --font "Arial" \
  --font-size 12
```

### Using Configuration File

```bash
md2docx input.md output.docx --config branding.json
```

### Generate Sample Configuration

```bash
md2docx --generate-config branding.json
```

## CLI Options

| Option | Description |
|--------|-------------|
| `input` | Input Markdown file path |
| `output` | Output Word document path (.docx) |
| `--config`, `-c` | Path to JSON branding configuration file |
| `--title` | Document title |
| `--author` | Document author |
| `--company` | Company name |
| `--header` | Header text for all pages (center zone) |
| `--footer` | Footer text for all pages (left zone) |
| `--no-page-numbers` | Disable page numbers in footer |
| `--font` | Body font name (e.g., 'Arial', 'Times New Roman') |
| `--font-size` | Body font size in points |
| `--ast` | Parse and display the AST only (no Word generation) |
| `--generate-config` | Generate a sample branding configuration file |
| `--verbose`, `-v` | Enable verbose output |

## AST Inspection

View the Abstract Syntax Tree without generating a Word document:

```bash
md2docx input.md --ast
```

## Supported Markdown Features

- **Headings** (h1-h6) using Word's built-in Heading styles
- **Text formatting**: bold, italic, strikethrough, inline code
- **Lists**: ordered (numbered) and unordered (bullet) with nesting
- **Tables** with column alignment (left, center, right)
- **Code blocks** with syntax highlighting support
- **Block quotes** using Word's Quote style
- **Links** as clickable hyperlinks
- **Horizontal rules**
- **Special characters and emoji**

## Examples

### Basic Conversion

```bash
md2docx report.md report.docx
```

### Corporate Document with Branding

```bash
md2docx quarterly-report.md report.docx \
  --config corporate-branding.json \
  --title "Q4 2024 Report" \
  --author "Finance Team"
```

### Debug AST Structure

```bash
md2docx complex-document.md --ast | head -50
```
