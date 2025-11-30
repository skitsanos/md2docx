# MD2DOCX REST API Documentation

REST API for converting Markdown content to Word documents with custom branding support.

## Running the API

```bash
# Development mode with auto-reload
source .venv/bin/activate
PYTHONPATH=src uvicorn md2docx.api:app --reload --port 8000

# Or using task
task app:api -- --reload --port 8000
```

## API Endpoints

### Health Check

```
GET /health
```

Returns service status and version information.

**Response:**
```json
{
  "status": "healthy",
  "version": "0.1.0",
  "service": "md2docx-api"
}
```

### Parse Markdown to AST

```
POST /parse
Content-Type: multipart/form-data
```

Parse Markdown content into Abstract Syntax Tree.

**Form Fields:**
- `markdown` (required): Markdown content to parse

**Response:**
```json
{
  "ast": [...],
  "node_count": 10
}
```

**Example:**
```bash
curl -X POST http://localhost:8000/parse \
  -F "markdown=# Hello\n\nWorld"
```

### Convert Markdown to Word

```
POST /convert
Content-Type: application/json
```

Convert Markdown content to Word document with optional branding.

**Request Body:**
```json
{
  "markdown": "# Title\n\nContent here...",
  "branding": {
    "title": "My Document",
    "author": "Author Name",
    "header": {
      "left_text": "Company",
      "text": "Document Title"
    },
    "footer": {
      "include_page_number": true
    }
  },
  "filename": "output.docx"
}
```

**Parameters:**
- `markdown` (required): Markdown content to convert
- `branding` (optional): Branding configuration object (same structure as JSON config file)
- `filename` (optional): Output filename (default: "document.docx")

**Response:**
Binary Word document file (`.docx`)

**Example:**
```bash
curl -X POST http://localhost:8000/convert \
  -H "Content-Type: application/json" \
  -d '{
    "markdown": "# Report\n\n## Summary\n\nThis is the summary.",
    "branding": {
      "title": "Q4 Report",
      "author": "Finance Team",
      "header": {
        "left_text": "ACME Corp",
        "text": "Quarterly Report",
        "right_text": "Confidential"
      },
      "footer": {
        "left_text": "Internal Use Only",
        "include_page_number": true,
        "page_number_position": "center"
      },
      "heading1": {
        "color": "#1E3A5F"
      }
    },
    "filename": "report.docx"
  }' -o report.docx
```

### Convert Uploaded File

```
POST /convert/file
Content-Type: multipart/form-data
```

Upload a Markdown file and convert it to Word document.

**Form Fields:**
- `file` (required): Markdown file (.md, .markdown, .txt)
- `branding` (optional): JSON string with branding configuration

**Example:**
```bash
# Without branding
curl -X POST http://localhost:8000/convert/file \
  -F "file=@document.md" \
  -o document.docx

# With branding
curl -X POST http://localhost:8000/convert/file \
  -F "file=@document.md" \
  -F 'branding={"title":"My Doc","author":"Me"}' \
  -o document.docx
```

### Get Sample Branding Configuration

```
GET /branding/sample
```

Returns a complete sample branding configuration.

**Response:**
```json
{
  "title": "My Document",
  "author": "Author Name",
  "company": "Company Name",
  "page": {
    "width": 8.5,
    "height": 11,
    "margin_top": 1,
    "margin_bottom": 1,
    "margin_left": 1,
    "margin_right": 1
  },
  "body_font": {
    "name": "Calibri",
    "size": 11,
    "color": "#000000"
  },
  "heading1": {
    "font_name": "Calibri",
    "font_size": 24,
    "color": "#2F5496",
    "bold": true,
    "space_before": 18,
    "space_after": 12
  },
  "header": {
    "left_text": "Company Name",
    "text": "Document Title",
    "right_text": "",
    "font_name": "Calibri",
    "font_size": 9,
    "color": "#808080"
  },
  "footer": {
    "left_text": "Confidential",
    "text": "",
    "right_text": "",
    "font_name": "Calibri",
    "font_size": 9,
    "color": "#808080",
    "include_page_number": true,
    "page_number_position": "right"
  }
}
```

## Interactive Documentation

When the API is running, access:
- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc

## Error Handling

API returns standard HTTP status codes:

| Code | Description |
|------|-------------|
| 200 | Success |
| 400 | Bad Request (invalid input) |
| 413 | Payload Too Large |
| 500 | Internal Server Error |

Error response format:
```json
{
  "detail": "Error message here"
}
```

## Branding Configuration

The `branding` object uses the same structure as the CLI JSON configuration file. See [BRANDING.md](BRANDING.md) for complete reference.

## Remote Logos / Images

Remote images are only fetched from hosts in the allowlist defined by the `MD2DOCX_ALLOWED_IMAGE_HOSTS` environment variable (comma-separated). Requests using logo URLs outside this list will be rejected.

Example:

```bash
export MD2DOCX_ALLOWED_IMAGE_HOSTS=cdn.example.com,assets.example.org
```

### Minimal Branding

```json
{
  "title": "Document Title",
  "author": "Author Name"
}
```

### Header/Footer Only

```json
{
  "header": {
    "text": "Company Confidential"
  },
  "footer": {
    "include_page_number": true
  }
}
```

### Custom Heading Colors

```json
{
  "heading1": {
    "color": "#FF5733"
  },
  "heading2": {
    "color": "#C70039"
  }
}
```

## Docker Deployment

```bash
# Build image
task docker-build

# Run container
task docker-run

# Or with docker-compose
task up
```

API will be available at http://localhost:8000
