# Branding Configuration

Complete reference for customizing Word document appearance.

## Configuration File Structure

```json
{
  "title": "Document Title",
  "author": "Author Name",
  "company": "Company Name",
  "page": { ... },
  "body_font": { ... },
  "heading1": { ... },
  "heading2": { ... },
  "heading3": { ... },
  "heading4": { ... },
  "heading5": { ... },
  "heading6": { ... },
  "code_font": { ... },
  "code_background_color": "#F5F5F5",
  "link_color": "#0563C1",
  "link_underline": true,
  "header": { ... },
  "footer": { ... }
}
```

## Document Metadata

```json
{
  "title": "My Document",
  "author": "Author Name",
  "company": "Company Name"
}
```

These values are embedded in the Word document properties.

## Page Setup

```json
{
  "page": {
    "width": 8.5,
    "height": 11,
    "margin_top": 1,
    "margin_bottom": 1,
    "margin_left": 1,
    "margin_right": 1
  }
}
```

All measurements are in **inches**.

### Common Page Sizes

| Size | Width | Height |
|------|-------|--------|
| Letter (US) | 8.5 | 11 |
| A4 | 8.27 | 11.69 |
| Legal | 8.5 | 14 |

## Body Font

```json
{
  "body_font": {
    "name": "Calibri",
    "size": 11,
    "color": "#000000"
  }
}
```

- `name`: Font family name (e.g., "Calibri", "Arial", "Times New Roman")
- `size`: Font size in points
- `color`: Hex color code (e.g., "#000000" for black)

## Headings

Each heading level (1-6) can be customized independently:

```json
{
  "heading1": {
    "font_name": "Calibri",
    "font_size": 24,
    "color": "#2F5496",
    "bold": true,
    "italic": false,
    "space_before": 18,
    "space_after": 12
  }
}
```

- `font_name`: Heading font family
- `font_size`: Size in points
- `color`: Hex color code
- `bold`: Boolean for bold text
- `italic`: Boolean for italic text
- `space_before`: Spacing before heading in points
- `space_after`: Spacing after heading in points

### Default Heading Sizes

| Level | Default Size | Description |
|-------|-------------|-------------|
| heading1 | 24pt | Main title |
| heading2 | 20pt | Section |
| heading3 | 16pt | Subsection |
| heading4 | 14pt | Sub-subsection |
| heading5 | 12pt | Minor heading |
| heading6 | 11pt | Smallest heading |

## Code Styling

```json
{
  "code_font": {
    "name": "Courier New",
    "size": 10
  },
  "code_background_color": "#F5F5F5"
}
```

- `code_font.name`: Monospace font for code blocks and inline code
- `code_font.size`: Font size in points
- `code_background_color`: Background shading for code blocks (hex color)

## Link Styling

```json
{
  "link_color": "#0563C1",
  "link_underline": true
}
```

- `link_color`: Color for hyperlinks
- `link_underline`: Boolean to underline links

### Remote Logo Security

Remote logos/images in branding (e.g., `header.logo_path`) have multiple security controls:

| Control | Environment Variable | Default |
|---------|---------------------|---------|
| Host allowlist | `MD2DOCX_ALLOWED_IMAGE_HOSTS` | (empty - blocks all) |
| Max image size | `MD2DOCX_MAX_IMAGE_SIZE` | 3 MB (3145728 bytes) |

**Host Allowlist**: Only images from approved hosts are fetched.
```bash
export MD2DOCX_ALLOWED_IMAGE_HOSTS=cdn.example.com,assets.example.org
```

**Size Limit**: Images larger than the configured maximum are rejected to prevent memory exhaustion.

**Redirect Blocking**: HTTP redirects are blocked to prevent SSRF attacks via redirect.

Requests using logo URLs outside the allowlist or exceeding the size limit will be rejected with an error.

## Headers and Footers

Headers and footers support **three zones** (left, center, right) using tab-separated content.

### Header Configuration

```json
{
  "header": {
    "left_text": "Company Name",
    "text": "Document Title",
    "right_text": "Date",
    "font_name": "Calibri",
    "font_size": 9,
    "color": "#808080"
  }
}
```

- `left_text`: Left-aligned text
- `text`: Center-aligned text
- `right_text`: Right-aligned text
- `font_name`: Header font family
- `font_size`: Font size in points
- `color`: Text color (hex)

### Footer Configuration

```json
{
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

- `include_page_number`: Boolean to add page numbers
- `page_number_position`: Where to place page number ("left", "center", or "right")

### Example: Corporate Header/Footer

```json
{
  "header": {
    "left_text": "ACME Corporation",
    "text": "Quarterly Report Q4 2024",
    "right_text": "Confidential"
  },
  "footer": {
    "left_text": "Copyright 2024 ACME Corp",
    "text": "",
    "right_text": "",
    "include_page_number": true,
    "page_number_position": "center"
  }
}
```

Result:
- Header: `ACME Corporation    Quarterly Report Q4 2024    Confidential`
- Footer: `Copyright 2024 ACME Corp    Page 1`

## Color Formats

Colors can be specified in multiple formats:

- **Hex string**: `"#FF0000"` or `"FF0000"` (red)
- **RGB array**: `[255, 0, 0]` (red)

## Complete Example

```json
{
  "title": "Technical Specification",
  "author": "Engineering Team",
  "company": "TechCorp Inc.",
  "page": {
    "width": 8.5,
    "height": 11,
    "margin_top": 1,
    "margin_bottom": 1,
    "margin_left": 1.25,
    "margin_right": 1.25
  },
  "body_font": {
    "name": "Georgia",
    "size": 11,
    "color": "#333333"
  },
  "heading1": {
    "font_name": "Helvetica",
    "font_size": 28,
    "color": "#1E3A5F",
    "bold": true,
    "space_before": 24,
    "space_after": 16
  },
  "heading2": {
    "font_name": "Helvetica",
    "font_size": 22,
    "color": "#2C5282",
    "bold": true,
    "space_before": 20,
    "space_after": 12
  },
  "code_font": {
    "name": "Consolas",
    "size": 9
  },
  "code_background_color": "#F0F4F8",
  "link_color": "#3182CE",
  "link_underline": true,
  "header": {
    "left_text": "TechCorp Inc.",
    "text": "Technical Specification",
    "right_text": "v1.0",
    "font_name": "Helvetica",
    "font_size": 8,
    "color": "#666666"
  },
  "footer": {
    "left_text": "CONFIDENTIAL",
    "text": "",
    "right_text": "",
    "font_name": "Helvetica",
    "font_size": 8,
    "color": "#999999",
    "include_page_number": true,
    "page_number_position": "center"
  }
}
```
