# Convert with Logo in Header
# Test document generation with logo from URL

POST http://localhost:8000/convert
Content-Type: application/json
{
  "markdown": "# Company Report\n\n## Executive Summary\n\nThis document includes a company logo in the header.",
  "branding": {
    "title": "Company Report",
    "author": "Executive Team",
    "header": {
      "logo_path": "https://cdn.plufinder.com/assets/img/favicon/favicon-32x32.png",
      "logo_position": "left",
      "logo_width": 1.5,
      "text": "Annual Report",
      "right_text": "2024"
    },
    "footer": {
      "left_text": "© 2024 Company",
      "include_page_number": true,
      "page_number_position": "center"
    }
  },
  "filename": "report-with-logo.docx"
}

HTTP 200
[Asserts]
header "Content-Type" contains "wordprocessingml.document"
bytes count > 10000
