# Convert with Branding Configuration
# Test document generation with custom branding

POST http://localhost:8000/convert
Content-Type: application/json
{
  "markdown": "# Report\n\n## Summary\n\nThis is the summary with **important** information.\n\n### Details\n\n- Point 1\n- Point 2\n- Point 3",
  "branding": {
    "title": "Test Report",
    "author": "API Test",
    "company": "Test Company",
    "heading1": {
      "color": "#2F5496"
    },
    "header": {
      "left_text": "Company Name",
      "text": "Document Title",
      "right_text": "2024"
    },
    "footer": {
      "left_text": "Confidential",
      "include_page_number": true,
      "page_number_position": "right"
    }
  },
  "filename": "branded-report.docx"
}

HTTP 200
[Asserts]
header "Content-Type" contains "wordprocessingml.document"
header "Content-Disposition" contains "branded-report.docx"
bytes count > 10000
