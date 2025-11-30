# Convert Simple Markdown to Word
# Basic conversion without branding

POST http://localhost:8000/convert
Content-Type: application/json
{
  "markdown": "# Hello World\n\nThis is a **test** document.",
  "filename": "test.docx"
}

HTTP 200
[Asserts]
header "Content-Type" contains "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
header "Content-Disposition" contains "test.docx"
bytes count > 10000
