# Parse Complex Markdown
# Test parsing Markdown with multiple elements

POST http://localhost:8000/parse
[FormParams]
markdown: **Bold text** and *italic text*

HTTP 200
[Asserts]
jsonpath "$.node_count" >= 1
