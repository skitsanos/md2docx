# Parse Simple Markdown
# Test parsing basic Markdown to AST

POST http://localhost:8000/parse
[FormParams]
markdown: Hello World

HTTP 200
[Asserts]
jsonpath "$.node_count" >= 1
