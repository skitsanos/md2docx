# Convert with All Markdown Features
# Comprehensive test with tables, lists, code, links, etc.

POST http://localhost:8000/convert
Content-Type: application/json
{
  "markdown": "# Comprehensive Test\n\n## Text Formatting\n\nThis is **bold**, *italic*, and ***bold italic*** text. Here's ~~strikethrough~~ and `inline code`.\n\n## Lists\n\n### Ordered\n\n1. First\n2. Second\n3. Third\n\n### Unordered\n\n- Alpha\n- Beta\n- Gamma\n\n## Tables\n\n| Name | Value | Status |\n|:-----|:-----:|-------:|\n| Item 1 | 100 | Active |\n| Item 2 | 200 | Pending |\n\n## Code Block\n\n```python\ndef hello():\n    print('Hello, World!')\n```\n\n## Quote\n\n> This is a block quote with important information.\n\n## Links\n\nVisit [Example](https://example.com) for more info.\n\n---\n\n## Conclusion\n\nEnd of document.",
  "branding": {
    "title": "Feature Test",
    "author": "QA Team",
    "heading1": {
      "color": "#FF5733"
    },
    "code_background_color": "#F5F5F5"
  },
  "filename": "comprehensive.docx"
}

HTTP 200
[Asserts]
header "Content-Type" contains "wordprocessingml.document"
bytes count > 15000
