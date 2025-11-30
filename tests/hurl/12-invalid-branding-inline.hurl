# Invalid Branding Configuration (inline convert)
# Expect a 400 when branding values cannot be parsed

POST http://localhost:8000/convert
Content-Type: application/json
{
  "markdown": "# Bad Branding\n\nThis request uses an invalid color.",
  "branding": {
    "heading1": {
      "color": "#GGGGGG"
    }
  }
}

HTTP 400
[Asserts]
jsonpath "$.detail" == "Invalid branding configuration"
