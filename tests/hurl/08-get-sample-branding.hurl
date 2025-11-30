# Get Sample Branding Configuration
# Test retrieving the sample branding config

GET http://localhost:8000/branding/sample

HTTP 200
[Asserts]
jsonpath "$.title" exists
jsonpath "$.author" exists
jsonpath "$.page" exists
jsonpath "$.page.width" == 8.5
jsonpath "$.page.height" == 11
jsonpath "$.body_font" exists
jsonpath "$.body_font.name" == "Calibri"
jsonpath "$.heading1" exists
jsonpath "$.heading1.font_size" == 24
jsonpath "$.header" exists
jsonpath "$.footer" exists
jsonpath "$.footer.include_page_number" == true
