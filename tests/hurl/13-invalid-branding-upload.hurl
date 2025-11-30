# Invalid Branding JSON in file upload
# Expect a 400 when branding is not valid JSON

POST http://localhost:8000/convert/file
[MultipartFormData]
file: file,samples/comprehensive-test.md;
branding: {not json}

HTTP 400
[Asserts]
jsonpath "$.detail" == "Branding must be valid JSON"
