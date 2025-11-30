# File Upload Conversion
# Test uploading a Markdown file for conversion
# Note: This test requires a sample file to exist

POST http://localhost:8000/convert/file
[MultipartFormData]
file: file,samples/comprehensive-test.md;

HTTP 200
[Asserts]
header "Content-Type" contains "wordprocessingml.document"
header "Content-Disposition" contains ".docx"
bytes count > 10000
