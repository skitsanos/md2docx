# File Upload with Branding
# Test file upload with JSON branding configuration

POST http://localhost:8000/convert/file
[MultipartFormData]
file: file,samples/comprehensive-test.md;
branding: {"title":"Uploaded Document","author":"Test User","header":{"text":"Header Text"},"footer":{"include_page_number":true}}

HTTP 200
[Asserts]
header "Content-Type" contains "wordprocessingml.document"
bytes count > 10000
