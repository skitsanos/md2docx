# Error Handling - Invalid Parse Request
# Test error response when markdown parameter is missing

POST http://localhost:8000/parse

HTTP 422
[Asserts]
jsonpath "$.detail" exists
