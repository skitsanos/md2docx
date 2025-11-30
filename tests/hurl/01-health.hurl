# Health Check Endpoint
# Simple test to verify the API is running

GET http://localhost:8000/health

HTTP 200
[Asserts]
jsonpath "$.status" == "healthy"
jsonpath "$.service" == "md2docx-api"
