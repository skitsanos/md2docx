# Hurl API Tests

This directory contains [Hurl](https://hurl.dev/) tests for the MD2DOCX REST API.

## Prerequisites

Install Hurl:

```bash
# macOS
brew install hurl

# Linux
curl --location --remote-name https://github.com/Orange-OpenSource/hurl/releases/latest/download/hurl_x.x.x_amd64.deb
sudo dpkg -i hurl_x.x.x_amd64.deb

# Windows
winget install hurl
```

## Running Tests

Make sure the API server is running:

```bash
# Start the server
task app:api -- --reload

# Or directly
uvicorn md2docx.api:app --reload --port 8000
```

Then run the tests:

```bash
# Run all tests
hurl tests/hurl/*.hurl

# Run specific test
hurl tests/hurl/01-health.hurl

# Run with verbose output
hurl --verbose tests/hurl/*.hurl

# Run and save outputs
hurl --test tests/hurl/*.hurl

# Run in parallel (if supported)
ls tests/hurl/*.hurl | xargs -P 4 -I {} hurl {}
```

## Test Files

| File | Description |
|------|-------------|
| `01-health.hurl` | Health check endpoint |
| `02-parse-simple.hurl` | Parse simple Markdown |
| `03-parse-complex.hurl` | Parse complex Markdown with multiple elements |
| `04-convert-simple.hurl` | Basic conversion without branding |
| `05-convert-with-branding.hurl` | Conversion with full branding config |
| `06-convert-with-logo.hurl` | Conversion with logo in header |
| `07-convert-all-features.hurl` | Comprehensive test with all Markdown features |
| `08-get-sample-branding.hurl` | Get sample branding configuration |
| `09-error-invalid-markdown.hurl` | Error handling test |
| `10-file-upload.hurl` | File upload conversion |
| `11-file-upload-with-branding.hurl` | File upload with branding |
| `12-invalid-branding-inline.hurl` | Invalid branding config in JSON body |
| `13-invalid-branding-upload.hurl` | Invalid branding JSON in multipart upload |

## Notes

- Use `--file-root .` when running the suite so upload tests can read files from `samples/`.
- Ensure `MD2DOCX_ALLOWED_IMAGE_HOSTS` includes any remote logo hosts used in tests.
- **Rate Limiting**: The API has rate limiting enabled (default: 30 requests/minute per IP). When running many tests in quick succession, you may hit the limit. Either:
  - Increase the limit: `export MD2DOCX_RATE_LIMIT=1000/minute`
  - Add delays between tests
  - Run tests sequentially rather than in parallel

## Test Coverage

These tests cover:

- ✅ Health check
- ✅ Markdown parsing (simple and complex)
- ✅ Document conversion (plain and branded)
- ✅ Logo support (URL and SVG)
- ✅ All Markdown features (tables, lists, code, links, etc.)
- ✅ Branding configuration
- ✅ File upload
- ✅ Error handling

## Continuous Integration

Add to your CI pipeline:

```yaml
# GitHub Actions example
- name: Run API Tests
  env:
    MD2DOCX_RATE_LIMIT: 1000/minute  # Increase for CI
    MD2DOCX_ALLOWED_IMAGE_HOSTS: cdn.plufinder.com
  run: |
    uvicorn md2docx.api:app --host 0.0.0.0 --port 8000 &
    sleep 5
    hurl --test tests/hurl/*.hurl
```

## Test Results

Hurl provides detailed output:
- ✅ Pass: Test assertions succeeded
- ❌ Fail: Test assertions failed with details
- Summary: Total tests, passed, failed, duration
