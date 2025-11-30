"""
FastAPI REST API for Markdown to Word document conversion.

This module provides HTTP endpoints for converting Markdown content
to Word documents with custom branding support.
"""

import asyncio
import json
import logging
import os
import re
import uuid
from io import BytesIO
from typing import Optional, Any

from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded
from starlette.middleware.base import BaseHTTPMiddleware

from . import __version__
from .parser import MarkdownParser
from .generator import WordGenerator
from .branding import BrandingConfig

# Configure module logger
logger = logging.getLogger(__name__)

# Maximum accepted markdown size (in bytes) for both raw and uploaded content
MAX_MARKDOWN_SIZE = 5 * 1024 * 1024  # 5 MB

# Request timeout in seconds
REQUEST_TIMEOUT = int(os.getenv("MD2DOCX_REQUEST_TIMEOUT", "60"))

# Rate limiting configuration (requests per minute)
RATE_LIMIT = os.getenv("MD2DOCX_RATE_LIMIT", "30/minute")

# CORS configuration
CORS_ORIGINS = [
    origin.strip()
    for origin in os.getenv("MD2DOCX_CORS_ORIGINS", "").split(",")
    if origin.strip()
]

# Initialize rate limiter
limiter = Limiter(key_func=get_remote_address)


class RequestIDMiddleware(BaseHTTPMiddleware):
    """Middleware to add unique request IDs for tracing."""

    async def dispatch(self, request: Request, call_next):
        request_id = str(uuid.uuid4())
        request.state.request_id = request_id
        logger.debug("Request %s: %s %s", request_id, request.method, request.url.path)
        response = await call_next(request)
        response.headers["X-Request-ID"] = request_id
        return response


class TimeoutMiddleware(BaseHTTPMiddleware):
    """Middleware to enforce request timeouts."""

    async def dispatch(self, request: Request, call_next):
        try:
            return await asyncio.wait_for(
                call_next(request),
                timeout=REQUEST_TIMEOUT
            )
        except asyncio.TimeoutError:
            request_id = getattr(request.state, "request_id", "unknown")
            logger.warning("Request %s timed out after %ds", request_id, REQUEST_TIMEOUT)
            return JSONResponse(
                status_code=504,
                content={"detail": "Request timeout"}
            )


# Create FastAPI app
app = FastAPI(
    title="MD2DOCX API",
    description="Convert Markdown content to Word documents with custom branding",
    version=__version__,
    docs_url="/docs",
    redoc_url="/redoc",
)

# Register rate limiter
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# Add CORS middleware
if CORS_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=CORS_ORIGINS,
        allow_credentials=True,
        allow_methods=["GET", "POST"],
        allow_headers=["Content-Type", "Authorization"],
    )
else:
    # Default restrictive CORS for development
    logger.info("No CORS origins configured. Set MD2DOCX_CORS_ORIGINS to enable CORS.")

# Add custom middleware (order matters - first added = last executed)
app.add_middleware(TimeoutMiddleware)
app.add_middleware(RequestIDMiddleware)


# Pydantic models for API requests/responses
class ConvertRequest(BaseModel):
    """Request model for converting Markdown content."""
    markdown: str = Field(..., description="Markdown content to convert")
    branding: Optional[dict[str, Any]] = Field(
        default=None,
        description="Branding configuration (same structure as JSON config file)"
    )
    filename: Optional[str] = Field(
        default="document.docx",
        description="Output filename"
    )


class ParseResponse(BaseModel):
    """Response model for AST parsing."""
    ast: list[dict]
    node_count: int


class HealthResponse(BaseModel):
    """Health check response."""
    status: str
    version: str
    service: str


def _sanitize_filename(name: Optional[str], default: str = "document.docx") -> str:
    """
    Sanitize filenames for HTTP headers and filesystem safety.

    Removes path separators and control characters; falls back to default on empty result.
    """
    if not name:
        return default
    cleaned = re.sub(r"[\\/\r\n\t]+", "_", name).strip()
    return cleaned or default


@app.get("/health", response_model=HealthResponse, tags=["System"])
async def health_check():
    """
    Health check endpoint.

    Returns the service status and version information.
    """
    return HealthResponse(
        status="healthy",
        version=__version__,
        service="md2docx-api"
    )


@app.post("/parse", response_model=ParseResponse, tags=["Parsing"])
@limiter.limit(RATE_LIMIT)
async def parse_markdown(request: Request, markdown: str = Form(..., description="Markdown content to parse")):
    """
    Parse Markdown content into an Abstract Syntax Tree (AST).

    This endpoint parses the provided Markdown text and returns
    the AST representation using mistune.
    """
    try:
        if len(markdown.encode("utf-8")) > MAX_MARKDOWN_SIZE:
            raise HTTPException(status_code=413, detail="Markdown payload is too large")

        parser = MarkdownParser()
        ast = parser.parse(markdown)
        return ParseResponse(ast=ast, node_count=len(ast))
    except HTTPException:
        raise
    except Exception:
        logger.exception("Failed to parse markdown content")
        raise HTTPException(status_code=400, detail="Failed to parse Markdown content")


@app.post("/convert", tags=["Conversion"])
@limiter.limit(RATE_LIMIT)
async def convert_markdown(request: Request, convert_request: ConvertRequest):
    """
    Convert Markdown content to a Word document.

    This endpoint takes Markdown content and optional branding configuration,
    generates a Word document, and returns it as a downloadable file.

    The branding configuration uses the same structure as the JSON config file:

    ```json
    {
      "title": "My Document",
      "author": "Author Name",
      "company": "Company Name",
      "page": {
        "width": 8.5,
        "height": 11,
        "margin_top": 1,
        "margin_bottom": 1
      },
      "body_font": {
        "name": "Calibri",
        "size": 11,
        "color": "#000000"
      },
      "heading1": {
        "font_name": "Calibri",
        "font_size": 24,
        "color": "#2F5496",
        "bold": true
      },
      "header": {
        "left_text": "Company",
        "text": "Title",
        "right_text": "Date"
      },
      "footer": {
        "left_text": "Confidential",
        "include_page_number": true,
        "page_number_position": "right"
      }
    }
    ```
    """
    try:
        # Create branding config from request
        if convert_request.branding:
            branding = BrandingConfig.from_dict(convert_request.branding)
        else:
            branding = BrandingConfig()

        if len(convert_request.markdown.encode("utf-8")) > MAX_MARKDOWN_SIZE:
            raise HTTPException(status_code=413, detail="Markdown payload is too large")

        # Generate document
        generator = WordGenerator(branding)
        document = generator.generate_from_markdown(convert_request.markdown)

        # Convert to bytes
        doc_bytes = generator.to_bytes(document)

        safe_filename = _sanitize_filename(convert_request.filename)

        # Return as streaming response
        return StreamingResponse(
            BytesIO(doc_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{safe_filename}"'
            },
        )

    except HTTPException:
        raise
    except ValueError as e:
        logger.warning("Branding validation failed: %s", e)
        raise HTTPException(status_code=400, detail="Invalid branding configuration")
    except Exception:
        logger.exception("Failed to generate document from inline markdown")
        raise HTTPException(status_code=500, detail="Failed to generate document")


@app.post("/convert/file", tags=["Conversion"])
@limiter.limit(RATE_LIMIT)
async def convert_markdown_file(
    request: Request,
    file: UploadFile = File(..., description="Markdown file to convert"),
    branding: Optional[str] = Form(default=None, description="Branding configuration as JSON string"),
):
    """
    Convert an uploaded Markdown file to a Word document.

    Upload a .md or .txt file and receive a Word document with
    optional branding configuration applied.

    The branding parameter should be a JSON string with the same structure
    as the branding configuration file.
    """
    # Validate file type
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided")

    allowed_extensions = [".md", ".markdown", ".txt"]
    if not any(file.filename.lower().endswith(ext) for ext in allowed_extensions):
        raise HTTPException(
            status_code=400,
            detail=f"Invalid file type. Allowed: {', '.join(allowed_extensions)}",
        )

    try:
        # Read file content
        content = await file.read(MAX_MARKDOWN_SIZE + 1)
        if len(content) > MAX_MARKDOWN_SIZE:
            raise HTTPException(status_code=413, detail="Uploaded file is too large")
        markdown_content = content.decode("utf-8")

        # Parse branding config from JSON string
        if branding:
            branding_dict = json.loads(branding)
            branding_config = BrandingConfig.from_dict(branding_dict)
        else:
            branding_config = BrandingConfig()

        # Generate document
        generator = WordGenerator(branding_config)
        document = generator.generate_from_markdown(markdown_content)

        # Convert to bytes
        doc_bytes = generator.to_bytes(document)

        # Generate output filename
        output_filename = _sanitize_filename(file.filename.rsplit(".", 1)[0] + ".docx")

        # Return as streaming response
        return StreamingResponse(
            BytesIO(doc_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"'
            },
        )

    except HTTPException:
        raise
    except UnicodeDecodeError:
        raise HTTPException(status_code=400, detail="File must be UTF-8 encoded text")
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Branding must be valid JSON")
    except ValueError as e:
        logger.warning("Branding validation failed: %s", e)
        raise HTTPException(status_code=400, detail="Invalid branding configuration")
    except Exception:
        logger.exception("Failed to generate document from uploaded markdown")
        raise HTTPException(status_code=500, detail="Failed to generate document")


@app.get("/branding/sample", tags=["Configuration"])
async def get_sample_branding():
    """
    Get a sample branding configuration.

    Returns a complete example of branding configuration that can be
    used as a template for customization. This matches the structure
    used in JSON configuration files.
    """
    sample = {
        "title": "My Document",
        "author": "Author Name",
        "company": "Company Name",
        "page": {
            "width": 8.5,
            "height": 11,
            "margin_top": 1,
            "margin_bottom": 1,
            "margin_left": 1,
            "margin_right": 1,
        },
        "body_font": {
            "name": "Calibri",
            "size": 11,
            "color": "#000000",
        },
        "heading1": {
            "font_name": "Calibri",
            "font_size": 24,
            "color": "#2F5496",
            "bold": True,
            "space_before": 18,
            "space_after": 12,
        },
        "heading2": {
            "font_name": "Calibri",
            "font_size": 20,
            "color": "#2F5496",
            "bold": True,
            "space_before": 16,
            "space_after": 10,
        },
        "heading3": {
            "font_name": "Calibri",
            "font_size": 16,
            "color": "#2F5496",
            "bold": True,
            "space_before": 14,
            "space_after": 8,
        },
        "code_font": {
            "name": "Courier New",
            "size": 10,
        },
        "code_background_color": "#F5F5F5",
        "link_color": "#0563C1",
        "link_underline": True,
        "header": {
            "left_text": "Company Name",
            "text": "Document Title",
            "right_text": "",
            "font_name": "Calibri",
            "font_size": 9,
            "color": "#808080",
        },
        "footer": {
            "left_text": "Confidential",
            "text": "",
            "right_text": "",
            "font_name": "Calibri",
            "font_size": 9,
            "color": "#808080",
            "include_page_number": True,
            "page_number_position": "right",
        },
    }
    return JSONResponse(content=sample)


# Main entry point for running the API
def run_server(host: str = "0.0.0.0", port: int = 8000, reload: bool = False):
    """
    Run the FastAPI server.

    Args:
        host: Host to bind to.
        port: Port to listen on.
        reload: Enable auto-reload for development.
    """
    import uvicorn
    uvicorn.run("md2docx.api:app", host=host, port=port, reload=reload)


if __name__ == "__main__":
    run_server(reload=True)
