"""
Main entry point for running the md2docx API server.

Usage:
    python -m md2docx [--host HOST] [--port PORT] [--reload]
"""

import argparse
import sys


def main():
    """Run the md2docx API server."""
    parser = argparse.ArgumentParser(
        description="Run the MD2DOCX API server",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--host",
        type=str,
        default="0.0.0.0",
        help="Host to bind the server to (default: 0.0.0.0)",
    )

    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port to listen on (default: 8000)",
    )

    parser.add_argument(
        "--reload",
        action="store_true",
        help="Enable auto-reload for development",
    )

    args = parser.parse_args()

    print(f"Starting MD2DOCX API server on {args.host}:{args.port}")
    if args.reload:
        print("Auto-reload enabled (development mode)")

    from .api import run_server
    run_server(host=args.host, port=args.port, reload=args.reload)


if __name__ == "__main__":
    sys.exit(main())
