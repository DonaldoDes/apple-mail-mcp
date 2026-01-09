#!/usr/bin/env python3
"""Entry point for apple-mail-mcp CLI."""

from apple_mail_mcp.apple_mail_mcp import mcp


def main():
    """Run the Apple Mail MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
