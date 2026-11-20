"""CLI entry point: ``outlook-mcp`` and ``python -m outlook_mcp``."""

from __future__ import annotations

import logging
import os
import sys

from outlook_mcp.config import ToolConfigError
from outlook_mcp.server import build_server


def main() -> None:
    # stdio servers must NEVER log to stdout — that's the JSON-RPC stream.
    logging.basicConfig(
        level=os.environ.get("OUTLOOK_MCP_LOG", "INFO"),
        stream=sys.stderr,
        format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    )
    log = logging.getLogger("outlook_mcp")
    # The MCP transport's DEBUG traces can include incoming tool arguments
    # (e.g. the subject supplied when creating a private appointment).
    if logging.getLogger("mcp").isEnabledFor(logging.DEBUG):
        logging.getLogger("mcp").setLevel(logging.INFO)

    try:
        mcp, bridge = build_server()
    except ToolConfigError as exc:
        log.error("Invalid Outlook MCP configuration: %s", exc)
        sys.exit(2)
    log.info("Starting outlook_mcp (stdio transport)")

    try:
        bridge.start()
    except Exception as exc:  # noqa: BLE001
        log.error("Failed to attach to Outlook: %s", exc)
        sys.exit(2)

    try:
        mcp.run()
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
