"""Optional interactive read check on a disposable classic Outlook profile.

Run manually with uv run python scripts/live_reconnect_test.py.
See docs/outlook-readiness-smoke.md for setup and the full procedure.
"""

import asyncio

from outlook_mcp.bridge import OutlookBridge
from outlook_mcp.client.account import whoami
from outlook_mcp.config import load_tool_policy


async def main() -> None:
    policy = load_tool_policy()
    if not policy.allows("account", "outlook_whoami"):
        raise SystemExit("Enable outlook_whoami in the synthetic test configuration.")
    input("Confirm this Windows session uses a disposable test profile. Press Enter to start: ")
    bridge = OutlookBridge()
    try:
        bridge.start()
        await bridge.call(whoami, policy.accounts)
        input("Initial read succeeded. Close test Outlook normally, then press Enter: ")
        await bridge.call(whoami, policy.accounts)
        print("Second identity read succeeded. Confirm the intended profile in Outlook.")
    finally:
        bridge.stop()


if __name__ == "__main__":
    asyncio.run(main())
