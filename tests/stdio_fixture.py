"""Run the normal STDIO entry point with a synthetic Outlook boundary."""

import sys
from pathlib import Path
from unittest.mock import patch

from outlook_mcp.entrypoint import main


class SyntheticBridge:
    def __init__(self):
        self.log = Path(sys.argv[1])
        self.log.write_text("constructed\n", encoding="utf-8")

    def start(self):
        pass

    def stop(self):
        pass

    async def call(self, function, *args, **kwargs):
        with self.log.open("a", encoding="utf-8") as stream:
            print(function.__name__, file=stream)
        return {"user": "Synthetic User"}


if __name__ == "__main__":
    with patch("outlook_mcp.server.OutlookBridge", SyntheticBridge):
        main()
