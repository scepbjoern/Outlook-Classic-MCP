"""Run the normal STDIO entry point with a synthetic Outlook boundary."""

import sys
from pathlib import Path
from types import SimpleNamespace
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
        if function.__name__ == "call_in_mailbox":
            kwargs.pop("retry_on_disconnect")
            accounts = [
                SimpleNamespace(SmtpAddress=f"{key}@example.test",
                                DeliveryStore=SimpleNamespace(StoreID=f"store-{key}"))
                for key in ("a", "b")
            ]
            namespace = SimpleNamespace(Accounts=SimpleNamespace(
                Count=len(accounts), Item=lambda index: accounts[index - 1],
            ))
            return function(None, namespace, *args, **kwargs)
        return {"user": "Synthetic User"}


if __name__ == "__main__":
    with patch("outlook_mcp.server.OutlookBridge", SyntheticBridge):
        main()
