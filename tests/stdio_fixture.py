"""Run the normal STDIO entry point with a synthetic Outlook boundary."""

import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from outlook_mcp.entrypoint import main
from test_mailbox import make_profile


class SyntheticBridge:
    def __init__(self):
        self.log = Path(sys.argv[1])
        self.log.write_text("constructed\n", encoding="utf-8")
        self.namespace = make_profile()
        for store in self.namespace.Stores:
            event = store.GetDefaultFolder.return_value.Items.Add.return_value
            event.EntryID = "synthetic-shared-id"
            event.Subject = event.Body = "synthetic-private-content"
            event.Sensitivity = 2
            event.Start, event.End = datetime(2026, 11, 23, 10), datetime(2026, 11, 23, 11)

    def start(self):
        pass

    def stop(self):
        pass

    async def call(self, function, *args, **kwargs):
        with self.log.open("a", encoding="utf-8") as stream:
            print(function.__name__, file=stream)
        kwargs.pop("retry_on_disconnect", None)
        return function(None, self.namespace, *args, **kwargs)


if __name__ == "__main__":
    with patch("outlook_mcp.server.OutlookBridge", SyntheticBridge):
        main()
