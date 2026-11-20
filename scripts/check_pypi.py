"""Permit upload only after PyPI explicitly reports the requested version missing."""

import json
import os
import re
from urllib.error import HTTPError
from urllib.request import urlopen


def require_missing_version(name, version):
    url = f"https://pypi.org/pypi/{name}/{version}/json"
    try:
        with urlopen(url, timeout=15):
            raise SystemExit("Version already exists or registry returned an unexpected success")
    except HTTPError as error:
        if (error.code != 404 or error.geturl() != url
                or json.load(error) != {"message": "Not Found"}):
            raise SystemExit("Unexpected PyPI response, publication stopped") from error


if __name__ == "__main__":
    import tomllib
    from pathlib import Path

    project = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))["project"]
    name, version = project["name"], project["version"]
    if (name != os.environ["PYPI_PROJECT"] or re.sub(r"[-_.]+", "-", name.lower()) == "outlook-classic-mcp"):
        raise SystemExit("Select a fork-owned PyPI project before publication")
    if os.environ["GITHUB_REF"] != f"refs/tags/v{version}":
        raise SystemExit("Select the reviewed tag v<package version> for publication")
    require_missing_version(name, version)
