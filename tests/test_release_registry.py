"""Synthetic registry responses covering the publication gate."""

import io
from unittest.mock import MagicMock, patch
from urllib.error import HTTPError, URLError

import pytest

from scripts.check_pypi import require_missing_version


URL = "https://pypi.org/pypi/example-project/1.0/json"


@pytest.mark.parametrize(("code", "body", "url", "allowed"), [
    (404, b'{"message":"Not Found"}', URL, True),
    (403, b'{}', URL, False),
    (429, b'{}', URL, False),
    (503, b'{}', URL, False),
    (404, b'upstream unavailable', URL, False),
    (404, b'{"message":"Not Found"}', URL + "/elsewhere", False),
])
def test_registry_status(code, body, url, allowed):
    response = HTTPError(url, code, "", {}, io.BytesIO(body))
    with patch("scripts.check_pypi.urlopen", side_effect=response):
        if allowed:
            require_missing_version("example-project", "1.0")
        else:
            with pytest.raises((SystemExit, ValueError)):
                require_missing_version("example-project", "1.0")


def test_existing_version_and_transport_failure_stop():
    with patch("scripts.check_pypi.urlopen", return_value=MagicMock()), pytest.raises(SystemExit):
        require_missing_version("example-project", "1.0")
    with patch("scripts.check_pypi.urlopen", side_effect=URLError("offline")), pytest.raises(URLError):
        require_missing_version("example-project", "1.0")
