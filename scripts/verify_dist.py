"""Check the two release archives against the source checkout. Run with Python 3.11+."""

import hashlib
import sys
import tarfile
import tomllib
import zipfile
from email.parser import BytesParser
from pathlib import Path


def check(condition, message):
    if not condition:
        raise SystemExit(message)


def verify(dist):
    root = Path(__file__).resolve().parents[1]
    project = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    name, version = project["name"], project["version"]
    stem = f"{name.replace('-', '_')}-{version}"
    sdist, wheel = dist / f"{stem}.tar.gz", dist / f"{stem}-py3-none-any.whl"
    artifacts = {path for path in dist.iterdir() if path.name != ".gitignore"}
    check(artifacts == {sdist, wheel}, "Expected exactly one sdist and one wheel")
    src = root / "src"
    payload = {
        path.relative_to(src).as_posix(): path.read_bytes()
        for path in (src / "outlook_mcp").rglob("*")
        if path.suffix == ".py" or (path.parent == src / "outlook_mcp/ui" and path.suffix == ".html")
    }
    info = f"{stem}.dist-info"
    metadata = {"METADATA", "RECORD", "WHEEL", "entry_points.txt", "licenses/LICENSE"}
    with zipfile.ZipFile(wheel) as archive:
        files = [f.filename for f in archive.infolist() if not f.is_dir()]
        check(sorted(files) == sorted(set(payload) | {f"{info}/{f}" for f in metadata}),
              "Unexpected or missing wheel file")
        for path, data in payload.items():
            check(archive.read(path) == data, f"Wheel differs from checkout at {path}")
        check(archive.read(f"{info}/licenses/LICENSE") == (root / "LICENSE").read_bytes(),
              "Wheel license differs from checkout")
        meta = BytesParser().parsebytes(archive.read(f"{info}/METADATA"))
        check(meta["Name"] == name and meta["Version"] == version, "Wheel identity mismatch")
        check(f"Homepage, {project['urls']['Homepage']}" in meta.get_all("Project-URL", []),
              "Wheel source URL mismatch")
    with tarfile.open(sdist, "r:gz") as archive:
        members = [f for f in archive if not f.isdir()]
        expected = {f"src/{p}" for p in payload} | {
            "PKG-INFO", "pyproject.toml", "pyproject.toml.orig", "LICENSE", "README.md",
        }
        check(all(f.isfile() for f in members), "Unexpected special file in sdist")
        check(sorted(f.name for f in members) == sorted(f"{stem}/{p}" for p in expected),
              "Unexpected or missing sdist file")
        for path, data in payload.items():
            check(archive.extractfile(f"{stem}/src/{path}").read() == data,
                  f"Sdist differs from checkout at {path}")
    for path in (sdist, wheel):
        print(hashlib.sha256(path.read_bytes()).hexdigest(), path.name)


if __name__ == "__main__":
    verify(Path(sys.argv[1]))
