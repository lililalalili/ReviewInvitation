from __future__ import annotations

import argparse
from pathlib import Path
import zipfile

EXCLUDE_PARTS = {
    ".git",
    ".venv",
    "__pycache__",
    "RunLogs",
    "PubMed",
    "dist",
}
EXCLUDE_SUFFIXES = {
    ".pyc",
    ".pyo",
    ".xlsm",
    ".xlsx",
    ".xls",
    ".log",
    ".sqlite",
    ".sqlite3",
    ".db",
    ".xml",
}
EXCLUDE_NAMES = {
    ".env",
}


def should_include(path: Path) -> bool:
    if any(part in EXCLUDE_PARTS for part in path.parts):
        return False
    if path.name in EXCLUDE_NAMES:
        return False
    if path.suffix.lower() in EXCLUDE_SUFFIXES:
        return False
    return True


def build_release_zip(repo_root: Path, output_zip: Path) -> Path:
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in repo_root.rglob("*"):
            if path.is_dir():
                continue
            rel = path.relative_to(repo_root)
            if rel == output_zip.relative_to(repo_root) or not should_include(rel):
                continue
            zf.write(path, rel.as_posix())
    return output_zip


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create release zip excluding sensitive/runtime files")
    parser.add_argument("--output", default="dist/nb_review_invitation_agent_release.zip")
    args = parser.parse_args(argv)

    repo_root = Path(__file__).resolve().parents[1]
    out = repo_root / args.output
    build_release_zip(repo_root, out)
    print(f"Created: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
