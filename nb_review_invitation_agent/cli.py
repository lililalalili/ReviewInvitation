from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

from .config import load_config


def _load_legacy_module():
    repo_root = Path(__file__).resolve().parents[1]
    legacy_path = repo_root / "legacy" / "review_invitation_agent_windows_ubuntu_ollama_v14.py"
    spec = importlib.util.spec_from_file_location("legacy_v14", legacy_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load legacy script at {legacy_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="NB Review Invitation Agent")
    p.add_argument("--dry-run", action="store_true", help="No-op execution for CI/bootstrap checks")
    p.add_argument("--no-gui", action="store_true", help="Disable GUI launch (reserved for future tasks)")
    p.add_argument("--fake-providers", action="store_true", help="Use offline fake providers in tests/CI")
    p.add_argument("--legacy-v14", action="store_true", help="Run legacy v14 monolithic script")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    cfg = load_config(dry_run=args.dry_run, no_gui=args.no_gui, fake_providers=args.fake_providers)
    if args.dry_run and not args.legacy_v14:
        print(f"[DRY-RUN] baseline CLI OK. base_dir={cfg.base_dir}")
        return 0

    if args.legacy_v14:
        legacy = _load_legacy_module()
        rc = legacy.run_with_lock()
        return 0 if rc is None else int(rc)

    print("Baseline package scaffold is ready. Use --legacy-v14 for production behavior until later tasks land.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
