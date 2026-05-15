from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class RuntimeConfig:
    base_dir: Path
    log_xlsx: Path
    target_xlsm: Path
    pubmed_dir: Path
    run_log_dir: Path
    state_db: Path
    dry_run: bool = False
    no_gui: bool = False
    fake_providers: bool = False



def load_config(*, dry_run: bool = False, no_gui: bool = False, fake_providers: bool = False) -> RuntimeConfig:
    base_dir = Path(os.getenv("REVIEW_INVITATION_BASE_DIR", r"D:\Agents\ReviewInvtation"))
    run_log_dir = base_dir / "RunLogs"
    return RuntimeConfig(
        base_dir=base_dir,
        log_xlsx=base_dir / "ReviewInvitationLog.xlsx",
        target_xlsm=base_dir / "NB_Author_2026.xlsm",
        pubmed_dir=base_dir / "PubMed",
        run_log_dir=run_log_dir,
        state_db=run_log_dir / "review_invitation_agent_state.sqlite3",
        dry_run=dry_run,
        no_gui=no_gui,
        fake_providers=fake_providers,
    )
