from nb_review_invitation_agent.cli import main


def test_cli_dry_run_returns_zero():
    rc = main(["--dry-run", "--fake-providers", "--no-gui"])
    assert rc == 0


def test_cli_default_returns_zero():
    rc = main([])
    assert rc == 0


def test_cli_legacy_none_return_maps_to_zero(monkeypatch):
    class FakeLegacy:
        @staticmethod
        def run_with_lock():
            return None

    monkeypatch.setattr("nb_review_invitation_agent.cli._load_legacy_module", lambda: FakeLegacy())
    rc = main(["--legacy-v14"])
    assert rc == 0


def test_cli_legacy_numeric_return_passthrough(monkeypatch):
    class FakeLegacy:
        @staticmethod
        def run_with_lock():
            return 2

    monkeypatch.setattr("nb_review_invitation_agent.cli._load_legacy_module", lambda: FakeLegacy())
    rc = main(["--legacy-v14"])
    assert rc == 2


def test_cli_gui_launch(monkeypatch):
    called = {"ok": False}

    def fake_launch(_cfg):
        called["ok"] = True

    monkeypatch.setattr("nb_review_invitation_agent.cli.launch_gui", fake_launch)
    rc = main(["--gui", "--fake-providers"])
    assert rc == 0
    assert called["ok"] is True
