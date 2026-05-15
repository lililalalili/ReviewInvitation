from nb_review_invitation_agent.cli import main


def test_cli_dry_run_returns_zero():
    rc = main(["--dry-run", "--fake-providers", "--no-gui"])
    assert rc == 0


def test_cli_default_returns_zero():
    rc = main([])
    assert rc == 0
