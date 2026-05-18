from pathlib import Path


def test_windows_launch_scripts_exist():
    for name in [
        'run_gui_windows.bat',
        'run_dry_run_windows.bat',
        'run_tests_windows.bat',
    ]:
        assert Path(name).exists(), f'missing {name}'


def test_release_docs_exist():
    for name in [
        'docs/USER_GUIDE.md',
        'docs/DEVELOPER_GUIDE.md',
        'docs/WINDOWS_SMOKE_TEST_CHECKLIST.md',
        'docs/RELEASE_CHECKLIST.md',
    ]:
        assert Path(name).exists(), f'missing {name}'
