from pathlib import Path
import zipfile

from scripts.make_release_zip import build_release_zip, should_include


def test_should_include_exclusion_rules():
    assert not should_include(Path('.git/config'))
    assert not should_include(Path('.venv/pyvenv.cfg'))
    assert not should_include(Path('.env'))
    assert not should_include(Path('RunLogs/log.txt'))
    assert not should_include(Path('PubMed/input.xml'))
    assert not should_include(Path('dist/old.zip'))
    assert not should_include(Path('__pycache__/a.pyc'))
    assert not should_include(Path('any_author_file.xlsm'))
    assert not should_include(Path('any_log_file.xlsx'))
    assert not should_include(Path('old_style.xls'))
    assert not should_include(Path('debug.log'))
    assert not should_include(Path('state.sqlite'))
    assert not should_include(Path('state.sqlite3'))
    assert not should_include(Path('state.db'))
    assert not should_include(Path('data.xml'))
    assert should_include(Path('README.md'))


def test_build_release_zip_excludes_private_and_runtime_files(tmp_path):
    repo = tmp_path / 'repo'
    repo.mkdir()
    (repo / 'README.md').write_text('ok', encoding='utf-8')
    (repo / '.env').write_text('secret', encoding='utf-8')
    (repo / 'arbitrary_runtime.xlsm').write_text('binary', encoding='utf-8')
    (repo / 'arbitrary_runtime.xlsx').write_text('binary', encoding='utf-8')
    (repo / 'audit.log').write_text('x', encoding='utf-8')
    (repo / 'state.sqlite3').write_text('x', encoding='utf-8')
    (repo / 'RunLogs').mkdir()
    (repo / 'RunLogs' / 'run.log').write_text('x', encoding='utf-8')
    (repo / 'PubMed').mkdir()
    (repo / 'PubMed' / 'authors.xml').write_text('x', encoding='utf-8')
    (repo / 'dist').mkdir()
    (repo / 'dist' / 'old-release.zip').write_text('x', encoding='utf-8')

    out = repo / 'dist' / 'release.zip'
    build_release_zip(repo, out)

    with zipfile.ZipFile(out) as zf:
        names = set(zf.namelist())

    assert 'README.md' in names
    assert '.env' not in names
    assert 'arbitrary_runtime.xlsm' not in names
    assert 'arbitrary_runtime.xlsx' not in names
    assert 'audit.log' not in names
    assert 'state.sqlite3' not in names
    assert 'RunLogs/run.log' not in names
    assert 'PubMed/authors.xml' not in names
    assert 'dist/old-release.zip' not in names
