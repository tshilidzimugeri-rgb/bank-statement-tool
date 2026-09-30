"""The upload page, driven through Streamlit's test runner: its side menu,
and each upload making a new workbook."""
import dataclasses
import sys
from datetime import timedelta
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))
from synthetic import build, generate, render_pdf  # noqa: E402

from statement_tool.excel_writer import read_transactions  # noqa: E402

APP = Path(__file__).parent.parent / "src" / "statement_tool" / "app.py"


@pytest.fixture(autouse=True)
def _restore_package_modules():
    """The page imports this package afresh when it starts; put back the
    modules the other tests already hold, or their classes stop matching."""
    def ours():
        return [n for n in sys.modules if n == "statement_tool" or n.startswith("statement_tool.")]
    saved = {n: sys.modules[n] for n in ours()}
    yield
    for n in ours():
        del sys.modules[n]
    sys.modules.update(saved)


def _two_statements_of_one_account(tmp_path):
    first = generate(4)
    rec = first.recipe
    start = rec.end + timedelta(days=1)
    end = (start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    rows = sorted(((min(end, d + (start - rec.start)), *rest) for d, *rest in rec.day_rows), key=lambda r: r[0])
    second = build(dataclasses.replace(rec, start=start, end=end, opening=first.closing, day_rows=rows))
    return render_pdf(first, tmp_path / "first.pdf"), render_pdf(second, tmp_path / "second.pdf")


def _go(at, page):
    at.sidebar.radio(key="page").set_value(page).run()
    assert not at.exception, [e.message for e in at.exception]


def _upload(at, pdfs):
    _go(at, "Upload statements")
    at.file_uploader[0].set_value([(p.name, p.read_bytes(), "application/pdf") for p in pdfs])
    at.button[0].click().run()
    assert not at.exception, [e.message for e in at.exception]
    return Path(at.session_state["last_workbook"])


def _open(at, workbook):
    _go(at, "Continue a workbook")
    at.file_uploader[0].set_value([(workbook.name, workbook.read_bytes(), "application/octet-stream")])
    at.button[0].click().run()
    assert not at.exception, [e.message for e in at.exception]


def _sources(workbook):
    """The uploaded file names in the workbook (saved as "<hash>_<name>")."""
    return {r["Source File"].split("_", 1)[1] for r in read_transactions(workbook)}


def test_each_upload_makes_a_new_workbook_and_continuing_one_is_its_own_page(tmp_path, monkeypatch):
    # The test runner counts as online, which needs a password. On this
    # computer the page behaves the same, without one.
    monkeypatch.setenv("APP_PASSWORD", "test")
    first, second = _two_statements_of_one_account(tmp_path)
    at = AppTest.from_file(str(APP), default_timeout=300)
    at.session_state["signed_in"] = True
    at.run()
    assert not at.exception, [e.message for e in at.exception]

    book = _upload(at, [first])
    assert _sources(book) == {first.name}
    kept = tmp_path / "downloaded.xlsx"
    kept.write_bytes(book.read_bytes())

    # A second statement of the same account, uploaded on its own later in
    # the same visit: the new workbook holds only it.
    assert _sources(_upload(at, [second])) == {second.name}

    # Uploaded together, two statements of one account get a workbook each.
    _upload(at, [first, second])
    books = sorted((Path(at.session_state["work_dir"]) / "output").glob("*.xlsx"))
    assert sorted(_sources(b) for b in books) == [{first.name}, {second.name}]

    # Opened on "Continue a workbook", the earlier workbook gets it added.
    _open(at, kept)
    assert _sources(_upload(at, [second])) == {first.name, second.name}

    # Every page of the menu shows this workbook.
    for page in ("Financials", "VAT", "Checks", "Download"):
        _go(at, page)
        assert at.header[0].value.startswith(page)
