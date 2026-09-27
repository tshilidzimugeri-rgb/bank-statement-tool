"""The upload page as it runs online, driven through Streamlit's test runner."""
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


def _upload(at, pdfs=(), workbooks=()):
    boxes = {u.label.split(" (")[0]: u for u in at.file_uploader}
    if workbooks:
        boxes["Workbook you downloaded from this page last time"].set_value(
            [(p.name, p.read_bytes(), "application/octet-stream") for p in workbooks])
    boxes["Bank statement PDFs"].set_value([(p.name, p.read_bytes(), "application/pdf") for p in pdfs])
    at.button[0].click()
    at.run()
    assert not at.exception, [e.message for e in at.exception]
    return Path(at.session_state["last_workbook"])


def _sources(workbook):
    """The uploaded file names in the workbook (saved as "<hash>_<name>")."""
    return {r["Source File"].split("_", 1)[1] for r in read_transactions(workbook)}


def test_each_online_upload_makes_a_new_workbook(tmp_path, monkeypatch):
    monkeypatch.setenv("APP_PASSWORD", "test")
    first, second = _two_statements_of_one_account(tmp_path)
    at = AppTest.from_file(str(APP), default_timeout=300)
    at.session_state["signed_in"] = True
    at.run()
    assert any("each upload makes a new workbook" in i.value for i in at.info), "not running as online"

    book = _upload(at, [first])
    assert _sources(book) == {first.name}
    kept = tmp_path / "downloaded.xlsx"
    kept.write_bytes(book.read_bytes())

    # A second statement of the same account, uploaded on its own later in
    # the same visit: the workbook holds only it.
    assert _sources(_upload(at, [second])) == {second.name}

    # Uploaded together with the earlier workbook, it's added to it.
    assert _sources(_upload(at, [second], workbooks=[kept])) == {first.name, second.name}
