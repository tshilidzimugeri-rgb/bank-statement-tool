from statement_tool.config import ClientRule
from statement_tool.gmail_client import (
    build_search_query,
    get_message_info,
    safe_local_filename,
)


def test_build_search_query_includes_lookback_and_client_terms():
    rules = [ClientRule(name="Acme", match=["acme", "acmetrading.co.za"])]
    query = build_search_query(rules, 30)
    assert "has:attachment" in query
    assert "filename:pdf" in query
    assert "newer_than:30d" in query
    assert "acme" in query
    assert "acmetrading.co.za" in query


def test_build_search_query_quotes_multiword_terms():
    rules = [ClientRule(name="Jane", match=["jane dlamini"])]
    query = build_search_query(rules, 30)
    assert '"jane dlamini"' in query


def test_build_search_query_with_no_client_rules_still_searches_pdfs():
    query = build_search_query([], 14)
    assert query == "has:attachment filename:pdf newer_than:14d"


def test_safe_local_filename_strips_unsafe_windows_characters():
    name = safe_local_filename("18c9f", 'weird:name*"?.pdf')
    assert name == "18c9f_weird_name___.pdf"
    assert not any(c in name for c in '<>:"/\\|?*')


class _FakeAttachmentsResource:
    def get(self, userId, messageId, id):
        raise AssertionError("not needed for this test")


class _FakeMessagesResource:
    def __init__(self, message):
        self._message = message

    def get(self, userId, id, format):
        class _Req:
            def execute(_self, **kwargs):
                return self._message
        return _Req()

    def attachments(self):
        return _FakeAttachmentsResource()


class _FakeService:
    def __init__(self, message):
        self._message = message

    def users(self):
        outer = self

        class _Users:
            def messages(_self):
                return _FakeMessagesResource(outer._message)
        return _Users()


def test_get_message_info_extracts_headers_and_pdf_attachments_only():
    message = {
        "payload": {
            "headers": [
                {"name": "From", "value": "Acme Bank <statements@acmetrading.co.za>"},
                {"name": "Subject", "value": "Your March statement"},
            ],
            "parts": [
                {"filename": "", "mimeType": "text/plain", "body": {"data": "..."}},
                {
                    "filename": "statement.pdf",
                    "mimeType": "application/pdf",
                    "body": {"attachmentId": "att-1", "size": 12345},
                },
                {
                    "filename": "logo.png",
                    "mimeType": "image/png",
                    "body": {"attachmentId": "att-2", "size": 500},
                },
            ],
        }
    }
    service = _FakeService(message)
    info = get_message_info(service, "msg-1")

    assert info.sender == "Acme Bank <statements@acmetrading.co.za>"
    assert info.subject == "Your March statement"
    assert [a.filename for a in info.pdf_attachments] == ["statement.pdf"]
    assert info.pdf_attachments[0].attachment_id == "att-1"
