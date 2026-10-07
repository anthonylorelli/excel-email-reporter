"""Unit tests for OutlookDraftManager."""

from unittest.mock import MagicMock, patch
from pathlib import Path
from src.config import Config
from src.outlook_manager import OutlookDraftManager


def test_outlook_manager_com_draft():
    config = Config({"email": {"sender": "teacher@school.org"}})
    manager = OutlookDraftManager(config)

    # Mock win32com
    mock_mail = MagicMock()
    mock_app = MagicMock()
    mock_app.CreateItem.return_value = mock_mail

    with patch.object(manager, "get_or_create_com_app", return_value=mock_app):
        success = manager.create_draft_via_com(
            recipient_email="student@school.org",
            recipient_name="Alice Student",
            subject="Progress Report",
            html_body="<p>Grade: 95%</p>",
            attachments=None,
        )

        assert success is True
        mock_app.CreateItem.assert_called_once_with(0)
        assert mock_mail.To == '"Alice Student" <student@school.org>'
        assert mock_mail.Subject == "Progress Report"
        assert mock_mail.HTMLBody == "<p>Grade: 95%</p>"
        mock_mail.Save.assert_called_once()


def test_outlook_manager_imap_draft():
    config = Config({"email": {"sender": "teacher@school.org"}})
    manager = OutlookDraftManager(config)
    manager.imap_user = "teacher@school.org"
    manager.imap_password = "password123"

    mock_imap = MagicMock()
    mock_imap.append.return_value = ("OK", [b"APPEND completed"])

    with patch.object(manager, "_get_imap_client", return_value=mock_imap):
        manager._imap_drafts_folder = "Drafts"
        success = manager.create_draft_via_imap(
            recipient_email="student@school.org",
            recipient_name="Bob Student",
            subject="Weekly Progress",
            html_body="<p>Grade: 88%</p>",
            text_body="Grade: 88%",
            attachments=None,
        )

        assert success is True
        mock_imap.append.assert_called_once()
        args, kwargs = mock_imap.append.call_args
        assert args[0] == "Drafts"
        assert args[1] == "\\Draft"


def test_outlook_manager_fallback_when_neither_available():
    config = Config({})
    manager = OutlookDraftManager(config)
    manager.imap_user = ""
    manager.imap_password = ""

    with patch.object(manager.graph_manager, "is_available", return_value=False):
        with patch.object(manager, "is_com_available", return_value=(False, "Class not registered")):
            result = manager.save_draft_to_outlook(
                recipient_email="test@test.com",
                recipient_name="Test",
                subject="Sub",
                html_body="body",
                text_body="body",
            )

            assert result["success"] is False
            assert "Classic Outlook COM unavailable" in result["error"]


def test_outlook_manager_graph_draft_success():
    config = Config({})
    manager = OutlookDraftManager(config)

    mock_graph_res = {"success": True, "method": "graph", "message_id": "msg_123", "error": None}
    with patch.object(manager.graph_manager, "is_available", return_value=True):
        with patch.object(manager.graph_manager, "create_draft", return_value=mock_graph_res) as mock_draft:
            result = manager.save_draft_to_outlook(
                recipient_email="student@school.org",
                recipient_name="Alice",
                subject="Progress Report",
                html_body="<p>Grade</p>",
                text_body="Grade",
            )

            assert result["success"] is True
            assert result["method"] == "graph"
            assert result["message_id"] == "msg_123"
            mock_draft.assert_called_once()


def test_outlook_manager_graph_fails_falls_back_to_com():
    config = Config({})
    manager = OutlookDraftManager(config)

    graph_fail_res = {"success": False, "method": "graph", "error": "SSO cancelled"}
    with patch.object(manager.graph_manager, "is_available", return_value=True):
        with patch.object(manager.graph_manager, "create_draft", return_value=graph_fail_res):
            with patch.object(manager, "is_com_available", return_value=(True, None)):
                with patch.object(manager, "create_draft_via_com", return_value=True) as mock_com:
                    result = manager.save_draft_to_outlook(
                        recipient_email="student@school.org",
                        recipient_name="Alice",
                        subject="Progress Report",
                        html_body="<p>Grade</p>",
                        text_body="Grade",
                    )

                    assert result["success"] is True
                    assert result["method"] == "com"
                    mock_com.assert_called_once()

