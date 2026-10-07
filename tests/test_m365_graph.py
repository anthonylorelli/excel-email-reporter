"""Unit tests for M365GraphDraftManager (Windows SSO Token Broker & Graph API)."""

from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from src.config import Config
from src.m365_graph_manager import M365GraphDraftManager, DEFAULT_CLIENT_ID


def test_m365_manager_init_defaults():
    config = Config({})
    manager = M365GraphDraftManager(config)

    assert manager.client_id == DEFAULT_CLIENT_ID
    assert manager.tenant_id == "common"
    assert "https://login.microsoftonline.com/common" in manager.authority
    assert "https://graph.microsoft.com/Mail.ReadWrite" in manager.scopes
    assert manager.sso_enabled is True


def test_m365_manager_silent_token_acquisition():
    config = Config({})
    manager = M365GraphDraftManager(config)

    mock_app = MagicMock()
    mock_app.get_accounts.return_value = [{"username": "testuser@live.com"}]
    mock_app.acquire_token_silent.return_value = {
        "access_token": "mock_silent_token_12345",
        "id_token_claims": {"preferred_username": "testuser@live.com"},
    }

    with patch.object(manager, "_get_app", return_value=mock_app):
        token, err = manager.acquire_token(interactive=False)

        assert token == "mock_silent_token_12345"
        assert err is None
        mock_app.acquire_token_silent.assert_called_once()


def test_m365_manager_interactive_token_acquisition():
    config = Config({})
    manager = M365GraphDraftManager(config)

    mock_app = MagicMock()
    mock_app.get_accounts.return_value = []
    mock_app.acquire_token_interactive.return_value = {
        "access_token": "mock_interactive_token_67890",
        "id_token_claims": {"preferred_username": "testuser@live.com"},
    }

    with patch.object(manager, "_get_app", return_value=mock_app):
        token, err = manager.acquire_token(interactive=True)

        assert token == "mock_interactive_token_67890"
        assert err is None
        mock_app.acquire_token_interactive.assert_called_once()


def test_m365_manager_create_draft_success(tmp_path):
    config = Config({})
    manager = M365GraphDraftManager(config)
    manager._access_token = "valid_token"

    # Create dummy attachment
    dummy_att = tmp_path / "summary.pdf"
    dummy_att.write_bytes(b"%PDF-1.4 dummy content")

    mock_resp = MagicMock()
    mock_resp.status_code = 201
    mock_resp.json.return_value = {"id": "graph_msg_abc123"}

    with patch("requests.post", return_value=mock_resp) as mock_post:
        res = manager.create_draft(
            recipient_email="student@school.org",
            recipient_name="Student Bob",
            subject="Progress Report: Student Bob",
            html_body="<p>Grade: 92%</p>",
            text_body="Grade: 92%",
            attachments=[dummy_att],
        )

        assert res["success"] is True
        assert res["message_id"] == "graph_msg_abc123"
        assert res["method"] == "graph"

        mock_post.assert_called_once()
        args, kwargs = mock_post.call_args
        assert args[0] == "https://graph.microsoft.com/v1.0/me/messages"
        payload = kwargs["json"]
        assert payload["subject"] == "Progress Report: Student Bob"
        assert payload["toRecipients"][0]["emailAddress"]["address"] == "student@school.org"
        assert payload["isDraft"] is True
        assert len(payload["attachments"]) == 1
        assert payload["attachments"][0]["name"] == "summary.pdf"


def test_m365_manager_token_refresh_on_401():
    config = Config({})
    manager = M365GraphDraftManager(config)
    manager._access_token = "initial_expired_token"

    mock_resp_401 = MagicMock()
    mock_resp_401.status_code = 401
    mock_resp_401.text = "Unauthorized"

    mock_resp_201 = MagicMock()
    mock_resp_201.status_code = 201
    mock_resp_201.json.return_value = {"id": "new_msg_xyz"}

    with patch.object(manager, "acquire_token", return_value=("refreshed_token", None)):
        with patch("requests.post", side_effect=[mock_resp_401, mock_resp_201]) as mock_post:
            res = manager.create_draft(
                recipient_email="student@school.org",
                recipient_name="Student Bob",
                subject="Progress Report",
                html_body="<p>Grade: 85%</p>",
            )

            assert res["success"] is True
            assert res["message_id"] == "new_msg_xyz"
            assert mock_post.call_count == 2


def test_m365_manager_handles_interactive_auth_failure():
    config = Config({})
    manager = M365GraphDraftManager(config)

    mock_app = MagicMock()
    mock_app.get_accounts.return_value = []
    mock_app.acquire_token_interactive.return_value = {
        "error": "access_denied",
        "error_description": "User cancelled the sign-in dialog.",
    }

    with patch.object(manager, "_get_app", return_value=mock_app):
        token, err = manager.acquire_token(interactive=True)

        assert token is None
        assert "User cancelled the sign-in dialog" in err

        # Subsequent draft attempt should fail cleanly without re-prompting
        draft_res = manager.create_draft(
            recipient_email="student@school.org",
            recipient_name="Student",
            subject="Report",
            html_body="<p>Content</p>",
        )
        assert draft_res["success"] is False
        assert "User cancelled the sign-in dialog" in draft_res["error"]
