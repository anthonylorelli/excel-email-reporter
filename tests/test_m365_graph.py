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
    assert manager.account_mode == "work"
    assert manager.tenant_id == "organizations"
    assert "https://login.microsoftonline.com/organizations" in manager.authority
    assert "https://graph.microsoft.com/Mail.ReadWrite" in manager.scopes
    assert manager.sso_enabled is True


def test_m365_manager_account_selection_options():
    # Personal mode via keyword
    cfg_personal = Config({"email": {"m365_account": "personal"}})
    mgr_personal = M365GraphDraftManager(cfg_personal)
    assert mgr_personal.account_mode == "personal"
    assert mgr_personal.tenant_id == "consumers"
    assert "https://login.microsoftonline.com/consumers" in mgr_personal.authority

    # Organizational mode via email address
    cfg_work_email = Config({"email": {"m365_account": "teacher@connectionsnw.org"}})
    mgr_work_email = M365GraphDraftManager(cfg_work_email)
    assert mgr_work_email.account_mode == "work"
    assert mgr_work_email.login_hint == "teacher@connectionsnw.org"
    assert mgr_work_email.tenant_id == "organizations"

    # Personal mode via live.com email address
    cfg_live_email = Config({"email": {"m365_account": "anthony.lorelli@live.com"}})
    mgr_live_email = M365GraphDraftManager(cfg_live_email)
    assert mgr_live_email.account_mode == "personal"
    assert mgr_live_email.login_hint == "anthony.lorelli@live.com"
    assert mgr_live_email.tenant_id == "consumers"


def test_m365_manager_silent_token_acquisition():
    config = Config({})
    manager = M365GraphDraftManager(config)

    mock_app = MagicMock()
    # Ensure personal cached account is skipped when in default 'work' mode, and organizational account is selected
    mock_app.get_accounts.return_value = [
        {"username": "personal@live.com", "realm": "9188040d-6c67-4c5b-b112-36a304b66dad"},
        {"username": "teacher@connectionsnw.org", "realm": "org-tenant-guid"},
    ]
    mock_app.acquire_token_silent.return_value = {
        "access_token": "mock_silent_token_12345",
        "id_token_claims": {"preferred_username": "teacher@connectionsnw.org"},
    }

    with patch.object(manager, "_get_app", return_value=mock_app):
        token, err = manager.acquire_token(interactive=False)

        assert token == "mock_silent_token_12345"
        assert err is None
        mock_app.acquire_token_silent.assert_called_once_with(
            manager.scopes,
            account={"username": "teacher@connectionsnw.org", "realm": "org-tenant-guid"},
        )


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


def test_m365_manager_wam_contract_violation_falls_back_to_browser():
    config = Config({})
    manager = M365GraphDraftManager(config)

    mock_broker_app = MagicMock()
    mock_broker_app._enable_broker = True
    mock_broker_app.get_accounts.return_value = []
    mock_broker_app.acquire_token_interactive.return_value = {
        "error": "broker_error",
        "error_description": "(pii). Status: Response_Status.Status_ApiContractViolation, Error code: 2156265475, Tag: 557973634",
    }

    mock_browser_app = MagicMock()
    mock_browser_app._enable_broker = False
    mock_browser_app.acquire_token_interactive.return_value = {
        "access_token": "browser_fallback_token_999",
        "id_token_claims": {"preferred_username": "teacher@connectionsnw.org"},
    }

    def get_app_side_effect(enable_broker=None):
        if enable_broker is False:
            return mock_browser_app
        return mock_broker_app

    with patch.object(manager, "_acquire_outlook_wam_token", return_value=None):
        with patch.object(manager, "_get_app", side_effect=get_app_side_effect):
            token, err = manager.acquire_token(interactive=True)

            assert token == "browser_fallback_token_999"
            assert err is None
            mock_broker_app.acquire_token_interactive.assert_called_once()
            mock_browser_app.acquire_token_interactive.assert_called_once()


def test_m365_manager_personal_account_owa_draft_creation():
    config = Config({"email": {"m365_account": "personal"}})
    manager = M365GraphDraftManager(config)

    mock_broker_app = MagicMock()
    mock_broker_app._enable_broker = True
    mock_broker_app.get_accounts.return_value = []

    wam_token_res = {
        "access_token": "EwBIBOl3BAAUcnotHNLX_personal_msa_token",
        "id_token_claims": {"preferred_username": "student@live.com"},
    }

    mock_owa_resp = MagicMock()
    mock_owa_resp.status_code = 200
    mock_owa_resp.json.return_value = {
        "Body": {
            "ResponseMessages": {
                "Items": [
                    {
                        "ResponseClass": "Success",
                        "ResponseCode": "NoError",
                        "Items": [
                            {
                                "ItemId": {
                                    "Id": "owa_draft_id_123",
                                    "ChangeKey": "ck_456",
                                }
                            }
                        ],
                    }
                ]
            }
        }
    }

    with patch.object(manager, "_get_app", return_value=mock_broker_app):
        with patch.object(manager, "_acquire_outlook_wam_token", return_value=wam_token_res):
            with patch("requests.post", return_value=mock_owa_resp) as mock_post:
                res = manager.create_draft(
                    recipient_email="parent@live.com",
                    recipient_name="Parent",
                    subject="Progress Report",
                    html_body="<p>Report content</p>",
                )

                assert res["success"] is True
                assert res["message_id"] == "owa_draft_id_123"
                assert res["method"] == "graph"
                mock_post.assert_called_once()
                args, kwargs = mock_post.call_args
                assert "outlook.office.com/owa/service.svc?action=CreateItem" in args[0]
                assert kwargs["headers"]["Action"] == "CreateItem"



