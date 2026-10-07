"""Microsoft 365 Graph Draft Manager: populates draft emails directly into M365 via Graph API and Windows SSO Token Broker."""

import base64
import logging
import mimetypes
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests

try:
    import msal  # type: ignore
except ImportError:
    msal = None  # type: ignore

try:
    import pymsalruntime  # type: ignore
except ImportError:
    pymsalruntime = None  # type: ignore

# Microsoft Graph Command Line Tools public client ID (supports WAM & delegated Graph scopes)
DEFAULT_CLIENT_ID = "14d82eec-204b-449e-b7d6-3e00e84b80e4"
DEFAULT_SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite"]
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common"


class M365GraphDraftManager:
    """Manages Microsoft 365 authentication via Windows SSO Token Broker (WAM)

    and creates draft messages in the user's mailbox via Microsoft Graph API.
    """

    def __init__(self, config: Optional[Any] = None):
        self.config = config
        self.logger = logging.getLogger(__name__)

        self.client_id = os.getenv(
            "M365_CLIENT_ID",
            getattr(config, "m365_client_id", DEFAULT_CLIENT_ID) if config else DEFAULT_CLIENT_ID,
        )
        self.tenant_id = os.getenv(
            "M365_TENANT_ID",
            getattr(config, "m365_tenant_id", "common") if config else "common",
        )
        self.authority = os.getenv(
            "M365_AUTHORITY",
            getattr(config, "m365_authority", f"https://login.microsoftonline.com/{self.tenant_id}")
            if config
            else f"https://login.microsoftonline.com/{self.tenant_id}",
        )

        scopes_env = os.getenv("M365_SCOPES")
        if scopes_env:
            self.scopes = [s.strip() for s in scopes_env.split(",") if s.strip()]
        elif config and getattr(config, "m365_scopes", None):
            self.scopes = config.m365_scopes
        else:
            self.scopes = DEFAULT_SCOPES

        # Token cache location (stored in output directory, git-ignored)
        drafts_dir = getattr(config, "output_drafts_dir", "output/drafts") if config else "output/drafts"
        cache_path_env = os.getenv("M365_CACHE_PATH")
        if cache_path_env:
            self.cache_path = Path(cache_path_env)
        else:
            self.cache_path = Path(drafts_dir).parent / ".msal_cache.bin"

        sso_env = os.getenv("M365_SSO")
        if sso_env is not None:
            self.sso_enabled = sso_env.lower() in ("true", "1", "yes")
        elif config and hasattr(config, "m365_sso"):
            self.sso_enabled = bool(config.m365_sso)
        else:
            self.sso_enabled = True

        self._token_cache: Optional[Any] = None
        self._app: Optional[Any] = None
        self._access_token: Optional[str] = None
        self._failed_previously: bool = False
        self._failure_reason: Optional[str] = None

    def is_available(self) -> bool:
        """Returns True if MSAL and requests are available and SSO is enabled."""
        if not self.sso_enabled:
            return False
        if msal is None or requests is None:
            return False
        return True

    def is_broker_supported(self) -> bool:
        """Returns True if Windows native Token Broker (WAM) is supported."""
        return sys.platform == "win32" and pymsalruntime is not None

    def _get_app(self) -> Any:
        """Retrieves or instantiates the MSAL PublicClientApplication."""
        if self._app is None:
            if msal is None:
                raise RuntimeError("MSAL library is not installed.")

            self._token_cache = msal.SerializableTokenCache()
            if self.cache_path.exists():
                try:
                    self._token_cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
                except Exception as ex:
                    self.logger.warning("Failed to load MSAL token cache: %s", ex)

            enable_broker = self.is_broker_supported()

            try:
                self._app = msal.PublicClientApplication(
                    client_id=self.client_id,
                    authority=self.authority,
                    token_cache=self._token_cache,
                    enable_broker_on_windows=enable_broker,
                )
            except Exception as ex:
                if enable_broker:
                    self.logger.warning(
                        "Failed to initialize MSAL with WAM broker, falling back to standard client: %s",
                        ex,
                    )
                    self._app = msal.PublicClientApplication(
                        client_id=self.client_id,
                        authority=self.authority,
                        token_cache=self._token_cache,
                        enable_broker_on_windows=False,
                    )
                else:
                    raise

        return self._app

    def _save_cache(self) -> None:
        """Persists the token cache to disk if changes occurred."""
        if self._token_cache and getattr(self._token_cache, "has_state_changed", False):
            try:
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                self.cache_path.write_text(self._token_cache.serialize(), encoding="utf-8")
            except Exception as ex:
                self.logger.warning("Failed to persist MSAL token cache: %s", ex)

    def acquire_token(self, interactive: bool = True) -> Tuple[Optional[str], Optional[str]]:
        """Acquires a Microsoft Graph access token silently or via Windows SSO / interactive prompt.

        Returns:
            Tuple[Optional[str], Optional[str]]: (access_token, error_message)
        """
        if self._access_token:
            return self._access_token, None

        if self._failed_previously and not interactive:
            return None, self._failure_reason

        if not self.is_available():
            return None, "MSAL is not available or M365 SSO is disabled in configuration."

        try:
            app = self._get_app()
        except Exception as ex:
            self._failed_previously = True
            self._failure_reason = str(ex)
            return None, str(ex)

        # 1. Attempt silent token acquisition from cache or broker
        try:
            accounts = app.get_accounts()
            if accounts:
                result = app.acquire_token_silent(self.scopes, account=accounts[0])
                if result and "access_token" in result:
                    self._access_token = result["access_token"]
                    self._save_cache()
                    return self._access_token, None
        except Exception as ex:
            self.logger.debug("Silent token acquisition encountered: %s", ex)

        # If silent acquisition failed and interactive is not permitted, return error
        if not interactive:
            self._failed_previously = True
            self._failure_reason = "No cached account or token found; interactive authentication required."
            return None, self._failure_reason

        # 2. Interactive acquisition via Windows SSO Token Broker (WAM)
        try:
            print("\n  [SSO] Connecting to Microsoft 365 via Windows SSO Token Broker...")
            print("        A Windows account prompt may appear to confirm sign-in without needing a password.")

            parent_window = None
            if self.is_broker_supported() and hasattr(msal.PublicClientApplication, "CONSOLE_WINDOW_HANDLE"):
                parent_window = msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE

            result = app.acquire_token_interactive(
                scopes=self.scopes,
                prompt="select_account",
                parent_window_handle=parent_window,
            )
        except Exception as ex:
            self._failed_previously = True
            self._failure_reason = f"Interactive authentication exception: {ex}"
            return None, self._failure_reason

        if result and "access_token" in result:
            self._access_token = result["access_token"]
            self._save_cache()
            user_display = (
                result.get("id_token_claims", {}).get("preferred_username")
                or (accounts[0].get("username") if accounts else "Windows SSO Account")
            )
            print(f"  [SSO] Authenticated successfully as: {user_display}\n")
            return self._access_token, None
        else:
            err_msg = (
                result.get("error_description", result.get("error", "Unknown authentication error"))
                if result
                else "No response from authentication broker"
            )
            self._failed_previously = True
            self._failure_reason = f"Authentication error: {err_msg}"
            return None, self._failure_reason

    def create_draft(
        self,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        text_body: str = "",
        attachments: Optional[List[Path]] = None,
    ) -> Dict[str, Any]:
        """Creates a draft message in Microsoft 365 via Microsoft Graph API.

        Returns:
            Dict with keys: success (bool), method ("graph"), message_id (str|None), error (str|None)
        """
        token, err = self.acquire_token(interactive=not self._failed_previously)
        if not token:
            return {
                "success": False,
                "method": "graph",
                "message_id": None,
                "error": err or "Failed to obtain Microsoft Graph access token",
            }

        message_payload: Dict[str, Any] = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body,
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient_email,
                        "name": recipient_name or recipient_email,
                    }
                }
            ],
            "isDraft": True,
        }

        if getattr(self.config, "email_cc", None):
            message_payload["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in self.config.email_cc
            ]
        if getattr(self.config, "email_bcc", None):
            message_payload["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in self.config.email_bcc
            ]

        # Attachments
        if attachments:
            file_attachments = []
            for att in attachments:
                if att.exists():
                    ctype, _ = mimetypes.guess_type(str(att))
                    if not ctype:
                        ctype = "application/octet-stream"
                    b64_content = base64.b64encode(att.read_bytes()).decode("utf-8")
                    file_attachments.append({
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": att.name,
                        "contentType": ctype,
                        "contentBytes": b64_content,
                    })
            if file_attachments:
                message_payload["attachments"] = file_attachments

        url = "https://graph.microsoft.com/v1.0/me/messages"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        try:
            resp = requests.post(url, json=message_payload, headers=headers, timeout=30)
            if resp.status_code == 201:
                msg_data = resp.json()
                return {
                    "success": True,
                    "method": "graph",
                    "message_id": msg_data.get("id"),
                    "error": None,
                }
            elif resp.status_code == 401:
                # Token might have expired; invalidate token and retry once
                self._access_token = None
                new_token, new_err = self.acquire_token(interactive=False)
                if new_token:
                    headers["Authorization"] = f"Bearer {new_token}"
                    retry_resp = requests.post(url, json=message_payload, headers=headers, timeout=30)
                    if retry_resp.status_code == 201:
                        return {
                            "success": True,
                            "method": "graph",
                            "message_id": retry_resp.json().get("id"),
                            "error": None,
                        }
                return {
                    "success": False,
                    "method": "graph",
                    "message_id": None,
                    "error": f"Graph API unauthorized (401): {resp.text}",
                }
            else:
                return {
                    "success": False,
                    "method": "graph",
                    "message_id": None,
                    "error": f"Graph API HTTP {resp.status_code}: {resp.text}",
                }
        except Exception as req_ex:
            return {
                "success": False,
                "method": "graph",
                "message_id": None,
                "error": f"Graph API request error: {req_ex}",
            }

    def close(self) -> None:
        """Clean up and save cache."""
        self._save_cache()
