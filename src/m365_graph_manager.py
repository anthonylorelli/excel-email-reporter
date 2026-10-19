"""Microsoft 365 Graph Draft Manager: populates draft emails directly into M365 via Graph API and Windows SSO Token Broker."""

import base64
import json
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

# Microsoft Office first-party FOCI client ID (provisioned in all M365 organizational tenants and pre-authorized for Graph Mail.ReadWrite via WAM)
DEFAULT_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
# Microsoft Power Query for Excel first-party client ID (pre-authorized for Graph Mail.ReadWrite with http://localhost redirect URI for browser fallback)
BROWSER_FALLBACK_CLIENT_ID = "a672d62c-fc7b-4e81-a576-e60dc46e951d"
DEFAULT_SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite"]
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/organizations"

# New Outlook for Windows pre-authorized WAM client ID (supports personal Microsoft accounts & Outlook.com)
OUTLOOK_WAM_CLIENT_ID = "5d661950-3475-41cd-a2c3-d671a3162bc1"
OUTLOOK_WAM_SCOPES = ["https://outlook.office.com/.default"]
OUTLOOK_OWA_SERVICE_URL = "https://outlook.office.com/owa/service.svc"

MSA_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad"
PERSONAL_DOMAINS = ("@live.com", "@outlook.com", "@hotmail.com", "@msn.com", "@passport.com")


class M365GraphDraftManager:
    """Manages Microsoft 365 & Outlook.com authentication via Windows SSO Token Broker (WAM)

    and creates draft messages in the user's mailbox via Microsoft Graph API or Outlook OWA service.
    """

    def __init__(self, config: Optional[Any] = None):
        self.config = config
        self.logger = logging.getLogger(__name__)

        raw_account = os.getenv(
            "M365_ACCOUNT",
            getattr(config, "m365_account", "work") if config else "work",
        ).strip()
        self.m365_account = raw_account or "work"
        self.login_hint: Optional[str] = None

        acct_lower = self.m365_account.lower()
        if acct_lower in ("personal", "consumers", "consumer", "msa"):
            self.account_mode = "personal"
        elif "@" in self.m365_account:
            self.login_hint = self.m365_account
            if acct_lower.endswith(PERSONAL_DOMAINS):
                self.account_mode = "personal"
            else:
                self.account_mode = "work"
        else:
            self.account_mode = "work"

        self.client_id = os.getenv(
            "M365_CLIENT_ID",
            getattr(config, "m365_client_id", DEFAULT_CLIENT_ID) if config else DEFAULT_CLIENT_ID,
        )

        default_tenant = "consumers" if self.account_mode == "personal" else "organizations"
        raw_tenant = os.getenv(
            "M365_TENANT_ID",
            getattr(config, "m365_tenant_id", default_tenant) if config else default_tenant,
        )
        if self.account_mode == "personal":
            self.tenant_id = "consumers"
        elif raw_tenant in ("common", "", None):
            self.tenant_id = "organizations"
        else:
            self.tenant_id = raw_tenant

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
        self._use_owa: bool = False
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

    @staticmethod
    def _is_personal_account_entry(account: Dict[str, Any]) -> bool:
        """Returns True if an MSAL account dictionary belongs to a personal Microsoft account."""
        username = (account.get("username") or "").lower()
        home_id = (account.get("home_account_id") or "").lower()
        realm = (account.get("realm") or "").lower()
        if MSA_TENANT_ID in home_id or realm in (MSA_TENANT_ID, "consumers"):
            return True
        if username.endswith(PERSONAL_DOMAINS):
            return True
        return False

    def _select_cached_account(self, accounts: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        """Selects a cached MSAL account matching the configured account_mode and login_hint."""
        if not accounts:
            return None
        for acc in accounts:
            uname = (acc.get("username") or "").lower()
            if self.login_hint and uname != self.login_hint.lower():
                continue
            # When home_account_id or domain clearly identifies personal vs work, filter accordingly
            is_personal = self._is_personal_account_entry(acc)
            if self.account_mode == "personal" and not is_personal:
                continue
            if self.account_mode == "work" and is_personal and not self.login_hint:
                continue
            return acc
        return None

    def _get_app(self, enable_broker: Optional[bool] = None) -> Any:
        """Retrieves or instantiates the MSAL PublicClientApplication."""
        if self._app is None or enable_broker is not None:
            if msal is None:
                raise RuntimeError("MSAL library is not installed.")

            if self._token_cache is None:
                self._token_cache = msal.SerializableTokenCache()
                if self.cache_path.exists():
                    try:
                        self._token_cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
                    except Exception as ex:
                        self.logger.warning("Failed to load MSAL token cache: %s", ex)

            if enable_broker is None:
                use_broker_env = os.getenv("M365_USE_BROKER")
                if use_broker_env is not None:
                    enable_broker = use_broker_env.lower() in ("true", "1", "yes") and self.is_broker_supported()
                else:
                    enable_broker = self.is_broker_supported()

            authority = self.authority
            client_id = self.client_id
            if enable_broker is False and client_id == DEFAULT_CLIENT_ID:
                client_id = BROWSER_FALLBACK_CLIENT_ID
                if authority.endswith("/common"):
                    authority = "https://login.microsoftonline.com/organizations"

            try:
                self._app = msal.PublicClientApplication(
                    client_id=client_id,
                    authority=authority,
                    token_cache=self._token_cache,
                    enable_broker_on_windows=enable_broker,
                )
            except Exception as ex:
                if enable_broker:
                    self.logger.warning(
                        "Failed to initialize MSAL with WAM broker, falling back to standard client: %s",
                        ex,
                    )
                    fallback_id = BROWSER_FALLBACK_CLIENT_ID if self.client_id == DEFAULT_CLIENT_ID else self.client_id
                    self._app = msal.PublicClientApplication(
                        client_id=fallback_id,
                        authority=authority,
                        token_cache=self._token_cache,
                        enable_broker_on_windows=False,
                    )
                else:
                    raise

        return self._app

    def _acquire_outlook_wam_token(
        self,
        prompt: Optional[str] = None,
        authority: Optional[str] = None,
    ) -> Optional[Dict[str, Any]]:
        """Acquires a pre-authorized Outlook token via Windows WAM for personal Microsoft accounts."""
        if not self.is_broker_supported() or msal is None:
            return None
        target_authority = authority or (
            "https://login.microsoftonline.com/consumers"
            if self.account_mode == "personal"
            else "https://login.microsoftonline.com/organizations"
        )
        try:
            parent_window = getattr(msal.PublicClientApplication, "CONSOLE_WINDOW_HANDLE", None)
            outlook_app = msal.PublicClientApplication(
                client_id=OUTLOOK_WAM_CLIENT_ID,
                authority=target_authority,
                token_cache=self._token_cache,
                enable_broker_on_windows=True,
            )
            matched_acc = self._select_cached_account(outlook_app.get_accounts())
            if matched_acc and prompt == "none":
                res = outlook_app.acquire_token_silent(OUTLOOK_WAM_SCOPES, account=matched_acc)
                if res and "access_token" in res:
                    return res
            interactive_kwargs: Dict[str, Any] = {
                "scopes": OUTLOOK_WAM_SCOPES,
                "prompt": prompt,
                "parent_window_handle": parent_window,
            }
            if self.login_hint:
                interactive_kwargs["login_hint"] = self.login_hint
            res = outlook_app.acquire_token_interactive(**interactive_kwargs)
            if isinstance(res, dict) and "access_token" in res:
                claims = res.get("id_token_claims") or {}
                uname = (claims.get("preferred_username") or "").lower()
                tid = (claims.get("tid") or "").lower()
                is_personal = tid == MSA_TENANT_ID or uname.endswith(PERSONAL_DOMAINS)
                if self.account_mode == "work" and is_personal and not self.login_hint:
                    return None
                if self.account_mode == "personal" and uname and not is_personal:
                    return None
                return res
        except Exception as ex:
            self.logger.debug("Outlook WAM token acquisition encountered: %s", ex)
        return None

    def _save_cache(self) -> None:
        """Persists the token cache to disk if changes occurred."""
        if self._token_cache and getattr(self._token_cache, "has_state_changed", False):
            try:
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                self.cache_path.write_text(self._token_cache.serialize(), encoding="utf-8")
            except Exception as ex:
                self.logger.warning("Failed to persist MSAL token cache: %s", ex)

    def acquire_token(self, interactive: bool = True) -> Tuple[Optional[str], Optional[str]]:
        """Acquires a Microsoft 365 / Outlook access token silently or via Windows SSO / browser OAuth.

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

        broker_active = getattr(app, "_enable_broker", False) is True

        # 1. Attempt silent token acquisition from cache
        matched_account = None
        try:
            accounts = app.get_accounts()
            matched_account = self._select_cached_account(accounts)
            if matched_account:
                result = app.acquire_token_silent(self.scopes, account=matched_account)
                if result and "access_token" in result:
                    self._access_token = result["access_token"]
                    self._use_owa = self._access_token.startswith("Ew")
                    self._save_cache()
                    return self._access_token, None
        except Exception as ex:
            self.logger.debug("Silent token acquisition encountered: %s", ex)

        # 1b. If using real WAM broker in personal mode, check silent WAM SSO
        if broker_active and self.account_mode == "personal":
            wam_silent = self._acquire_outlook_wam_token(prompt="none")
            if wam_silent and "access_token" in wam_silent:
                self._access_token = wam_silent["access_token"]
                self._use_owa = True
                self._save_cache()
                user_display = (
                    wam_silent.get("id_token_claims", {}).get("preferred_username")
                    or "Microsoft Account"
                )
                print(f"\n  [SSO] Authenticated via Windows Outlook SSO as: {user_display}\n")
                return self._access_token, None

        # If silent acquisition failed and interactive is not permitted, return error
        if not interactive:
            self._failed_previously = True
            self._failure_reason = "No cached account or token found; interactive authentication required."
            return None, self._failure_reason

        # 2. Interactive acquisition
        result = None
        acct_label = "personal Microsoft" if self.account_mode == "personal" else "organizational Microsoft 365"
        if self.login_hint:
            acct_label = f"{acct_label} ({self.login_hint})"

        if broker_active and self.account_mode == "personal":
            print(f"\n  [SSO] Connecting to {acct_label} account...")
            wam_res = self._acquire_outlook_wam_token(prompt=None)
            if wam_res and "access_token" in wam_res:
                self._access_token = wam_res["access_token"]
                self._use_owa = True
                self._save_cache()
                user_display = (
                    wam_res.get("id_token_claims", {}).get("preferred_username")
                    or self.login_hint
                    or "Personal Microsoft Account"
                )
                print(f"  [SSO] Authenticated successfully as: {user_display}\n")
                return self._access_token, None

        try:
            print(f"\n  [SSO] Connecting to {acct_label} account...")
            print("        Please select your account in the sign-in prompt.")

            parent_window = None
            if self.is_broker_supported() and hasattr(msal.PublicClientApplication, "CONSOLE_WINDOW_HANDLE"):
                parent_window = msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE

            interactive_kwargs: Dict[str, Any] = {
                "scopes": self.scopes,
                "prompt": "select_account",
                "parent_window_handle": parent_window,
            }
            if self.login_hint:
                interactive_kwargs["login_hint"] = self.login_hint

            result = app.acquire_token_interactive(**interactive_kwargs)
        except Exception as ex:
            if not broker_active:
                self._failed_previously = True
                self._failure_reason = f"Interactive authentication exception: {ex}"
                return None, self._failure_reason
            self.logger.debug("WAM broker interactive exception (%s), trying fallbacks.", ex)

        # 3. If WAM broker failed, fall back to browser OAuth (using pre-authorized localhost client ID)
        if broker_active and (not result or "access_token" not in result):
            err_code = result.get("error", "") if isinstance(result, dict) else ""
            err_desc = result.get("error_description", "") if isinstance(result, dict) else ""
            if err_code == "broker_error" or "ApiContractViolation" in err_desc or result is None:
                if self.account_mode == "personal":
                    wam_interactive = self._acquire_outlook_wam_token(prompt=None)
                    if wam_interactive and "access_token" in wam_interactive:
                        self._access_token = wam_interactive["access_token"]
                        self._use_owa = True
                        self._save_cache()
                        user_display = (
                            wam_interactive.get("id_token_claims", {}).get("preferred_username")
                            or self.login_hint
                            or "Personal Microsoft Account"
                        )
                        print(f"  [SSO] Authenticated successfully as: {user_display}\n")
                        return self._access_token, None

                print("  [SSO] Windows WAM broker unavailable for this account; opening browser sign-in...")
                try:
                    browser_app = self._get_app(enable_broker=False)
                    browser_kwargs: Dict[str, Any] = {
                        "scopes": self.scopes,
                        "prompt": "select_account",
                    }
                    if self.login_hint:
                        browser_kwargs["login_hint"] = self.login_hint
                    result = browser_app.acquire_token_interactive(**browser_kwargs)
                except Exception as browser_ex:
                    self._failed_previously = True
                    self._failure_reason = f"Browser OAuth fallback exception: {browser_ex}"
                    return None, self._failure_reason

        if result and "access_token" in result:
            self._access_token = result["access_token"]
            self._use_owa = self._access_token.startswith("Ew")
            self._save_cache()
            user_display = (
                result.get("id_token_claims", {}).get("preferred_username")
                or (matched_account.get("username") if matched_account else "Microsoft 365 Account")
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

    def _create_draft_via_owa(
        self,
        token: str,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        attachments: Optional[List[Path]] = None,
    ) -> Dict[str, Any]:
        """Creates a draft message in Outlook.com / Exchange Online via OWA service.svc CreateItem."""
        item_dict: Dict[str, Any] = {
            "__type": "Message:#Exchange",
            "Subject": subject,
            "Body": {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "HTML",
                "Value": html_body,
            },
            "ToRecipients": [
                {
                    "__type": "EmailAddressWrapper:#Exchange",
                    "EmailAddress": recipient_email,
                    "Name": recipient_name or recipient_email,
                }
            ],
        }

        if getattr(self.config, "email_cc", None):
            item_dict["CcRecipients"] = [
                {"__type": "EmailAddressWrapper:#Exchange", "EmailAddress": addr}
                for addr in self.config.email_cc
            ]
        if getattr(self.config, "email_bcc", None):
            item_dict["BccRecipients"] = [
                {"__type": "EmailAddressWrapper:#Exchange", "EmailAddress": addr}
                for addr in self.config.email_bcc
            ]

        create_payload = {
            "__type": "CreateItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
            },
            "Body": {
                "__type": "CreateItemRequest:#Exchange",
                "MessageDisposition": "SaveOnly",
                "Items": [item_dict],
            },
        }

        post_str = json.dumps(create_payload)
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
            "Action": "CreateItem",
            "X-OWA-UrlPostData": post_str,
        }

        try:
            resp = requests.post(
                f"{OUTLOOK_OWA_SERVICE_URL}?action=CreateItem&app=Mail",
                data=post_str.encode("utf-8"),
                headers=headers,
                timeout=30,
            )
            if resp.status_code != 200:
                return {
                    "success": False,
                    "method": "graph",
                    "message_id": None,
                    "error": f"Outlook OWA HTTP {resp.status_code}: {resp.text[:300]}",
                }

            body = resp.json()
            resp_items = (
                body.get("Body", {})
                .get("ResponseMessages", {})
                .get("Items", [])
            )
            if not resp_items or resp_items[0].get("ResponseClass") != "Success":
                msg_text = resp_items[0].get("MessageText", "Unknown OWA error") if resp_items else "Empty OWA response"
                return {
                    "success": False,
                    "method": "graph",
                    "message_id": None,
                    "error": f"Outlook OWA CreateItem error: {msg_text}",
                }

            created_items = resp_items[0].get("Items", [])
            item_id_obj = created_items[0].get("ItemId", {}) if created_items else {}
            msg_id = item_id_obj.get("Id")

            # Upload attachments if present
            if attachments and item_id_obj:
                file_atts = []
                for att in attachments:
                    if att.exists():
                        ctype, _ = mimetypes.guess_type(str(att))
                        if not ctype:
                            ctype = "application/octet-stream"
                        b64_content = base64.b64encode(att.read_bytes()).decode("utf-8")
                        file_atts.append({
                            "__type": "FileAttachment:#Exchange",
                            "Name": att.name,
                            "ContentType": ctype,
                            "Content": b64_content,
                            "IsInline": False,
                        })
                if file_atts:
                    att_payload = {
                        "__type": "CreateAttachmentJsonRequest:#Exchange",
                        "Header": {
                            "__type": "JsonRequestHeaders:#Exchange",
                            "RequestServerVersion": "V2018_01_08",
                        },
                        "Body": {
                            "__type": "CreateAttachmentRequest:#Exchange",
                            "ParentItemId": item_id_obj,
                            "Attachments": file_atts,
                        },
                    }
                    att_str = json.dumps(att_payload)
                    att_headers = {
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json; charset=utf-8",
                        "Action": "CreateAttachment",
                        "X-OWA-UrlPostData": att_str,
                    }
                    requests.post(
                        f"{OUTLOOK_OWA_SERVICE_URL}?action=CreateAttachment&app=Mail",
                        data=att_str.encode("utf-8"),
                        headers=att_headers,
                        timeout=30,
                    )

            return {
                "success": True,
                "method": "graph",
                "message_id": msg_id,
                "error": None,
            }
        except Exception as owa_ex:
            return {
                "success": False,
                "method": "graph",
                "message_id": None,
                "error": f"Outlook OWA request error: {owa_ex}",
            }

    def create_draft(
        self,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        text_body: str = "",
        attachments: Optional[List[Path]] = None,
    ) -> Dict[str, Any]:
        """Creates a draft message in Microsoft 365 / Outlook.com via Graph API or OWA service.

        Returns:
            Dict with keys: success (bool), method ("graph"), message_id (str|None), error (str|None)
        """
        token, err = self.acquire_token(interactive=not self._failed_previously)
        if not token:
            return {
                "success": False,
                "method": "graph",
                "message_id": None,
                "error": err or "Failed to obtain Microsoft 365 access token",
            }

        if self._use_owa or token.startswith("Ew"):
            return self._create_draft_via_owa(
                token=token,
                recipient_email=recipient_email,
                recipient_name=recipient_name,
                subject=subject,
                html_body=html_body,
                attachments=attachments,
            )

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
