"""Configuration loader and manager for excel-email-reporter."""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import yaml
from dotenv import load_dotenv


class Config:
    """Manages application configuration combining YAML and environment variables."""

    def __init__(self, config_data: Dict[str, Any], env_prefix: str = ""):
        self.raw = config_data
        self.project = config_data.get("project", {})
        self.data_sources = config_data.get("data_sources", {})
        self.summary_metrics = config_data.get("summary_metrics", [])
        self.breakdown = config_data.get("breakdown", {})
        self.email = config_data.get("email", {})

        # Resolve SMTP settings with environment variables having higher precedence
        self.smtp_host = os.getenv("SMTP_HOST", self.email.get("smtp_host", "localhost"))
        self.smtp_port = int(os.getenv("SMTP_PORT", str(self.email.get("smtp_port", 587))))
        self.smtp_user = os.getenv("SMTP_USERNAME", self.email.get("smtp_user", ""))
        self.smtp_password = os.getenv("SMTP_PASSWORD", self.email.get("smtp_password", ""))
        self.smtp_use_tls = os.getenv(
            "SMTP_USE_TLS", str(self.email.get("smtp_use_tls", True))
        ).lower() in ("true", "1", "yes")
        self.smtp_use_ssl = os.getenv(
            "SMTP_USE_SSL", str(self.email.get("smtp_use_ssl", False))
        ).lower() in ("true", "1", "yes")

        # Email addressing
        self.email_sender = os.getenv(
            "EMAIL_SENDER", self.email.get("sender", "reports@example.com")
        )
        recipients_env = os.getenv("EMAIL_RECIPIENTS")
        if recipients_env:
            self.email_recipients = [r.strip() for r in recipients_env.split(",") if r.strip()]
        else:
            self.email_recipients = self.email.get("recipients", [])

        cc_env = os.getenv("EMAIL_CC")
        self.email_cc = [r.strip() for r in cc_env.split(",") if r.strip()] if cc_env else self.email.get("cc", [])
        
        bcc_env = os.getenv("EMAIL_BCC")
        self.email_bcc = [r.strip() for r in bcc_env.split(",") if r.strip()] if bcc_env else self.email.get("bcc", [])

        self.grade_calculation = config_data.get("grade_calculation", {})

        # Subject & templates
        self.audience = self.email.get("audience", "general")
        self.email_subject = self.email.get("subject", "Excel Automated Report - {date}")
        self.subject_template = self.email.get("subject_template", "Progress Report: {student_name} - {course_name}")
        self.template_html = self.email.get("template_html", "templates/report_template.html")
        self.template_text = self.email.get("template_text", "templates/report_template.txt")
        self.draft_target = os.getenv("DRAFT_TARGET", self.email.get("draft_target", "eml")).lower()
        self.outlook_method = os.getenv("OUTLOOK_METHOD", self.email.get("outlook_method", "auto")).lower()
        self.m365_account = os.getenv("M365_ACCOUNT", self.email.get("m365_account", "work")).strip()
        self.m365_client_id = os.getenv("M365_CLIENT_ID", self.email.get("m365_client_id", "14d82eec-204b-449e-b7d6-3e00e84b80e4"))
        self.m365_tenant_id = os.getenv("M365_TENANT_ID", self.email.get("m365_tenant_id", "organizations"))
        self.m365_sso = os.getenv("M365_SSO", str(self.email.get("m365_sso", True))).lower() in ("true", "1", "yes")
        self.imap_host = os.getenv("IMAP_HOST", self.email.get("imap_host", "outlook.office365.com"))
        self.imap_port = int(os.getenv("IMAP_PORT", str(self.email.get("imap_port", 993))))
        self.dashboard_template = self.email.get("dashboard_template", "templates/drafts_dashboard.html")
        self.output_drafts_dir = self.email.get("output_drafts_dir", "output/drafts")
        self.attachments = self.email.get("attachments", {})


def get_bundled_defaults_dir() -> Path:
    """Returns the path to the bundled defaults directory inside the package."""
    return Path(__file__).resolve().parent / "defaults"


def resolve_resource_path(relative_path: str) -> Optional[Path]:
    """Resolves a config or template path from CWD, repo root, or bundled package defaults."""
    p = Path(relative_path)
    if p.exists():
        return p

    repo_fallback = Path(__file__).resolve().parent.parent / relative_path
    if repo_fallback.exists():
        return repo_fallback

    bundled_fallback = get_bundled_defaults_dir() / relative_path
    if bundled_fallback.exists():
        return bundled_fallback

    return None


def load_config(config_path: Optional[str] = None, env_path: Optional[str] = None) -> Config:
    """Load configuration from YAML file and .env file.
    
    Args:
        config_path: Path to YAML config file. If None, looks for config/config.yaml or config.example.yaml.
        env_path: Path to .env file. If None, uses default .env resolution.
        
    Returns:
        Config instance.
    """
    if env_path and Path(env_path).exists():
        load_dotenv(env_path)
    else:
        load_dotenv()

    # Search candidates for config
    candidates = [
        config_path,
        "config/config.yaml",
        "config/config.yml",
        "config/grades_config.yaml",
        "config/config.example.yaml",
    ]

    selected_path: Optional[Path] = None
    for cand in candidates:
        if cand:
            resolved = resolve_resource_path(cand)
            if resolved is not None:
                selected_path = resolved
                break

    data: Dict[str, Any] = {}
    if selected_path:
        with open(selected_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

    return Config(data)
