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
        self.dashboard_template = self.email.get("dashboard_template", "templates/drafts_dashboard.html")
        self.output_drafts_dir = self.email.get("output_drafts_dir", "output/drafts")
        self.attachments = self.email.get("attachments", {})


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
        "config/config.example.yaml",
    ]

    selected_path = None
    for cand in candidates:
        if cand and Path(cand).exists():
            selected_path = cand
            break

    data: Dict[str, Any] = {}
    if selected_path:
        with open(selected_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

    return Config(data)
