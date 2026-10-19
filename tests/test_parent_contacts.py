"""Unit tests for parent contact loading, matching, and CC draft distribution."""

import email
import json
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch
import pandas as pd
import pytest

from src.config import Config
from src.email_sender import EmailSender
from src.excel_extractor import ExcelExtractor
from src.grade_calculator import GradeCalculator
from src.m365_graph_manager import M365GraphDraftManager
from src.outlook_manager import OutlookDraftManager


def test_extractor_load_parent_contacts_single_and_multiple_sheets(tmp_path):
    """Verifies that load_parent_contacts extracts student-parent mappings across multiple sheets and files."""
    file1 = tmp_path / "contacts_course1.xlsx"
    file2 = tmp_path / "contacts_course2.xlsx"

    # Sheet 1: Single parent email
    df1 = pd.DataFrame([
        {
            "Registrant First Name": "Alice",
            "Registrant Last Name": "Walker",
            "Email Address": "alice_mom@example.com",
            "Primary First Name": "Mary",
            "Primary Last Name": "Walker",
        },
        {
            # Comma / Semicolon separated multiple parents
            "Registrant First Name": "Bob",
            "Registrant Last Name": "Smith",
            "Email Address": "bob_mom@example.com; bob_dad@example.com, bob_guardian@example.com",
            "Primary First Name": "Robert",
            "Primary Last Name": "Smith",
        },
        {
            # Missing / NaN student name or email
            "Registrant First Name": None,
            "Registrant Last Name": "Invalid",
            "Email Address": "invalid@example.com",
        },
    ])

    with pd.ExcelWriter(file1, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Roster1", index=False)

    # File 2: Sheet with alternative column names (Student First Name / Parent Email)
    df2 = pd.DataFrame([
        {
            "Student First Name": "Charlie",
            "Student Last Name": "Brown",
            "Parent Email": "snoopy@example.com",
        }
    ])
    with pd.ExcelWriter(file2, engine="openpyxl") as writer:
        df2.to_excel(writer, sheet_name="Roster2", index=False)

    extractor = ExcelExtractor({})
    contacts = extractor.load_parent_contacts(contacts_dir=str(tmp_path))

    assert ("alice", "walker") in contacts
    assert contacts[("alice", "walker")] == ["alice_mom@example.com"]

    assert ("bob", "smith") in contacts
    assert contacts[("bob", "smith")] == [
        "bob_mom@example.com",
        "bob_dad@example.com",
        "bob_guardian@example.com",
    ]

    assert ("charlie", "brown") in contacts
    assert contacts[("charlie", "brown")] == ["snoopy@example.com"]


def test_grade_calculator_associates_parent_emails():
    """Verifies that GradeCalculator attaches parent emails to StudentGradeProfile based on name matching."""
    weights = {"Homework": 0.5, "Essays": 0.5}
    calculator = GradeCalculator(weights)

    data = [
        # Student 1: Matches parent contact
        {
            "Full Name": "Alice Walker",
            "First Name": "Alice",
            "Last Name": "Walker",
            "Email Address": "alice.student@school.edu",
            "Assignments": "HW 1",
            "Grading Category": "Homework",
            "Points": 10.0,
            "Max Points": 10.0,
            "Status": "Returned",
        },
        # Student 2: Case insensitivity check (uppercase in grade sheet)
        {
            "Full Name": "BOB SMITH",
            "First Name": "BOB",
            "Last Name": "SMITH",
            "Email Address": "bob.student@school.edu",
            "Assignments": "HW 1",
            "Grading Category": "Homework",
            "Points": 9.0,
            "Max Points": 10.0,
            "Status": "Returned",
        },
        # Student 3: No parent contact registered
        {
            "Full Name": "Dan NoParent",
            "First Name": "Dan",
            "Last Name": "NoParent",
            "Email Address": "dan@school.edu",
            "Assignments": "HW 1",
            "Grading Category": "Homework",
            "Points": 8.0,
            "Max Points": 10.0,
            "Status": "Returned",
        },
    ]

    df = pd.DataFrame(data)
    parent_contacts = {
        ("alice", "walker"): ["parent_alice@example.com"],
        ("bob", "smith"): ["mom_bob@example.com", "dad_bob@example.com"],
    }

    profiles = calculator.process_records(df, course_name="Lit 101", parent_contacts=parent_contacts)
    profile_map = {p.full_name: p for p in profiles}

    assert profile_map["Alice Walker"].parent_emails == ["parent_alice@example.com"]
    assert profile_map["BOB SMITH"].parent_emails == ["mom_bob@example.com", "dad_bob@example.com"]
    assert profile_map["Dan NoParent"].parent_emails == []


def test_email_sender_eml_cc_headers(tmp_path):
    """Verifies that EmailSender writes Cc headers merging global config CC and per-draft parent emails."""
    config = Config({
        "email": {
            "sender": "teacher@school.org",
            "cc": ["principal@school.org"],
        }
    })
    sender = EmailSender(config)

    eml_file = tmp_path / "test_student_draft.eml"
    sender.create_draft_eml(
        recipient_email="student@school.org",
        recipient_name="Student Name",
        subject="Progress Report",
        html_body="<p>Grade Report</p>",
        text_body="Grade Report",
        output_eml_path=eml_file,
        cc_emails=["parent1@family.com", "principal@school.org", "parent2@family.com"],  # principal is duplicate
    )

    assert eml_file.exists()
    msg = email.message_from_bytes(eml_file.read_bytes())

    assert msg["To"] == '"Student Name" <student@school.org>'
    assert msg["Subject"] == "Progress Report"
    assert msg["X-Unsent"] == "1"

    cc_header = msg["Cc"]
    # De-duplicated and contains all recipients
    assert "principal@school.org" in cc_header
    assert "parent1@family.com" in cc_header
    assert "parent2@family.com" in cc_header
    # Should only contain principal@school.org once
    assert cc_header.count("principal@school.org") == 1


def test_outlook_manager_com_cc():
    """Verifies that Outlook COM draft creation populates mail.CC with merged parent CCs."""
    config = Config({"email": {"sender": "teacher@school.org", "cc": ["admin@school.org"]}})
    manager = OutlookDraftManager(config)

    mock_mail = MagicMock()
    mock_app = MagicMock()
    mock_app.CreateItem.return_value = mock_mail

    with patch.object(manager, "get_or_create_com_app", return_value=mock_app):
        manager.create_draft_via_com(
            recipient_email="student@school.org",
            recipient_name="Alice Student",
            subject="Progress Report",
            html_body="<p>Grade: 95%</p>",
            cc_emails=["parent1@home.org", "parent2@home.org"],
        )

        assert mock_mail.CC == "admin@school.org; parent1@home.org; parent2@home.org"


def test_outlook_manager_imap_cc():
    """Verifies that Outlook IMAP draft creation adds Cc header to MIME message."""
    config = Config({"email": {"sender": "teacher@school.org", "cc": ["dept_head@school.org"]}})
    manager = OutlookDraftManager(config)
    manager.imap_user = "teacher@school.org"
    manager.imap_password = "password123"

    mock_imap = MagicMock()
    mock_imap.append.return_value = ("OK", [b"APPEND completed"])

    with patch.object(manager, "_get_imap_client", return_value=mock_imap):
        manager._imap_drafts_folder = "Drafts"
        manager.create_draft_via_imap(
            recipient_email="student@school.org",
            recipient_name="Bob Student",
            subject="Weekly Progress",
            html_body="<p>Grade: 88%</p>",
            text_body="Grade: 88%",
            cc_emails=["parent@home.org"],
        )

        args, _ = mock_imap.append.call_args
        raw_mime = args[3]
        parsed = email.message_from_bytes(raw_mime)
        assert parsed["Cc"] == "dept_head@school.org, parent@home.org"


def test_m365_graph_manager_create_draft_graph_cc():
    """Verifies that Microsoft Graph API create_draft sends ccRecipients in the payload."""
    config = Config({"email": {"cc": ["counselor@school.org"]}})
    manager = M365GraphDraftManager(config)
    manager._access_token = "valid_graph_token"
    manager._use_owa = False

    mock_resp = MagicMock()
    mock_resp.status_code = 201
    mock_resp.json.return_value = {"id": "msg_graph_cc_123"}

    with patch("requests.post", return_value=mock_resp) as mock_post:
        res = manager.create_draft(
            recipient_email="student@school.org",
            recipient_name="Student",
            subject="Report",
            html_body="<p>Body</p>",
            cc_emails=["parent@home.org"],
        )

        assert res["success"] is True
        assert res["message_id"] == "msg_graph_cc_123"

        _, kwargs = mock_post.call_args
        payload = kwargs["json"]
        assert "ccRecipients" in payload
        cc_addrs = [r["emailAddress"]["address"] for r in payload["ccRecipients"]]
        assert cc_addrs == ["counselor@school.org", "parent@home.org"]


def test_m365_graph_manager_create_draft_owa_cc():
    """Verifies that Outlook OWA create_draft sends CcRecipients in the payload."""
    config = Config({"email": {"cc": ["counselor@school.org"]}})
    manager = M365GraphDraftManager(config)
    manager._access_token = "Ew_personal_token"
    manager._use_owa = True

    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {
        "Body": {
            "ResponseMessages": {
                "Items": [
                    {
                        "ResponseClass": "Success",
                        "Items": [{"ItemId": {"Id": "owa_msg_cc_456"}}],
                    }
                ]
            }
        }
    }

    with patch("requests.post", return_value=mock_resp) as mock_post:
        res = manager.create_draft(
            recipient_email="student@school.org",
            recipient_name="Student",
            subject="Report",
            html_body="<p>Body</p>",
            cc_emails=["parent@home.org"],
        )

        assert res["success"] is True
        assert res["message_id"] == "owa_msg_cc_456"

        _, kwargs = mock_post.call_args
        payload = json.loads(kwargs["data"].decode("utf-8"))
        item = payload["Body"]["Items"][0]
        assert "CcRecipients" in item
        cc_addrs = [r["EmailAddress"] for r in item["CcRecipients"]]
        assert cc_addrs == ["counselor@school.org", "parent@home.org"]
