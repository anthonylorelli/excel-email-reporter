"""Tests for report generation and metric calculations."""

import pandas as pd
from src.config import Config
from src.report_generator import ReportGenerator


def test_metrics_and_html_generation():
    config_dict = {
        "project": {"name": "Test Sales Report"},
        "summary_metrics": [
            {"id": "total_rev", "label": "Total Revenue", "column": "revenue", "type": "sum", "format": "${:,.2f}"},
            {"id": "avg_rev", "label": "Average Order", "column": "revenue", "type": "mean", "format": "${:,.2f}"},
        ],
        "breakdown": {
            "group_by": "region",
            "metrics": [
                {"column": "revenue", "agg": "sum", "header": "Total Revenue", "format": "${:,.2f}"}
            ]
        },
        "email": {
            "template_html": "templates/report_template.html",
            "template_text": "templates/report_template.txt",
        }
    }
    cfg = Config(config_dict)
    generator = ReportGenerator(cfg)

    df = pd.DataFrame([
        {"date": "2026-09-01", "region": "North", "revenue": 1000.0, "status": "Completed"},
        {"date": "2026-09-02", "region": "South", "revenue": 3000.0, "status": "Completed"},
    ])

    metrics = generator.calculate_metrics(df)
    assert len(metrics) == 2
    assert metrics[0]["formatted_value"] == "$4,000.00"
    assert metrics[1]["formatted_value"] == "$2,000.00"

    breakdown = generator.calculate_breakdown(df)
    assert breakdown["group_by"] == "region"
    assert len(breakdown["rows"]) == 2

    # Verify HTML template rendering
    class MockSummary:
        filename = "test.xlsx"
        included_count = 2

    rendered = generator.render(df, [MockSummary()], title="Test Title")
    assert "Test Title" in rendered["html"]
    assert "$4,000.00" in rendered["html"]
    assert "North" in rendered["html"]
    assert "South" in rendered["html"]
    assert "Test Title" in rendered["text"]


def test_bundled_defaults_exist():
    from src.config import get_bundled_defaults_dir, load_config

    defaults_dir = get_bundled_defaults_dir()
    assert (defaults_dir / "config" / "grades_config.yaml").exists()
    assert (defaults_dir / "config" / "config.example.yaml").exists()
    assert (defaults_dir / "templates" / "student_report.html").exists()
    assert (defaults_dir / "templates" / "student_report.txt").exists()
    assert (defaults_dir / "templates" / "drafts_dashboard.html").exists()

    cfg = load_config(str(defaults_dir / "config" / "grades_config.yaml"))
    assert cfg.project.get("type") == "student_grades"

