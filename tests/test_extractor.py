"""Tests for Excel data extractor."""

import tempfile
from pathlib import Path
import pandas as pd
import pytest
from src.excel_extractor import ExcelExtractor


@pytest.fixture
def sample_excel_file():
    """Creates a temporary test Excel workbook."""
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = Path(tmpdir) / "test_sales.xlsx"
        df = pd.DataFrame([
            {"Date": "2026-09-01", "Region": "North", "Revenue": 1500.0, "Status": "Completed"},
            {"Date": "2026-09-02", "Region": "South", "Revenue": 2500.0, "Status": "Completed"},
            {"Date": "2026-09-03", "Region": "North", "Revenue": 800.0, "Status": "Cancelled"},
        ])
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sales_Data", index=False)
        yield file_path


def test_extractor_processes_file(sample_excel_file):
    config = {
        "sheet_name": "Sales_Data",
        "header_row": 1,
        "columns": {
            "date": "Date",
            "region": "Region",
            "revenue": "Revenue",
            "status": "Status",
        },
        "filters": {
            "status": ["Completed"],
        },
    }

    extractor = ExcelExtractor(config)
    df, raw_count = extractor.extract_from_file(sample_excel_file)

    assert raw_count == 3
    assert len(df) == 2  # Cancelled row filtered out
    assert "revenue" in df.columns
    assert "region" in df.columns
    assert df["revenue"].sum() == 4000.0
