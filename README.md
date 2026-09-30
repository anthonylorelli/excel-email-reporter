# Excel Email Reporter (`excel-email-reporter`)

A modular, automated data reporting pipeline that ingests tabular data across a series of Excel spreadsheets (`.xlsx`, `.xls`), performs filtering, normalization, and metric calculations, and delivers beautiful, responsive HTML/plain-text executive reports with attachments via email (SMTP).

---

## Features

- **Student Grade & Progress Reporting**: Calculates accurate running weighted grades from Microsoft Teams Assignment workbooks across categories (Homework, Exams, Teams Discussion, Participation).
- **Missing Work Policy**: Automatically excludes unsubmitted / ungraded assignments until graded, and highlights pending items for students.
- **Individual Email Drafts**: Generates standard RFC 822 `.eml` drafts with `X-Unsent: 1` that open directly into Outlook / Windows Mail in compose mode, plus companion `.html` previews.
- **Master Review Dashboard**: Creates an interactive `index.html` review dashboard allowing instructors to inspect class averages, medians, and every student draft.
- **Batch Spreadsheet Ingestion**: Scans directories for matching Excel workbooks (e.g. `*.xlsx`), ignoring lock/temporary files (`~$*.xlsx`).
- **Flexible Schema & Column Mapping**: Maps arbitrary Excel headers to standardized internal schema fields.
- **Executive KPI Cards & Grouped Breakdowns**: Generates cross-sectional pivot tables and performance statistics.
- **Responsive HTML & Multi-part MIME Email**: Generates email-safe HTML compatible with Outlook, Apple Mail, and Gmail, paired with a plain-text fallback.
- **Automated Attachments**: Generates and attaches consolidated Excel workbooks or attaches original source files.
- **Safe Dry-Run Mode**: Preview HTML reports in your browser and inspect email headers before sending live emails.
- **Automated Scheduling**: Designed for effortless scheduling via Windows Task Scheduler or cron.

---

## Directory Structure

```text
excel-email-reporter/
├── config/
│   └── config.example.yaml        # Spreadsheet mappings, metrics, and report rules
├── templates/
│   ├── report_template.html       # Responsive HTML email template
│   └── report_template.txt        # Plain-text alternative fallback
├── src/
│   ├── __init__.py
│   ├── cli.py                     # Command-line interface and orchestrator
│   ├── config.py                  # YAML and .env configuration manager
│   ├── excel_extractor.py         # Workbook reader, cleaner, and aggregator
│   ├── report_generator.py        # Metric calculator and Jinja2 renderer
│   └── email_sender.py            # SMTP MIME builder and dispatcher
├── samples/
│   ├── generate_sample_data.py    # Generates mock Excel workbooks for testing
│   └── data/                      # Sample workbooks (generated)
├── tests/
│   ├── test_extractor.py          # Extraction and filtering unit tests
│   └── test_report.py             # Metric calculation and rendering tests
├── .env.example                   # SMTP server and credential template
├── .gitignore
├── pyproject.toml
├── requirements.txt
└── README.md
```

---

## Quick Start

### 1. Set Up Python Environment

```powershell
# Navigate to repository
cd C:\Users\antho\source\repos\excel-email-reporter

# Create and activate virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

### 2. Configure Credentials

Copy `.env.example` to `.env` and fill in your email server credentials:

```powershell
Copy-Item .env.example .env
```

Edit `.env`:
```ini
SMTP_HOST=smtp.gmail.com          # or smtp.office365.com
SMTP_PORT=587
SMTP_USE_TLS=true
SMTP_USERNAME=your-account@gmail.com
SMTP_PASSWORD=your-app-password
EMAIL_SENDER="Sales Reports <your-account@gmail.com>"
EMAIL_RECIPIENTS=executive@company.com,team@company.com
```

### 3. Generate Sample Spreadsheets

Bootstrap sample regional sales workbooks to test the full pipeline immediately:

```powershell
python -m src --generate-samples
```

This creates realistic mock spreadsheets under `samples/data/`.

### 4. Run a Dry Run (Generate Preview)

By default, the pipeline runs in **dry-run** mode to avoid sending unwanted emails during testing:

```powershell
python -m src --config config/config.example.yaml
```

This will:
1. Ingest all workbooks in `samples/data/*.xlsx`.
2. Compute the summary metrics and grouped breakdown table.
3. Save the styled HTML email preview to `output/report_preview_<date>.html`.
4. Generate the consolidated Excel attachment in `output/Consolidated_Report_<date>.xlsx`.
5. Display the simulated email payload in your terminal.

You can double-click or open `output/report_preview_<date>.html` in your browser to inspect the visual output!

### 5. Generate Student Grade Reports & Draft Emails

To process student assignment workbooks (e.g. from Microsoft Teams) and generate individual student progress report drafts:

```powershell
python -m src --config config/grades_config.yaml
```

This will:
1. Ingest class gradebooks from `data/grades/*.xlsx` (or path specified in config).
2. Weight grades according to configured categories:
   - **Homework:** 35%
   - **Exams:** 35%
   - **Teams Discussion:** 20%
   - **Attendance/Participation:** 10%
3. Exclude missing/ungraded assignments from the current running score until graded.
4. Generate standard `.eml` draft files in `output/drafts/<Course>/<Student_Name>_grade_report.eml` with `X-Unsent: 1`. Double-clicking any `.eml` file opens Outlook / Windows Mail directly in compose mode with the student's email and personalized report ready to review!
5. Generate companion `.html` previews and a master `output/drafts/index.html` review dashboard for the instructor.

### 6. Send Real Email (General Reporting)

When you are ready to dispatch general aggregated reports to configured recipients:

```powershell
python -m src --config config/config.example.yaml --send
```

---

## CLI Options

| Argument | Description | Default |
|---|---|---|
| `--config`, `-c` | Path to YAML configuration file | `config/config.yaml` or `config.example.yaml` |
| `--files`, `-f` | Specific list of Excel files to process | Scans directory pattern from config |
| `--dry-run` | Simulate email sending and write HTML preview | `True` (safe by default) |
| `--send` | Send actual email via SMTP | `False` |
| `--output-dir`, `-o` | Directory for generated attachments and HTML preview | `output` |
| `--generate-samples` | Generate sample mock Excel workbooks | `False` |

---

## Configuration Reference (`config.example.yaml`)

### Data Sources
```yaml
data_sources:
  directory: "samples/data"
  file_pattern: "*.xlsx"
  sheet_name: "Sales_Data"      # Target sheet name or 0 for first sheet
  header_row: 1                 # 1-indexed row where headers begin
  columns:                      # Maps your Excel column headers to internal fields
    date: "Date"
    region: "Region"
    representative: "Sales Rep"
    category: "Product Category"
    units: "Units Sold"
    revenue: "Revenue"
    status: "Status"
  filters:                      # Row filters
    status: ["Completed", "Shipped", "Delivered"]
```

### Summary Metrics
```yaml
summary_metrics:
  - id: "total_revenue"
    label: "Total Revenue"
    type: "sum"                 # sum, mean, count, min, max
    column: "revenue"
    format: "${:,.2f}"
```

### Grouped Breakdown
```yaml
breakdown:
  group_by: "region"
  metrics:
    - column: "revenue"
      agg: "sum"
      header: "Total Revenue"
      format: "${:,.2f}"
```

---

## Scheduling Automated Reports (Windows Task Scheduler)

To run this report every Monday morning at 8:00 AM:

1. Open PowerShell as Administrator.
2. Create a scheduled task using the following snippet:

```powershell
$Action = New-ScheduledTaskAction -Execute "C:\Users\antho\source\repos\excel-email-reporter\.venv\Scripts\python.exe" `
    -Argument "-m src --config C:\Users\antho\source\repos\excel-email-reporter\config\config.yaml --send" `
    -WorkingDirectory "C:\Users\antho\source\repos\excel-email-reporter"

$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 8:00am

Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName "AutomatedExcelEmailReport" -Description "Extracts Excel data and sends formatted email report"
```

---

## Running Tests

Execute unit tests with `pytest`:

```powershell
pytest
```
