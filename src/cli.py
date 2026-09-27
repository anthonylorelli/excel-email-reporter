"""Command-line interface for excel-email-reporter."""

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from src.config import load_config
from src.excel_extractor import ExcelExtractor
from src.report_generator import ReportGenerator
from src.email_sender import EmailSender


def run_pipeline(
    config_path: Optional[str] = None,
    explicit_files: Optional[List[str]] = None,
    dry_run: bool = True,
    save_preview: bool = True,
    output_dir: str = "output",
) -> int:
    """Runs the full Excel extraction, report generation, and email pipeline."""
    print("=" * 65)
    print("  EXCEL EMAIL REPORTER PIPELINE")
    print("=" * 65)

    # 1. Load configuration
    config = load_config(config_path)
    print(f"[1/4] Loaded configuration (Project: {config.project.get('name', 'Default')})")

    # 2. Extract and consolidate Excel files
    print("[2/4] Scanning and parsing Excel spreadsheets...")
    extractor = ExcelExtractor(config.data_sources)
    df, summaries = extractor.process_all(explicit_files)

    if not summaries:
        print("  [!] Warning: No Excel files were found to process.")
        print(f"      Search pattern: {config.data_sources.get('directory', 'data')}/{config.data_sources.get('file_pattern', '*.xlsx')}")
        return 1

    total_rows = sum(s.included_count for s in summaries)
    print(f"  Processed {len(summaries)} file(s), ingested {total_rows} record(s):")
    for s in summaries:
        print(f"   - {s.filename}: {s.included_count} rows included ({s.row_count} total)")

    if df.empty:
        print("  [!] Consolidated dataset is empty after filters.")
        return 1

    # 3. Generate report
    print("\n[3/4] Generating styled email report...")
    generator = ReportGenerator(config)
    rendered = generator.render(df, summaries)

    # Prepare attachments
    attachments: List[Path] = []
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)

    date_str = datetime.now().strftime("%Y-%m-%d")

    # Export consolidated Excel if requested
    attach_cfg = getattr(config, "attachments", {})
    if attach_cfg.get("attach_consolidated_excel", True):
        filename_pattern = attach_cfg.get("consolidated_filename", "Consolidated_Report_{date}.xlsx")
        excel_filename = filename_pattern.format(date=date_str)
        excel_out_path = out_dir_path / excel_filename
        extractor.export_consolidated_excel(df, excel_out_path)
        attachments.append(excel_out_path)
        print(f"  Generated consolidated attachment: {excel_out_path}")

    # Attach source files if configured
    if attach_cfg.get("attach_source_files", False):
        for s in summaries:
            attachments.append(Path(s.filepath))

    # Save HTML preview
    if save_preview or dry_run:
        preview_path = out_dir_path / f"report_preview_{date_str}.html"
        preview_path.write_text(rendered["html"], encoding="utf-8")
        print(f"  Saved HTML report preview to: {preview_path.resolve()}")

    # 4. Dispatch Email
    print("\n[4/4] Dispatching email...")
    subject_template = getattr(config, "email_subject", "Excel Report - {date}")
    subject = subject_template.format(date=date_str)

    sender = EmailSender(config)
    try:
        sender.send(
            subject=subject,
            html_body=rendered["html"],
            text_body=rendered["text"],
            attachments=attachments,
            dry_run=dry_run,
        )
    except Exception as ex:
        print(f"\n[ERROR] Failed to send email: {ex}")
        return 1

    print("[SUCCESS] Pipeline completed successfully!")
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Extract tabular data from multiple Excel sheets and email formatted reports."
    )
    parser.add_argument(
        "--config", "-c",
        type=str,
        default=None,
        help="Path to YAML configuration file (default: config/config.yaml or config.example.yaml)",
    )
    parser.add_argument(
        "--files", "-f",
        nargs="+",
        help="Explicit list of Excel file paths to process",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Simulate email sending without connecting to SMTP server and save preview HTML",
    )
    parser.add_argument(
        "--send",
        action="store_true",
        default=False,
        help="Actually send email to configured recipients via SMTP (overrides default dry-run)",
    )
    parser.add_argument(
        "--output-dir", "-o",
        type=str,
        default="output",
        help="Directory to save generated attachments and preview HTML",
    )
    parser.add_argument(
        "--generate-samples",
        action="store_true",
        help="Generate sample Excel spreadsheets for demonstration and testing",
    )

    args = parser.parse_args()

    if args.generate_samples:
        from samples.generate_sample_data import generate_all_samples
        generate_all_samples()
        print("Sample data generated in samples/data/")
        return 0

    # Default to dry-run unless --send is explicitly passed
    is_dry_run = not args.send or args.dry_run

    exit_code = run_pipeline(
        config_path=args.config,
        explicit_files=args.files,
        dry_run=is_dry_run,
        save_preview=True,
        output_dir=args.output_dir,
    )
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
