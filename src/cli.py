"""Command-line interface for excel-email-reporter."""

import argparse
import re
import shutil
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from src.config import get_bundled_defaults_dir, load_config, resolve_resource_path
from src.excel_extractor import ExcelExtractor
from src.report_generator import ReportGenerator
from src.email_sender import EmailSender
from src.grade_calculator import GradeCalculator, StudentGradeProfile
from src.outlook_manager import OutlookDraftManager


def run_standard_pipeline(
    config: Any,
    explicit_files: Optional[List[str]] = None,
    dry_run: bool = True,
    save_preview: bool = True,
    output_dir: str = "output",
) -> int:
    """Runs the standard sales/operations summary aggregation pipeline."""
    print("=" * 65)
    print("  EXCEL EMAIL REPORTER — GENERAL AGGREGATION")
    print("=" * 65)

    print(f"[1/4] Loaded configuration (Project: {config.project.get('name', 'Default')})")

    # Extract and consolidate Excel files
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

    # Generate report
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

    # Dispatch Email
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


def run_grade_pipeline(
    config: Any,
    explicit_files: Optional[List[str]] = None,
    output_dir: str = "output/drafts",
    draft_target: Optional[str] = None,
    contacts_dir: Optional[str] = None,
) -> int:
    """Processes student assignment data, calculates weighted grades, and creates individual email drafts."""
    print("=" * 65)
    print("  STUDENT GRADE REPORTING — DRAFT EMAIL GENERATION")
    print("=" * 65)

    target_mode = (draft_target or getattr(config, "draft_target", "eml")).lower()
    if target_mode not in ["eml", "outlook"]:
        target_mode = "eml"

    print(f"[1/4] Initializing Grade Engine for: {config.project.get('name', 'Student Grades')}")
    print(f"      Draft Target: {target_mode.upper()} ({'Populate directly in Outlook' if target_mode == 'outlook' else 'Save .eml draft files on disk'})")
    cat_weights = config.grade_calculation.get("categories", {})
    print("      Active Categories & Weights:")
    for cat, w in cat_weights.items():
        print(f"       * {cat}: {int(w * 100)}%")
    print("      Missing Work Policy: Excluded from calculation until graded")

    # Discover and process workbooks
    print("\n[2/4] Ingesting student grade spreadsheets...")
    extractor = ExcelExtractor(config.data_sources)
    target_files = extractor.discover_files(explicit_files)

    if not target_files:
        print("  [!] Warning: No grade workbooks found.")
        print(f"      Checked: {config.data_sources.get('directory', 'data/grades')}/{config.data_sources.get('file_pattern', '*.xlsx')}")
        return 1

    print(f"  Found {len(target_files)} workbook(s) to process:")
    for f in target_files:
        print(f"   - {f.name}")

    parent_contacts = extractor.load_parent_contacts(contacts_dir)
    if parent_contacts:
        print(f"  Loaded parent contact emails for {len(parent_contacts)} student(s) (will be added on CC)")

    calculator = GradeCalculator(category_weights=cat_weights)
    generator = ReportGenerator(config)
    sender = EmailSender(config)
    outlook_mgr = OutlookDraftManager(config) if target_mode == "outlook" else None

    out_base = Path(output_dir)
    out_base.mkdir(parents=True, exist_ok=True)

    courses_summary: List[Dict[str, Any]] = []
    total_drafts_created = 0
    total_outlook_drafts = 0

    print("\n[3/4] Calculating weighted grades and generating email drafts...")

    try:
        for file_path in target_files:
            df, raw_count = extractor.extract_from_file(file_path)
            course_name = df["_course_name"].iloc[0] if "_course_name" in df.columns and not df.empty else file_path.stem
            
            profiles = calculator.process_records(
                df,
                course_name=course_name,
                parent_contacts=parent_contacts,
            )
            if not profiles:
                print(f"  [!] No valid student records extracted from: {file_path.name}")
                continue

            course_safe_name = re.sub(r'[^a-zA-Z0-9_\- ]', '', course_name).strip()
            course_dir = out_base / course_safe_name
            course_dir.mkdir(parents=True, exist_ok=True)

            students_data: List[Dict[str, Any]] = []
            scores = [p.overall_percent for p in profiles]
            class_avg = round(statistics.mean(scores), 1) if scores else 0.0
            class_med = round(statistics.median(scores), 1) if scores else 0.0

            for profile in profiles:
                rendered = generator.render_student_report(
                    student=profile,
                    template_html=config.template_html,
                    template_text=config.template_text,
                )

                student_safe_name = re.sub(r'[^a-zA-Z0-9_\- ]', '', profile.full_name).strip().replace(" ", "_")
                eml_path = course_dir / f"{student_safe_name}_grade_report.eml"
                html_path = course_dir / f"{student_safe_name}_grade_report.html"

                subject_pattern = getattr(config, "subject_template", "Progress Report: {student_name} - {course_name}")
                subject = subject_pattern.format(
                    student_name=profile.first_name or profile.full_name,
                    course_name=course_name,
                )

                in_outlook = False
                if target_mode == "outlook" and outlook_mgr is not None:
                    res = outlook_mgr.save_draft_to_outlook(
                        recipient_email=profile.email,
                        recipient_name=profile.full_name,
                        subject=subject,
                        html_body=rendered["html"],
                        text_body=rendered["text"],
                        cc_emails=profile.parent_emails,
                    )
                    if res["success"]:
                        in_outlook = True
                        total_outlook_drafts += 1
                    else:
                        if len(students_data) == 0:
                            print(f"  [!] Outlook direct save notice: {res['error']}")
                            print("      Creating companion .eml draft files on disk as fallback.")

                # If target is 'eml' or Outlook direct save wasn't available, write .eml file
                if not in_outlook or target_mode == "eml":
                    sender.create_draft_eml(
                        recipient_email=profile.email,
                        recipient_name=profile.full_name,
                        subject=subject,
                        html_body=rendered["html"],
                        text_body=rendered["text"],
                        output_eml_path=eml_path,
                        cc_emails=profile.parent_emails,
                    )

                # Save companion standalone HTML preview for easy inspection
                html_path.write_text(rendered["html"], encoding="utf-8")

                students_data.append({
                    "full_name": profile.full_name,
                    "first_name": profile.first_name,
                    "email": profile.email,
                    "parent_emails": ", ".join(profile.parent_emails),
                    "overall_percent": profile.overall_percent,
                    "letter_grade": profile.letter_grade,
                    "total_graded": profile.total_graded,
                    "total_missing": profile.total_missing,
                    "in_outlook": in_outlook,
                    "html_rel_path": f"{course_safe_name}/{html_path.name}",
                    "eml_rel_path": f"{course_safe_name}/{eml_path.name}" if eml_path.exists() else "",
                })
                total_drafts_created += 1

            courses_summary.append({
                "name": course_name,
                "filename": file_path.name,
                "students": students_data,
                "class_average": class_avg,
                "class_median": class_med,
            })

            print(f"  Processed course '{course_name}':")
            print(f"   * Students: {len(students_data)}")
            print(f"   * Class Average: {class_avg}% | Median: {class_med}%")
            if total_outlook_drafts > 0:
                print(f"   * Populated in Outlook: {sum(1 for s in students_data if s['in_outlook'])} drafts")
            print(f"   * Drafts Directory: {course_dir.resolve()}")

    finally:
        if outlook_mgr:
            outlook_mgr.close()

    # 4. Generate master review dashboard
    print("\n[4/4] Creating Drafts Review Dashboard...")
    dashboard_html = generator.render_drafts_dashboard(
        courses=courses_summary,
        dashboard_template=config.dashboard_template,
    )
    dash_path = out_base / "index.html"
    dash_path.write_text(dashboard_html, encoding="utf-8")

    print("\n" + "=" * 65)
    print("  DRAFTS GENERATION COMPLETE")
    print("=" * 65)
    print(f"  Total Student Drafts Processed: {total_drafts_created}")
    if total_outlook_drafts > 0:
        print(f"  Populated directly in Outlook:  {total_outlook_drafts} drafts in your Drafts folder")
    print(f"  Master Review Dashboard:        {dash_path.resolve()}")
    print("\n  Tip: Open the Dashboard in your browser or double-click any .eml")
    print("       file to open and review the draft in Outlook / Windows Mail!")
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Extract tabular data from Excel spreadsheets and generate formatted email reports or drafts."
    )
    parser.add_argument(
        "--config", "-c",
        type=str,
        default=None,
        help="Path to YAML configuration file",
    )
    parser.add_argument(
        "--files", "-f",
        nargs="+",
        help="Explicit list of Excel file paths to process",
    )
    parser.add_argument(
        "--grades",
        action="store_true",
        default=False,
        help="Run student grade calculation pipeline with individual draft reports",
    )
    parser.add_argument(
        "--draft-target", "--target",
        choices=["eml", "outlook"],
        default=None,
        help="Drafts destination: 'eml' (save .eml files to disk) or 'outlook' (populate directly in Outlook Drafts folder)",
    )
    parser.add_argument(
        "--outlook",
        action="store_const",
        const="outlook",
        dest="draft_target",
        help="Shortcut for --draft-target outlook (populate drafts directly in Outlook)",
    )
    parser.add_argument(
        "--eml",
        action="store_const",
        const="eml",
        dest="draft_target",
        help="Shortcut for --draft-target eml (save .eml files to disk)",
    )
    parser.add_argument(
        "--account", "-a",
        type=str,
        default=None,
        help="Microsoft account to use with --outlook: 'work' (default, organizational M365), 'personal' (Outlook.com/Live), or a specific email address (e.g. user@domain.com)",
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
        help="Actually send email via SMTP (overrides default dry-run)",
    )
    parser.add_argument(
        "--output-dir", "-o",
        type=str,
        default=None,
        help="Directory to save generated attachments, previews, or drafts",
    )
    parser.add_argument(
        "--contacts-dir",
        type=str,
        default=None,
        help="Directory containing parent contact spreadsheets (default: data/contacts)",
    )
    parser.add_argument(
        "--generate-samples",
        action="store_true",
        help="Generate sample Excel spreadsheets for demonstration and testing",
    )
    parser.add_argument(
        "--init",
        action="store_true",
        help="Initialize local config/, templates/, data/grades/, data/contacts/, and output/drafts/ folders in the current directory",
    )
    from src import __version__
    parser.add_argument(
        "--version", "-V",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    args = parser.parse_args()

    if args.init:
        defaults_dir = get_bundled_defaults_dir()
        for subdir in ["config", "templates"]:
            src_dir = defaults_dir / subdir
            dst_dir = Path(subdir)
            dst_dir.mkdir(parents=True, exist_ok=True)
            if src_dir.exists():
                for item in src_dir.iterdir():
                    target = dst_dir / item.name
                    if not target.exists():
                        shutil.copy2(item, target)
                        print(f"  [+] Created {target}")
                    else:
                        print(f"  [=] Skipped {target} (already exists)")
        for folder in [Path("data/grades"), Path("data/contacts"), Path("output/drafts")]:
            folder.mkdir(parents=True, exist_ok=True)
            print(f"  [+] Ensured directory {folder}/")
        print(f"\nWorkspace initialized in: {Path.cwd()}")
        return 0

    if args.generate_samples:
        from samples.generate_sample_data import generate_all_samples
        generate_all_samples()
        print("Sample data generated in samples/data/")
        return 0

    # Determine config to load
    config_file = args.config
    if config_file is None:
        if args.grades or (not Path("config/config.yaml").exists() and resolve_resource_path("config/grades_config.yaml") is not None):
            config_file = "config/grades_config.yaml"
        else:
            config_file = "config/config.yaml"

    config = load_config(config_file)
    if args.account:
        config.m365_account = args.account.strip()

    # Route pipeline based on project type or CLI flag
    is_grades = args.grades or config.project.get("type") == "student_grades" or config.audience == "individual_drafts"

    if is_grades:
        out_dir = args.output_dir or getattr(config, "output_drafts_dir", "output/drafts")
        contacts_dir = args.contacts_dir or config.data_sources.get("contacts_directory") or getattr(config, "contacts_directory", None)
        exit_code = run_grade_pipeline(
            config=config,
            explicit_files=args.files,
            output_dir=out_dir,
            draft_target=args.draft_target,
            contacts_dir=contacts_dir,
        )
    else:
        is_dry_run = not args.send or args.dry_run
        out_dir = args.output_dir or "output"
        exit_code = run_standard_pipeline(
            config=config,
            explicit_files=args.files,
            dry_run=is_dry_run,
            save_preview=True,
            output_dir=out_dir,
        )

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
