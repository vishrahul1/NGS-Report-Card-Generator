#!/usr/bin/env python3
"""
Report Card Mail Merge System – CLI Entry Point

Usage:
  python main.py ingest    --file exam.xlsx
  python main.py status
  python main.py generate  --batch "P_G6_TULIPS" --template templates/Sample_IIT_Report_Card.docx
  python main.py generate  --all   --template templates/Sample_IIT_Report_Card.docx
  python main.py export-pdf [--input output/docx] [--output output/pdf] [--merge]
  python main.py reset     --confirm
"""

import os
import sys
import csv
import shutil

import click

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CSV_PATH, OUTPUT_DIR, TEMPLATE_DIR


# ── CLI Group ──────────────────────────────────────────────────────────────────

@click.group()
def cli():
    """NEO GEETANJALI SCHOOLS – Report Card Generator"""


# ── ingest ────────────────────────────────────────────────────────────────────

@cli.command()
@click.option("--file", "file_path", required=True, help="Path to the .xlsx exam file")
@click.option("--csv",  "csv_path",  default=CSV_PATH, show_default=True,
              help="Path to the output CSV (default: data/student_exam_data.csv)")
def ingest(file_path: str, csv_path: str):
    """Phase 1 – Ingest an Excel exam file into the persistent CSV store."""
    from ingest.excel_to_csv import ingest_excel
    try:
        ingest_excel(file_path, csv_path)
    except (FileNotFoundError, ValueError) as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)


# ── status ────────────────────────────────────────────────────────────────────

@cli.command()
@click.option("--csv", "csv_path", default=CSV_PATH, show_default=True)
def status(csv_path: str):
    """Show a summary of all ingested student/exam data."""
    if not os.path.exists(csv_path):
        click.echo("No data found. Run 'python main.py ingest --file <xlsx>' first.")
        return

    with open(csv_path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        click.echo("CSV exists but is empty.")
        return

    usernames   = {r["Username"] for r in rows}
    batches     = {r.get("Batch", "") for r in rows if r.get("Batch")}
    exams       = {r.get("Exam", "") for r in rows if r.get("Exam")}

    click.echo("\n=== Student Exam Data Summary ===")
    click.echo(f"Total Records  : {len(rows)}")
    click.echo(f"Unique Students: {len(usernames)}")
    click.echo(f"Batches        : {', '.join(sorted(batches))}")
    click.echo(f"Exams Loaded   : {', '.join(sorted(exams))}")

    click.echo("\nBreakdown by Batch:")
    batch_students: dict[str, set] = {}
    batch_exams:    dict[str, set] = {}
    for r in rows:
        b = r.get("Batch", "UNKNOWN")
        batch_students.setdefault(b, set()).add(r["Username"])
        batch_exams.setdefault(b, set()).add(r.get("Exam", ""))

    for b in sorted(batch_students):
        s_count = len(batch_students[b])
        e_count = len(batch_exams[b])
        click.echo(f"  {b}: {s_count} student(s) × {e_count} exam(s) = {s_count * e_count} records (expected)")


# ── generate ──────────────────────────────────────────────────────────────────

@cli.command()
@click.option("--batch",    default=None, help="Filter by batch name")
@click.option("--all",      "all_batches", is_flag=True, help="Generate for all batches")
@click.option("--template", default=os.path.join(TEMPLATE_DIR, "Sample_IIT_Report_Card.docx"),
              show_default=True, help="Path to the Word template (.docx)")
@click.option("--csv",      "csv_path", default=CSV_PATH, show_default=True)
def generate(batch, all_batches, template, csv_path):
    """Phase 2 – Generate per-student .docx report cards."""
    if not batch and not all_batches:
        click.echo("[ERROR] Specify --batch <name> or --all", err=True)
        sys.exit(1)

    from generate.template_filler import generate_report_cards
    try:
        generate_report_cards(
            csv_path=csv_path,
            template_path=template,
            batch_filter=None if all_batches else batch,
        )
    except (FileNotFoundError, ValueError) as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)


# ── export-pdf ────────────────────────────────────────────────────────────────

@cli.command("export-pdf")
@click.option("--input",  "input_dir",  default=os.path.join(OUTPUT_DIR, "docx"), show_default=True)
@click.option("--output", "output_dir", default=os.path.join(OUTPUT_DIR, "pdf"),  show_default=True)
@click.option("--merge",  is_flag=True, help="Merge all PDFs into ALL_REPORT_CARDS.pdf")
def export_pdf(input_dir, output_dir, merge):
    """Convert .docx report cards to PDF (requires LibreOffice)."""
    from generate.pdf_exporter import export_all
    export_all(input_dir=input_dir, output_dir=output_dir, merge=merge)


# ── reset ─────────────────────────────────────────────────────────────────────

@cli.command()
@click.option("--confirm", is_flag=True, required=True,
              help="Must pass --confirm to actually delete data")
@click.option("--csv",    "csv_path", default=CSV_PATH)
def reset(confirm, csv_path):
    """Delete the CSV data store and all generated output files."""
    deleted = []
    if os.path.exists(csv_path):
        os.remove(csv_path)
        deleted.append(csv_path)

    for sub in ["docx", "pdf"]:
        d = os.path.join(OUTPUT_DIR, sub)
        if os.path.isdir(d):
            shutil.rmtree(d)
            os.makedirs(d)
            deleted.append(d + "/")

    if deleted:
        click.echo("Deleted:\n" + "\n".join(f"  {p}" for p in deleted))
    else:
        click.echo("Nothing to delete.")


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    cli()
