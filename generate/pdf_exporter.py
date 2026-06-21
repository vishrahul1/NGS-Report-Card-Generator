"""
Phase 2C – PDF Exporter
Converts .docx files to .pdf using LibreOffice headless,
then optionally merges them into a single PDF.
"""

import os
import subprocess
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import OUTPUT_DIR

PDF_DIR = os.path.join(OUTPUT_DIR, "pdf")
DOCX_DIR = os.path.join(OUTPUT_DIR, "docx")
MERGED_PDF_NAME = "ALL_REPORT_CARDS.pdf"


def _find_libreoffice() -> str:
    """Locate the LibreOffice executable."""
    candidates = [
        "libreoffice",
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if shutil.which(c) or os.path.exists(c):
            return c
    return ""


def convert_docx_to_pdf(docx_paths: list[str], output_dir: str = PDF_DIR) -> list[str]:
    """
    Convert a list of .docx files to PDF using LibreOffice headless.
    Returns list of generated PDF paths.
    """
    lo = _find_libreoffice()
    if not lo:
        print(
            "[ERROR] LibreOffice not found.\n"
            "  Install from https://www.libreoffice.org/download/download/\n"
            "  Then ensure 'soffice' or 'libreoffice' is on your PATH."
        )
        return []

    os.makedirs(output_dir, exist_ok=True)
    generated = []

    for docx_path in docx_paths:
        if not os.path.exists(docx_path):
            print(f"  [SKIP] Not found: {docx_path}")
            continue

        print(f"  Converting: {os.path.basename(docx_path)}")
        try:
            result = subprocess.run(
                [lo, "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode != 0:
                print(f"  [ERROR] {result.stderr.strip() or result.stdout.strip()}")
                continue

            base = os.path.splitext(os.path.basename(docx_path))[0]
            pdf_path = os.path.join(output_dir, base + ".pdf")
            if os.path.exists(pdf_path):
                generated.append(pdf_path)
            else:
                print(f"  [WARN] PDF not found after conversion: {pdf_path}")
        except subprocess.TimeoutExpired:
            print(f"  [ERROR] Timeout converting {docx_path}")
        except Exception as exc:
            print(f"  [ERROR] {exc}")

    return generated


def merge_pdfs(pdf_paths: list[str], output_path: str) -> bool:
    """
    Merge a list of PDF files into one using PyPDF2.
    Returns True on success.
    """
    try:
        import PyPDF2
    except ImportError:
        print("[ERROR] PyPDF2 not installed. Run: pip install PyPDF2")
        return False

    if not pdf_paths:
        print("[WARN] No PDFs to merge.")
        return False

    merger = PyPDF2.PdfMerger()
    for path in sorted(pdf_paths):
        if os.path.exists(path):
            merger.append(path)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "wb") as f:
        merger.write(f)

    merger.close()
    print(f"  Merged {len(pdf_paths)} PDFs → {output_path}")
    return True


def export_all(
    input_dir: str = DOCX_DIR,
    output_dir: str = PDF_DIR,
    merge: bool = False,
) -> None:
    """
    Convert all .docx files in input_dir to PDF, optionally merging them.
    """
    docx_files = [
        os.path.join(input_dir, f)
        for f in sorted(os.listdir(input_dir))
        if f.lower().endswith(".docx")
    ] if os.path.isdir(input_dir) else []

    if not docx_files:
        print(f"No .docx files found in '{input_dir}'.")
        return

    print(f"\nConverting {len(docx_files)} file(s) to PDF...")
    pdf_paths = convert_docx_to_pdf(docx_files, output_dir)
    print(f"Done. {len(pdf_paths)} PDF(s) written to '{output_dir}'")

    if merge and pdf_paths:
        merged_path = os.path.join(output_dir, MERGED_PDF_NAME)
        merge_pdfs(pdf_paths, merged_path)
