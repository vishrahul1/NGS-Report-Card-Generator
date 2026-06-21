# NGS Report Card Generator

A Python-based report card generation system for Neo Geetanjali Schools. The project can ingest Excel exam data, summarize student records, generate Word report cards, and export PDFs.

## Features

- Upload and ingest exam Excel files
- View and inspect student data
- Generate report cards for selected or all batches
- Export generated reports to PDF
- CLI and Streamlit UI support

## Project Structure

- `app.py` - Streamlit web interface
- `main.py` - CLI entry point
- `config.py` - project configuration and paths
- `ingest/` - Excel ingestion logic
- `generate/` - report generation and PDF export logic
- `data/` - input/output CSV data storage
- `templates/` - Word templates
- `output/` - generated DOCX/PDF outputs

## Requirements

Install dependencies with:

```bash
pip install -r requirements.txt
```

## Running the App

### Streamlit UI

```bash
streamlit run app.py
```

### CLI

```bash
python main.py --help
```

Example commands:

```bash
python main.py ingest --file path/to/exam.xlsx
python main.py status
python main.py generate --all --template templates/Sample_IIT_Report_Card.docx
```

## Notes

- The UI launcher script is `run_ui.bat` for Windows users.
- Generated files are stored under `output/`.
