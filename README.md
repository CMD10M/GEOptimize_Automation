# GEOOOOO

This is a first-pass internal tool for generating a GEOptimize preliminary GHX feasibility report from:

1. A project proposal PDF
2. A GLD output PDF

The app sends both PDFs to the Claude API, requests strict JSON, then fills the branded Word template.

## What colleagues will do

1. Open the app.
2. Upload the proposal PDF.
3. Upload the GLD output PDF.
4. Click **Generate Report**.
5. Download the completed `.docx` report.

## Setup

### 1. Install Python
Install Python 3.11 or newer.

### 2. Install dependencies
From this folder, run:

```bash 
pip install -r requirements.txt
```

### 3. Add your Anthropic API key
Copy `.env.example` to `.env` and paste your API key:

```text
ANTHROPIC_API_KEY=sk-ant-api03-...
CLAUDE_MODEL=claude-opus-4-7
```

You can also enter the key directly in the Streamlit sidebar.

## Run the app on a Mac

```bash
python3 -m streamlit run app.py
```

Then open the local URL shown in the terminal.

## Windows quick start

Double-click `run_app_windows.bat` after Python is installed. The first run may take a few minutes while dependencies install.

## Manual JSON fallback

If you prefer to use Claude manually:

1. Paste the prompt from `report_prompt.py` into Claude.
2. Save Claude's JSON output as `report_data.json` in this folder.
3. Run:

```bash
python generate_from_json.py
```

The final report will be saved in the `output/` folder.

## Notes

- Keep the Word template outside Claude. Claude should produce JSON only.
- The Word template controls formatting and branding.
- The Claude prompt controls engineering judgment and report content.
- Review all generated reports before sending externally.

## API reference context

Anthropic supports PDF inputs as document content blocks in the Messages API. PDFs are processed as text plus page images, allowing Claude to analyze text, tables, charts, and visual layout.
