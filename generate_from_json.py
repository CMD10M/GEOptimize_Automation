import json
from pathlib import Path
from docxtpl import DocxTemplate
from report_prompt import REQUIRED_KEYS

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "GEOptimize_GHX_Feasibility_Autofill_Template.docx"
JSON_PATH = BASE_DIR / "report_data.json"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_PATH = OUTPUT_DIR / "GHX_Feasibility_Report.docx"


def main() -> None:
    if not JSON_PATH.exists():
        raise FileNotFoundError(f"Missing {JSON_PATH}. Save Claude's JSON output as report_data.json first.")

    with open(JSON_PATH, "r", encoding="utf-8") as f:
        context = json.load(f)

    missing = [k for k in REQUIRED_KEYS if k not in context]
    if missing:
        raise ValueError(f"JSON is missing required keys: {missing}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    doc = DocxTemplate(str(TEMPLATE_PATH))
    doc.render(context)
    doc.save(str(OUTPUT_PATH))
    print(f"Created: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
