import base64
import json
import os
import re
import tempfile
from pathlib import Path

import streamlit as st
from anthropic import Anthropic
from docxtpl import DocxTemplate
from dotenv import load_dotenv

from report_prompt import SYSTEM_PROMPT, USER_PROMPT, REQUIRED_KEYS

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "GEOptimize_GHX_Feasibility_Autofill_Template.docx"


def pdf_block(uploaded_file):
    data = base64.standard_b64encode(uploaded_file.getvalue()).decode("utf-8")
    return {
        "type": "document",
        "source": {
            "type": "base64",
            "media_type": "application/pdf",
            "data": data,
        },
    }


def extract_text(message) -> str:
    parts = []
    for block in message.content:
        if getattr(block, "type", None) == "text":
            parts.append(block.text)
    return "\n".join(parts).strip()


def parse_json_response(text: str) -> dict:
    cleaned = text.strip()
    # Handles accidental fenced JSON despite prompt instructions.
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", cleaned, flags=re.DOTALL)
    if fenced:
        cleaned = fenced.group(1)
    else:
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start != -1 and end != -1 and end > start:
            cleaned = cleaned[start : end + 1]
    return json.loads(cleaned)


def validate_context(context: dict) -> dict:
    for key in REQUIRED_KEYS:
        context.setdefault(key, "Not provided")
    return context


def render_docx(context: dict) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / "GHX_Feasibility_Report.docx"
        doc = DocxTemplate(str(TEMPLATE_PATH))
        doc.render(context)
        doc.save(str(out_path))
        return out_path.read_bytes()


st.set_page_config(page_title="GHX Feasibility Report Generator", layout="centered")
st.title("GHX Feasibility Report Generator")
st.caption("Upload the proposal and GLD output PDF, then generate a formatted Word report.")

with st.sidebar:
    st.header("Claude Settings")
    api_key = st.text_input(
        "Anthropic API key",
        value=os.getenv("ANTHROPIC_API_KEY", ""),
        type="password",
        help="You can also place this in a .env file as ANTHROPIC_API_KEY.",
    )
    model = st.text_input(
        "Claude model",
        value=os.getenv("CLAUDE_MODEL", "claude-opus-4-7"),
        help="Use the model ID shown in your Claude Console. For cost savings, use a Sonnet model if preferred.",
    )
    max_tokens = st.number_input("Max output tokens", min_value=2000, max_value=20000, value=8000, step=500)

proposal_pdf = st.file_uploader("Proposal PDF", type=["pdf"])
gld_pdf = st.file_uploader("GLD output PDF", type=["pdf"])

if st.button("Generate Report", type="primary"):
    if not api_key:
        st.error("Enter your Anthropic API key in the sidebar or .env file.")
        st.stop()
    if not proposal_pdf or not gld_pdf:
        st.error("Upload both the proposal PDF and the GLD output PDF.")
        st.stop()

    client = Anthropic(api_key=api_key)

    content = [
        pdf_block(proposal_pdf),
        pdf_block(gld_pdf),
        {"type": "text", "text": USER_PROMPT},
    ]

    with st.status("Analyzing documents with Claude and generating report...", expanded=True) as status:
        st.write("Sending PDFs to Claude...")
        message = client.messages.create(
            model=model,
            max_tokens=int(max_tokens),
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": content}],
        )
        st.write("Parsing JSON response...")
        raw_text = extract_text(message)
        st.text_area("Raw Claude response", raw_text, height=300)
        try:
            context = validate_context(parse_json_response(raw_text))
            from datetime import datetime

            context["report_date"] = datetime.now().strftime("%B %d, %Y")
        except Exception as exc:
            st.error("Claude returned output that could not be parsed as JSON. Copy the raw output below and retry with a stricter prompt if needed.")
            st.text_area("Raw Claude output", raw_text, height=400)
            st.exception(exc)
            st.stop()

        st.write("Filling Word template...")
        docx_bytes = render_docx(context)
        status.update(label="Report generated", state="complete")

    st.success("Done. Download the completed Word report below.")
    st.download_button(
        label="Download GHX Feasibility Report (.docx)",
        data=docx_bytes,
        file_name="GHX_Feasibility_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    with st.expander("View generated JSON"):
        st.json(context)
