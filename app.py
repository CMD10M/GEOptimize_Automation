import base64
import json
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
import time
from anthropic._exceptions import (
    OverloadedError,
    RateLimitError,
    APIConnectionError,
)
from anthropic import Anthropic
from docx import Document
from docxtpl import DocxTemplate
from dotenv import load_dotenv
from docx.shared import Pt

from report_prompt import SYSTEM_PROMPT, USER_PROMPT, REQUIRED_KEYS
from proposal_prompt import PROPOSAL_SYSTEM_PROMPT, PROPOSAL_USER_PROMPT

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

FEASIBILITY_TEMPLATE_PATH = (
    BASE_DIR
    / "templates"
    / "GEOptimize_GHX_Feasibility_Autofill_Template.docx"
)

PROPOSAL_TEMPLATE_PATH = (
    BASE_DIR
    / "templates"
    / "GEOptimize_Proposal_Master_Template.docx"
)



PROPOSAL_DATABASE_DIR = BASE_DIR / "proposal_database"
PROPOSAL_DATABASE_DIR.mkdir(exist_ok=True)


def pdf_block_from_bytes(data: bytes):
    encoded = base64.standard_b64encode(data).decode("utf-8")
    return {
        "type": "document",
        "source": {
            "type": "base64",
            "media_type": "application/pdf",
            "data": encoded,
        },
    }


def pdf_block(uploaded_file):
    return pdf_block_from_bytes(uploaded_file.getvalue())


def text_block(label: str, text: str):
    return {"type": "text", "text": f"\n\n--- {label} ---\n{text}"}


def extract_docx_text(data: bytes) -> str:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(data)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
        parts = []
        for para in doc.paragraphs:
            if para.text.strip():
                parts.append(para.text.strip())
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
        return "\n".join(parts)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def uploaded_file_to_content_block(uploaded_file, label: str):
    suffix = Path(uploaded_file.name).suffix.lower()
    data = uploaded_file.getvalue()
    if suffix == ".pdf":
        return pdf_block_from_bytes(data)
    if suffix == ".docx":
        return text_block(label, extract_docx_text(data))
    if suffix in [".txt", ".md"]:
        return text_block(label, data.decode("utf-8", errors="replace"))
    return text_block(label, f"Unsupported file type for direct extraction: {uploaded_file.name}")


def database_file_to_content_block(path: Path):
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return pdf_block_from_bytes(path.read_bytes())
    if suffix == ".docx":
        return text_block(f"Existing proposal example: {path.name}", extract_docx_text(path.read_bytes()))
    if suffix in [".txt", ".md"]:
        return text_block(f"Existing proposal example: {path.name}", path.read_text(encoding="utf-8", errors="replace"))
    return None


def extract_text(message) -> str:
    parts = []
    for block in message.content:
        if getattr(block, "type", None) == "text":
            parts.append(block.text)
    return "\n".join(parts).strip()


def parse_json_response(text: str) -> dict:
    cleaned = text.strip()
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


def render_feasibility_docx(context: dict) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / "GHX_Feasibility_Report.docx"
        doc = DocxTemplate(str(FEASIBILITY_TEMPLATE_PATH))
        doc.render(context)
        doc.save(str(out_path))
        return out_path.read_bytes()


def remove_paragraph(paragraph):
    element = paragraph._element
    element.getparent().remove(element)


def as_list(value):
    if isinstance(value, list):
        return value

    if isinstance(value, str):
        return [
            line.strip("•- ").strip()
            for line in value.splitlines()
            if line.strip()
        ]

    return []


def replace_marker_with_scope(doc: Document, marker: str, scope_sections):
    if not isinstance(scope_sections, list):
        scope_sections = []

    for paragraph in list(doc.paragraphs):
        if marker in paragraph.text:
            for section in scope_sections:
                heading = section.get("heading", "").strip()
                intro = section.get("intro", "").strip()
                tasks = section.get("tasks", [])

                if heading:
                    p = paragraph.insert_paragraph_before(heading)
                    p.style = "Heading 2"

                if intro:
                    p = paragraph.insert_paragraph_before(intro)
                    p.style = "Normal"
                for task in tasks:
                    p = paragraph.insert_paragraph_before(str(task).strip())
                    try:
                        p.style = "GEOptimize Bullet"
                    except KeyError:
                        p.style = "Normal"
                        p.text = "• " + p.text

            remove_paragraph(paragraph)
            break


def replace_marker_with_bullets(doc: Document, marker: str, items):
    items = as_list(items)

    for paragraph in list(doc.paragraphs):
        if marker in paragraph.text:
            for item in items:
                p = paragraph.insert_paragraph_before(str(item).strip())
                try:
                    p.style = "GEOptimize Bullet"
                except KeyError:
                    p.style = "Normal"
                    p.text = "• " + p.text

            remove_paragraph(paragraph)
            break

def replace_marker_with_terms(doc: Document, marker: str, terms_text: str):
    for paragraph in list(doc.paragraphs):
        if marker in paragraph.text:

            for line in terms_text.splitlines():
                clean = line.strip()

                if not clean:
                    paragraph.insert_paragraph_before("")

                elif re.match(r"^\d+\.\s", clean):
                    p = paragraph.insert_paragraph_before(clean)
                    p.style = "Legal"

                elif re.match(r"^\([a-z]\)", clean):
                    p = paragraph.insert_paragraph_before(clean)
                    p.style = "Legal"

                else:
                    p = paragraph.insert_paragraph_before(clean)
                    p.style = "Legal"

            remove_paragraph(paragraph)
            break


def render_proposal_docx(context: dict, terms_region: str) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        rendered_path = Path(tmpdir) / "proposal_rendered.docx"
        final_path = Path(tmpdir) / "GEOptimize_Proposal_Draft.docx"

        doc = DocxTemplate(str(PROPOSAL_TEMPLATE_PATH))
        doc.render(context)
        doc.save(str(rendered_path))

        final_doc = Document(str(rendered_path))

        replace_marker_with_scope(
            final_doc,
            "[[SCOPE_OF_WORK]]",
            context.get("scope_sections", [])
        )

        replace_marker_with_bullets(
            final_doc,
            "[[INFORMATION_REQUIRED]]",
            context.get("information_required", [])
        )

        replace_marker_with_bullets(
            final_doc,
            "[[DELIVERABLES]]",
            context.get("deliverables", [])
        )

        if terms_region == "USA":
            terms_path = BASE_DIR / "templates" / "terms" / "us_terms.txt"
        else:
            terms_path = BASE_DIR / "templates" / "terms" / "canada_terms.txt"

        terms_text = terms_path.read_text(encoding="utf-8")

        replace_marker_with_terms(
            final_doc,
            "[[TERMS_AND_CONDITIONS]]",
            terms_text
        )

        final_doc.save(str(final_path))
        return final_path.read_bytes()


def call_claude(client: Anthropic, model: str, max_tokens: int, system_prompt: str, content: list):
    retries = 3

    for attempt in range(retries):
        try:
            return client.messages.create(
                model=model,
                max_tokens=int(max_tokens),
                system=system_prompt,
                messages=[{"role": "user", "content": content}],
            )

        except Exception as exc:
            error_text = str(exc)

            overloaded = (
                "529" in error_text
                or "Overloaded" in error_text
                or "rate_limit" in error_text
            )

            if not overloaded or attempt == retries - 1:
                raise exc

            wait_seconds = 5 * (attempt + 1)

            st.warning(
                f"Claude API busy. Retrying in {wait_seconds} seconds..."
            )

            time.sleep(wait_seconds)


st.set_page_config(page_title="GEOptimize Automation", layout="centered")
st.title("GEOptimize Automation")

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
    max_tokens = st.number_input("Max output tokens", min_value=2000, max_value=30000, value=10000, step=500)
    page = st.radio("Tool", ["Feasibility Report Generator", "Proposal Generator"])

if page == "Feasibility Report Generator":
    st.header("GHX Feasibility Report Generator")
    st.caption("Upload the proposal and GLD output PDF, then generate a formatted Word report.")

    proposal_pdf = st.file_uploader("Proposal PDF", type=["pdf"], key="feasibility_proposal")
    gld_pdf = st.file_uploader("GLD output PDF", type=["pdf"], key="feasibility_gld")

    if st.button("Generate Feasibility Report", type="primary"):
        if not api_key:
            st.error("Enter your Anthropic API key in the sidebar or .env file.")
            st.stop()
        if not proposal_pdf or not gld_pdf:
            st.error("Upload both the proposal PDF and the GLD output PDF.")
            st.stop()

        client = Anthropic(api_key=api_key)
        content = [pdf_block(proposal_pdf), pdf_block(gld_pdf), {"type": "text", "text": USER_PROMPT}]

        with st.status("Analyzing documents with Claude and generating report...", expanded=True) as status:
            st.write("Sending PDFs to Claude...")
            message = call_claude(client, model, max_tokens, SYSTEM_PROMPT, content)
            st.write("Parsing JSON response...")
            raw_text = extract_text(message)
            try:
                context = validate_context(parse_json_response(raw_text))
                context["report_date"] = datetime.now().strftime("%B %d, %Y")
            except Exception as exc:
                st.error("Claude returned output that could not be parsed as JSON.")
                st.text_area("Raw Claude output", raw_text, height=400)
                st.exception(exc)
                st.stop()

            st.write("Filling Word template...")
            docx_bytes = render_feasibility_docx(context)
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

else:
    st.header("Proposal Generator")
    st.caption("Use existing proposal examples, uploaded project documents, and notes to draft a new proposal.")

    st.subheader("1. Existing proposal database")
    st.write("Place reusable proposal examples in the local `proposal_database` folder. Supported formats: PDF, DOCX, TXT, MD.")
    database_files = sorted([
        p for p in PROPOSAL_DATABASE_DIR.iterdir()
        if p.suffix.lower() in [".pdf", ".docx", ".txt", ".md"]
    ])
    if database_files:
        selected_database_files = st.multiselect(
            "Select proposal examples to use as style/scope references",
            options=[p.name for p in database_files],
            default=[],
        )
    else:
        selected_database_files = []
        st.info("No proposal examples found yet. Add files to the `proposal_database` folder, or upload examples below.")

    uploaded_examples = st.file_uploader(
        "Optional: upload additional example proposals",
        type=["pdf", "docx", "txt", "md"],
        accept_multiple_files=True,
        key="uploaded_examples",
    )

    st.subheader("2. New project documents")
    project_docs = st.file_uploader(
        "Upload site plans, RFPs, client notes, building documents, or related project files",
        type=["pdf", "docx", "txt", "md"],
        accept_multiple_files=True,
        key="project_docs",
    )

    st.subheader("3. Project notes")
    project_notes = st.text_area(
        "Enter any additional project information, scope direction, pricing notes, deadlines, exclusions, or assumptions",
        height=220,
        placeholder="Example: Client is evaluating preliminary GHX sizing only. Include energy model review as optional scope. Pricing to be left as placeholder.",
    )
    st.subheader("4. Terms & Conditions")

    terms_region = st.radio(
        "Select Terms & Conditions",
        ["Canada", "USA"],
        horizontal=True,
    )

    if st.button("Generate Proposal Draft", type="primary"):
        if not api_key:
            st.error("Enter your Anthropic API key in the sidebar or .env file.")
            st.stop()
        if not project_docs and not project_notes.strip():
            st.error("Upload at least one new project document or enter project notes.")
            st.stop()

        client = Anthropic(api_key=api_key)
        content = []

        selected_paths = [PROPOSAL_DATABASE_DIR / name for name in selected_database_files]
        for path in selected_paths:
            block = database_file_to_content_block(path)
            if block:
                content.append(block)

        for uploaded in uploaded_examples or []:
            content.append(uploaded_file_to_content_block(uploaded, f"Uploaded example proposal: {uploaded.name}"))

        for uploaded in project_docs or []:
            content.append(uploaded_file_to_content_block(uploaded, f"New project document: {uploaded.name}"))

        if project_notes.strip():
            content.append(text_block("User project notes", project_notes.strip()))

        content.append({"type": "text", "text": PROPOSAL_USER_PROMPT})

        with st.status("Drafting proposal with Claude...", expanded=True) as status:
            st.write("Sending proposal references and project documents to Claude...")
            message = call_claude(client, model, max_tokens, PROPOSAL_SYSTEM_PROMPT, content)
            st.write("Parsing proposal JSON response...")
            raw_text = extract_text(message)
            try:
                context = parse_json_response(raw_text)
                context["proposal_date"] = datetime.now().strftime("%B %d, %Y")

                # Backward-compatible client field cleanup
                if not context.get("client_name") or context.get("client_name") == "Not provided":
                    context["client_name"] = context.get("client_company", "Not provided")

                if not context.get("client_address"):
                    context["client_address"] = "Not provided"

                if not context.get("client_contact"):
                    context["client_contact"] = "Not provided"


            except Exception as exc:
                st.error("Claude returned output that could not be parsed as JSON.")
                st.text_area("Raw Claude output", raw_text, height=400)
                st.exception(exc)
                st.stop()

            st.write("Creating Word proposal draft...")
            docx_bytes = render_proposal_docx(context, terms_region)
            status.update(label="Proposal draft generated", state="complete")

        st.success("Done. Download the proposal draft below.")
        st.download_button(
            label="Download Proposal Draft (.docx)",
            data=docx_bytes,
            file_name="GEOptimize_Proposal_Draft.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        with st.expander("View generated JSON"):
            st.json(context)
