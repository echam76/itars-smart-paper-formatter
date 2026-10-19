"""English-language interface for ITARS 6.0 Smart Paper Formatter."""
from pathlib import Path

import streamlit as st

from formatter import format_document

st.set_page_config(page_title="ITARS Smart Paper Formatter | ITARS 6.0", page_icon="📄", layout="centered")
st.title("ITARS Smart Paper Formatter")
st.caption("International TVET Academic and Research Symposium (ITARS 6.0)")
st.info(
    "Upload your Microsoft Word manuscript (.docx) to apply basic formatting based on "
    "the official ITARS 6.0 reference template, and generate a compliance report."
)
st.warning(
    "**Prototype version:** Automated formatting does not guarantee full compliance. "
    "Review the abstract, figures, tables, equations, and APA 7th edition references before submission."
)

with st.expander("Privacy and processing limitations"):
    st.write(
        "The application processes uploaded documents in memory and does not intentionally "
        "store manuscripts in a database. The hosting provider may have its own logging or "
        "retention practices. Please do not upload confidential manuscripts until the "
        "organiser has approved the hosting and privacy arrangements."
    )

TEMPLATE = Path(__file__).with_name("ITARS_6_Reference_Template.docx")
st.download_button(
    "Download the official reference template",
    data=TEMPLATE.read_bytes(),
    file_name="ITARS_6_Reference_Template.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)

uploaded = st.file_uploader(
    "Upload your manuscript (Microsoft Word)",
    type=["docx"],
    help="Microsoft Word (.docx) only; maximum 20 MB.",
)

# Ensure a previous author's output is not displayed after a different manuscript is uploaded.
file_id = (uploaded.name, uploaded.size, uploaded.file_id) if uploaded else None
if st.session_state.get("last_file_id") != file_id:
    st.session_state.pop("formatted", None)
    st.session_state.pop("report", None)
    st.session_state["last_file_id"] = file_id

if uploaded and st.button("Format Manuscript", type="primary"):
    try:
        output, report = format_document(uploaded.getvalue(), label="ITARS 6.0", year="2027")
        st.session_state["formatted"] = output
        st.session_state["report"] = report
        st.success("Basic formatting completed. Please review the compliance report before submitting your paper.")
    except Exception as exc:
        st.error(f"Unable to process this document: {exc}")

if "formatted" in st.session_state:
    st.download_button(
        "⬇️ Download Formatted Word Document",
        data=st.session_state["formatted"],
        file_name="ITARS6_Formatted_Paper.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    st.subheader("Formatting and Compliance Report")
    report = st.session_state["report"]
    # Accept both the earlier Malay report schema and the English report schema.
    normalized_checks = []
    for item in report.get("checks", []):
        raw_status = item.get("status", item.get("Status", "Review required"))
        status = "Pass" if str(raw_status).casefold() in ("pass", "lulus") else "Review required"
        name = item.get("check", item.get("Semakan", "Formatting check"))
        details = item.get("details", item.get("Butiran", "No details provided"))
        normalized_checks.append({"status": status, "check": name, "details": details})
        st.write(("✅" if status == "Pass" else "⚠️") + f" **{name}** — {status}")
        st.caption(details)
    lines = [
        "ITARS 6.0 — FORMATTING AND COMPLIANCE REPORT",
        "Note: Automated checks are limited and do not indicate acceptance of the paper.",
        "",
    ]
    lines.extend(f"{item['status']} | {item['check']} | {item['details']}" for item in normalized_checks)
    st.download_button(
        "⬇️ Download Compliance Report (TXT)",
        data="\n".join(lines).encode("utf-8"),
        file_name="ITARS6_Compliance_Report.txt",
        mime="text/plain",
    )

st.divider()
st.caption(
    "This formatting tool is separate from the official paper submission portal. "
    "Download and carefully review your Word document before submission."
)
