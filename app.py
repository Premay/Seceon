from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException
from pdfminer.pdfparser import PDFSyntaxError

from soc_pdf_to_excel import DEFAULT_TEMPLATE, build_workbook, parse_pdf


st.set_page_config(page_title="SOC PDF to Excel", page_icon="📊", layout="wide")
st.title("SOC Report PDF to Excel")
st.write(
    "Upload a Seceon SOC report PDF. The dashboard template is included with the app, "
    "and the generated workbook is available to download."
)

with st.form("conversion"):
    pdf_upload = st.file_uploader("SOC report (PDF)", type=["pdf"])
    submitted = st.form_submit_button("Convert")

if submitted:
    if pdf_upload is None:
        st.error("Upload a PDF report.")
        st.stop()

    template_path = Path(__file__).with_name(DEFAULT_TEMPLATE)
    if not template_path.is_file():
        st.error(
            f"Dashboard template not found: {DEFAULT_TEMPLATE}. "
            "Add the template file next to app.py before deploying."
        )
        st.stop()

    st.session_state.pop("conversion_result", None)
    try:
        with st.spinner("Reading the report and building the workbook..."):
            parsed = parse_pdf(BytesIO(pdf_upload.getvalue()))
            if not (parsed["widgets"] or parsed["kpis"] or parsed["stats"]):
                raise ValueError(
                    "No report cards were detected. Check that this is a text-based "
                    "Seceon SOC report; scanned PDFs are not supported."
                )

            output = BytesIO()
            checks, org, rid, period = build_workbook(
                parsed,
                template_path,
                output,
                pdf_upload.name,
            )
    except (PDFSyntaxError, InvalidFileException, BadZipFile, ValueError, OSError) as exc:
        st.error(f"Could not convert these files: {exc}")
        st.stop()

    st.session_state["conversion_result"] = {
        "output": output.getvalue(),
        "checks": checks,
        "org": org,
        "rid": rid,
        "period": period,
        "widget_count": len(parsed["widgets"]),
        "kpi_count": len(parsed["kpis"]),
        "stat_count": len(parsed["stats"]),
    }

result = st.session_state.get("conversion_result")
if result:
    report_id = f" ({result['rid']})" if result["rid"] else ""
    st.success(f"Conversion complete: {result['org']}{report_id} — {result['period']}")
    st.caption(
        f"Detected {result['widget_count']} chart widgets, "
        f"{result['kpi_count']} KPI cards, and {result['stat_count']} stat cards."
    )
    if result["checks"]:
        check_table = pd.DataFrame(result["checks"], columns=["Status", "Check", "Detail"])
        st.subheader("Validation checks")
        st.dataframe(check_table, use_container_width=True, hide_index=True)

    st.download_button(
        "Download Excel workbook",
        data=result["output"],
        file_name="SOC_Executive_Dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.info(
    "This converter relies on the PDF's text layer and chart geometry; it does not "
    "perform OCR. Review the workbook's Checks sheet and verify the results before use."
)
