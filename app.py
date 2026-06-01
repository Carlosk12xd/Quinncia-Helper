import io
from pathlib import Path

import streamlit as st

from quinncia_pptx_updater import update_powerpoint

st.set_page_config(page_title="Quinncia PowerPoint Updater", page_icon="📊", layout="centered")

st.title("Quinncia PowerPoint Updater")
st.caption("Upload the report PowerPoint and the Quinncia metrics CSV/Excel export. The app updates slide 2 only: Appendix: Entire MSB.")

with st.expander("What this app changes", expanded=True):
    st.write(
        "This app pulls the first section named **Quinncia Metrics (All Students)** and updates the "
        "values in the slide 2 Appendix table. It leaves the rest of the deck alone. It also makes "
        "the updated table body text black."
    )

pptx_file = st.file_uploader("1. Upload the PowerPoint template/report", type=["pptx"])
metrics_file = st.file_uploader("2. Upload the Quinncia metrics spreadsheet", type=["csv", "xlsx", "xlsm", "xls"])

if pptx_file and metrics_file:
    default_name = Path(pptx_file.name).stem + " - UPDATED.pptx"
    output_name = st.text_input("Output file name", value=default_name)

    if st.button("Update PowerPoint", type="primary"):
        try:
            updated_bytes, summary = update_powerpoint(
                pptx_bytes=pptx_file.getvalue(),
                metrics_bytes=metrics_file.getvalue(),
                metrics_filename=metrics_file.name,
            )

            st.success(
                f"Done. Updated {summary.updated_rows} rows and {summary.updated_cells} values on slide {summary.slide_updated}."
            )

            if summary.missing_programs:
                st.warning(
                    "These PowerPoint program rows were not found in the Quinncia data: "
                    + ", ".join(summary.missing_programs)
                )

            st.download_button(
                label="Download updated PowerPoint",
                data=io.BytesIO(updated_bytes),
                file_name=output_name if output_name.lower().endswith(".pptx") else output_name + ".pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as e:
            st.error(str(e))
else:
    st.info("Upload both files to enable the updater.")
