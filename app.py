import datetime as dt
from pathlib import Path

import streamlit as st

from converter import extract_rows_from_workbook, apply_template_columns

st.set_page_config(page_title="BCEHS Schedule Converter", layout="centered")

TEMPLATE_PATH = Path("assets/bcehs-schedule-template.csv")

st.title("BCEHS Schedule Converter")
st.caption("Upload a BCEHS schedule (.xlsx/.xlsm). Click ACP or PCP to generate the populated template CSV.")

# Default to current year
year = st.number_input(
    "Year for dates",
    min_value=2020,
    max_value=2100,
    value=dt.date.today().year,
    step=1,
)

bcehs_file = st.file_uploader("BCEHS schedule (.xlsx or .xlsm)", type=["xlsx", "xlsm"])
debug = st.checkbox("Also produce debug file", value=True)

if not TEMPLATE_PATH.exists():
    st.error(f"Missing built-in template at: {TEMPLATE_PATH}. Add it to the repo.")
    st.stop()


def run_conversion(mode: str):
    """
    mode should be 'ACP' or 'PCP'
    """
    if bcehs_file is None:
        st.warning("Please upload the BCEHS schedule file first.")
        return

    try:
        xlsx_bytes = bcehs_file.getvalue()

        extracted = extract_rows_from_workbook(xlsx_bytes, int(year), mode)
        out_df, debug_df = apply_template_columns(extracted, str(TEMPLATE_PATH))

        st.success(f"{mode} conversion complete! Rows exported: {len(out_df)}")

        out_name = f"bcehs-populated-template_{mode}.csv"
        out_csv = out_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            f"Download {mode} populated template CSV",
            data=out_csv,
            file_name=out_name,
            mime="text/csv",
        )

        if debug:
            dbg_name = f"bcehs-debug_{mode}.csv"
            dbg_csv = debug_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                f"Download {mode} debug CSV",
                data=dbg_csv,
                file_name=dbg_name,
                mime="text/csv",
            )

        st.subheader("Preview (first 50 rows)")
        st.dataframe(out_df.head(50), use_container_width=True)

    except Exception as e:
        st.error(f"{mode} conversion failed.")
        st.exception(e)


col1, col2 = st.columns(2)
with col1:
    if st.button("Convert ACP", type="primary", disabled=(bcehs_file is None)):
        run_conversion("ACP")

with col2:
    if st.button("Convert PCP", type="primary", disabled=(bcehs_file is None)):
        run_conversion("PCP")
