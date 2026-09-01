import datetime as dt
from pathlib import Path

import streamlit as st

from converter import (
    extract_rows_from_workbook,
    extract_open_shift_rows_from_workbook,
    apply_template_columns,
)


st.set_page_config(
    page_title="BCEHS Schedule Converter",
    layout="centered",
)

TEMPLATE_PATH = Path("assets/bcehs-schedule-template.csv")

st.title("BCEHS Schedule Converter")
st.caption(
    "Upload a BCEHS schedule (.xlsx, .xlsm, or .xls). "
    "Click ACP or PCP to generate the populated template CSV."
)

year = st.number_input(
    "Year for dates",
    min_value=2020,
    max_value=2100,
    value=dt.date.today().year,
    step=1,
)

bcehs_file = st.file_uploader(
    "BCEHS schedule (.xlsx, .xlsm, or .xls)",
    type=["xlsx", "xlsm", "xls"],
)

debug = st.checkbox(
    "Also produce debug file",
    value=True,
)

if not TEMPLATE_PATH.exists():
    st.error(
        f"Missing built-in template at: {TEMPLATE_PATH}. "
        "Add it to the repo."
    )
    st.stop()


def run_conversion(mode: str):
    if bcehs_file is None:
        st.warning("Please upload the BCEHS schedule file first.")
        return

    try:
        workbook_bytes = bcehs_file.getvalue()

        extracted = extract_rows_from_workbook(
            workbook_bytes,
            int(year),
            mode,
        )

        out_df, debug_df = apply_template_columns(
            extracted,
            str(TEMPLATE_PATH),
        )

        open_df = None
        if mode == "ACP":
            open_extracted = extract_open_shift_rows_from_workbook(
                workbook_bytes,
                int(year),
                mode,
            )
            open_df, _ = apply_template_columns(
                open_extracted,
                str(TEMPLATE_PATH),
            )

        if mode == "ACP":
            st.success(
                f"ACP conversion complete! Assigned rows: {len(out_df)}. "
                f"Open TBD rows: {len(open_df)}."
            )
        else:
            st.success(
                f"{mode} conversion complete! Rows exported: {len(out_df)}"
            )

        out_name = f"bcehs-populated-template_{mode}.csv"
        out_csv = out_df.to_csv(index=False).encode("utf-8")

        st.download_button(
            f"Download {mode} populated template CSV",
            data=out_csv,
            file_name=out_name,
            mime="text/csv",
        )

        if mode == "ACP" and open_df is not None:
            open_csv = open_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download ACP Open Shift Template CSV",
                data=open_csv,
                file_name="ACP Open Shift Template.csv",
                mime="text/csv",
            )

        if debug:
            debug_name = f"bcehs-debug_{mode}.csv"
            debug_csv = debug_df.to_csv(index=False).encode("utf-8")

            st.download_button(
                f"Download {mode} debug CSV",
                data=debug_csv,
                file_name=debug_name,
                mime="text/csv",
            )

        st.subheader("Assigned schedule preview (first 50 rows)")
        st.dataframe(
            out_df.head(50),
            use_container_width=True,
        )

        if mode == "ACP" and open_df is not None:
            st.subheader("ACP open shifts preview (first 50 rows)")
            st.dataframe(
                open_df.head(50),
                use_container_width=True,
            )

    except Exception as exc:
        st.error(f"{mode} conversion failed.")
        st.exception(exc)


col1, col2 = st.columns(2)

with col1:
    if st.button(
        "Convert ACP",
        type="primary",
        disabled=(bcehs_file is None),
    ):
        run_conversion("ACP")

with col2:
    if st.button(
        "Convert PCP",
        type="primary",
        disabled=(bcehs_file is None),
    ):
        run_conversion("PCP")
