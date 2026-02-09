import datetime as dt

import pandas as pd
import streamlit as st

from converter import extract_rows_from_workbook, apply_template_columns

st.set_page_config(page_title="BCEHS Schedule Converter", layout="centered")

st.title("BCEHS Schedule Converter")
st.caption("Upload a BCEHS schedule (.xlsx) and your template (.csv). Download the populated template output.")

col1, col2 = st.columns(2)
with col1:
    year = st.number_input("Year for dates", min_value=2020, max_value=2100, value=dt.date.today().year, step=1)
with col2:
    label = st.selectbox("Label (optional)", ["", "ACP", "PCP"], index=0)

bcehs_file = st.file_uploader("BCEHS schedule (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("Template (.csv)", type=["csv"])

debug = st.checkbox("Also produce debug file (recommended at first)", value=True)

if st.button("Convert", type="primary", disabled=(bcehs_file is None or template_file is None)):
    try:
        xlsx_bytes = bcehs_file.getvalue()
        template_bytes = template_file.getvalue()

        extracted = extract_rows_from_workbook(xlsx_bytes, int(year))
        out_df, debug_df = apply_template_columns(extracted, template_bytes)

        st.success(f"Converted! Rows exported: {len(out_df)}")

        # output CSV bytes
        out_csv = out_df.to_csv(index=False).encode("utf-8")
        out_name = f"bcehs-populated-template{('_' + label) if label else ''}.csv"
        st.download_button("Download populated template CSV", data=out_csv, file_name=out_name, mime="text/csv")

        if debug:
            dbg_csv = debug_df.to_csv(index=False).encode("utf-8")
            dbg_name = f"bcehs-debug{('_' + label) if label else ''}.csv"
            st.download_button("Download debug CSV", data=dbg_csv, file_name=dbg_name, mime="text/csv")

        # quick preview
        st.subheader("Preview")
        st.dataframe(out_df.head(50), use_container_width=True)

    except Exception as e:
        st.error("Conversion failed. If you download the debug file later, it can help diagnose.")
        st.exception(e)
