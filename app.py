from datetime import date

import pandas as pd
import streamlit as st

from report import run_report
from xlsx_builder import COLUMNS, build_workbook

st.set_page_config(page_title="DXC India Insuff Report", page_icon="📋", layout="wide", initial_sidebar_state="collapsed")

today = date.today()
ss = st.session_state
ss.setdefault("start", date(2024, 1, 1))
ss.setdefault("end", today)
ss.setdefault("rows", None)


def load():
    try:
        with st.spinner("Querying database…"):
            rows = run_report(ss.start, ss.end)
    except Exception as e:  # surface DB errors in the page instead of a stack trace
        ss.rows = None
        st.error(f"Query failed: {e}")
        return
    ss.rows = rows


st.title("DXC India Insuff Report")
st.caption("Candidate Insuff Calling Details — latest follow-up per insufficient check (client 2307)")

# ---------- filters ----------
c_from, c_to, c_run, c_dl = st.columns([2, 2, 1.5, 2.5], vertical_alignment="bottom")
c_from.date_input("From", key="start")
c_to.date_input("To", key="end")
if c_run.button("Run report", type="primary", width="stretch"):
    load()  # runs in the page body so the "Querying database…" spinner shows

rows = ss.rows
if rows is None:
    st.info("Pick a date range and click **Run report**.")
    st.stop()

c_dl.download_button(
    "Download Excel (with PivotTables)",
    data=build_workbook(rows),
    file_name=f"DXC India Insuff Report {date.today():%d-%b-%y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
)

df = pd.DataFrame(rows, columns=COLUMNS)
q = st.text_input("Search all columns", placeholder="name, ARS no, remark…")
if q:
    mask = df.astype(str).apply(lambda col: col.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
    df = df[mask]
st.caption(f"{len(df)} of {len(rows)} rows")
st.dataframe(df, hide_index=True, width="stretch", height=600)
