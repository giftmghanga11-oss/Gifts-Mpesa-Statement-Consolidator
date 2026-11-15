import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st

from converter import build_workbook, load_branches, mapping_csv, output_name, process_zip

st.set_page_config(page_title="Gift Mwanyumba M-Pesa Statement Consolidator", page_icon="📒", layout="wide")
st.title("📒 Gift Mwanyumba M-Pesa statement consolidator")
st.caption("Upload a zip of branch statements and get one workbook with a sheet per branch.")

zf = st.file_uploader("Zipped folder of branch statements", type=["zip"])
if not zf:
    st.stop()

data = zf.getvalue()
zid = hashlib.md5(data).hexdigest()[:8]
mapping = load_branches(Path(__file__).parent / "branches.csv")
try:
    res = process_zip(data, mapping)
except Exception as e:
    st.error(f"Couldn't read that zip: {e}")
    st.stop()

sts = res["statements"]
if not sts:
    st.error("No valid M-Pesa statements found in the zip.")
    st.stop()

new = [s for s in sts if s.is_new]
c = st.columns(4)
c[0].metric("Branches", len(sts))
c[1].metric("Duplicates skipped", len(res["skipped"]))
c[2].metric("New branches", len(new))
c[3].metric("Files rejected", len(res["rejected"]))

from collections import Counter
main = Counter(s.period for s in sts).most_common(1)[0][0]
odd = [s for s in sts if s.period != main]
if odd:
    st.warning(f"Most statements are for {main[0][:10]}, but these are for a different day: "
               + ", ".join(f"{s.sheet_name} ({s.period[0][:10]})" for s in odd))

if new:
    st.subheader("New branches")
    st.caption("These short codes aren't in branches.csv. Confirm or edit the sheet names.")
    ed = st.data_editor(
        pd.DataFrame({"Short code": [s.short_code for s in new],
                      "Account holder": [s.holder for s in new],
                      "Sheet name": [s.sheet_name for s in new]}),
        disabled=["Short code", "Account holder"], hide_index=True,
        width="stretch", key=f"new_{zid}")
    for s, n in zip(new, ed["Sheet name"]):
        s.sheet_name = str(n).strip() or s.sheet_name

st.subheader("Branches in the workbook")
st.dataframe(pd.DataFrame({
    "Sheet": [s.sheet_name for s in sts], "Short code": [s.short_code for s in sts],
    "Transactions": [s.n_txn for s in sts], "Totals check": [s.check for s in sts],
    "File": [s.file for s in sts]}), hide_index=True, width="stretch")

for title, items, two_col in [("Duplicates skipped", res["skipped"], True),
                              ("Rejected files", res["rejected"], True),
                              ("Ignored (not a statement file type)", res["ignored"], False)]:
    if items:
        with st.expander(f"{title} ({len(items)})"):
            if two_col:
                st.dataframe(pd.DataFrame(items, columns=["File", "Reason"]), hide_index=True, width="stretch")
            else:
                st.write("\n".join(f"- {i}" for i in items))

blank = st.checkbox("Keep an empty 'Sheet1' at the front (as in your current file)", value=True)
st.download_button("⬇️ Download consolidated workbook", build_workbook(sts, blank),
                   file_name=output_name(sts), type="primary",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if new:
    st.download_button("⬇️ Download updated branches.csv", mapping_csv(sts, mapping),
                       file_name="branches.csv", mime="text/csv")
    st.caption("Replace branches.csv in your GitHub repo with this file so the new names are remembered.")
