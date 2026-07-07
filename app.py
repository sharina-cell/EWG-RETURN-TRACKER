"""
app.py  —  OMRON / EWG Returns Reconciliation
===============================================
Streamlit UI for the OMRON Return Report tracker update tool.

Uploads accepted per widget:
  Tracker      : .xlsx
  TC Report    : .csv, .xls, .xlsx
  Marketplace  : .xls, .xlsx, .zip  (zip is auto-extracted)

Logic:
  - Rows already in tracker that are now Refunded / Cancelled → deleted
  - New active return orders not yet in tracker → appended
  - Invoice Number  → from TC Order Report (order_id match)
  - SKU             → from TC Order Report, fallback to MP file
  - TikTok          → only "In Process" / "To Process" rows added
  - Manual columns (Tracking Status, Return Confirmation by MGL, Receiving
    Date, Fault Description, Action, Dispute, Status) → always left blank
"""

from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from omron_report_engine import run_reconciliation

# ---------------------------------------------------------------------------
# Zip / file helpers
# ---------------------------------------------------------------------------

_EXCEL_EXT    = {".xls", ".xlsx"}
_RETURN_KW    = ["return_refund", "return refund", "returns", "refund"]
_CANCEL_KW    = ["cancelled", "cancel"]


def _extract_from_zip(uploaded_zip) -> io.BytesIO | None:
    """Pull the return/refund Excel out of a Seller Centre zip download."""
    raw = uploaded_zip.read() if hasattr(uploaded_zip, "read") else uploaded_zip
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        excel = [m for m in z.namelist() if Path(m).suffix.lower() in _EXCEL_EXT]
        hits  = [m for m in excel
                 if any(k in m.lower() for k in _RETURN_KW)
                 and not any(k in m.lower() for k in _CANCEL_KW)]
        target = hits[0] if hits else (excel[0] if excel else None)
        if target is None:
            return None
        buf = io.BytesIO(z.read(target))
        buf.name = Path(target).name
        return buf


def _resolve(uploaded_file) -> io.BytesIO | None:
    """Return a BytesIO (with .name) ready for the engine, or None."""
    if uploaded_file is None:
        return None
    if uploaded_file.name.lower().endswith(".zip"):
        buf = _extract_from_zip(uploaded_file)
        if buf is None:
            st.warning(f"No usable Excel file found inside **{uploaded_file.name}**.")
        return buf
    raw = uploaded_file.read()
    buf = io.BytesIO(raw)
    buf.name = uploaded_file.name
    return buf


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="OMRON Returns Reconciliation",
    page_icon="🔄",
    layout="wide",
)

st.title("🔄 OMRON Returns Reconciliation")
st.caption(
    "Cross-check marketplace return exports against the OMRON Return Report "
    "tracker — remove refunded/cancelled rows and auto-fill new active returns."
)

with st.expander("ℹ️ How to use", expanded=False):
    st.markdown(
        """
        1. **Upload the OMRON tracker** (`Omron-Return_report-TC_and_MP_*.xlsx`).
        2. **Upload the TC Order Report** (`.csv` or `.xlsx`) — used to fill
           **Invoice Number** and **SKU** for new rows.
        3. **Upload marketplace return files** for whichever platforms you have.
           Every uploader accepts **.xls**, **.xlsx**, or **.zip** (the
           return/refund file is auto-extracted from the zip; cancelled-order
           files inside are ignored).
        4. Click **▶️ Run**. The app will:
           - **Delete** tracker rows that are now Refunded or Cancelled in the
             MP files.
           - **Add** new active return rows, filling: Platform, Return Request
             Date, Order Number, Invoice Date, Invoice Number, Tracking Number,
             SKU, Qty, Return Reason.
           - **Leave blank** all manual columns (Tracking Status, Return
             Confirmation by MGL, Receiving Date, Fault Description, Action,
             Dispute, Redressing, Status).
        5. Preview the changes, then download the updated tracker.

        **TikTok:** only *In Process* and *To Process* returns are added —
        Completed and Refund Rejected are skipped automatically.
        """
    )

st.divider()

# --- Upload widgets ---
row1_l, row1_r = st.columns(2)

with row1_l:
    st.subheader("1 · OMRON Tracker")
    tracker_file = st.file_uploader(
        "Omron-Return_report-TC_and_MP_*.xlsx",
        type=["xlsx"],
        help="The existing OMRON Return Report tracker workbook.",
    )

with row1_r:
    st.subheader("2 · TC Order Report")
    tc_raw = st.file_uploader(
        "TC Order Report — .csv or .xlsx",
        type=["csv", "xls", "xlsx"],
        help="Enriches new rows with Invoice Number and SKU.",
    )

st.subheader("3 · Marketplace return files")
st.caption("Each uploader accepts .xls · .xlsx · .zip")

mp1, mp2, mp3 = st.columns(3)
with mp1:
    shopee_raw = st.file_uploader(
        "🛍️ Shopee return export",
        type=["xls", "xlsx", "zip"],
        key="shopee",
    )
with mp2:
    lazada_raw = st.file_uploader(
        "🟠 Lazada return export",
        type=["xls", "xlsx", "zip"],
        key="lazada",
    )
with mp3:
    tiktok_raw = st.file_uploader(
        "🎵 TikTok return export",
        type=["xls", "xlsx", "zip"],
        key="tiktok",
    )

st.divider()

run_btn = st.button(
    "▶️ Run reconciliation",
    type="primary",
    disabled=tracker_file is None,
)

if tracker_file is None:
    st.info("Upload the OMRON tracker workbook to get started.")
    st.stop()

if not run_btn:
    st.stop()

# --- Validate ---
if not any([shopee_raw, lazada_raw, tiktok_raw]):
    st.warning("Please upload at least one marketplace return file.")
    st.stop()

# --- Resolve uploads ---
shopee_buf  = _resolve(shopee_raw)
lazada_buf  = _resolve(lazada_raw)
tiktok_buf  = _resolve(tiktok_raw)
tc_buf      = _resolve(tc_raw)
tracker_buf = io.BytesIO(tracker_file.read())

# --- Run ---
with st.spinner("Cross-checking order IDs and rebuilding tracker…"):
    try:
        workbook, results, removed = run_reconciliation(
            tracker_file=tracker_buf,
            shopee_file=shopee_buf,
            lazada_file=lazada_buf,
            tiktok_file=tiktok_buf,
            tc_file=tc_buf,
        )
    except Exception as exc:
        st.error(f"Something went wrong: {exc}")
        st.stop()

st.success("Reconciliation complete!")

# --- Summary metrics ---
total_new = sum(len(r.new_rows) for r in results)
n_metrics  = max(len(results) + 1, 2)
mcols = st.columns(n_metrics)

with mcols[0]:
    st.metric("🗑️ Rows removed", removed,
              help="Refunded or cancelled rows deleted from the tracker")

for mcol, res in zip(mcols[1:], results):
    with mcol:
        st.metric(
            label=f"{res.marketplace}",
            value=f"{len(res.new_rows)} new",
            delta=f"{len(res.already_tracked)} already tracked",
            delta_color="off",
        )

# Warnings
for res in results:
    for w in res.warnings:
        st.warning(f"**{res.marketplace}**: {w}")

# --- New rows preview ---
st.divider()
st.subheader("New rows preview")

if total_new == 0:
    st.info("No new return orders — everything in the uploaded files is already tracked.")
else:
    for res in results:
        if not res.new_rows:
            continue
        st.markdown(f"**{res.marketplace}** — {len(res.new_rows)} new row(s)")
        st.dataframe(
            [
                {
                    "Order ID":         r.order_id,
                    "Platform":         r.platform,
                    "Return Req. Date": r.return_request_date,
                    "Invoice Date":     r.invoice_date,
                    "Invoice Number":   r.invoice_number,
                    "Tracking":         r.tracking_number,
                    "SKU":              r.sku,
                    "Qty":              r.qty,
                    "Return Reason":    r.return_reason,
                    "Notes":            r.notes,
                }
                for r in res.new_rows
            ],
            use_container_width=True,
        )

# --- Download ---
st.divider()
out_buf = io.BytesIO()
workbook.save(out_buf)
out_buf.seek(0)

out_name = (
    tracker_file.name.rsplit(".", 1)[0]
    + f"_updated_{datetime.now().strftime('%Y%m%d')}.xlsx"
)

st.download_button(
    "⬇️ Download updated tracker",
    data=out_buf,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
