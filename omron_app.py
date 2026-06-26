"""
app.py  —  OMRON Returns Reconciliation
========================================
Streamlit UI for the OMRON Return Report tracker update tool.

Uploads accepted: .xls / .xlsx / .zip / .csv (for TC report)
- Zip files are auto-extracted to find the return/refund Excel inside.
- TikTok: only "In Process" and "To Process" rows are added.
- Refunded / Cancelled rows in the tracker are removed automatically.
"""

from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from report_engine import run_reconciliation, ReconciliationResult

# ---------------------------------------------------------------------------
# Zip helper
# ---------------------------------------------------------------------------

EXCEL_EXTENSIONS = {".xls", ".xlsx"}
RETURN_KEYWORDS  = ["return_refund", "return refund", "returns", "refund"]
CANCEL_KEYWORDS  = ["cancelled", "cancel"]


def extract_from_zip(uploaded_zip) -> io.BytesIO | None:
    raw = uploaded_zip.read() if hasattr(uploaded_zip, "read") else uploaded_zip
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        excel_members = [m for m in z.namelist() if Path(m).suffix.lower() in EXCEL_EXTENSIONS]
        return_members = [
            m for m in excel_members
            if any(k in m.lower() for k in RETURN_KEYWORDS)
            and not any(k in m.lower() for k in CANCEL_KEYWORDS)
        ]
        target = return_members[0] if return_members else (excel_members[0] if excel_members else None)
        if target is None:
            return None
        buf = io.BytesIO(z.read(target))
        buf.name = Path(target).name
        return buf


def resolve_upload(uploaded_file) -> io.BytesIO | None:
    """Return a BytesIO (with .name) ready for the engine, or None."""
    if uploaded_file is None:
        return None
    name = uploaded_file.name.lower()
    if name.endswith(".zip"):
        buf = extract_from_zip(uploaded_file)
        if buf is None:
            st.warning(f"No Excel file found inside {uploaded_file.name}.")
        return buf
    raw = uploaded_file.read()
    buf = io.BytesIO(raw)
    buf.name = uploaded_file.name
    return buf


# ---------------------------------------------------------------------------
# Page layout
# ---------------------------------------------------------------------------

st.set_page_config(page_title="OMRON Returns Reconciliation", page_icon="🔄", layout="wide")

st.title("🔄 OMRON Returns Reconciliation")
st.caption(
    "Cross-check OMRON marketplace return files against the Return Report tracker, "
    "remove refunded/cancelled rows, and auto-fill new active returns."
)

with st.expander("ℹ️ How this works", expanded=False):
    st.markdown(
        """
        1. **Upload the OMRON tracker** — `Omron-Return_report-TC_and_MP_*.xlsx`
        2. **Upload marketplace return files** — Shopee, Lazada, and/or TikTok.
           Each uploader accepts **.xls / .xlsx / .zip**.
           If you upload a zip (e.g. from Shopee Seller Centre), the return/refund
           file is auto-extracted; cancelled-order files inside the zip are ignored.
        3. **Optionally upload the TC Order Report** (.csv or .xlsx) — used to fill
           Invoice Number and SKU when the marketplace file doesn't carry them.
        4. Click **Run**. The app will:
           - **Remove** rows that are now Refunded / Cancelled in the MP files
           - **Add** new active return rows not yet in the tracker
           - Fill: Platform, Return Request Date, Order Number, Invoice Date,
             Invoice Number, Tracking Number, SKU, JDE Model, Qty, Return Reason
           - Leave blank: Tracking Status, Return Confirmation by MGL, Receiving
             Date, Fault Description, Action, Dispute, Status (manual columns)

        **TikTok:** only *In Process* and *To Process* returns are added.
        Completed and Refund Rejected are skipped automatically.
        """
    )

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. OMRON Tracker")
    tracker_file = st.file_uploader(
        "Omron-Return_report-TC_and_MP_*.xlsx",
        type=["xlsx"],
        help="The existing OMRON Return Report tracker workbook.",
    )
with col2:
    st.subheader("2. TC Order Report (optional)")
    tc_file = st.file_uploader(
        "TC Order Report (.csv or .xlsx)",
        type=["csv", "xls", "xlsx"],
        help="Enriches new rows with Invoice Number and SKU.",
    )

st.subheader("3. Marketplace return files — any combination")
st.caption("Each uploader accepts .xls, .xlsx, or .zip")

mcol1, mcol2, mcol3 = st.columns(3)
with mcol1:
    shopee_raw = st.file_uploader(
        "🛍️ Shopee return export",
        type=["xls", "xlsx", "zip"],
        key="shopee",
    )
with mcol2:
    lazada_raw = st.file_uploader(
        "🟠 Lazada return export",
        type=["xls", "xlsx", "zip"],
        key="lazada",
    )
with mcol3:
    tiktok_raw = st.file_uploader(
        "🎵 TikTok return export",
        type=["xls", "xlsx", "zip"],
        key="tiktok",
    )

st.divider()

run = st.button("▶️ Run reconciliation", type="primary", disabled=tracker_file is None)

if tracker_file is None:
    st.info("Upload the OMRON tracker to get started.")

if run:
    if not any([shopee_raw, lazada_raw, tiktok_raw]):
        st.warning("Upload at least one marketplace return file before running.")
        st.stop()

    shopee_file = resolve_upload(shopee_raw)
    lazada_file = resolve_upload(lazada_raw)
    tiktok_file = resolve_upload(tiktok_raw)
    tc_buf      = resolve_upload(tc_file) if tc_file else None

    tracker_buf = io.BytesIO(tracker_file.read())

    with st.spinner("Cross-checking order IDs..."):
        try:
            workbook, results, removed = run_reconciliation(
                tracker_file=tracker_buf,
                shopee_file=shopee_file,
                lazada_file=lazada_file,
                tiktok_file=tiktok_file,
                tc_file=tc_buf,
            )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.stop()

    st.success("Done!")

    # Summary metrics
    total_new = sum(len(r.new_rows) for r in results)
    metric_cols = st.columns(max(len(results) + 1, 2))
    with metric_cols[0]:
        st.metric("Rows removed", removed, help="Refunded or cancelled rows deleted from tracker")
    for col, res in zip(metric_cols[1:], results):
        with col:
            st.metric(
                label=res.marketplace,
                value=f"{len(res.new_rows)} new",
                delta=f"{len(res.already_tracked)} already tracked",
                delta_color="off",
            )

    for res in results:
        for w in res.warnings:
            st.warning(f"**{res.marketplace}**: {w}")

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
                        "Order ID":           r.order_id,
                        "Platform":           r.platform,
                        "Return Req. Date":   r.return_request_date,
                        "Invoice Date":       r.invoice_date,
                        "Invoice Number":     r.invoice_number,
                        "Tracking":           r.tracking_number,
                        "SKU":                r.sku,
                        "Qty":                r.qty,
                        "Return Reason":      r.return_reason,
                        "Notes":              r.notes,
                    }
                    for r in res.new_rows
                ],
                use_container_width=True,
            )

    # Download
    buf = io.BytesIO()
    workbook.save(buf)
    buf.seek(0)
    orig_name = tracker_file.name.rsplit(".", 1)[0]
    out_name = f"{orig_name}_updated_{datetime.now().strftime('%Y%m%d')}.xlsx"

    st.divider()
    st.download_button(
        "⬇️ Download updated tracker",
        data=buf,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
