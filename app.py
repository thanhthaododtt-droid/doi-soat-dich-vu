import streamlit as st
import pandas as pd
import io
from datetime import datetime

# ========== CẤU HÌNH ỨNG DỤNG ==========
st.set_page_config(page_title="Đối soát MS365 - Domain + SKU + Quantity (Final Fixed)", layout="wide")
st.title("📊 CÔNG CỤ ĐỐI SOÁT MS365 - Domain + SKU + Quantity (FINAL SINGLE LINE)")

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file NCC (sheet 'SEPT 25-MAT BAO')", type=["xlsx", "xls"])
with col2:
    internal_file = st.file_uploader("📥 Upload file PO nội bộ", type=["xlsx", "xls"])

# ========== HÀM CHUẨN HÓA ==========
def normalize(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()

# ========== XỬ LÝ ==========
if st.button("🚀 Tiến hành đối soát"):
    if not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ hai file (NCC + PO nội bộ).")
        st.stop()

    try:
        # --- Đọc dữ liệu ---
        df_ncc = pd.read_excel(vendor_file, sheet_name="SEPT 25-MAT BAO", dtype=object)
        df_po = pd.read_excel(internal_file, dtype=object)

        # --- Chuẩn hóa NCC ---
        df_ncc = df_ncc.rename(columns={
            "Domain Name": "Domain_Name",
            "SKU Name": "SKU_Name",
            "Billable Quantity": "Billable_Quantity",
            "Subscription ID": "Subscription_ID",
            "Partner Cost (USD)": "Partner_Cost_USD",
            "Partner Cost (VND)": "Partner_Cost_VND"
        })
        df_ncc["Domain_norm"] = df_ncc["Domain_Name"].apply(normalize)
        df_ncc["SKU_norm"] = df_ncc["SKU_Name"].apply(normalize)
        df_ncc["Billable_Quantity"] = pd.to_numeric(df_ncc["Billable_Quantity"], errors="coerce").fillna(0)

        # --- Chuẩn hóa PO ---
        df_po["Domain_norm"] = df_po["Domain"].apply(normalize)
        df_po["SKU_norm"] = df_po["Product"].apply(normalize)
        df_po["Quantity"] = pd.to_numeric(df_po["Quantity"], errors="coerce").fillna(0)

        # --- Gộp 2 file theo Domain + SKU ---
        merged = pd.merge(
            df_po,
            df_ncc[["Domain_norm", "SKU_norm", "Billable_Quantity", "Subscription_ID", "Partner_Cost_USD", "Partner_Cost_VND"]],
            on=["Domain_norm", "SKU_norm"],
            how="outer",
            suffixes=("_PO", "_NCC"),
            indicator=True
        )

        # --- Xác định trạng thái đối soát ---
        match_status, qty_diff = [], []
        for _, row in merged.iterrows():
            q_po = row.get("Quantity", 0)
            q_ncc = row.get("Billable_Quantity", 0)

            if pd.isna(q_po): q_po = 0
            if pd.isna(q_ncc): q_ncc = 0

            if q_po == q_ncc and q_po > 0:
                match_status.append("✅ Khớp hoàn toàn")
                qty_diff.append(0)
            elif q_po > 0 and q_ncc > 0 and q_po != q_ncc:
                if q_po > q_ncc:
                    match_status.append("⚠️ Sai lệch Quantity (PO > NCC)")
                else:
                    match_status.append("⚠️ Sai lệch Quantity (NCC > PO)")
                qty_diff.append(abs(q_po - q_ncc))
            elif q_po > 0 and q_ncc == 0:
                match_status.append("❌ Thiếu ở NCC")
                qty_diff.append(q_po)
            elif q_ncc > 0 and q_po == 0:
                match_status.append("❌ Thiếu ở PO")
                qty_diff.append(q_ncc)
            else:
                match_status.append("⚠️ Không xác định")
                qty_diff.append(0)

        merged["Quantity_Diff"] = qty_diff
        merged["Match_Status"] = match_status

        # --- Báo cáo tổng hợp (Summary) ---
        summary = merged.groupby("SKU_norm", dropna=False).agg({
            "Quantity": "sum",
            "Billable_Quantity": "sum",
            "Partner_Cost_USD": "sum",
            "Partner_Cost_VND": "sum"
        }).reset_index().rename(columns={
            "SKU_norm": "SKU_Name (Normalized)",
            "Quantity": "Total_Quantity_PO",
            "Billable_Quantity": "Total_Quantity_NCC",
            "Partner_Cost_USD": "Total_Cost_USD",
            "Partner_Cost_VND": "Total_Cost_VND"
        })

        # --- Payment Summary ---
        total_po = len(df_po)
        total_match = sum(merged["Match_Status"] == "✅ Khớp hoàn toàn")
        total_diff_po = sum(merged["Match_Status"] == "⚠️ Sai lệch Quantity (PO > NCC)")
        total_diff_ncc = sum(merged["Match_Status"] == "⚠️ Sai lệch Quantity (NCC > PO)")
        total_missing_ncc = sum(merged["Match_Status"] == "❌ Thiếu ở NCC")
        total_missing_po = sum(merged["Match_Status"] == "❌ Thiếu ở PO")
        total_usd = merged.loc[merged["Match_Status"] == "✅ Khớp hoàn toàn", "Partner_Cost_USD"].sum()
        total_vnd = merged.loc[merged["Match_Status"] == "✅ Khớp hoàn toàn", "Partner_Cost_VND"].sum()

        payment_summary = pd.DataFrame({
            "Chỉ tiêu": [
                "Tổng số PO",
                "Số dòng khớp hoàn toàn",
                "Sai lệch Quantity (PO > NCC)",
                "Sai lệch Quantity (NCC > PO)",
                "Thiếu ở NCC",
                "Thiếu ở PO",
                "Tổng Partner Cost (USD)",
                "Tổng Partner Cost (VND)",
                "Ngày đối soát"
            ],
            "Giá trị": [
                total_po,
                total_match,
                total_diff_po,
                total_diff_ncc,
                total_missing_ncc,
                total_missing_po,
                total_usd,
                total_vnd,
                datetime.now().strftime("%d/%m/%Y %H:%M")
            ]
        })

        # --- Xuất Excel ---
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Full_Matched_Detail")
            summary.to_excel(writer, index=False, sheet_name="Summary")
            payment_summary.to_excel(writer, index=False, sheet_name="Payment_Summary")
            df_ncc.to_excel(writer, index=False, sheet_name="NCC_Data")
        towrite.seek(0)

        # --- Giao diện Streamlit ---
        st.success("✅ Đối soát hoàn tất! File kết quả đã sẵn sàng tải xuống.")
        st.download_button(
            label="⬇️ Tải file Excel kết quả đối soát tổng hợp",
            data=towrite,
            file_name=f"doi_soat_MS365_singleline_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ Lỗi trong quá trình xử lý: {e}")
