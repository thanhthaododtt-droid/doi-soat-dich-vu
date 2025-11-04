import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io
from datetime import datetime

# ========== CẤU HÌNH ỨNG DỤNG ==========
st.set_page_config(page_title="Đối soát MS365 - Chuẩn 3 điều kiện", layout="wide")
st.title("📊 CÔNG CỤ ĐỐI SOÁT MS365 - Domain + SKU + Quantity (FINAL)")

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file NCC (sheet 'SEPT 25-MAT BAO')", type=["xlsx", "xls"])
with col2:
    internal_file = st.file_uploader("📥 Upload file PO nội bộ", type=["xlsx", "xls"])

# ========== HÀM TIỆN ÍCH ==========
def normalize(s):
    if pd.isna(s): return ""
    return str(s).strip().lower()

def fuzzy(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ========== XỬ LÝ ==========
if st.button("🚀 Tiến hành đối soát"):
    if not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ hai file.")
        st.stop()

    try:
        # --- Đọc dữ liệu ---
        df_ncc = pd.read_excel(vendor_file, sheet_name="SEPT 25-MAT BAO", dtype=object)
        df_po = pd.read_excel(internal_file, dtype=object)

        # --- Chuẩn hóa file NCC ---
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

        # --- Chuẩn hóa file PO ---
        df_po["Domain_norm"] = df_po["Domain"].apply(normalize)
        df_po["SKU_norm"] = df_po["Product"].apply(normalize)
        df_po["Quantity"] = pd.to_numeric(df_po["Quantity"], errors="coerce").fillna(0)

        # --- Tạo khóa chính ---
        df_ncc["Key_full"] = df_ncc["Domain_norm"] + "|" + df_ncc["SKU_norm"] + "|" + df_ncc["Billable_Quantity"].astype(str)
        df_ncc["Key_partial"] = df_ncc["Domain_norm"] + "|" + df_ncc["SKU_norm"]

        df_po["Key_full"] = df_po["Domain_norm"] + "|" + df_po["SKU_norm"] + "|" + df_po["Quantity"].astype(str)
        df_po["Key_partial"] = df_po["Domain_norm"] + "|" + df_po["SKU_norm"]

        # --- Merge full outer để giữ cả 2 phía ---
        df_ncc_key = df_ncc[[
            "Key_full", "Key_partial", "Subscription_ID",
            "Partner_Cost_USD", "Partner_Cost_VND"
        ]]

        merged = pd.merge(df_po, df_ncc_key, on="Key_full", how="outer", indicator=True)

        # === XÁC ĐỊNH TRẠNG THÁI ===
        status, score = [], []
        df_po_keys_partial = set(df_po["Key_partial"])
        df_ncc_keys_partial = set(df_ncc["Key_partial"])

        for _, row in merged.iterrows():
            key_p = row.get("Key_partial", "")
            if row["_merge"] == "both":
                status.append("✅ Khớp hoàn toàn")
                score.append(100)
            elif row["_merge"] == "left_only":  # Có ở PO, không có ở NCC
                if key_p in df_ncc_keys_partial:
                    status.append("⚠️ Sai lệch Quantity")
                    score.append(75)
                else:
                    status.append("❌ Thiếu ở NCC")
                    score.append(0)
            elif row["_merge"] == "right_only":  # Có ở NCC, không có ở PO
                if key_p in df_po_keys_partial:
                    status.append("⚠️ Sai lệch Quantity")
                    score.append(75)
                else:
                    status.append("❌ Thiếu ở PO")
                    score.append(0)
            else:
                status.append("⚠️ Không xác định")
                score.append(0)

        merged["Match_Status"] = status
        merged["Match_Score (%)"] = score
        merged.drop(columns=["_merge"], inplace=True)

        # --- Báo cáo tổng hợp (Summary) ---
        summary = merged.groupby("SKU_norm", dropna=False).agg({
            "Quantity": "sum",
            "Partner_Cost_USD": "sum",
            "Partner_Cost_VND": "sum"
        }).reset_index().rename(columns={
            "SKU_norm": "SKU_Name (Normalized)",
            "Quantity": "Total_Quantity",
            "Partner_Cost_USD": "Total_Cost_USD",
            "Partner_Cost_VND": "Total_Cost_VND"
        })

        # --- Sheet Payment Summary ---
        total_po = len(df_po)
        total_match = sum(merged["Match_Status"] == "✅ Khớp hoàn toàn")
        total_diff = sum(merged["Match_Status"].isin(["⚠️ Sai lệch Quantity"]))
        total_missing_ncc = sum(merged["Match_Status"] == "❌ Thiếu ở NCC")
        total_missing_po = sum(merged["Match_Status"] == "❌ Thiếu ở PO")
        total_usd = merged.loc[merged["Match_Status"] == "✅ Khớp hoàn toàn", "Partner_Cost_USD"].sum()
        total_vnd = merged.loc[merged["Match_Status"] == "✅ Khớp hoàn toàn", "Partner_Cost_VND"].sum()

        payment_summary = pd.DataFrame({
            "Chỉ tiêu": [
                "Tổng số PO",
                "Số dòng khớp hoàn toàn",
                "Số dòng sai lệch Quantity",
                "Thiếu ở NCC",
                "Thiếu ở PO",
                "Tổng Partner Cost (USD)",
                "Tổng Partner Cost (VND)",
                "Ngày đối soát"
            ],
            "Giá trị": [
                total_po,
                total_match,
                total_diff,
                total_missing_ncc,
                total_missing_po,
                total_usd,
                total_vnd,
                datetime.now().strftime("%d/%m/%Y %H:%M")
            ]
        })

        # --- Xuất file Excel ---
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
            label="⬇️ Tải file Excel đối soát tổng hợp",
            data=towrite,
            file_name=f"doi_soat_MS365_final_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ Lỗi trong quá trình xử lý: {e}")
