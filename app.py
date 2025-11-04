import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io
from datetime import datetime

st.set_page_config(page_title="Đối soát MS365 - Chuẩn kế toán", layout="wide")
st.title("📊 Công cụ đối soát MS365 - Chuẩn kế toán (Domain + SKU + Quantity)")

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file NCC (TD gửi - sheet 'SEPT 25-MAT BAO')", type=["xlsx", "xls"])
with col2:
    internal_file = st.file_uploader("📥 Upload file PO nội bộ", type=["xlsx", "xls"])

def normalize_text(s):
    if pd.isna(s): return ""
    return str(s).strip().lower()

def fuzzy(a, b):
    return SequenceMatcher(None, a, b).ratio()

if st.button("🚀 Tiến hành đối soát"):
    if not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ hai file.")
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
        df_ncc["Domain_norm"] = df_ncc["Domain_Name"].apply(normalize_text)
        df_ncc["SKU_norm"] = df_ncc["SKU_Name"].apply(normalize_text)
        df_ncc["Billable_Quantity"] = pd.to_numeric(df_ncc["Billable_Quantity"], errors="coerce").fillna(0)

        # --- Chuẩn hóa PO ---
        df_po["Domain_norm"] = df_po["Domain"].apply(normalize_text)
        df_po["SKU_norm"] = df_po["Product"].apply(normalize_text)
        df_po["Quantity"] = pd.to_numeric(df_po["Quantity"], errors="coerce").fillna(0)

        # --- Merge full outer để không mất dữ liệu ---
        df_ncc_key = df_ncc[["Domain_norm", "SKU_norm", "Billable_Quantity", 
                             "Subscription_ID", "Partner_Cost_USD", "Partner_Cost_VND"]]
        df_ncc_key = df_ncc_key.rename(columns={
            "Billable_Quantity": "Quantity",
            "Subscription_ID": "NCC_Subscription_ID",
            "Partner_Cost_USD": "NCC_Partner_Cost_USD",
            "Partner_Cost_VND": "NCC_Partner_Cost_VND"
        })

        merged = pd.merge(df_po, df_ncc_key,
                          on=["Domain_norm", "SKU_norm", "Quantity"],
                          how="outer",
                          indicator=True)

        # --- Tạo trạng thái đối soát ---
        status = []
        score_list = []
        for _, row in merged.iterrows():
            if row["_merge"] == "both":
                status.append("✅ Khớp hoàn toàn")
                score_list.append(100)
            elif row["_merge"] == "left_only":
                status.append("❌ Thiếu ở NCC")
                score_list.append(0)
            else:
                status.append("❌ Thiếu ở PO")
                score_list.append(0)
        merged["Match_Status"] = status
        merged["Match_Score (%)"] = score_list

        merged.drop(columns=["_merge"], inplace=True)

        # --- Xuất báo cáo tổng hợp ---
        summary = merged.groupby("SKU_norm", dropna=False).agg({
            "Quantity": "sum",
            "NCC_Partner_Cost_USD": "sum",
            "NCC_Partner_Cost_VND": "sum"
        }).reset_index().rename(columns={
            "SKU_norm": "SKU_Name (Normalized)",
            "Quantity": "Total_Quantity",
            "NCC_Partner_Cost_USD": "Total_Cost_USD",
            "NCC_Partner_Cost_VND": "Total_Cost_VND"
        })

        # --- Xuất file Excel ---
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Full_Matched_Detail")
            summary.to_excel(writer, index=False, sheet_name="Summary")
            df_ncc.to_excel(writer, index=False, sheet_name="NCC_Data")
        towrite.seek(0)

        st.success("✅ Đối soát hoàn tất! File xuất đã sẵn sàng tải xuống.")
        st.download_button(
            label="⬇️ Tải file Excel kết quả đối soát tổng hợp",
            data=towrite,
            file_name=f"doi_soat_MS365_final_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ Lỗi trong quá trình xử lý: {e}")
