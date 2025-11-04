import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io
from datetime import datetime

st.set_page_config(page_title="Đối soát MS365 theo Domain (Auto Detect)", layout="wide")
st.title("📊 Công cụ đối soát MS365 - Match theo Domain (Tự nhận dạng cột NCC & PO)")

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file NCC (TD gửi)", type=["xlsx", "xls"])
with col2:
    internal_file = st.file_uploader("📥 Upload file PO nội bộ", type=["xlsx", "xls"])

def normalize_text(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()

def fuzzy(a, b):
    return SequenceMatcher(None, a, b).ratio()

def find_best_col(columns, keywords):
    """Tìm cột gần đúng nhất theo từ khóa"""
    for c in columns:
        c_norm = c.strip().lower()
        for k in keywords:
            if k in c_norm:
                return c
    # fallback: fuzzy match
    best_col, best_score = None, 0
    for c in columns:
        for k in keywords:
            score = fuzzy(c.lower(), k)
            if score > best_score:
                best_col, best_score = c, score
    return best_col

if st.button("🚀 Tiến hành đối soát"):
    if not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ hai file.")
        st.stop()

    try:
        # === Đọc dữ liệu ===
        df_ncc = pd.read_excel(vendor_file, header=2)
        df_po = pd.read_excel(internal_file)

        # === Dò cột trong NCC ===
        cols_ncc = [str(c).strip() for c in df_ncc.columns]
        domain_col = find_best_col(cols_ncc, ["domain", "tên miền"])
        sku_col = find_best_col(cols_ncc, ["sku", "gói", "plan", "service"])
        usd_col = find_best_col(cols_ncc, ["usd"])
        vnd_col = find_best_col(cols_ncc, ["vnd"])

        st.write(f"🧩 Đã phát hiện cột NCC: Domain → `{domain_col}`, SKU → `{sku_col}`, USD → `{usd_col}`, VND → `{vnd_col}`")

        if not domain_col or not sku_col:
            st.error("❌ Không thể tìm thấy cột Domain hoặc SKU trong file NCC. Hãy kiểm tra tên cột trong Excel.")
            st.stop()

        # Chuẩn hóa dữ liệu NCC
        df_ncc = df_ncc.rename(columns={
            domain_col: "NCC_Domain_Name",
            sku_col: "NCC_SKU_Name",
            usd_col: "NCC_Partner_Cost_USD",
            vnd_col: "NCC_Partner_Cost_VND"
        })
        df_ncc["Domain_norm"] = df_ncc["NCC_Domain_Name"].apply(normalize_text)

        # === Dò cột Domain trong PO ===
        cols_po = [str(c).strip() for c in df_po.columns]
        po_domain_col = find_best_col(cols_po, ["domain", "tên miền"])
        st.write(f"🧩 Đã phát hiện cột Domain trong PO: `{po_domain_col}`")

        if not po_domain_col:
            st.error("❌ Không thể tìm thấy cột Domain trong file PO nội bộ.")
            st.stop()

        df_po["Domain_norm"] = df_po[po_domain_col].apply(normalize_text)

        # === Match theo Domain ===
        results = []
        for _, po_row in df_po.iterrows():
            po_domain = po_row["Domain_norm"]
            best_match = None
            best_score = 0

            for _, ncc_row in df_ncc.iterrows():
                score = fuzzy(po_domain, ncc_row["Domain_norm"])
                if score > best_score:
                    best_score = score
                    best_match = ncc_row

            result = po_row.to_dict()
            if best_match is not None and best_score >= 0.85:
                result["NCC_Domain_Name"] = best_match["NCC_Domain_Name"]
                result["NCC_SKU_Name"] = best_match["NCC_SKU_Name"]
                result["NCC_Partner_Cost_USD"] = best_match.get("NCC_Partner_Cost_USD", "")
                result["NCC_Partner_Cost_VND"] = best_match.get("NCC_Partner_Cost_VND", "")
                result["Match_Score (%)"] = round(best_score * 100, 1)
                result["Trạng thái"] = "✅ Đã khớp"
            else:
                result["NCC_Domain_Name"] = ""
                result["NCC_SKU_Name"] = ""
                result["NCC_Partner_Cost_USD"] = ""
                result["NCC_Partner_Cost_VND"] = ""
                result["Match_Score (%)"] = round(best_score * 100, 1)
                result["Trạng thái"] = "❌ Không có trong NCC"
            results.append(result)

        df_result = pd.DataFrame(results)

        # === Xuất Excel ===
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="Full_Matched_Detail")
            df_ncc.to_excel(writer, index=False, sheet_name="NCC_Data")
        towrite.seek(0)

        st.success("✅ Đối soát hoàn tất! File xuất đã sẵn sàng tải xuống.")
        st.download_button(
            label="⬇️ Tải file Excel kết quả đối soát theo Domain",
            data=towrite,
            file_name=f"doi_soat_MS365_domain_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Lỗi trong quá trình xử lý: {e}")
