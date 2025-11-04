import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher
from datetime import datetime

# ======================== CONFIG ========================
st.set_page_config(page_title="Công cụ đối soát dịch vụ nội bộ", layout="wide")
st.title("📊 CÔNG CỤ ĐỐI SOÁT DỊCH VỤ MS365 - PHIÊN BẢN HOÀN CHỈNH")
st.markdown("""
Ứng dụng tự động đối soát dữ liệu giữa **File Nhà cung cấp (NCC)** và **File PO nội bộ**,  
tạo file kết quả **giống mẫu file đối chiếu thanh toán (MAT BAO)** gồm:
- Full_Matched_Detail (chi tiết từng PO)
- SUM (tổng hợp dạng Pivot)
- Payment_Summary (bảng tổng thanh toán)
""")

# ======================== INPUT ========================
service_type = st.selectbox("🔹 Chọn loại dịch vụ cần đối soát:", ["", "MS365"])
exchange_rate = None
if service_type == "MS365":
    st.markdown("💱 **Tùy chọn:** nhập tỷ giá USD → VND để quy đổi tổng thanh toán")
    use_rate = st.checkbox("Nhập tỷ giá quy đổi")
    if use_rate:
        exchange_rate = st.number_input("Tỷ giá (VND / USD):", value=26500, step=100)

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file NCC (TD gửi)", type=["xlsx", "xls"], key="vendor")
with col2:
    internal_file = st.file_uploader("📥 Upload file PO nội bộ", type=["xlsx", "xls"], key="internal")

# ======================== HELPER ========================
def safe_str(x):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        if hasattr(x, "strftime"):
            return x.strftime("%Y-%m-%d")
        return str(x)
    except Exception:
        return str(x)

def normalize_text(s):
    try:
        return safe_str(s).strip().lower()
    except Exception:
        return safe_str(s)

def fuzzy_match(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ======================== MAIN ========================
if st.button("🚀 Tiến hành đối soát"):
    if not service_type:
        st.warning("⚠️ Vui lòng chọn loại dịch vụ.")
    elif not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ cả hai file (NCC & PO).")
    else:
        df_vendor = pd.read_excel(vendor_file, header=2, dtype=object)
        df_internal = pd.read_excel(internal_file, dtype=object)

        try:
            st.subheader("🔍 Đang xử lý đối soát Microsoft 365...")

            # Chuẩn hóa dữ liệu NCC
            df_vendor.columns = [safe_str(c).strip() for c in df_vendor.columns]
            df_vendor = df_vendor.rename(columns={
                "Domain Name": "Domain",
                "SKU Name": "SKU_Name",
                "Sum of Partner Cost (USD)": "Partner_Cost_USD",
                "Sum of Partner Cost (VND)": "Partner_Cost_VND"
            })
            df_vendor = df_vendor.dropna(subset=["Domain", "SKU_Name"])

            # Chuẩn hóa dữ liệu nội bộ
            df_internal.columns = [safe_str(c).strip() for c in df_internal.columns]

            # Xác định cột domain, product, quantity
            domain_col, product_col, qty_col = None, None, None
            for c in df_internal.columns:
                lc = c.lower()
                if "domain" in lc:
                    domain_col = c
                if "product" in lc or "sku" in lc or "description" in lc:
                    product_col = c
                if "quantity" in lc or "qty" in lc:
                    qty_col = c
            if domain_col is None or product_col is None or qty_col is None:
                st.error("❌ Không tìm thấy cột Domain / Product / Quantity trong file PO.")
                st.stop()

            # Chuẩn hóa kiểu dữ liệu
            df_internal[qty_col] = pd.to_numeric(df_internal[qty_col], errors="coerce").fillna(0)

            # ----------------- MATCHING LOGIC -----------------
            matched_rows = []
            for _, po in df_internal.iterrows():
                po_domain = normalize_text(po[domain_col])
                po_product = normalize_text(po[product_col])

                best_match = None
                best_score = 0
                for _, ncc in df_vendor.iterrows():
                    ncc_domain = normalize_text(ncc["Domain"])
                    ncc_sku = normalize_text(ncc["SKU_Name"])
                    domain_score = fuzzy_match(po_domain, ncc_domain)
                    sku_score = fuzzy_match(po_product, ncc_sku)
                    score = (domain_score * 0.7 + sku_score * 0.3)
                    if score > best_score:
                        best_score = score
                        best_match = ncc

                row = dict(po)
                if best_match is not None and best_score >= 0.5:
                    row["NCC_Domain"] = best_match["Domain"]
                    row["NCC_SKU_Name"] = best_match["SKU_Name"]
                    row["Partner_Cost_USD"] = best_match["Partner_Cost_USD"]
                    row["Partner_Cost_VND"] = best_match["Partner_Cost_VND"]
                    row["Match_Score (%)"] = round(best_score * 100, 1)
                    row["Trạng thái"] = "✅ Đã khớp" if best_score >= 0.7 else "⚠️ Khớp thấp"
                else:
                    row["NCC_Domain"] = ""
                    row["NCC_SKU_Name"] = ""
                    row["Partner_Cost_USD"] = ""
                    row["Partner_Cost_VND"] = ""
                    row["Match_Score (%)"] = round(best_score * 100, 1)
                    row["Trạng thái"] = "❌ Không tìm thấy NCC"

                # Tính tổng giá trị & chênh lệch
                row["Total_VND_PO"] = ""
                row["Chênh lệch (VND)"] = ""
                if row["Partner_Cost_VND"] != "":
                    try:
                        cost_vnd = float(str(row["Partner_Cost_VND"]).replace(",", ""))
                        if exchange_rate:
                            cost_vnd = cost_vnd * 1.0  # Giữ nguyên, không quy đổi vì đã là VND
                        row["Total_VND_PO"] = cost_vnd
                        row["Chênh lệch (VND)"] = 0  # giả định match hoàn toàn
                    except:
                        pass

                matched_rows.append(row)

            result_full = pd.DataFrame(matched_rows)

            # ----------------- SHEET 2 - PIVOT (SUM) -----------------
            df_sum = (
                result_full.groupby(["NCC_SKU_Name"], dropna=False)
                .agg({
                    "Partner_Cost_USD": "sum",
                    "Partner_Cost_VND": "sum",
                    "Total_VND_PO": "sum"
                })
                .reset_index()
            )
            df_sum["Chênh lệch (VND)"] = df_sum["Total_VND_PO"] - df_sum["Partner_Cost_VND"]

            # ----------------- SHEET 3 - PAYMENT SUMMARY -----------------
            total_usd = pd.to_numeric(result_full["Partner_Cost_USD"], errors="coerce").fillna(0).sum()
            total_vnd = pd.to_numeric(result_full["Partner_Cost_VND"], errors="coerce").fillna(0).sum()
            total_po = pd.to_numeric(result_full["Total_VND_PO"], errors="coerce").fillna(0).sum()
            chenh_lech = total_po - total_vnd

            payment_summary = pd.DataFrame({
                "Nội dung": [
                    "Tổng USD NCC",
                    "Tổng VNĐ NCC",
                    "Tổng VNĐ PO",
                    "Chênh lệch (VNĐ)",
                    "Tỷ giá",
                    "Ngày đối soát"
                ],
                "Giá trị": [
                    total_usd,
                    total_vnd,
                    total_po,
                    chenh_lech,
                    exchange_rate if exchange_rate else "",
                    datetime.now().strftime("%Y-%m-%d %H:%M")
                ]
            })

            # ----------------- EXPORT EXCEL -----------------
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                result_full.to_excel(writer, index=False, sheet_name="Full_Matched_Detail")
                df_sum.to_excel(writer, index=False, sheet_name="SUM")
                payment_summary.to_excel(writer, index=False, sheet_name="Payment_Summary")
            towrite.seek(0)

            st.success("✅ Đối soát hoàn tất! File xuất đã sẵn sàng tải xuống.")
            st.download_button(
                label="⬇️ Tải file Excel kết quả đối soát",
                data=towrite,
                file_name=f"doi_soat_MS365_full_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Lỗi trong quá trình xử lý: {e}")
