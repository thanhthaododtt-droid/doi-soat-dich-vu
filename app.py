import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher
from datetime import datetime

# ------------------ CONFIG ------------------
st.set_page_config(page_title="Công cụ đối soát dịch vụ nội bộ", layout="wide")

st.title("📊 CÔNG CỤ ĐỐI SOÁT DỊCH VỤ NỘI BỘ")
st.markdown(
    """
Ứng dụng hỗ trợ đối chiếu dữ liệu giữa **file Nhà cung cấp** và **file Nội bộ (PO)**  
Áp dụng cho các dịch vụ CNTT như MS365, SSL, ODS License, Google Workspace, TMQT, Chứng thư CKS.
"""
)

# ------------------ INPUT ------------------
service_type = st.selectbox(
    "🔹 Chọn loại dịch vụ cần đối soát:",
    ["", "MS365", "ODS License", "SSL", "Google Workspace", "TMQT", "Chứng thư CKS"]
)

# Tùy chọn nhập tỷ giá (chỉ áp dụng cho MS365)
exchange_rate = None
if service_type == "MS365":
    st.markdown("💱 **Tùy chọn:** nhập tỷ giá USD → VND để quy đổi tổng thanh toán")
    use_rate = st.checkbox("Nhập tỷ giá quy đổi")
    if use_rate:
        exchange_rate = st.number_input("Tỷ giá (VND / USD):", value=26500, step=100)

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file Nhà cung cấp", type=["xlsx", "xls", "csv"], key="vendor")
with col2:
    internal_file = st.file_uploader("📥 Upload file Nội bộ (PO)", type=["xlsx", "xls", "csv"], key="internal")

# ------------------ HELPER ------------------
def safe_str(x):
    """Chắc chắn trả về chuỗi, tránh lỗi nếu x là datetime/float/int/NaN"""
    try:
        if x is None:
            return ""
        if isinstance(x, float) and pd.isna(x):
            return ""
        if hasattr(x, "strftime"):
            return x.strftime("%Y-%m-%d")
        return str(x)
    except Exception:
        try:
            return str(x)
        except Exception:
            return ""

def read_file(f, service_type=None):
    """Đọc file Excel/CSV, xử lý riêng cho MS365 (header ở dòng 3)"""
    if f is None:
        return None
    try:
        if service_type == "MS365":
            df = pd.read_excel(f, header=2, dtype=object)
        else:
            if f.name.endswith(".csv"):
                df = pd.read_csv(f, dtype=object)
            else:
                df = pd.read_excel(f, dtype=object)
        df.columns = [safe_str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return None

def normalize_text(s):
    """Chuẩn hóa text an toàn, tránh lỗi khi gặp datetime hoặc số"""
    try:
        s2 = safe_str(s)
        return s2.strip().lower()
    except Exception:
        return safe_str(s)

def fuzzy_match(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ------------------ MAIN ------------------
if st.button("🚀 Tiến hành đối soát"):
    if not service_type:
        st.warning("⚠️ Vui lòng chọn loại dịch vụ.")
    elif not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ cả hai file (Nhà cung cấp & Nội bộ).")
    else:
        df_vendor = read_file(vendor_file, service_type)
        df_internal = read_file(internal_file, service_type)

        # ------------------ MS365 ------------------
        if service_type == "MS365":
            st.subheader("🔍 Đang xử lý đối soát Microsoft 365...")

            try:
                # Chuẩn hóa dữ liệu NCC
                df_vendor.columns = [safe_str(c).strip() for c in df_vendor.columns]
                df_vendor = df_vendor.rename(columns={
                    "Row Labels": "Plan",
                    "Sum of Partner Cost (USD)": "USD",
                    "Sum of Partner Cost (VND)": "VND"
                })
                df_vendor = df_vendor.dropna(subset=["Plan"])
                df_vendor = df_vendor[df_vendor["Plan"] != "Row Labels"]

                # Chuẩn hóa dữ liệu nội bộ
                df_internal.columns = [safe_str(c).strip() for c in df_internal.columns]
                desc_col = None
                qty_col = None
                for c in df_internal.columns:
                    lc = safe_str(c).lower()
                    if "description" in lc or "product" in lc or "recurring" in lc or "plan" in lc:
                        desc_col = c
                    if "quantity" in lc or "qty" in lc:
                        qty_col = c
                if desc_col is None:
                    desc_col = df_internal.columns[0]
                if qty_col is None:
                    df_internal["__qty__"] = 1
                    qty_col = "__qty__"

                df_internal[qty_col] = pd.to_numeric(df_internal[qty_col].apply(lambda x: safe_str(x)), errors="coerce").fillna(0)

                # Fuzzy match chi tiết giữa NCC và nội bộ
                matched_details = []
                for _, vrow in df_vendor.iterrows():
                    v_plan = normalize_text(vrow.get("Plan", ""))
                    best_match = None
                    best_score = 0
                    for _, irow in df_internal.iterrows():
                        i_plan = normalize_text(irow.get(desc_col, ""))
                        score = fuzzy_match(v_plan, i_plan)
                        if score > best_score:
                            best_score = score
                            best_match = irow

                    combined = {}
                    for c in df_vendor.columns:
                        combined[f"NCC_{c}"] = vrow.get(c, "")
                    if best_match is not None:
                        for c in df_internal.columns:
                            combined[f"PO_{c}"] = best_match.get(c, "")
                    else:
                        for c in df_internal.columns:
                            combined[f"PO_{c}"] = ""

                    combined["Match_Score (%)"] = round(best_score * 100, 1)
                    combined["Ghi chú"] = "✅ Đã khớp" if best_score >= 0.6 else "❌ Không khớp"
                    matched_details.append(combined)

                result_full = pd.DataFrame(matched_details)

                # Tính quy đổi (nếu có tỷ giá)
                if exchange_rate:
                    result_full["VND_Quydoi"] = pd.to_numeric(result_full["NCC_USD"], errors="coerce").fillna(0) * exchange_rate
                    result_full["VND_Quydoi"] = result_full["VND_Quydoi"].astype(int)

                # Tổng hợp
                result_full["USD_num"] = pd.to_numeric(result_full["NCC_USD"], errors="coerce").fillna(0)
                result_full["VND_num"] = pd.to_numeric(result_full["NCC_VND"], errors="coerce").fillna(0)
                total_usd = result_full["USD_num"].sum()
                total_vnd = result_full["VND_num"].sum()
                total_qd = result_full["VND_Quydoi"].sum() if "VND_Quydoi" in result_full else 0
                chenh_lech = total_qd - total_vnd if exchange_rate else 0

                # Hiển thị kết quả
                st.success("✅ Đối soát hoàn tất!")
                st.dataframe(result_full, use_container_width=True)

                st.markdown("### 📊 Tổng hợp")
                st.write(f"**Tổng (USD):** {total_usd:,.2f}")
                st.write(f"**Tổng (VND - NCC):** {total_vnd:,.0f}")
                if exchange_rate:
                    st.write(f"**Tổng (VND quy đổi):** {total_qd:,.0f}")
                    st.write(f"**Chênh lệch:** {chenh_lech:,.0f}")

                # Xuất file Excel (3 sheet)
                towrite = io.BytesIO()
                with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                    result_full.to_excel(writer, index=False, sheet_name="Full_Matched_Detail")

                    summary = pd.DataFrame({
                        "Tổng USD": [total_usd],
                        "Tổng VND NCC": [total_vnd],
                        "Tổng VND Quy đổi": [total_qd],
                        "Chênh lệch (VND)": [chenh_lech],
                        "Tỷ giá": [exchange_rate if exchange_rate else ""],
                        "Ngày đối soát": [datetime.now().strftime("%Y-%m-%d %H:%M")]
                    })
                    summary.to_excel(writer, index=False, sheet_name="Summary")

                    # Sheet tổng hợp thanh toán (Payment_Summary)
                    payment_summary = pd.DataFrame({
                        "Nội dung": [
                            "Tổng USD NCC",
                            "Tổng VNĐ NCC",
                            "Tỷ giá quy đổi",
                            "Tổng VNĐ quy đổi",
                            "Chênh lệch (VNĐ)",
                            "Ngày đối soát"
                        ],
                        "Giá trị": [
                            total_usd,
                            total_vnd,
                            exchange_rate if exchange_rate else "",
                            total_qd,
                            chenh_lech,
                            datetime.now().strftime("%Y-%m-%d %H:%M")
                        ]
                    })
                    payment_summary.to_excel(writer, index=False, sheet_name="Payment_Summary")

                towrite.seek(0)

                st.download_button(
                    label="⬇️ Tải file Excel kết quả đối soát (3 sheet đầy đủ)",
                    data=towrite,
                    file_name=f"doi_soat_MS365_full_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Lỗi trong quá trình xử lý: {e}")

        # ------------------ OTHER SERVICES ------------------
        else:
            st.info(
                f"Hiện chưa định nghĩa logic đối soát riêng cho dịch vụ: **{service_type}**. "
                "Bạn có thể sử dụng tính năng này cho MS365 trước."
            )
