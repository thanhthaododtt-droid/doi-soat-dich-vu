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
Ứng dụng hỗ trợ đối chiếu dữ liệu giữa **file Nhà cung cấp (NCC)** và **file Nội bộ (PO)**  
Phiên bản này xuất ra **3 sheet**:  
1️⃣ Full_Matched_Detail (toàn bộ dữ liệu 2 chiều)  
2️⃣ Summary (tổng hợp theo từng gói)  
3️⃣ Payment_Summary (báo cáo thanh toán)
"""
)

# ------------------ INPUT ------------------
service_type = st.selectbox(
    "🔹 Chọn loại dịch vụ cần đối soát:",
    ["", "MS365", "ODS License", "SSL", "Google Workspace", "TMQT", "Chứng thư CKS"]
)

exchange_rate = None
if service_type == "MS365":
    st.markdown("💱 **Tùy chọn:** nhập tỷ giá USD → VND để quy đổi tổng thanh toán")
    use_rate = st.checkbox("Nhập tỷ giá quy đổi")
    if use_rate:
        exchange_rate = st.number_input("Tỷ giá (VND / USD):", value=26500, step=100)

col1, col2 = st.columns(2)
with col1:
    vendor_file = st.file_uploader("📤 Upload file Nhà cung cấp (NCC)", type=["xlsx", "xls", "csv"], key="vendor")
with col2:
    internal_file = st.file_uploader("📥 Upload file Nội bộ (PO)", type=["xlsx", "xls", "csv"], key="internal")

# ------------------ HELPER ------------------
def safe_str(x):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        if hasattr(x, "strftime"):
            return x.strftime("%Y-%m-%d")
        return str(x)
    except Exception:
        return str(x)

def read_file(f, service_type=None):
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
    try:
        return safe_str(s).strip().lower()
    except Exception:
        return safe_str(s)

def fuzzy_match(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ------------------ MAIN ------------------
if st.button("🚀 Tiến hành đối soát"):
    if not service_type:
        st.warning("⚠️ Vui lòng chọn loại dịch vụ.")
    elif not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ cả hai file (NCC & PO).")
    else:
        df_vendor = read_file(vendor_file, service_type)
        df_internal = read_file(internal_file, service_type)

        if service_type == "MS365":
            st.subheader("🔍 Đang xử lý đối soát Microsoft 365...")

            try:
                # Chuẩn hóa dữ liệu NCC
                df_vendor = df_vendor.rename(columns={
                    "Row Labels": "Plan",
                    "Sum of Partner Cost (USD)": "USD",
                    "Sum of Partner Cost (VND)": "VND"
                })
                df_vendor = df_vendor.dropna(subset=["Plan"])
                df_vendor = df_vendor[df_vendor["Plan"] != "Row Labels"]

                # Chuẩn hóa dữ liệu nội bộ
                desc_col, qty_col = None, None
                for c in df_internal.columns:
                    lc = safe_str(c).lower()
                    if "description" in lc or "product" in lc or "plan" in lc:
                        desc_col = c
                    if "quantity" in lc or "qty" in lc:
                        qty_col = c
                if desc_col is None:
                    desc_col = df_internal.columns[0]
                if qty_col is None:
                    df_internal["__qty__"] = 1
                    qty_col = "__qty__"

                df_internal[qty_col] = pd.to_numeric(df_internal[qty_col].apply(lambda x: safe_str(x)), errors="coerce").fillna(0)

                # --- Fuzzy match 2 chiều ---
                matched_rows = []
                used_po = set()
                for _, vrow in df_vendor.iterrows():
                    v_plan = normalize_text(vrow["Plan"])
                    best_match = None
                    best_score = 0
                    for idx, irow in df_internal.iterrows():
                        i_plan = normalize_text(irow[desc_col])
                        score = fuzzy_match(v_plan, i_plan)
                        if score > best_score:
                            best_score = score
                            best_match = (idx, irow)
                    combined = {}
                    for c in df_vendor.columns:
                        combined[f"NCC_{c}"] = vrow.get(c, "")
                    if best_match and best_score >= 0.4:
                        idx, irow = best_match
                        used_po.add(idx)
                        for c in df_internal.columns:
                            combined[f"PO_{c}"] = irow.get(c, "")
                        combined["Trạng thái đối soát"] = "✅ Đã khớp" if best_score >= 0.6 else "⚠️ Khớp thấp"
                    else:
                        for c in df_internal.columns:
                            combined[f"PO_{c}"] = ""
                        combined["Trạng thái đối soát"] = "⚠️ Thiếu ở PO"
                    combined["Match_Score (%)"] = round(best_score * 100, 1)
                    matched_rows.append(combined)

                # Thêm các PO chưa match
                for idx, irow in df_internal.iterrows():
                    if idx not in used_po:
                        combined = {}
                        for c in df_vendor.columns:
                            combined[f"NCC_{c}"] = ""
                        for c in df_internal.columns:
                            combined[f"PO_{c}"] = irow.get(c, "")
                        combined["Trạng thái đối soát"] = "❌ Thiếu ở NCC"
                        combined["Match_Score (%)"] = 0
                        matched_rows.append(combined)

                result_full = pd.DataFrame(matched_rows)

                # --- Tính tỷ giá và tổng hợp ---
                if exchange_rate:
                    result_full["VND_Quydoi"] = pd.to_numeric(result_full["NCC_USD"], errors="coerce").fillna(0) * exchange_rate
                    result_full["VND_Quydoi"] = result_full["VND_Quydoi"].astype(int)

                result_full["USD_num"] = pd.to_numeric(result_full["NCC_USD"], errors="coerce").fillna(0)
                result_full["VND_num"] = pd.to_numeric(result_full["NCC_VND"], errors="coerce").fillna(0)

                total_usd = result_full["USD_num"].sum()
                total_vnd = result_full["VND_num"].sum()
                total_qd = result_full["VND_Quydoi"].sum() if "VND_Quydoi" in result_full else 0
                chenh_lech = total_qd - total_vnd if exchange_rate else 0

                # --- Summary (Pivot dạng Plan) ---
                summary = (
                    result_full.groupby("NCC_Plan", as_index=False)
                    .agg({
                        "USD_num": "sum",
                        "VND_num": "sum",
                        "VND_Quydoi": "sum" if "VND_Quydoi" in result_full else "mean",
                    })
                )
                summary.rename(columns={
                    "NCC_Plan": "Plan",
                    "USD_num": "Tổng USD",
                    "VND_num": "Tổng VND",
                    "VND_Quydoi": "Tổng VND Quy đổi"
                }, inplace=True)

                # --- Xuất Excel ---
                towrite = io.BytesIO()
                with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                    result_full.to_excel(writer, index=False, sheet_name="Full_Matched_Detail")
                    summary.to_excel(writer, index=False, sheet_name="Summary")

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
                st.success("✅ Đối soát hoàn tất! Xuất dữ liệu 3 sheet đầy đủ.")
                st.download_button(
                    label="⬇️ Tải file Excel đối soát (Full + Summary + Payment)",
                    data=towrite,
                    file_name=f"doi_soat_MS365_full_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Lỗi trong quá trình xử lý: {e}")

        else:
            st.info("Hiện chỉ hỗ trợ đối soát cho **MS365** trong phiên bản này.")
