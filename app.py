import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher
from datetime import datetime

# ------------------ CONFIG ------------------
st.set_page_config(page_title="Công cụ đối soát dịch vụ nội bộ", layout="wide")

st.title("📊 CÔNG CỤ ĐỐI SOÁT DỊCH VỤ NỘI BỘ")
st.markdown("Ứng dụng dùng để đối chiếu dữ liệu giữa **file Nhà cung cấp** và **file Nội bộ (PO)** cho các dịch vụ CNTT.")

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
def read_file(f):
    if f is None:
        return None
    try:
        if f.name.endswith(".csv"):
            return pd.read_csv(f, dtype=str)
        else:
            return pd.read_excel(f, dtype=str)
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return None

def normalize_text(s):
    if pd.isna(s): return ""
    return str(s).strip().lower()

def fuzzy_match(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ------------------ MAIN ------------------
if st.button("🚀 Tiến hành đối soát"):
    if not service_type:
        st.warning("⚠️ Vui lòng chọn loại dịch vụ.")
    elif not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ cả hai file (Nhà cung cấp & Nội bộ).")
    else:
        df_vendor = read_file(vendor_file)
        df_internal = read_file(internal_file)

        if service_type == "MS365":
            st.subheader("🔍 Đang xử lý đối soát Microsoft 365...")
            try:
                # Lấy dữ liệu NCC
                df_vendor.columns = [c.strip() for c in df_vendor.columns]
                df_vendor = df_vendor.rename(columns={
                    "Row Labels": "Plan",
                    "Sum of Partner Cost (USD)": "USD",
                    "Sum of Partner Cost (VND)": "VND"
                })
                df_vendor = df_vendor.dropna(subset=["Plan"])
                df_vendor = df_vendor[df_vendor["Plan"] != "Row Labels"]

                # Lấy dữ liệu nội bộ
                df_internal.columns = [c.strip() for c in df_internal.columns]
                internal_group = (
                    df_internal.groupby("Description", as_index=False)
                    .agg({"Quantity": "sum"})
                    .rename(columns={"Description": "Plan", "Quantity": "Qty_Internal"})
                )

                # So khớp tên Plan (fuzzy)
                matched_rows = []
                for _, vendor_row in df_vendor.iterrows():
                    v_plan = normalize_text(vendor_row["Plan"])
                    best_match = None
                    best_score = 0
                    for _, internal_row in internal_group.iterrows():
                        i_plan = normalize_text(internal_row["Plan"])
                        score = fuzzy_match(v_plan, i_plan)
                        if score > best_score:
                            best_score = score
                            best_match = internal_row
                    if best_match is not None and best_score >= 0.6:
                        matched_rows.append({
                            "Plan": vendor_row["Plan"],
                            "USD": vendor_row["USD"],
                            "VND": vendor_row["VND"],
                            "Qty_Internal": best_match["Qty_Internal"],
                            "Match_Score": round(best_score * 100, 1)
                        })
                    else:
                        matched_rows.append({
                            "Plan": vendor_row["Plan"],
                            "USD": vendor_row["USD"],
                            "VND": vendor_row["VND"],
                            "Qty_Internal": None,
                            "Match_Score": round(best_score * 100, 1)
                        })

                result = pd.DataFrame(matched_rows)

                # Xử lý tỷ giá
                if exchange_rate:
                    result["VND_Quydoi"] = pd.to_numeric(result["USD"], errors="coerce").fillna(0) * exchange_rate
                    result["VND_Quydoi"] = result["VND_Quydoi"].astype(int)

                # Tổng hợp
                result["USD"] = pd.to_numeric(result["USD"], errors="coerce").fillna(0)
                result["VND"] = pd.to_numeric(result["VND"], errors="coerce").fillna(0)
                total_usd = result["USD"].sum()
                total_vnd = result["VND"].sum()
                total_qd = result["VND_Quydoi"].sum() if "VND_Quydoi" in result else None

                st.success("✅ Đối soát hoàn tất!")
                st.dataframe(result)

                st.markdown("### 📊 Tổng hợp")
                st.write(f"**Tổng (USD):** {total_usd:,.2f}")
                st.write(f"**Tổng (VND - NCC):** {total_vnd:,.0f}")
                if exchange_rate:
                    st.write(f"**Tổng (VND quy đổi):** {total_qd:,.0f}")

                # Xuất Excel
                towrite = io.BytesIO()
                with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                    result.to_excel(writer, index=False, sheet_name="MS365_Matched")
                    summary = pd.DataFrame({
                        "Tổng USD": [total_usd],
                        "Tổng VND NCC": [total_vnd],
                        "Tổng VND Quy đổi": [total_qd if total_qd else ""],
                        "Tỷ giá": [exchange_rate if exchange_rate else ""],
                        "Ngày đối soát": [datetime.now().strftime("%Y-%m-%d %H:%M")]
                    })
                    summary.to_excel(writer, index=False, sheet_name="Summary")
                towrite.seek(0)

                st.download_button(
                    label="⬇️ Tải file Excel kết quả đối soát",
                    data=towrite,
                    file_name=f"doi_soat_MS365_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Lỗi trong quá trình xử lý: {e}")
        else:
            st.info(f"Hiện chưa định nghĩa logic đối soát riêng cho dịch vụ: **{service_type}**. "
                    "Bạn có thể sử dụng tính năng này cho MS365 trước.")
