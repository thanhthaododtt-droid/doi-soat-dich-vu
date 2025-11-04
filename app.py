import streamlit as st
import pandas as pd

st.set_page_config(page_title="Công cụ đối soát dịch vụ", layout="wide")

st.title("📊 CÔNG CỤ ĐỐI SOÁT DỊCH VỤ NỘI BỘ")
st.markdown("Ứng dụng nội bộ dùng để đối chiếu dữ liệu giữa file **Nhà cung cấp** và **File nội bộ (PO)**.")

# --- Chọn loại dịch vụ ---
service_type = st.selectbox(
    "🔹 Chọn loại dịch vụ cần đối soát:",
    ["", "MS365", "ODS License", "SSL", "Google Workspace", "TMQT", "Chứng thư CKS"]
)

# --- Upload file ---
col1, col2 = st.columns(2)

with col1:
    vendor_file = st.file_uploader("📤 Upload file từ Nhà cung cấp", type=["xlsx", "xls", "csv"])
with col2:
    internal_file = st.file_uploader("📥 Upload file Nội bộ (PO)", type=["xlsx", "xls", "csv"])

# --- Xử lý ---
if st.button("🚀 Tiến hành đối soát"):
    if not service_type:
        st.warning("⚠️ Vui lòng chọn loại dịch vụ trước khi đối soát.")
    elif not vendor_file or not internal_file:
        st.warning("⚠️ Cần upload đủ cả hai file (Nhà cung cấp & Nội bộ).")
    else:
        def read_file(f):
            if f.name.endswith(".csv"):
                return pd.read_csv(f)
            else:
                return pd.read_excel(f)
        
        df_vendor = read_file(vendor_file)
        df_internal = read_file(internal_file)

        st.success(f"✅ Đã tải đủ dữ liệu cho loại dịch vụ **{service_type}**.")
        st.subheader("📂 File Nhà cung cấp (5 dòng đầu):")
        st.dataframe(df_vendor.head())

        st.subheader("📂 File Nội bộ (5 dòng đầu):")
        st.dataframe(df_internal.head())

        st.info("👉 Bước tiếp theo: thêm logic đối chiếu và xuất file kết quả Excel.")

st.markdown("---")
st.caption("© 2025 - Bộ phận Quản lý Dịch vụ | Ứng dụng Streamlit nội bộ")
