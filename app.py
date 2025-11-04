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
        # pandas NaN detection
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
            # header ở dòng 3 cho file NCC MS365 của bạn
            df = pd.read_excel(f, header=2, dtype=object)
        else:
            if f.name.endswith(".csv"):
                df = pd.read_csv(f, dtype=object)
            else:
                df = pd.read_excel(f, dtype=object)
        # đảm bảo tất cả column names là str, tránh trường hợp column name là datetime
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

        if df_vendor is None or df_internal is None:
            st.error("Không thể đọc file. Hãy kiểm tra định dạng file (xlsx/csv).")
        elif service_type == "MS365":
            st.subheader("🔍 Đang xử lý đối soát Microsoft 365...")
            try:
                # Chuẩn hóa dữ liệu NCC — dùng safe_str trên column values khi cần
                # map cột nếu có tên mặc định
                df_vendor = df_vendor.copy()
                df_vendor = df_vendor.rename(columns={
                    "Row Labels": "Plan",
                    "Sum of Partner Cost (USD)": "USD",
                    "Sum of Partner Cost (VND)": "VND"
                })
                if "Plan" not in df_vendor.columns:
                    # Hơi dự phòng: thử tìm cột chứa "row" và "label"
                    for c in df_vendor.columns:
                        lc = safe_str(c).lower()
                        if "row" in lc and "label" in lc:
                            df_vendor = df_vendor.rename(columns={c: "Plan"})
                            break

                if "Plan" not in df_vendor.columns:
                    raise Exception("Không tìm thấy cột Plan (Row Labels) trong file Nhà cung cấp. Vui lòng kiểm tra header (dòng 3).")

                # Drop rows không có Plan
                df_vendor = df_vendor[df_vendor["Plan"].notna()]

                # Chuẩn hóa nội dung các cột vendor (ép thành str để tránh lỗi)
                for col in ["Plan", "USD", "VND"]:
                    if col in df_vendor.columns:
                        df_vendor[col] = df_vendor[col].apply(lambda x: safe_str(x))

                # Chuẩn hóa dữ liệu nội bộ
                df_internal = df_internal.copy()

                # Tìm cột Description/Product/Quantity
                desc_col = None
                qty_col = None
                for c in df_internal.columns:
                    lc = safe_str(c).lower()
                    if "description" in lc or "product" in lc or "recurring" in lc or "plan" in lc:
                        desc_col = c if desc_col is None else desc_col
                    if "quantity" in lc or "qty" in lc or "amount" in lc:
                        qty_col = c if qty_col is None else qty_col

                if desc_col is None:
                    desc_col = df_internal.columns[0]
                if qty_col is None:
                    df_internal["__qty__"] = 1
                    qty_col = "__qty__"

                # Ép qty thành numeric an toàn
                df_internal[qty_col] = pd.to_numeric(df_internal[qty_col].apply(lambda x: safe_str(x)), errors="coerce").fillna(0)

                # Group internal
                internal_group = (
                    df_internal.groupby(desc_col, as_index=False)
                    .agg({qty_col: "sum"})
                    .rename(columns={desc_col: "Plan", qty_col: "Qty_Internal"})
                )

                # So khớp tên Plan (fuzzy)
                matched_rows = []
                # convert vendor USD/VND to safe numeric strings where needed later
                for _, vendor_row in df_vendor.iterrows():
                    v_plan = normalize_text(vendor_row.get("Plan", ""))
                    best_match = None
                    best_score = 0
                    for _, internal_row in internal_group.iterrows():
                        i_plan = normalize_text(internal_row.get("Plan", ""))
                        score = fuzzy_match(v_plan, i_plan)
                        if score > best_score:
                            best_score = score
                            best_match = internal_row
                    usd_val = safe_str(vendor_row.get("USD", ""))
                    vnd_val = safe_str(vendor_row.get("VND", ""))
                    matched_rows.append({
                        "Plan": safe_str(vendor_row.get("Plan", "")),
                        "USD": usd_val,
                        "VND": vnd_val,
                        "Qty_Internal": int(best_match["Qty_Internal"]) if best_match is not None else 0,
                        "Match_Score (%)": round(best_score * 100, 1)
                    })

                result = pd.DataFrame(matched_rows)

                # Xử lý tỷ giá (nếu có)
                if exchange_rate:
                    result["VND_Quydoi"] = pd.to_numeric(result["USD"].apply(lambda x: safe_str(x)), errors="coerce").fillna(0) * exchange_rate
                    result["VND_Quydoi"] = result["VND_Quydoi"].astype(int)

                # Tổng hợp
                result["USD_num"] = pd.to_numeric(result["USD"].apply(lambda x: safe_str(x)), errors="coerce").fillna(0)
                result["VND_num"] = pd.to_numeric(result["VND"].apply(lambda x: safe_str(x)), errors="coerce").fillna(0)
                total_usd = result["USD_num"].sum()
                total_vnd = result["VND_num"].sum()
                total_qd = result["VND_Quydoi"].sum() if "VND_Quydoi" in result else None

                # Hiển thị kết quả
                st.success("✅ Đối soát hoàn tất!")
                st.dataframe(result, use_container_width=True)

                st.markdown("### 📊 Tổng hợp")
                st.write(f"**Tổng (USD):** {total_usd:,.2f}")
                st.write(f"**Tổng (VND - NCC):** {total_vnd:,.0f}")
                if exchange_rate:
                    st.write(f"**Tổng (VND quy đổi):** {total_qd:,.0f}")

                # Xuất file Excel
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
            st.info(
                f"Hiện chưa định nghĩa logic đối soát riêng cho dịch vụ: **{service_type}**. "
                "Bạn có thể sử dụng tính năng này cho MS365 trước."
            )
