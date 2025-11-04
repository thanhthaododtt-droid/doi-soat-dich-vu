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

        # Xuất file Excel (gồm 3 sheet)
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
