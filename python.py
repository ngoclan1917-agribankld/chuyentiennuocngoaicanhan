import streamlit as st
import pandas as pd

# ==========================
# ⚙️ Cấu hình giao diện
# ==========================
st.set_page_config(page_title="Chatbot Trắc Nghiệm", page_icon="🤖", layout="wide")
st.title("🤖 Chatbot Trắc nghiệm")
st.markdown("📂 **Trái:** Vui lòng Tải các file câu hỏi theo nghiệp vụ — 💬 **Phải:** Tra cứu đáp án đúng.")

# ==========================
# 🧠 Session State
# ==========================
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = {}
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0  # để reset file_uploader khi xóa

# ==========================
# 📏 Tăng khoảng cách giữa 2 vùng
# ==========================
st.markdown(
    """
    <style>
    div[data-testid="column"]:first-child {
        margin-right: 60px !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ==========================
# 🧭 2 CỘT GIAO DIỆN
# ==========================
col1, col2 = st.columns([1, 2])

# ==========================
# 📂 CỘT TRÁI: TẢI FILE
# ==========================
with col1:
    st.subheader("📂 Tải file Excel")

    def read_file_from_header(file):
        """Tìm dòng chứa 'CÂU HỎI' và đọc dữ liệu từ đó trở xuống"""
        df_raw = pd.read_excel(file, header=None)
        header_row_idx = None
        for i, row in df_raw.iterrows():
            if any(str(cell).strip().upper() == "CÂU HỎI" for cell in row):
                header_row_idx = i
                break
        if header_row_idx is None:
            raise ValueError("❌ Không tìm thấy dòng tiêu đề có cột 'CÂU HỎI'.")
        df = pd.read_excel(file, header=header_row_idx)
        return df

    uploaded_files = st.file_uploader(
        "Chọn file Excel (có thể nhiều)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        for file in uploaded_files:
            if file.name not in st.session_state.uploaded_files:
                try:
                    df = read_file_from_header(file)
                    st.session_state.uploaded_files[file.name] = df
                except Exception as e:
                    st.error(f"Lỗi đọc file {file.name}: {e}")

    if st.session_state.uploaded_files:
        if st.button("🧹 Xóa tất cả file đã tải"):
            st.session_state.uploaded_files.clear()
            st.session_state.uploader_key += 1  # reset file_uploader
            st.rerun()

# ==========================
# 💬 CỘT PHẢI: CHATBOT
# ==========================
with col2:
    st.subheader("💬 Chatbot tra cứu đáp án")

    if st.session_state.uploaded_files:
        # Gộp dữ liệu từ tất cả file đã tải
        combined_df = pd.concat(st.session_state.uploaded_files.values(), ignore_index=True)
        combined_df.columns = [str(c).strip().upper() for c in combined_df.columns]

        user_input = st.text_input(
            "🔎 Nhập từ khóa câu hỏi và nhấn Enter hoặc bấm 'Tìm kiếm'"
        )
        search_btn = st.button("Tìm kiếm")

        def tim_cau_hoi(keyword, dataframe):
            kw = keyword.lower().strip()
            return dataframe[dataframe['CÂU HỎI'].str.lower().str.contains(kw, na=False)]

        if user_input or search_btn:
            if user_input:
                results = tim_cau_hoi(user_input, combined_df)
                if results.empty:
                    st.warning("❌ Không tìm thấy câu hỏi nào phù hợp.")
                else:
                    for _, row in results.iterrows():
                        try:
                            dap_an_dung = int(row['ĐÁP ÁN ĐÚNG'])
                            noi_dung_dap_an = row[f'ĐÁP ÁN {dap_an_dung}']
                            st.markdown(f"**📌 Câu hỏi:** {row['CÂU HỎI']}")
                            st.success(f"✅ **Đáp án đúng:** {noi_dung_dap_an}")
                            st.divider()
                        except Exception:
                            st.error("⚠️ File không đúng định dạng cột đáp án.")
    else:
        st.info("📌 Vui lòng tải ít nhất một file trước khi tra cứu.")

# ==========================
# 📘 HƯỚNG DẪN
# ==========================
with st.expander("📘 Hướng dẫn sử dụng"):
    st.write("- Có thể tải nhiều file Excel cùng lúc.")
    st.write("- Tự động phát hiện dòng tiêu đề có cột 'CÂU HỎI'.")
    st.write("- Sau khi tải, có thể xóa tất cả file bằng nút 🧹.")
    st.write("- Nhập từ khóa câu hỏi và nhấn Enter hoặc nút 'Tìm kiếm'.")
    st.write("- Cột bắt buộc: STT | CÂU HỎI | ĐÁP ÁN 1–4 | ĐÁP ÁN ĐÚNG.")
