import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import time

# Quan trọng: Đây là nơi ứng dụng Giao diện (Thân xác) gọi đến thư viện Logic (Não)
# Dòng này chỉ hoạt động sau khi bạn triển khai lên Streamlit Cloud với Secrets đúng cách.
try:
    from core_logic.processor import DataProcessor
    from core_logic.word_handler import process_single_document_to_buffer
except ImportError:
    st.error("Lỗi nghiêm trọng: Không thể nhập thư viện 'core_logic'. Hãy chắc chắn bạn đã triển khai đúng cách.")
    # Mô phỏng class để app không bị crash hoàn toàn khi chạy local
    class DataProcessor:
        def __init__(self, *args, **kwargs): pass
        def check(self): return [], [], [], pd.DataFrame()
        def export_clean_file_to_buffer(self): return BytesIO()
    def process_single_document_to_buffer(self, *args, **kwargs): return BytesIO(), set()

# --- Khởi tạo Session State ---
if 'df' not in st.session_state:
    st.session_state.df = None
    st.session_state.data_config = None
    st.session_state.data_checker = None
    st.session_state.selected_sheet = None
    st.session_state.uploaded_excel_name = ""

# --- Giao diện Streamlit ---
st.set_page_config(page_title="Trình tạo Hồ sơ Tự động", page_icon="📄", layout="wide")

# --- Sidebar ---
with st.sidebar:
    st.title("⚙️ Bảng Điều Khiển")
    st.markdown("---")
    
    uploaded_excel = st.file_uploader("1. Tải lên file Excel dữ liệu:", type=["xlsx", "xls"])
    
    if uploaded_excel:
        if uploaded_excel.name != st.session_state.uploaded_excel_name:
            st.session_state.uploaded_excel_name = uploaded_excel.name
            st.session_state.df = None # Reset khi có file mới
        
        try:
            xls = pd.ExcelFile(uploaded_excel)
            sheet_names = xls.sheet_names
            st.session_state.selected_sheet = st.selectbox("Chọn sheet dữ liệu:", sheet_names)
            
            if st.button("Tải Dữ liệu từ Sheet"):
                st.session_state.df = pd.read_excel(uploaded_excel, sheet_name=st.session_state.selected_sheet)
                st.session_state.data_config = None # Reset config khi tải lại data
                st.success(f"Đã tải {len(st.session_state.df)} dòng từ sheet '{st.session_state.selected_sheet}'.")

        except Exception as e:
            st.error(f"Lỗi đọc file Excel: {e}")
            st.session_state.df = None

    uploaded_templates = st.file_uploader("2. Tải lên file Word mẫu:", type=["docx"], accept_multiple_files=True)
    st.markdown("---")

# --- Khu vực chính ---
st.title("📄 Trình tạo Hồ sơ Tự động")
st.markdown(f"Chào Boss, chào mừng đến với phần mềm. Thời gian hiện tại: {datetime.now().strftime('%H:%M:%S, %d/%m/%Y')}")

if st.session_state.df is not None:
    tab1, tab2, tab3 = st.tabs(["1️⃣ Cấu hình & Kiểm tra", "2️⃣ Tạo Hồ sơ", "📈 Xem Dữ liệu"])

    with tab1:
        st.header("Cấu hình và Kiểm tra Dữ liệu")
        with st.form("data_check_form"):
            st.subheader("Cấu hình cột bắt buộc và giá trị hợp lệ")
            
            cols = st.columns(3)
            column_names = st.session_state.df.columns
            config_inputs = {}

            for i, col_name in enumerate(column_names):
                with cols[i % 3]:
                    st.markdown(f"**{col_name}**")
                    is_mandatory = st.checkbox("Bắt buộc", key=f"mand_{col_name}")
                    valid_values = st.text_input("Giá trị hợp lệ (cách nhau bởi dấu phẩy)", key=f"valid_{col_name}")
                    config_inputs[col_name] = (is_mandatory, valid_values)
            
            submitted = st.form_submit_button("Lưu Cấu hình")
            if submitted:
                st.session_state.data_config = {}
                for col_name, (is_mandatory, valid_values_str) in config_inputs.items():
                    st.session_state.data_config[col_name] = {
                        "mandatory": is_mandatory,
                        "valid_values": [v.strip() for v in valid_values_str.split(',')] if valid_values_str else None
                    }
                st.success("Đã lưu cấu hình kiểm tra!")

        if st.session_state.data_config:
            if st.button("🔎 Bắt đầu Kiểm tra Dữ liệu"):
                with st.spinner("Đang kiểm tra..."):
                    checker = DataProcessor(st.session_state.df, st.session_state.data_config, log_func=st.info)
                    errors, warnings, _, _ = checker.check()
                    st.session_state.data_checker = checker # Lưu lại để xuất file
                
                if not errors and not warnings:
                    st.success("✅ Dữ liệu hoàn toàn hợp lệ!")
                else:
                    st.warning(f"Phát hiện {len(errors)} lỗi và {len(warnings)} cảnh báo.")
                    if errors:
                        with st.expander("❌ Xem các lỗi nghiêm trọng", expanded=True):
                            for error in errors: st.error(error)
                    if warnings:
                        with st.expander("⚠️ Xem các cảnh báo"):
                            for warning in warnings: st.warning(warning)
            
            if st.session_state.data_checker:
                clean_file_buffer = st.session_state.data_checker.export_clean_file_to_buffer()
                st.download_button(
                    label="📥 Tải về File Excel đã làm sạch",
                    data=clean_file_buffer,
                    file_name="cleaned_data.xlsx"
                )

    with tab2:
        st.header("Tạo Hồ sơ hàng loạt")
        if not uploaded_templates:
            st.warning("Vui lòng tải lên ít nhất một file Word mẫu ở thanh bên.")
        else:
            if st.button("🚀 Bắt đầu Tạo Toàn bộ Hồ sơ", type="primary"):
                with st.spinner("Đang tạo các file Word..."):
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        progress_bar = st.progress(0)
                        total_rows = len(st.session_state.df)
                        
                        # Logic tạo hồ sơ
                        # Giả định: cột đầu tiên trong excel là tên template cần dùng
                        template_col_name = st.session_state.df.columns[0]
                        
                        for i, row in st.session_state.df.iterrows():
                            template_name_in_row = str(row[template_col_name]).strip()
                            
                            # Tìm template phù hợp
                            chosen_template_buffer = None
                            for template_file in uploaded_templates:
                                if template_name_in_row.lower() in template_file.name.lower():
                                    chosen_template_buffer = BytesIO(template_file.getvalue())
                                    break
                            
                            if chosen_template_buffer:
                                mapping = row.to_dict()
                                doc_buffer, warnings = process_single_document_to_buffer(chosen_template_buffer, mapping)
                                output_filename = f"{template_name_in_row}_{i+1}.docx"
                                zip_file.writestr(output_filename, doc_buffer.getvalue())
                            
                            progress_bar.progress((i + 1) / total_rows)
                            time.sleep(0.05) # Giả lập độ trễ

                    st.success("Hoàn thành! Tải file ZIP chứa tất cả hồ sơ ở dưới.")
                    st.download_button(
                        label="📥 Tải về File ZIP chứa Toàn bộ Hồ sơ",
                        data=zip_buffer,
                        file_name="Tat_ca_Ho_so.zip",
                        mime="application/zip"
                    )

    with tab3:
        st.header("Xem trước Dữ liệu")
        st.dataframe(st.session_state.df)

else:
    st.info("Chào mừng Boss! Vui lòng tải lên file Excel ở thanh bên trái để bắt đầu.")