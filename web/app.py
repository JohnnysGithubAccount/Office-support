import streamlit as st
import os
import shutil
import zipfile
from io import BytesIO
from backend import get_data, fill_invitations, merge_documents, combine_all_docx
import time
import atexit

# Tạo một thư mục tạm
TEMP_DIR = "temp_results"
os.makedirs(TEMP_DIR, exist_ok=True)

st.sidebar.title("Cài đặt Format")
selected_placeholder = st.sidebar.selectbox(
    "Chọn loại định dạng Placeholder:",
    options=["«Key_word»", "[Key_word]", "«Key word»", "[Key word]"]
)

file_prefix = st.sidebar.text_input(
    "Nhập tên cho file:",
    value="Document"
)

save_option = st.sidebar.selectbox(
    "Chọn cách lưu:",
    options=[
        "Nhiều file",
        "Một file (Để in)"
    ]
)

isTest = st.sidebar.selectbox(
    "Có muốn test thử trước 1 file không:",
    options=["Không", "Có"]
)

percentage_text = st.sidebar.empty()
time_estimation_text = st.sidebar.empty()


# Hàm dọn dẹp khi kết thúc
def cleanup_temp_folder():
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)  # Xóa toàn bộ thư mục tạm


atexit.register(cleanup_temp_folder)

# Tiêu đề ứng dụng
st.title("Tạo Tài Liệu Tự Động 📄")
st.write("**Upload file Excel và file Word để tạo tài liệu!**")

# Upload file template Word
uploaded_template = st.file_uploader("Upload file Word (.docx) làm template", type=["docx"])
# Upload file dữ liệu Excel
uploaded_excel = st.file_uploader("Upload file Excel (.xlsx) để lấy dữ liệu", type=["xlsx"])

if uploaded_template and uploaded_excel:
    st.success("File đã được tải lên thành công!")

    # Lưu file tạm
    template_path = os.path.join(TEMP_DIR, "temp_template.docx")
    with open(template_path, "wb") as f:
        f.write(uploaded_template.getbuffer())

    excel_path = os.path.join(TEMP_DIR, "temp_data.xlsx")
    with open(excel_path, "wb") as f:
        f.write(uploaded_excel.getbuffer())

    iterations, df, columns = get_data(excel_path)

    selected_columns_for_filename = []
    if save_option == "Nhiều file":
        selected_columns_for_filename = st.sidebar.multiselect(
            "Chọn các cột để thêm vào tên file:",
            options=columns,
            default=[]  # Không chọn cột nào mặc định
        )

    # Thư mục đầu ra
    output_folder = os.path.join(TEMP_DIR, "results")
    os.makedirs(output_folder, exist_ok=True)

    if st.button("Bắt đầu xử lý"):
        # Đọc file Excel để lấy dữ liệu
        if isTest == "Có":
            iterations = 1

        # Thanh tiến trình
        progress_bar = st.progress(0)

        start_time = time.time()  # Bắt đầu tính thời gian xử lý

        # Xử lý từng dòng dữ liệu
        documents = []
        file_names = []

        if save_option == "Nhiều file":
            # Clear outputfolder
            for file in os.listdir(output_folder):
                file_path = os.path.join(output_folder, file)
                if os.path.isfile(file_path):
                    os.remove(file_path)

        for idx in range(iterations):
            data = {}
            row = df.iloc[idx, :]
            for column in columns:
                data[column.strip()] = row[column]

            # Gọi hàm fill_invitations để xử lý từng tài liệu
            doc = fill_invitations(template_path, data, selected_placeholder)

            file_suffix = "_".join([str(row[col]) for col in selected_columns_for_filename if col in row.index])
            file_path = os.path.join(output_folder,
                                     f"{file_prefix}"
                                     f"{('_' + f'{idx}') if save_option == 'Một file (Để in)' else ''}"
                                     f"{'_' if file_suffix == '' else ''}"
                                     f"{file_suffix}.docx")
            file_names.append(file_path)
            # file_path = os.path.join(output_folder, f"{file_prefix}{idx + 1}.docx")
            doc.save(file_path)

            if save_option != "Nhiều file":
                documents.append(doc)

            # Cập nhật thanh tiến trình
            percentage_complete = int(((idx + 1) / iterations) * 100)
            progress_bar.progress(percentage_complete / 100)

            # Hiển thị phần trăm hoàn thành
            percentage_text.text(f"Đã hoàn thành: {percentage_complete}%")

            # Tính thời gian dự kiến còn lại
            elapsed_time = time.time() - start_time
            avg_time_per_file = elapsed_time / (idx + 1)
            estimated_time_left = avg_time_per_file * (iterations - idx - 1)
            time_estimation_text.text(f"Thời gian dự kiến còn lại: {round(estimated_time_left, 2)} giây")


        if save_option == "Nhiều file":
              # Tạo file zip để tải xuống
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_folder):
                    for file in files:
                        zipf.write(os.path.join(root, file), arcname=file)

            zip_buffer.seek(0)

            # Hiển thị nút tải file zip
            st.sidebar.success("Tài liệu đã được xử lý thành công! Tải file zip bên dưới:")
            st.sidebar.download_button(
                label="Tải file ZIP 📂",
                data=zip_buffer,
                file_name="documents.zip",
                mime="application/zip"
            )
        else:
            # Merge all documents into a single document
            # merged_doc = merge_documents(documents)
            #
            # # Save the merged document
            # merged_output_path = os.path.join(output_folder, "merged_document.docx")
            # merged_doc.save(merged_output_path)

            merged_output_path = f"{file_prefix}.docx"
            print(len(file_names))

            combine_all_docx(
                filename_master=file_names[0]
,               files_list=file_names[1:],
                output_path=merged_output_path
            )

            # Đọc file đã lưu để chuẩn bị tải xuống
            with open(merged_output_path, "rb") as f:
                merged_doc_data = f.read()

            # Hiển thị nút tải xuống file .docx
            st.sidebar.success("Tài liệu đã được xử lý thành công! Tải file .docx bên dưới:")
            st.sidebar.download_button(
                label="Tải file 📄",
                data=merged_doc_data,
                file_name=f"{file_prefix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )