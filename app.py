import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os

# Ẩn thanh công cụ Streamlit và các biểu tượng "Running", "Share"
st.markdown(
    """
    <style>
    [data-testid="stToolbar"] {
        visibility: hidden;
    }
    [data-testid="stStatusWidget"] {
        visibility: hidden;
    }
    [data-testid="stDecoration"] {
        visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.set_page_config(page_title="Trích xuất bảng điểm PDF", layout="wide")
st.title("📄 Trích xuất bảng điểm từ file PDF")
st.markdown("Tải lên file PDF chứa bảng điểm để trích xuất và lưu ra Excel.")

def split_name(fullname):
    """Split a full name into first/middle name (Họ đệm) and last name (Tên)."""
    if not fullname:
        return '', ''
    parts = fullname.strip().split()
    if len(parts) == 1:
        return '', parts[0]
    elif len(parts) == 2:
        return parts[0], parts[1]
    else:
        return ' '.join(parts[:-1]), parts[-1]

def extract_scores_from_pdf(file):
    rows = []
    has_thuongky = False
    has_giua_ky = False
    has_thuc_hanh = False
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                st.warning(f"Không tìm thấy văn bản trên trang {page_num + 1}.")
                continue
            
            lines = text.splitlines()
            for line in lines:
                line = line.strip()

                # Bỏ qua các dòng không chứa điểm
                if not re.search(r"\d+\.\d\d", line):
                    continue

                # Tách dòng theo khoảng trắng nhiều lần (2+)
                parts = re.split(r'\s{2,}', line)
                if len(parts) < 7:
                    continue  # Không đủ dữ liệu
                
                try:
                    stt = int(parts[0])
                    mssv = parts[1]
                    fullname = parts[2].strip()
                    scores = [p for p in parts[3:] if re.match(r'^\d+\.\d\d$', p)]
                    diem_chu = next((p for p in parts if re.match(r'^[ABCDF][+-]?$', p)), None)
                    ho_dem, ten = split_name(fullname)

                    # Khởi tạo row
                    row = {
                        "STT": stt,
                        "Mã số sinh viên": mssv,
                        "Họ đệm": ho_dem,
                        "Tên": ten,
                        "Điểm chữ": diem_chu
                    }

                    # Gán điểm theo độ dài
                    if len(scores) == 5:
                        has_thuongky = True
                        has_giua_ky = True
                        has_thuc_hanh = True
                        row.update({
                            "Điểm giữa kỳ": float(scores[0]),
                            "Điểm thường kỳ": float(scores[1]),
                            "Điểm thực hành": float(scores[2]),
                            "Điểm cuối kỳ": float(scores[3]),
                            "Điểm TB môn học": float(scores[4]),
                        })
                    elif len(scores) == 4:
                        has_thuongky = True
                        has_giua_ky = True
                        row.update({
                            "Điểm giữa kỳ": float(scores[0]),
                            "Điểm thường kỳ": float(scores[1]),
                            "Điểm cuối kỳ": float(scores[2]),
                            "Điểm TB môn học": float(scores[3]),
                        })
                    elif len(scores) == 3:
                        row.update({
                            "Điểm cuối kỳ": float(scores[0]),
                            "Điểm TB môn học": float(scores[1]),
                        })
                    else:
                        st.warning(f"Dòng không xác định được số điểm: {line}")
                        continue

                    rows.append(row)
                except Exception as e:
                    st.warning(f"Lỗi xử lý dòng: {line}. Lỗi: {str(e)}")
                    continue

    df = pd.DataFrame(rows)

    # Drop optional columns if not detected
    if not has_thuc_hanh and "Điểm thực hành" in df.columns:
        df.drop(columns=["Điểm thực hành"], inplace=True)
    if not has_giua_ky and "Điểm giữa kỳ" in df.columns:
        df.drop(columns=["Điểm giữa kỳ"], inplace=True)
    if not has_thuongky and "Điểm thường kỳ" in df.columns:
        df.drop(columns=["Điểm thường kỳ"], inplace=True)
    
    return df


# File upload interface
uploaded_file = st.file_uploader("📌 Tải file PDF bảng điểm:", type="pdf", accept_multiple_files=False, help="File PDF nên dưới 200MB.")
if uploaded_file is not None:
    try:
        df = extract_scores_from_pdf(uploaded_file)
        if not df.empty:
            st.success("✅ Đã trích xuất thành công!")
            st.dataframe(df, use_container_width=True)
            
            # Download button for Excel, using the uploaded file's name
            output = io.BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            
            # Get the uploaded file's name and replace .pdf with .xlsx
            file_name = uploaded_file.name
            excel_file_name = os.path.splitext(file_name)[0] + ".xlsx"
            
            st.download_button(
                label="📥 Tải xuống Excel",
                data=output,
                file_name=excel_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Không trích xuất được dữ liệu từ file PDF.")
    except Exception as e:
        st.error(f"Lỗi xử lý file PDF: {str(e)}")
