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
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
            lines = text.splitlines()
            for line in lines:
                line = line.strip()

                # Bỏ qua dòng tiêu đề hoặc chứa từ khóa không phải dữ liệu
                if re.search(r"STT|Họ đệm|Điểm|Hệ số|Mã số sinh viên|Xếp loại|Ghi chú", line, re.IGNORECASE):
                    continue

                # Tách theo khoảng trắng lớn (dấu hiệu phân cột)
                parts = re.split(r'\s{2,}', line)
                if len(parts) < 6:
                    continue  # Không đủ cột

                try:
                    stt = int(parts[0])
                    mssv = parts[1]
                    fullname = parts[2]
                    ho_dem, ten = split_name(fullname)

                    # Tìm điểm (chỉ lấy dạng float x.y), loại bỏ các "V", "--", "vắng thi", "Được dự thi"
                    score_values = [p for p in parts if re.match(r"\d+\.\d{2}", p)]
                    diem_chu = next((p for p in parts if re.match(r"^[ABCDF][+-]?$", p)), None)
                    xep_loai = parts[-2] if len(parts) >= 10 else ""
                    ghi_chu = parts[-1] if len(parts) >= 10 else ""

                    row = {
                        "STT": stt,
                        "Mã số sinh viên": mssv,
                        "Họ đệm": ho_dem,
                        "Tên": ten,
                        "Điểm chữ": diem_chu,
                        "Xếp loại": xep_loai,
                        "Ghi chú": ghi_chu,
                    }

                    if len(score_values) == 5:
                        row.update({
                            "Điểm giữa kỳ": float(score_values[0]),
                            "Điểm thường kỳ": float(score_values[1]),
                            "Điểm thực hành": float(score_values[2]),
                            "Điểm cuối kỳ": float(score_values[3]),
                            "Điểm tổng kết": float(score_values[4]),
                        })
                    elif len(score_values) == 4:
                        row.update({
                            "Điểm giữa kỳ": float(score_values[0]),
                            "Điểm thường kỳ": float(score_values[1]),
                            "Điểm cuối kỳ": float(score_values[2]),
                            "Điểm tổng kết": float(score_values[3]),
                        })
                    elif len(score_values) == 3:
                        row.update({
                            "Điểm cuối kỳ": float(score_values[0]),
                            "Điểm tổng kết": float(score_values[1]),
                        })

                    rows.append(row)

                except Exception as e:
                    st.warning(f"Lỗi dòng: {line} — {str(e)}")
                    continue

    df = pd.DataFrame(rows)
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
