import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os

# Ẩn thanh công cụ Streamlit
st.markdown(
    """
    <style>
    [data-testid="stToolbar"] {visibility: hidden;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    [data-testid="stDecoration"] {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True
)

st.set_page_config(page_title="Trích xuất bảng điểm PDF", layout="wide")
st.title("📄 Trích xuất bảng điểm từ file PDF")
st.markdown("Tải lên file PDF chứa bảng điểm để trích xuất và lưu ra Excel.")

def normalize_text(text):
    """Chuẩn hóa văn bản để xử lý lỗi OCR."""
    if not text:
        return text
    text = text.replace('Nguyêa', 'Nguyễn').replace('Hau', 'Hậu').replace('Ken', 'Kém')
    text = text.replace('Hoc lai', 'Học lại').replace('Hoc lai', 'Học lại')
    return text.strip()

def split_name(fullname):
    """Split a full name into first/middle name (Họ đệm) and last name (Tên)."""
    if not fullname or not isinstance(fullname, str):
        return '', ''
    parts = fullname.strip().split()
    if not parts:
        return '', ''
    if len(parts) == 1:
        return '', parts[0]
    return ' '.join(parts[:-1]), parts[-1]

def extract_scores_from_pdf(file):
    """Extract grade data from PDF using table extraction."""
    rows = []
    has_thuongky = False
    has_giua_ky = False
    has_thuc_hanh = False
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                st.warning(f"Không tìm thấy bảng trên trang {page_num + 1}.")
                continue
            
            for table in tables:
                for row in table[1:]:  # Bỏ qua dòng tiêu đề
                    if len(row) < 8:  # Kiểm tra số cột tối thiểu
                        st.warning(f"Dòng không đủ cột trên trang {page_num + 1}: {row}")
                        continue
                    try:
                        stt = int(row[0]) if row[0] else None
                        mssv = row[1] if row[1] else ''
                        fullname = normalize_text(row[2]) if row[2] else ''
                        diem_gk = float(row[3]) if row[3] else 0.0
                        diem_thuongky = float(row[4]) if row[4] else 0.0
                        diem_th = float(row[5]) if row[5] else None
                        diem_cuoi_ky = float(row[6]) if row[6] else 0.0
                        diem_tb = float(row[7]) if row[7] else 0.0
                        diem_chu = row[8] if row[8] else ''
                        ghi_chu = normalize_text(row[9]) if len(row) > 9 else ''
                        
                        if diem_chu not in ['A', 'B', 'C', 'D', 'F']:
                            st.warning(f"Điểm chữ không hợp lệ: {diem_chu} trên dòng: {row}")
                            continue
                        
                        ho_dem, ten = split_name(fullname)
                        row_data = {
                            "STT": stt,
                            "Mã số sinh viên": mssv,
                            "Họ đệm": ho_dem,
                            "Tên": ten,
                            "Điểm giữa kỳ": diem_gk,
                            "Điểm thường kỳ": diem_thuongky,
                            "Điểm cuối kỳ": diem_cuoi_ky,
                            "Điểm TB môn học": diem_tb,
                            "Điểm chữ": diem_chu,
                            "Ghi chú": ghi_chu
                        }
                        if diem_th is not None:
                            row_data["Điểm thực hành"] = diem_th
                            has_thuc_hanh = True
                        has_thuongky = True
                        has_giua_ky = True
                        rows.append(row_data)
                    except Exception as e:
                        st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {row}. Lỗi: {str(e)}")
                        continue
    
    df = pd.DataFrame(rows)
    if not has_thuc_hanh and "Điểm thực hành" in df.columns:
        df = df.drop(columns=["Điểm thực hành"])
    if not has_giua_ky and "Điểm giữa kỳ" in df.columns:
        df = df.drop(columns=["Điểm giữa kỳ"])
    if not has_thuongky and "Điểm thường kỳ" in df.columns:
        df = df.drop(columns=["Điểm thường kỳ"])
    return df

# File upload interface
uploaded_file = st.file_uploader("📌 Tải file PDF bảng điểm:", type="pdf", accept_multiple_files=False, help="File PDF nên dưới 200MB.")
if uploaded_file is not None:
    try:
        df = extract_scores_from_pdf(uploaded_file)
        if not df.empty:
            st.success("✅ Đã trích xuất thành công!")
            st.dataframe(df, use_container_width=True)
            st.info(f"Tổng số dòng trích xuất: {len(df)}")
            if "Ghi chú" in df.columns:
                hoc_lai_rows = df[df["Ghi chú"].str.contains("Học lại", case=False, na=False)]
                st.info(f"Số dòng có ghi chú 'Học lại': {len(hoc_lai_rows)}")
            
            # Download button for Excel
            output = io.BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
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
