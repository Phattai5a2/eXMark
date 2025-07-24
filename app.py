import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os

# Hide Streamlit toolbar and status widgets
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
    """Extract grade data from PDF, handling varying column sets including Ghi chú."""
    rows = []
    has_thuongky = False
    has_giua_ky = False
    has_thuc_hanh = False
    has_ghi_chu = False
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            # Try extracting text first
            text = page.extract_text()
            if text and text.strip():
                st.write(f"**Raw text from page {page_num + 1}:**")
                st.text(text)  # Display raw text for debugging
                lines = text.splitlines()
                for line in lines:
                    # Skip empty or header-like lines
                    if not line.strip() or line.startswith('STT') or line.startswith('Số TT'):
                        continue
                    # Regex patterns with optional Ghi chú
                    pattern_full = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+V\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+([ABCD])\s+(\S+)(?:\s+(.+))?"
                    pattern_no_th = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+V\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+([ABCD])\s+(\S+)(?:\s+(.+))?"
                    pattern_minimal = r"(\d+)\s+(\d+)\s+(.+?)\s+V\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+([ABCD])\s+(\S+)(?:\s+(.+))?"
                    
                    match = re.match(pattern_full, line)
                    if match:
                        has_thuongky = True
                        has_giua_ky = True
                        has_thuc_hanh = True
                        has_ghi_chu = bool(match.group(11))
                        try:
                            stt = int(match.group(1))
                            mssv = match.group(2)
                            fullname = match.group(3).strip()
                            diem_gk = float(match.group(4))
                            diem_thuongky = float(match.group(5))
                            diem_th = float(match.group(6))
                            diem_cuoi_ky = float(match.group(7))
                            diem_tb = float(match.group(8))
                            diem_chu = match.group(9)
                            ghi_chu = match.group(11) if match.group(11) else ""
                            
                            if diem_chu not in ['A', 'B', 'C', 'D']:
                                st.warning(f"Điểm chữ không hợp lệ trên dòng: {line}")
                                continue
                            
                            ho_dem, ten = split_name(fullname)
                            
                            row = {
                                "STT": stt,
                                "Mã số sinh viên": mssv,
                                "Họ đệm": ho_dem,
                                "Tên": ten,
                                "Điểm thường kỳ": diem_thuongky,
                                "Điểm giữa kỳ": diem_gk,
                                "Điểm thực hành": diem_th,
                                "Điểm cuối kỳ": diem_cuoi_ky,
                                "Điểm TB môn học": diem_tb,
                                "Điểm chữ": diem_chu,
                                "Ghi chú": ghi_chu
                            }
                            rows.append(row)
                        except Exception as e:
                            st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                            continue
                    else:
                        match = re.match(pattern_no_th, line)
                        if match:
                            has_thuongky = True
                            has_giua_ky = True
                            has_ghi_chu = bool(match.group(10))
                            try:
                                stt = int(match.group(1))
                                mssv = match.group(2)
                                fullname = match.group(3).strip()
                                diem_thuongky = float(match.group(4))
                                diem_gk = float(match.group(5))
                                diem_cuoi_ky = float(match.group(6))
                                diem_tb = float(match.group(7))
                                diem_chu = match.group(8)
                                ghi_chu = match.group(10) if match.group(10) else ""
                                
                                if diem_chu not in ['A', 'B', 'C', 'D']:
                                    st.warning(f"Điểm chữ không hợp lệ trên dòng: {line}")
                                    continue
                                
                                ho_dem, ten = split_name(fullname)
                                
                                row = {
                                    "STT": stt,
                                    "Mã số sinh viên": mssv,
                                    "Họ đệm": ho_dem,
                                    "Tên": ten,
                                    "Điểm thường kỳ": diem_thuongky,
                                    "Điểm giữa kỳ": diem_gk,
                                    "Điểm cuối kỳ": diem_cuoi_ky,
                                    "Điểm TB môn học": diem_tb,
                                    "Điểm chữ": diem_chu,
                                    "Ghi chú": ghi_chu
                                }
                                rows.append(row)
                            except Exception as e:
                                st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                                continue
                        else:
                            match = re.match(pattern_minimal, line)
                            if match:
                                has_ghi_chu = bool(match.group(8))
                                try:
                                    stt = int(match.group(1))
                                    mssv = match.group(2)
                                    fullname = match.group(3).strip()
                                    diem_cuoi_ky = float(match.group(4))
                                    diem_tb = float(match.group(5))
                                    diem_chu = match.group(6)
                                    ghi_chu = match.group(8) if match.group(8) else ""
                                    
                                    if diem_chu not in ['A', 'B', 'C', 'D']:
                                        st.warning(f"Điểm chữ không hợp lệ trên dòng: {line}")
                                        continue
                                    
                                    ho_dem, ten = split_name(fullname)
                                    
                                    row = {
                                        "STT": stt,
                                        "Mã số sinh viên": mssv,
                                        "Họ đệm": ho_dem,
                                        "Tên": ten,
                                        "Điểm cuối kỳ": diem_cuoi_ky,
                                        "Điểm TB môn học": diem_tb,
                                        "Điểm chữ": diem_chu,
                                        "Ghi chú": ghi_chu
                                    }
                                    rows.append(row)
                                except Exception as e:
                                    st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                                    continue
            
            # Try extracting tables as a fallback
            tables = page.extract_tables()
            if tables:
                st.write(f"**Found tables on page {page_num + 1}:**")
                for table_idx, table in enumerate(tables):
                    st.write(f"Table {table_idx + 1}:")
                    st.write(table)  # Display raw table for debugging
                    # Process table rows (assuming table has a header)
                    for row in table[1:]:  # Skip header row
                        if not row or len(row) < 6:  # Minimum columns for minimal pattern
                            continue
                        try:
                            stt = int(row[0]) if row[0] else None
                            mssv = row[1] if row[1] else ""
                            fullname = row[2].strip() if row[2] else ""
                            ho_dem, ten = split_name(fullname)
                            # Adjust column indices based on expected structure
                            col_offset = 3  # Start after STT, MSSV, Fullname
                            scores = [float(x) if x and x.replace('.', '').isdigit() else None for x in row[col_offset:]]
                            diem_chu = row[-3] if len(row) >= 3 else ""
                            ghi_chu = row[-1] if len(row) >= 1 and row[-1] not in ['A', 'B', 'C', 'D'] else ""
                            
                            row_data = {
                                "STT": stt,
                                "Mã số sinh viên": mssv,
                                "Họ đệm": ho_dem,
                                "Tên": ten,
                            }
                            score_idx = 0
                            if len(scores) >= 5:  # Full pattern
                                row_data["Điểm giữa kỳ"] = scores[score_idx]
                                row_data["Điểm thường kỳ"] = scores[score_idx + 1]
                                row_data["Điểm thực hành"] = scores[score_idx + 2]
                                row_data["Điểm cuối kỳ"] = scores[score_idx + 3]
                                row_data["Điểm TB môn học"] = scores[score_idx + 4]
                                has_thuongky = True
                                has_giua_ky = True
                                has_thuc_hanh = True
                                score_idx += 5
                            elif len(scores) >= 4:  # No thực hành
                                row_data["Điểm thường kỳ"] = scores[score_idx]
                                row_data["Điểm giữa kỳ"] = scores[score_idx + 1]
                                row_data["Điểm cuối kỳ"] = scores[score_idx + 2]
                                row_data["Điểm TB môn học"] = scores[score_idx + 3]
                                has_thuongky = True
                                has_giua_ky = True
                                score_idx += 4
                            elif len(scores) >= 2:  # Minimal
                                row_data["Điểm cuối kỳ"] = scores[score_idx]
                                row_data["Điểm TB môn học"] = scores[score_idx + 1]
                                score_idx += 2
                            
                            if diem_chu in ['A', 'B', 'C', 'D']:
                                row_data["Điểm chữ"] = diem_chu
                            else:
                                st.warning(f"Điểm chữ không hợp lệ trong bảng trên trang {page_num + 1}: {row}")
                                continue
                            
                            if ghi_chu:
                                row_data["Ghi chú"] = ghi_chu
                                has_ghi_chu = True
                            
                            rows.append(row_data)
                        except Exception as e:
                            st.warning(f"Lỗi xử lý dòng bảng trên trang {page_num + 1}: {row}. Lỗi: {str(e)}")
                            continue
            else:
                st.warning(f"Không tìm thấy bảng trên trang {page_num + 1}.")
    
    df = pd.DataFrame(rows)
    # Drop optional columns if they were not detected
    if not has_thuc_hanh and "Điểm thực hành" in df.columns:
        df = df.drop(columns=["Điểm thực hành"])
    if not has_giua_ky and "Điểm giữa kỳ" in df.columns:
        df = df.drop(columns=["Điểm giữa kỳ"])
    if not has_thuongky and "Điểm thường kỳ" in df.columns:
        df = df.drop(columns=["Điểm thường kỳ"])
    if not has_ghi_chu and "Ghi chú" in df.columns:
        df = df.drop(columns=["Ghi chú"])
    return df

# File upload interface
uploaded_file = st.file_uploader("📌 Tải file PDF bảng điểm:", type="pdf", accept_multiple_files=False, help="File PDF nên dưới 200MB.")
if uploaded_file is not None:
    try:
        df = extract_scores_from_pdf(uploaded_file)
        if not df.empty:
            st.success("✅ Đã trích xuất thành công!")
            st.dataframe(df, use_container_width=True)
            
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
            st.warning("⚠️ Không trích xuất được dữ liệu từ file PDF. Vui lòng kiểm tra định dạng PDF hoặc thử OCR nếu là bản scan.")
    except Exception as e:
        st.error(f"Lỗi xử lý file PDF: {str(e)}")
