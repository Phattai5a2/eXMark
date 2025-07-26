#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jul 26 14:53:38 2025

@author: phattai
"""

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
    """Extract only Điểm giữa kỳ, Điểm thường kỳ, Điểm thực hành, Điểm cuối kỳ from PDF."""
    rows = []
    unmatched_lines = []
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                st.warning(f"Không tìm thấy văn bản trên trang {page_num + 1}. Có thể cần OCR.")
                continue
            
            lines = text.splitlines()
            for line in lines:
                # Pattern: Capture STT, MSSV, Fullname, Điểm giữa kỳ, Điểm thường kỳ, Điểm thực hành, Điểm cuối kỳ
                # Make 'V' separator optional
                pattern = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d{1,2})\s+(\d+\.\d{1,2})\s+(?:V\s+)?(\d+\.\d{1,2})\s+(\d+\.\d{1,2})\s+.*$"
                
                match = re.match(pattern, line)
                if match:
                    try:
                        stt = int(match.group(1))
                        mssv = match.group(2)
                        fullname = match.group(3).strip()
                        diem_gk = float(match.group(4))  # Điểm giữa kỳ
                        diem_thuongky = float(match.group(5))  # Điểm thường kỳ
                        diem_th = float(match.group(6))  # Điểm thực hành
                        diem_cuoi_ky = float(match.group(7))  # Điểm cuối kỳ
                        
                        ho_dem, ten = split_name(fullname)
                        
                        rows.append({
                            "STT": stt,
                            "Mã số sinh viên": mssv,
                            "Họ đệm": ho_dem,
                            "Tên": ten,
                            "Điểm giữa kỳ": diem_gk,
                            "Điểm thường kỳ": diem_thuongky,
                            "Điểm thực hành": diem_th,
                            "Điểm cuối kỳ": diem_cuoi_ky
                        })
                    except Exception as e:
                        st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                        unmatched_lines.append(f"Page {page_num + 1}: {line}")
                else:
                    unmatched_lines.append(f"Page {page_num + 1}: {line}")
    
    # Log unmatched lines to Streamlit
    if unmatched_lines:
        st.warning("Các dòng không khớp:")
        for ul in unmatched_lines:
            st.text(ul)
        # Optionally save unmatched lines to a file for debugging
        with open("unmatched_lines.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(unmatched_lines))
    
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
            st.error("⚠️ Không trích xuất được dữ liệu từ file PDF. Kiểm tra các dòng không khớp hoặc thử OCR.")
    except Exception as e:
        st.error(f"Lỗi xử lý file PDF: {str(e)}")
