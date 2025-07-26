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
    """Extract score columns from PDF based on three conditions."""
    rows = []
    unmatched_lines = []
    has_giua_ky = False
    has_thuongky = False
    has_thuc_hanh = False
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                with open("unmatched_lines.txt", "a", encoding="utf-8") as f:
                    f.write(f"Page {page_num + 1}: No text extracted\n")
                continue
            
            lines = text.splitlines()
            for line in lines:
                # Pattern 1: All columns (Điểm giữa kỳ, Điểm thường kỳ, Điểm thực hành, Điểm cuối kỳ)
                pattern_full = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d{1,2})\s+(\d+\.\d{1,2})\s+(?:V\s+)?(\d+\.\d{1,2})\s+(\d+\.\d{1,2})\s+.*$"
                
                # Pattern 2: No Điểm thực hành (Điểm giữa kỳ, Điểm thường kỳ, Điểm cuối kỳ)
                pattern_no_th = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d{1,2})\s+(\d+\.\d{1,2})\s+(?:V\s+)?(\d+\.\d{1,2})\s+.*$"
                
                # Pattern 3: Only Điểm cuối kỳ (no Điểm giữa kỳ, Điểm thường kỳ, Điểm thực hành)
                pattern_minimal = r"(\d+)\s+(\d+)\s+(.+?)\s+(?:V\s+)?(\d+\.\d{1,2})\s+.*$"
                
                # Try matching patterns in order of complexity
                match = re.match(pattern_full, line)
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
                        has_giua_ky = True
                        has_thuongky = True
                        has_thuc_hanh = True
                    except Exception as e:
                        unmatched_lines.append(f"Page {page_num + 1}: {line} (Error: {str(e)})")
                    continue
                
                match = re.match(pattern_no_th, line)
                if match and not re.match(pattern_full, line):  # Ensure pattern_full doesn't match
                    try:
                        stt = int(match.group(1))
                        mssv = match.group(2)
                        fullname = match.group(3).strip()
                        diem_gk = float(match.group(4))  # Điểm giữa kỳ
                        diem_thuongky = float(match.group(5))  # Điểm thường kỳ
                        diem_cuoi_ky = float(match.group(6))  # Điểm cuối kỳ
                        
                        ho_dem, ten = split_name(fullname)
                        
                        rows.append({
                            "STT": stt,
                            "Mã số sinh viên": mssv,
                            "Họ đệm": ho_dem,
                            "Tên": ten,
                            "Điểm giữa kỳ": diem_gk,
                            "Điểm thường kỳ": diem_thuongky,
                            "Điểm cuối kỳ": diem_cuoi_ky
                        })
                        has_giua_ky = True
                        has_thuongky = True
                    except Exception as e:
                        unmatched_lines.append(f"Page {page_num + 1}: {line} (Error: {str(e)})")
                    continue
                
                match = re.match(pattern_minimal, line)
                if match and not (re.match(pattern_full, line) or re.match(pattern_no_th, line)):  # Ensure neither pattern_full nor pattern_no_th matches
                    try:
                        stt = int(match.group(1))
                        mssv = match.group(2)
                        fullname = match.group(3).strip()
                        diem_cuoi_ky = float(match.group(4))  # Điểm cuối kỳ
                        
                        ho_dem, ten = split_name(fullname)
                        
                        rows.append({
                            "STT": stt,
                            "Mã số sinh viên": mssv,
                            "Họ đệm": ho_dem,
                            "Tên": ten,
                            "Điểm cuối kỳ": diem_cuoi_ky
                        })
                    except Exception as e:
                        unmatched_lines.append(f"Page {page_num + 1}: {line} (Error: {str(e)})")
                    continue
                
                # Log unmatched lines to file
                unmatched_lines.append(f"Page {page_num + 1}: {line}")
    
    # Save unmatched lines to file for debugging
    if unmatched_lines:
        with open("unmatched_lines.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(unmatched_lines))
    
    # Define base columns
    columns = ["STT", "Mã số sinh viên", "Họ đệm", "Tên"]
    # Add score columns based on what was detected
    if has_giua_ky:
        columns.append("Điểm giữa kỳ")
    if has_thuongky:
        columns.append("Điểm thường kỳ")
    if has_thuc_hanh:
        columns.append("Điểm thực hành")
    columns.append("Điểm cuối kỳ")  # Always include Điểm cuối kỳ
    
    # Create DataFrame with only the relevant columns
    df = pd.DataFrame(rows, columns=columns)
    
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
            st.error("⚠️ Không trích xuất được dữ liệu từ file PDF. Vui lòng kiểm tra file hoặc định dạng.")
    except Exception as e:
        st.error(f"Lỗi xử lý file PDF: {str(e)}")
