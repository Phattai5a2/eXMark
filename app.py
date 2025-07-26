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
    """Extract grade data from PDF, handling varying column sets."""
    rows = []
    has_thuongky = False  # Flag for Điểm thường kỳ
    has_giua_ky = False   # Flag for Điểm giữa kỳ
    has_thuc_hanh = False # Flag for Điểm thực hành
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                st.warning(f"Không tìm thấy văn bản trên trang {page_num + 1}.")
                continue
            
            lines = text.splitlines()
            for line in lines:
                # Pattern 1: Full columns with numeric Điểm chữ, letter grade, and Ghi chú
                pattern_full = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+V\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+([ABCD])\s+(\S+)(?:\s+(.+))?$"
                
                # Pattern 2: No Điểm thực hành
                pattern_no_th = r"(\d+)\s+(\d+)\s+(.+?)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+V\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+([ABCD])\s+(\S+)(?:\s+(.+))?$"
                
                # Pattern 3: Only Điểm cuối kỳ, Điểm TB, numeric Điểm chữ, letter grade, and Ghi chú
                pattern_minimal = r"(\d+)\s+(\d+)\s+(.+?)\s+V\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+([ABCD])\s+(\S+)(?:\s+(.+))?$"
                
                # Try matching patterns in order of complexity
                match = re.match(pattern_full, line)
                if match:
                    has_thuongky = True
                    has_giua_ky = True
                    has_thuc_hanh = True
                    try:
                        stt = int(match.group(1))
                        mssv = match.group(2)
                        fullname = match.group(3).strip()
                        diem_gk = float(match.group(4))
                        diem_thuongky = float(match.group(5))
                        diem_th = float(match.group(6))  # Điểm thực hành
                        diem_cuoi_ky = float(match.group(7))  # Điểm cuối kỳ
                        diem_tb = float(match.group(8))  # Điểm TB môn học
                        diem_chu_numeric = float(match.group(9))  # Numeric Điểm chữ
                        diem_chu = match.group(10)  # Letter grade (A, B, C, D)
                        ghi_chu = match.group(11) if match.group(11) else ""  # Ghi chú
                        
                        if diem_chu not in ['A', 'B', 'C', 'D']:
                            st.warning(f"Điểm chữ (chữ cái) không hợp lệ trên dòng: {line}")
                            continue
                        
                        ho_dem, ten = split_name(fullname)
                        
                        rows.append({
                            "STT": stt,
                            "Mã số sinh viên": mssv,
                            "Họ đệm": ho_dem,
                            "Tên": ten,
                            "Điểm thường kỳ": diem_thuongky,
                            "Điểm giữa kỳ": diem_gk,
                            "Điểm thực hành": diem_th,
                            "Điểm cuối kỳ": diem_cuoi_ky,
                            "Điểm TB môn học": diem_tb,
                            "Điểm chữ (số)": diem_chu_numeric,
                            "Điểm chữ": diem_chu,
                            "Ghi chú": ghi_chu
                        })
                    except Exception as e:
                        st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                        continue
                else:
                    match = re.match(pattern_no_th, line)
                    if match:
                        has_thuongky = True
                        has_giua_ky = True
                        try:
                            stt = int(match.group(1))
                            mssv = match.group(2)
                            fullname = match.group(3).strip()
                            diem_thuongky = float(match.group(4))
                            diem_gk = float(match.group(5))
                            diem_cuoi_ky = float(match.group(6))  # Điểm cuối kỳ
                            diem_tb = float(match.group(7))  # Điểm TB môn học
                            diem_chu_numeric = float(match.group(8))  # Numeric Điểm chữ
                            diem_chu = match.group(9)  # Letter grade
                            ghi_chu = match.group(10) if match.group(10) else ""  # Ghi chú
                            
                            if diem_chu not in ['A', 'B', 'C', 'D']:
                                st.warning(f"Điểm chữ (chữ cái) không hợp lệ trên dòng: {line}")
                                continue
                            
                            ho_dem, ten = split_name(fullname)
                            
                            rows.append({
                                "STT": stt,
                                "Mã số sinh viên": mssv,
                                "Họ đệm": ho_dem,
                                "Tên": ten,
                                "Điểm thường kỳ": diem_thuongky,
                                "Điểm giữa kỳ": diem_gk,
                                "Điểm cuối kỳ": diem_cuoi_ky,
                                "Điểm TB môn học": diem_tb,
                                "Điểm chữ (số)": diem_chu_numeric,
                                "Điểm chữ": diem_chu,
                                "Ghi chú": ghi_chu
                            })
                        except Exception as e:
                            st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                            continue
                    else:
                        match = re.match(pattern_minimal, line)
                        if match:
                            try:
                                stt = int(match.group(1))
                                mssv = match.group(2)
                                fullname = match.group(3).strip()
                                diem_cuoi_ky = float(match.group(4))  # Điểm cuối kỳ
                                diem_tb = float(match.group(5))  # Điểm TB môn học
                                diem_chu_numeric = float(match.group(6))  # Numeric Điểm chữ
                                diem_chu = match.group(7)  # Letter grade
                                ghi_chu = match.group(8) if match.group(8) else ""  # Ghi chú
                                
                                if diem_chu not in ['A', 'B', 'C', 'D']:
                                    st.warning(f"Điểm chữ (chữ cái) không hợp lệ trên dòng: {line}")
                                    continue
                                
                                ho_dem, ten = split_name(fullname)
                                
                                rows.append({
                                    "STT": stt,
                                    "Mã số sinh viên": mssv,
                                    "Họ đệm": ho_dem,
                                    "Tên": ten,
                                    "Điểm cuối kỳ": diem_cuoi_ky,
                                    "Điểm TB môn học": diem_tb,
                                    "Điểm chữ (số)": diem_chu_numeric,
                                    "Điểm chữ": diem_chu,
                                    "Ghi chú": ghi_chu
                                })
                            except Exception as e:
                                st.warning(f"Lỗi xử lý dòng trên trang {page_num + 1}: {line}. Lỗi: {str(e)}")
                                continue
                
                # Debug unmatched lines
                if not match:
                    st.warning(f"Dòng không khớp trên trang {page_num + 1}: {line}")
    
    df = pd.DataFrame(rows)
    # Drop optional columns if they were not detected
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
