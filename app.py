import io
import re
import warnings
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

warnings.filterwarnings("ignore", message="Workbook contains no default style")

ALLOWED_USERS = {
    "PHCN", "NGOẠI TH", "XNTT", "NHH", "TTTM", "NGOẠI CT", "DA LIỄU", "PTGMHS",
    "NỘI TH", "PHỤ SẢN", "MẮT", "TTBVSK", "DƯỢC", "THẬN", "TRUYỀN NHIỄM", "TCKT",
    "NGOẠI UB", "NGOẠI XẠ TRỊ", "RHM", "KB", "CXK", "CSGN", "NGOẠI TN", "HHLS",
    "TMH", "HSTC", "HTL", "NGOẠI TKLN", "NHI", "ĐÔNG Y", "CẤP CỨU", "LKTK"
}

def normalize_text(value):
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip().upper()

def copy_style(source, target):
    if source.has_style:
        target._style = copy(source._style)
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)
    target.number_format = source.number_format

def center_cell(cell, wrap=False):
    alignment = copy(cell.alignment)
    alignment.horizontal = "center"
    alignment.vertical = "center"
    alignment.wrap_text = wrap
    cell.alignment = alignment

def find_header_row(ws):
    for row in range(1, min(ws.max_row, 20) + 1):
        values = [normalize_text(ws.cell(row, col).value) for col in range(1, ws.max_column + 1)]
        if "STT" in values and "NGƯỜI DÙNG" in values:
            return row
    raise Exception("Không tìm thấy dòng tiêu đề STT / NGƯỜI DÙNG.")

def find_columns(ws, header_row):
    user_col = None
    day_columns = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col).value
        text = normalize_text(value)
        if text == "NGƯỜI DÙNG":
            user_col = col
        if value is not None:
            value_text = str(value).strip()
            if value_text.isdigit():
                day = int(value_text)
                if 1 <= day <= 31:
                    day_columns[day] = col
    if user_col is None:
        raise Exception("Không tìm thấy cột NGƯỜI DÙNG.")
    if not day_columns:
        raise Exception("Không tìm thấy cột ngày.")
    return user_col, day_columns

def process_giam_dinh_web(uploaded_file, start_day, end_day):
    source_wb = load_workbook(uploaded_file)
    if "Người dùng" not in source_wb.sheetnames:
        raise Exception("File không có sheet 'Người dùng'.")
    source = source_wb["Người dùng"]
    header_row = find_header_row(source)
    user_col, day_columns = find_columns(source, header_row)

    selected_days = [d for d in range(start_day, end_day + 1) if d in day_columns]
    if not selected_days:
        raise Exception(f"Không tìm thấy cột ngày nào từ {start_day} đến {end_day}.")

    source_columns = [1, 2] + [day_columns[d] for d in selected_days]
    result_last_col = len(source_columns)
    result_last_letter = get_column_letter(result_last_col)

    result_wb = Workbook()
    result = result_wb.active
    result.title = "Người dùng"

    for row in range(1, 6):
        for new_col, old_col in enumerate(source_columns, start=1):
            source_cell = source.cell(row, old_col)
            target_cell = result.cell(row, new_col)
            target_cell.value = source_cell.value
            copy_style(source_cell, target_cell)

    title = source["A1"].value
    for col in range(2, result_last_col + 1):
        result.cell(1, col).value = None
    result["A1"] = title
    result.merge_cells(start_row=1, start_column=1, end_row=1, end_column=result_last_col)
    center_cell(result["A1"], wrap=False)
    result.row_dimensions[1].height = 32

    note = source["A2"].value
    for col in range(2, result_last_col + 1):
        result.cell(2, col).value = None
    result["A2"] = note
    result.merge_cells(start_row=2, start_column=1, end_row=2, end_column=result_last_col)

    result["A5"] = None
    result.merge_cells("A4:A5")
    center_cell(result["A4"], wrap=False)

    result["B5"] = None
    result.merge_cells("B4:B5")
    center_cell(result["B4"], wrap=False)

    output_row = 6
    stt = 1
    for source_row in range(6, source.max_row + 1):
        user_name = normalize_text(source.cell(source_row, user_col).value)
        if user_name not in ALLOWED_USERS:
            continue
        for new_col, old_col in enumerate(source_columns, start=1):
            source_cell = source.cell(source_row, old_col)
            target_cell = result.cell(output_row, new_col)
            target_cell.value = stt if new_col == 1 else source_cell.value
            copy_style(source_cell, target_cell)
        stt += 1
        output_row += 1

    for new_col, old_col in enumerate(source_columns, start=1):
        old_letter = get_column_letter(old_col)
        new_letter = get_column_letter(new_col)
        width = source.column_dimensions[old_letter].width
        if width:
            result.column_dimensions[new_letter].width = width

    result.column_dimensions["A"].width = 8
    result.column_dimensions["B"].width = 25
    result.freeze_panes = "C6"
    result.auto_filter.ref = f"A4:{result_last_letter}{result.max_row}"

    out_io = io.BytesIO()
    result_wb.save(out_io)
    out_io.seek(0)
    return out_io

def process_cntt_web(uploaded_file, sheet_name):
    wb_source = load_workbook(uploaded_file, data_only=True)
    ws_source = wb_source[sheet_name]

    software_columns = list(range(5, 26))
    software_headers = []
    for col in software_columns:
        header = ws_source.cell(105, col).value
        software_headers.append(str(header).strip() if header is not None else "")

    departments = []
    for row in range(106, ws_source.max_row + 1):
        dept = ws_source.cell(row, 4).value
        if dept is None:
            continue
        dept = str(dept).strip()
        if dept.upper() == "TỔNG CỘNG":
            break
        if dept:
            departments.append(dept)

    data = {dept: [None] * len(software_columns) for dept in departments}
    for row in range(106, ws_source.max_row + 1):
        dept = ws_source.cell(row, 4).value
        if dept is None:
            continue
        dept = str(dept).strip()
        if dept.upper() == "TỔNG CỘNG":
            break
        if dept not in data:
            continue
        for i, col in enumerate(software_columns):
            value = ws_source.cell(row, col).value
            data[dept][i] = None if value in (None, "", 0) else value

    other_errors = {dept: [] for dept in departments}
    for row in range(7, 106):
        dept = ws_source.cell(row, 2).value
        error = ws_source.cell(row, 35).value
        if dept is None or error is None:
            continue
        dept = str(dept).strip()
        error = str(error).strip()
        if dept in other_errors and error and error != "0":
            other_errors[dept].append(error)

    for dept in other_errors:
        unique_errors = []
        seen = set()
        for err in other_errors[dept]:
            if err not in seen:
                seen.add(err)
                unique_errors.append(err)
        other_errors[dept] = unique_errors

    valid_departments = []
    for dept in departments:
        has_sw = any(v not in (None, "", 0) for v in data[dept])
        has_err = any(e not in (None, "", "0") for e in other_errors[dept])
        if has_sw or has_err:
            valid_departments.append(dept)

    active_sw_idx = [i for i in range(len(software_columns)) if any(data[d][i] not in (None, "", 0) for d in valid_departments)]
    has_other_col = any(other_errors[d] for d in valid_departments)

    headers = ["STT", "KHOA, PHÒNG, TRUNG TÂM"] + [software_headers[i] for i in active_sw_idx]
    if has_other_col:
        headers.append("SỬA LỖI KHÁC")

    total_columns = len(headers)
    result = Workbook()
    ws = result.active
    ws.title = "BÁO CÁO TỔNG HỢP"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_columns)
    title_cell = ws.cell(1, 1)
    title_cell.value = "TỔNG HỢP CÔNG VIỆC P.CNTT"
    title_cell.font = Font(name="Arial", size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_row = 3
    data_start_row = 4

    for col, header in enumerate(headers, 1):
        cell = ws.cell(header_row, col)
        cell.value = header
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 75

    for index, dept in enumerate(valid_departments, 1):
        row = data_start_row + index - 1
        output_col = 1
        ws.cell(row, output_col, index)
        output_col += 1
        ws.cell(row, output_col, dept)
        output_col += 1
        for i in active_sw_idx:
            val = data[dept][i]
            if val not in (None, "", 0):
                ws.cell(row, output_col, val)
            output_col += 1
        if has_other_col:
            if other_errors[dept]:
                ws.cell(row, output_col, ", ".join(other_errors[dept]))
            output_col += 1
        for col in range(1, total_columns + 1):
            cell = ws.cell(row, col)
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(horizontal="left" if col in (2, total_columns) else "center", vertical="top", wrap_text=True)

    total_row = data_start_row + len(valid_departments)
    ws.cell(total_row, 2, "TỔNG CỘNG")

    for col in range(3, total_columns + 1):
        if ws.cell(header_row, col).value == "SỬA LỖI KHÁC":
            continue
        vals = [ws.cell(r, col).value for r in range(data_start_row, total_row) if isinstance(ws.cell(r, col).value, (int, float))]
        if vals and sum(vals) != 0:
            ws.cell(total_row, col, sum(vals))

    if has_other_col:
        other_total = 0
        for col in range(3, total_columns):
            if str(ws.cell(header_row, col).value or "").strip() == "Đăng bài trên website":
                continue
            val = ws.cell(total_row, col).value
            if isinstance(val, (int, float)):
                other_total += val
        for col in range(3, total_columns):
            if str(ws.cell(header_row, col).value or "").strip() == "Đăng bài trên website":
                ws.cell(total_row, col).value = None
                break
        if other_total != 0:
            ws.cell(total_row, total_columns, other_total)

    for col in range(1, total_columns + 1):
        cell = ws.cell(total_row, col)
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.border = border
        cell.alignment = Alignment(horizontal="left" if col in (2, total_columns) else "center", vertical="top", wrap_text=True)

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 28
    for col in range(3, total_columns + 1):
        h = ws.cell(header_row, col).value
        ws.column_dimensions[get_column_letter(col)].width = 55 if h == "SỬA LỖI KHÁC" else 17
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:{get_column_letter(total_columns)}{total_row}"

    out_io = io.BytesIO()
    result.save(out_io)
    out_io.seek(0)
    return out_io

# ============================================================
# GIAO DIỆN STREAMLIT WEB
# ============================================================
st.set_page_config(page_title="Báo cáo CNTT & Giám định BHYT", page_icon="📊", layout="wide")

st.title("📊 HỆ THỐNG TỔNG HỢP BÁO CÁO ONLINE")
st.caption("Ứng dụng tổng hợp báo cáo Cổng giám định BHYT và Công việc Phòng CNTT")

tab1, tab2 = st.tabs(["📋 Báo cáo Cổng giám định BHYT", "💻 Báo cáo công việc P.CNTT"])

with tab1:
    st.subheader("1. Báo cáo Cổng giám định BHYT")
    uploaded_giamdinh = st.file_uploader("Tải lên file Excel Cổng giám định (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="gd")
    date_range = st.radio("Chọn khoảng ngày báo cáo:", ["01 - 14", "15 - 31"], horizontal=True)

    if uploaded_giamdinh is not None:
        if st.button("🚀 Xử lý báo cáo Giám định", type="primary"):
            try:
                with st.spinner("Đang xử lý dữ liệu..."):
                    start_day, end_day = (1, 14) if date_range == "01 - 14" else (15, 31)
                    output_file = process_giam_dinh_web(uploaded_giamdinh, start_day, end_day)
                    st.success("✅ Đã xử lý xong báo cáo!")
                    st.download_button(
                        label="📥 Tải file kết quả (.xlsx)",
                        data=output_file,
                        file_name="DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra: {e}")

with tab2:
    st.subheader("2. Báo cáo tổng hợp công việc P.CNTT")
    uploaded_cntt = st.file_uploader("Tải lên file Excel báo cáo công việc (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="cntt")

    if uploaded_cntt is not None:
        try:
            wb_temp = load_workbook(uploaded_cntt, read_only=True)
            sheet_options = wb_temp.sheetnames
            selected_sheet = st.selectbox("Chọn Sheet cần xử lý:", options=sheet_options)
            
            if st.button("🚀 Xử lý báo cáo P.CNTT", type="primary"):
                with st.spinner("Đang tính toán và định dạng..."):
                    output_file_cntt = process_cntt_web(uploaded_cntt, selected_sheet)
                    st.success("✅ Đã tổng hợp thành công!")
                    st.download_button(
                        label="📥 Tải file Báo cáo P.CNTT (.xlsx)",
                        data=output_file_cntt,
                        file_name="TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"❌ Không thể đọc file Excel: {e}")
