import io
import re
import warnings
import urllib.request
import http.client
import tempfile
import os
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from flask import Flask, request, send_file, jsonify, render_template_string

warnings.filterwarnings("ignore", message="Workbook contains no default style")

# ============================================================
# CẤU HÌNH & CHUẨN HÓA DÙNG CHUNG
# ============================================================
ALLOWED_USERS = {
    "PHCN", "NGOẠI TH", "XNTT", "NHH", "TTTM", "NGOẠI CT", "DA LIỄU", "PTGMHS",
    "NỘI TH", "PHỤ SẢN", "MẮT", "TTBVSK", "DƯỢC", "THẬN", "TRUYỀN NHIỄM", "TCKT",
    "NGOẠI UB", "NGOẠI XẠ TRỊ", "RHM", "KB", "CXK", "CSGN", "NGOẠI TN", "HHLS",
    "TMH", "HSTC", "HTL", "NGOẠI TKLN", "NHI", "ĐÔNG Y", "CẤP CỨU", "LKTK"
}

DEFAULT_FIXED_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139"

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

# ============================================================
# LOGIC XỬ LÝ BÁO CÁO GIÁM ĐỊNH BHYT
# ============================================================
def process_giam_dinh_web(uploaded_file, start_day, end_day):
    source_wb = load_workbook(uploaded_file)
    if "Người dùng" not in source_wb.sheetnames:
        raise Exception("File không có sheet 'Người dùng'.")
    source = source_wb["Người dùng"]
    header_row = find_header_row(source)
    user_col, day_columns = find_columns(source, header_row)

    selected_days = [d for d in range(start_day, end_day + 1) if d in day_columns]
    if not selected_days:
        raise Exception(f"Không tìm thấy cột ngày nào từ ngày {start_day} đến ngày {end_day}.")

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

# ============================================================
# LOGIC XỬ LÝ BÁO CÁO P.CNTT & TẢI GOOGLE SHEETS
# ============================================================
def download_google_sheet(url):
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9\-_]+)', url)
    if not match:
        raise ValueError("URL không hợp lệ. Vui lòng cung cấp link Google Sheets hợp lệ.")
    spreadsheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"

    req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=90) as resp:
        chunks = []
        while True:
            try:
                chunk = resp.read(65536)
                if not chunk:
                    break
                chunks.append(chunk)
            except http.client.IncompleteRead as ir:
                if ir.partial:
                    chunks.append(ir.partial)
                break
        content = b"".join(chunks)
    return io.BytesIO(content)

def process_cntt_web(file_stream, sheet_name):
    wb_source = load_workbook(file_stream, data_only=True)
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
# FLASK WEB APPLICATION (TƯƠNG THÍCH HOÀN TOÀN VERCEL)
# ============================================================
app = Flask(__name__)
application = app  # Cho WSGI / Vercel
handler = app      # Cho Vercel Serverless Function

HTML_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Hệ Thống Tổng Hợp Báo Cáo Online - BVĐK Bắc Ninh Số 2</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --primary-light: #eff6ff;
      --bg-dark: #0f172a;
      --bg-sidebar: #0f172a;
      --bg-body: #f8fafc;
      --card-bg: #ffffff;
      --border: #e2e8f0;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --sidebar-text: #94a3b8;
      --sidebar-hover: #1e293b;
      --success: #10b981;
      --error: #ef4444;
      --radius: 12px;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Plus Jakarta Sans', sans-serif;
    }

    body {
      min-height: 100vh;
      background: var(--bg-body);
      color: var(--text-main);
      display: flex;
    }

    /* BỐ CỤC SIDEBAR BÊN TRÁI */
    .app-layout {
      display: flex;
      width: 100%;
      min-height: 100vh;
    }

    .sidebar {
      width: 280px;
      background: var(--bg-sidebar);
      color: #fff;
      display: flex;
      flex-direction: column;
      flex-shrink: 0;
      border-right: 1px solid rgba(255, 255, 255, 0.08);
      position: sticky;
      top: 0;
      height: 100vh;
      overflow-y: auto;
    }

    .sidebar-header {
      padding: 20px 18px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
    }

    .brand-badge {
      font-size: 0.75rem;
      font-weight: 700;
      color: #60a5fa;
      letter-spacing: 0.5px;
      margin-bottom: 4px;
    }

    .sidebar-header h2 {
      font-size: 1.1rem;
      font-weight: 800;
      color: #ffffff;
      line-height: 1.3;
    }

    .sidebar-header p {
      font-size: 0.78rem;
      color: var(--sidebar-text);
      margin-top: 3px;
    }

    .nav-menu {
      padding: 16px 10px;
      flex: 1;
    }

    .nav-label {
      font-size: 0.72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      color: #64748b;
      padding: 6px 12px 10px 12px;
    }

    .nav-item {
      width: 100%;
      background: transparent;
      border: none;
      outline: none;
      text-align: left;
      padding: 12px 14px;
      border-radius: 10px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 6px;
      transition: all 0.2s ease;
      color: var(--sidebar-text);
    }

    .nav-item:hover {
      background: var(--sidebar-hover);
      color: #fff;
    }

    .nav-item.active {
      background: var(--primary);
      color: #fff;
      box-shadow: 0 4px 12px rgba(37, 99, 235, 0.35);
    }

    .nav-item .icon {
      font-size: 1.3rem;
      line-height: 1;
    }

    .nav-item .text .title {
      font-size: 0.88rem;
      font-weight: 700;
      color: inherit;
    }

    .nav-item .text .desc {
      font-size: 0.72rem;
      opacity: 0.75;
      margin-top: 2px;
    }

    .sidebar-footer {
      padding: 16px;
      border-top: 1px solid rgba(255, 255, 255, 0.08);
      font-size: 0.75rem;
      color: #64748b;
      background: rgba(0, 0, 0, 0.2);
    }

    .status-indicator {
      display: flex;
      align-items: center;
      gap: 6px;
      color: #34d399;
      font-weight: 600;
      margin-bottom: 4px;
    }

    .status-indicator .dot {
      width: 8px;
      height: 8px;
      background: #10b981;
      border-radius: 50%;
      box-shadow: 0 0 8px #10b981;
    }

    /* BỐ CỤC KHUNG NỘI DUNG CHÍNH (BÊN PHẢI) */
    .main-content {
      flex: 1;
      display: flex;
      flex-direction: column;
      overflow-y: auto;
    }

    .top-bar {
      background: #ffffff;
      border-bottom: 1px solid var(--border);
      padding: 14px 28px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      position: sticky;
      top: 0;
      z-index: 10;
    }

    .top-bar .breadcrumb {
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .top-bar .badge {
      background: #dbeafe;
      color: #1d4ed8;
      font-size: 0.75rem;
      font-weight: 700;
      padding: 4px 10px;
      border-radius: 6px;
      text-transform: uppercase;
    }

    .top-bar h1 {
      font-size: 1.15rem;
      font-weight: 800;
      color: var(--text-main);
    }

    .content-body {
      padding: 28px;
      max-width: 960px;
      width: 100%;
    }

    .page-tab {
      display: none;
      animation: fadeIn 0.25s ease;
    }

    .page-tab.active {
      display: block;
    }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }

    /* CARDS */
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 24px 28px;
      margin-bottom: 24px;
      box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04);
    }

    .card-title {
      font-size: 1.05rem;
      font-weight: 700;
      margin-bottom: 6px;
      color: var(--text-main);
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .card-desc {
      font-size: 0.85rem;
      color: var(--text-muted);
      margin-bottom: 20px;
    }

    .form-group {
      margin-bottom: 20px;
    }

    .form-group label {
      display: block;
      font-size: 0.88rem;
      font-weight: 700;
      margin-bottom: 8px;
      color: #1e293b;
    }

    /* DROPZONE */
    .dropzone {
      border: 2px dashed #cbd5e1;
      border-radius: 12px;
      padding: 32px 20px;
      text-align: center;
      cursor: pointer;
      background: #f8fafc;
      transition: all 0.2s ease;
    }

    .dropzone:hover {
      border-color: var(--primary);
      background: #eff6ff;
    }

    .dropzone .icon {
      font-size: 2.2rem;
      margin-bottom: 8px;
    }

    .dropzone .title {
      font-weight: 700;
      font-size: 0.95rem;
      color: #1e293b;
    }

    .dropzone .desc {
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-top: 4px;
    }

    .file-chosen {
      display: none;
      margin-top: 10px;
      padding: 10px 14px;
      background: #ecfdf5;
      border: 1px solid #a7f3d0;
      border-radius: 8px;
      font-size: 0.85rem;
      color: #065f46;
      font-weight: 600;
    }

    /* RADIO GRID */
    .radio-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 10px;
    }

    .radio-card {
      border: 1px solid var(--border);
      background: #ffffff;
      padding: 12px 14px;
      border-radius: 8px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 0.88rem;
      font-weight: 600;
      color: #334155;
      transition: all 0.2s ease;
    }

    .radio-card:hover {
      border-color: var(--primary);
      background: #f8fafc;
    }

    .radio-card input {
      accent-color: var(--primary);
      width: 16px;
      height: 16px;
    }

    /* INPUTS */
    .text-input, .select-input {
      width: 100%;
      background: #ffffff;
      border: 1px solid var(--border);
      color: var(--text-main);
      padding: 12px 14px;
      border-radius: 8px;
      font-size: 0.9rem;
      outline: none;
      transition: border-color 0.2s ease;
    }

    .text-input:focus, .select-input:focus {
      border-color: var(--primary);
      box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
    }

    .btn-quick {
      margin-top: 8px;
      background: #eff6ff;
      color: var(--primary);
      border: 1px dashed #93c5fd;
      padding: 6px 14px;
      border-radius: 6px;
      font-size: 0.8rem;
      cursor: pointer;
      font-weight: 700;
      transition: all 0.2s;
    }

    .btn-quick:hover {
      background: #dbeafe;
    }

    /* SUBMIT BUTTON */
    .btn-submit {
      width: 100%;
      background: var(--primary);
      color: #fff;
      border: none;
      padding: 14px;
      border-radius: 10px;
      font-size: 0.98rem;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
    }

    .btn-submit:hover:not(:disabled) {
      background: var(--primary-hover);
      transform: translateY(-1px);
    }

    .btn-submit:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    /* ALERTS */
    .alert {
      padding: 14px 16px;
      border-radius: 8px;
      font-size: 0.88rem;
      margin-top: 16px;
      display: none;
      align-items: center;
      gap: 10px;
    }

    .alert.success {
      background: #ecfdf5;
      border: 1px solid #a7f3d0;
      color: #065f46;
    }

    .alert.error {
      background: #fef2f2;
      border: 1px solid #fecaca;
      color: #991b1b;
    }

    .spinner {
      border: 3px solid rgba(255, 255, 255, 0.3);
      border-top: 3px solid #fff;
      border-radius: 50%;
      width: 18px;
      height: 18px;
      animation: spin 0.8s linear infinite;
      display: none;
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    /* GUIDE BOX */
    .guide-box {
      background: #f8fafc;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 16px 20px;
      font-size: 0.85rem;
      color: #334155;
      line-height: 1.6;
    }

    .guide-box ol {
      padding-left: 20px;
      margin-top: 6px;
    }

    .guide-box li {
      margin-bottom: 4px;
    }

    @media (max-width: 768px) {
      .app-layout { flex-direction: column; }
      .sidebar { width: 100%; height: auto; position: static; }
      .content-body { padding: 16px; }
    }
  </style>
</head>
<body>
  <div class="app-layout">
    <!-- SIDEBAR BÊN TRÁI -->
    <aside class="sidebar">
      <div class="sidebar-header">
        <div class="brand-badge">🏥 BVĐK BẮC NINH SỐ 2</div>
        <h2>HỆ THỐNG BÁO CÁO</h2>
        <p>Phiên bản trực tuyến Online</p>
      </div>

      <nav class="nav-menu">
        <div class="nav-label">DANH MỤC TÁC VỤ</div>

        <button class="nav-item active" onclick="switchWebTab('giamdinh', this)">
          <span class="icon">🏥</span>
          <div class="text">
            <div class="title">1. Giám Định BHYT</div>
            <div class="desc">Lọc người dùng & xuất Excel A4</div>
          </div>
        </button>

        <button class="nav-item" onclick="switchWebTab('cntt-file', this)">
          <span class="icon">💻</span>
          <div class="text">
            <div class="title">2. P.CNTT (File Excel)</div>
            <div class="desc">Tổng hợp từ file báo cáo máy tính</div>
          </div>
        </button>

        <button class="nav-item" onclick="switchWebTab('cntt-gsheet', this)">
          <span class="icon">⚡</span>
          <div class="text">
            <div class="title">3. P.CNTT (Google Sheets)</div>
            <div class="desc">Tự động tải Google Sheets cố định</div>
          </div>
        </button>

        <button class="nav-item" onclick="switchWebTab('guide', this)">
          <span class="icon">📖</span>
          <div class="text">
            <div class="title">4. Hướng Dẫn Sử Dụng</div>
            <div class="desc">Thông tin quy trình xử lý</div>
          </div>
        </button>
      </nav>

      <div class="sidebar-footer">
        <div class="status-indicator">
          <span class="dot"></span>
          <span>Máy chủ Vercel: Hoạt động</span>
        </div>
        <p>Phòng CNTT • BVĐK Bắc Ninh số 2</p>
      </div>
    </aside>

    <!-- KHUNG NỘI DUNG CHÍNH (BÊN PHẢI) -->
    <main class="main-content">
      <div class="top-bar">
        <div class="breadcrumb">
          <span class="badge">ĐANG MỞ</span>
          <h1 id="topbar-title">🏥 1. Báo Cáo Cổng Giám Định BHYT</h1>
        </div>
      </div>

      <div class="content-body">
        <!-- TAB 1: GIÁM ĐỊNH BHYT -->
        <section id="tab-giamdinh" class="page-tab active">
          <div class="card">
            <div class="card-title">📋 Lọc Danh Sách Theo Dõi Giám Định BHYT</div>
            <div class="card-desc">Tự động lọc danh sách 32 khoa phòng được phép, chọn ngày và xuất file Excel chuẩn khổ A4 dọc.</div>

            <form id="form-giamdinh" onsubmit="handleGiamDinh(event)">
              <div class="form-group">
                <label>1️⃣ Tải lên file Excel Cổng Giám Định (.xlsx, .xlsm):</label>
                <div class="dropzone" onclick="document.getElementById('file-gd').click()">
                  <div class="icon">📊</div>
                  <div class="title">Bấm để chọn file hoặc kéo thả vào đây</div>
                  <div class="desc">File chứa sheet 'Người dùng' xuất từ cổng giám định</div>
                  <input type="file" id="file-gd" accept=".xlsx,.xlsm" style="display: none" onchange="onFileSelected(this, 'chosen-gd')">
                </div>
                <div id="chosen-gd" class="file-chosen"></div>
              </div>

              <div class="form-group">
                <label>2️⃣ Chọn khoảng ngày báo cáo:</label>
                <div class="radio-grid">
                  <label class="radio-card">
                    <input type="radio" name="date_range" value="01-14" checked>
                    <span>Ngày 01 đến 14</span>
                  </label>
                  <label class="radio-card">
                    <input type="radio" name="date_range" value="15-31">
                    <span>Ngày 15 đến 31</span>
                  </label>
                  <label class="radio-card">
                    <input type="radio" name="date_range" value="full">
                    <span>Full tháng (01 - 31)</span>
                  </label>
                </div>
              </div>

              <button type="submit" id="btn-gd" class="btn-submit">
                <span class="spinner" id="spin-gd"></span>
                <span id="txt-gd">🚀 Xử Lý & Tải Xuống File Báo Cáo Giám Định</span>
              </button>

              <div id="alert-gd" class="alert"></div>
            </form>
          </div>
        </section>

        <!-- TAB 2: P.CNTT TỪ FILE EXCEL -->
        <section id="tab-cntt-file" class="page-tab">
          <div class="card">
            <div class="card-title">💻 Báo Cáo Công Việc Phòng CNTT (Từ File Cục Bộ)</div>
            <div class="card-desc">Đọc dữ liệu từ file Excel báo cáo công việc trên máy, tự động tổng hợp lỗi phần mềm và sửa lỗi khác theo dòng 7.</div>

            <form id="form-cntt-file" onsubmit="handleCNTTFile(event)">
              <div class="form-group">
                <label>1️⃣ Chọn file Excel báo cáo P.CNTT (.xlsx, .xlsm):</label>
                <div class="dropzone" onclick="document.getElementById('file-cntt').click()">
                  <div class="icon">📁</div>
                  <div class="title">Bấm để chọn file báo cáo sửa chữa P.CNTT</div>
                  <input type="file" id="file-cntt" accept=".xlsx,.xlsm" style="display: none" onchange="onCNTTFileSelected(this)">
                </div>
                <div id="chosen-cntt" class="file-chosen"></div>
              </div>

              <div class="form-group">
                <label>2️⃣ Chọn Sheet ngày cần tổng hợp:</label>
                <select id="cntt-file-sheet" class="select-input">
                  <option value="">-- Vui lòng chọn file Excel trước để nạp danh sách Sheet --</option>
                </select>
              </div>

              <button type="submit" id="btn-cntt-file" class="btn-submit">
                <span class="spinner" id="spin-cntt-file"></span>
                <span id="txt-cntt-file">🚀 Tổng Hợp Báo Cáo P.CNTT (A4 Ngang)</span>
              </button>

              <div id="alert-cntt-file" class="alert"></div>
            </form>
          </div>
        </section>

        <!-- TAB 3: P.CNTT TỪ GOOGLE SHEETS -->
        <section id="tab-cntt-gsheet" class="page-tab">
          <div class="card">
            <div class="card-title">⚡ Báo Cáo P.CNTT (Google Sheets Tự Động)</div>
            <div class="card-desc">Lấy dữ liệu trực tiếp từ Google Sheets của Phòng CNTT, quét tất cả sheet ngày và xuất báo cáo chuẩn A4 ngang.</div>

            <form id="form-cntt-gsheet" onsubmit="handleCNTTGSheet(event)">
              <div class="form-group">
                <label>1️⃣ Đường Link Google Sheets (Đã điền sẵn link cố định):</label>
                <input type="text" id="gsheet-url" class="text-input" value="https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139">
                <div style="margin-top: 8px; display: flex; gap: 8px;">
                  <button type="button" class="btn-quick" onclick="useFixedLink()">
                    ⚡ Dùng Lại Link Cố Định
                  </button>
                  <button type="button" class="btn-quick" style="background: #ecfdf5; color: #059669; border-color: #a7f3d0;" onclick="fetchGSheets()">
                    🔄 Nạp Danh Sách Sheet Ngày
                  </button>
                </div>
              </div>

              <div class="form-group">
                <label>2️⃣ Chọn Sheet ngày cần xử lý:</label>
                <select id="gsheet-sheet-select" class="select-input">
                  <option value="">-- Bấm nút 'Nạp Danh Sách Sheet Ngày' ở trên --</option>
                </select>
              </div>

              <button type="submit" id="btn-cntt-gsheet" class="btn-submit">
                <span class="spinner" id="spin-cntt-gsheet"></span>
                <span id="txt-cntt-gsheet">🚀 Tổng Hợp Từ Google Sheets</span>
              </button>

              <div id="alert-cntt-gsheet" class="alert"></div>
            </form>
          </div>
        </section>

        <!-- TAB 4: HƯỚNG DẪN -->
        <section id="tab-guide" class="page-tab">
          <div class="card">
            <div class="card-title">📖 Hướng Dẫn Sử Dụng Hệ Thống</div>
            <div class="guide-box">
              <b>1. Báo cáo Giám Định BHYT:</b>
              <ol>
                <li>Chuẩn bị file Excel xuất từ cổng Giám định BHYT (.xlsx, .xlsm).</li>
                <li>Chọn khoảng ngày: 01-14 (nửa đầu tháng), 15-31 (nửa cuối tháng), hoặc Full tháng.</li>
                <li>Bấm <b>Xử Lý & Tải Xuống</b>. File kết quả sẽ tự động lưu về máy theo chuẩn A4 dọc.</li>
              </ol>
              <br>
              <b>2. Báo cáo Công việc Phòng CNTT:</b>
              <ol>
                <li>Hệ thống hỗ trợ cả file Excel tải lên hoặc Link Google Sheets trực tiếp.</li>
                <li>Hệ thống tự động đọc danh sách các Sheet ngày từ file/link.</li>
                <li>Dữ liệu được lọc từ dòng 7 trở đi, tổng hợp số lượt sửa phần mềm và lỗi sửa chữa khác.</li>
                <li>Kết quả xuất ra file <b>TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx</b> chuẩn khổ A4 nằm ngang fit cột.</li>
              </ol>
            </div>
          </div>
        </section>
      </div>
    </main>
  </div>

  <script>
    const FIXED_URL = "https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139";

    const titles = {
      'giamdinh': '🏥 1. Báo Cáo Cổng Giám Định BHYT',
      'cntt-file': '💻 2. Báo Cáo P.CNTT (Từ File Cục Bộ)',
      'cntt-gsheet': '⚡ 3. Báo Cáo P.CNTT (Google Sheets Tự Động)',
      'guide': '📖 4. Hướng Dẫn & Thông Tin Quy Trình'
    };

    function switchWebTab(tabKey, btn) {
      document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      document.querySelectorAll('.page-tab').forEach(p => p.classList.remove('active'));
      const activePage = document.getElementById('tab-' + tabKey);
      if (activePage) activePage.classList.add('active');

      document.getElementById('topbar-title').textContent = titles[tabKey] || 'Hệ Thống Báo Cáo';

      if (tabKey === 'cntt-gsheet') {
        const sel = document.getElementById('gsheet-sheet-select');
        if (sel.options.length <= 1) {
          fetchGSheets();
        }
      }
    }

    function useFixedLink() {
      document.getElementById('gsheet-url').value = FIXED_URL;
      fetchGSheets();
    }

    function onFileSelected(input, targetId) {
      const target = document.getElementById(targetId);
      if (input.files && input.files[0]) {
        target.innerHTML = `<span>📄 <b>${input.files[0].name}</b> (${Math.round(input.files[0].size/1024)} KB)</span>`;
        target.style.display = 'block';
      }
    }

    async function onCNTTFileSelected(input) {
      onFileSelected(input, 'chosen-cntt');
      if (input.files && input.files[0]) {
        const formData = new FormData();
        formData.append('file', input.files[0]);
        await loadSheetsIntoSelect(formData, 'cntt-file-sheet', 'alert-cntt-file');
      }
    }

    async function fetchGSheets() {
      const url = document.getElementById('gsheet-url').value.trim();
      if (!url) {
        alert('Vui lòng nhập link Google Sheets!');
        return;
      }
      const formData = new FormData();
      formData.append('url', url);
      await loadSheetsIntoSelect(formData, 'gsheet-sheet-select', 'alert-cntt-gsheet');
    }

    async function loadSheetsIntoSelect(formData, selectId, alertId) {
      const sel = document.getElementById(selectId);
      sel.innerHTML = '<option value="">⏳ Đang nạp danh sách các sheet ngày...</option>';
      try {
        const res = await fetch('/api/get-sheets', {
          method: 'POST',
          body: formData
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);

        sel.innerHTML = '';
        data.sheets.forEach((s, idx) => {
          const opt = document.createElement('option');
          opt.value = s;
          opt.textContent = `📅 Sheet: ${s}`;
          if (idx === data.sheets.length - 1) {
            opt.selected = true; // Chọn sheet mới nhất theo mặc định
          }
          sel.appendChild(opt);
        });
      } catch (err) {
        sel.innerHTML = '<option value="">❌ Không đọc được sheet. Vui lòng kiểm tra lại!</option>';
        showAlert(alertId, 'error', 'Lỗi: ' + err.message);
      }
    }

    function showAlert(id, type, msg) {
      const el = document.getElementById(id);
      el.className = 'alert ' + type;
      el.textContent = msg;
      el.style.display = 'flex';
    }

    function hideAlert(id) {
      const el = document.getElementById(id);
      if (el) el.style.display = 'none';
    }

    // Xử lý Giám Định
    async function handleGiamDinh(e) {
      e.preventDefault();
      hideAlert('alert-gd');

      const fileInput = document.getElementById('file-gd');
      if (!fileInput.files || !fileInput.files[0]) {
        showAlert('alert-gd', 'error', '⚠️ Vui lòng chọn file Excel Cổng Giám Định!');
        return;
      }

      const rangeVal = document.querySelector('input[name="date_range"]:checked').value;
      const formData = new FormData();
      formData.append('file', fileInput.files[0]);
      formData.append('date_range', rangeVal);

      setLoading('gd', true);

      try {
        const res = await fetch('/api/process-giam-dinh', {
          method: 'POST',
          body: formData
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }

        const blob = await res.blob();
        downloadBlob(blob, 'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx');
        showAlert('alert-gd', 'success', '✅ Đã xuất và tải xuống file Giám Định thành công!');
      } catch (err) {
        showAlert('alert-gd', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('gd', false);
      }
    }

    // Xử lý P.CNTT từ File
    async function handleCNTTFile(e) {
      e.preventDefault();
      hideAlert('alert-cntt-file');

      const fileInput = document.getElementById('file-cntt');
      const sheetName = document.getElementById('cntt-file-sheet').value;

      if (!fileInput.files || !fileInput.files[0]) {
        showAlert('alert-cntt-file', 'error', '⚠️ Vui lòng chọn file Excel báo cáo!');
        return;
      }
      if (!sheetName) {
        showAlert('alert-cntt-file', 'error', '⚠️ Vui lòng chọn Sheet cần xử lý!');
        return;
      }

      const formData = new FormData();
      formData.append('file', fileInput.files[0]);
      formData.append('sheet_name', sheetName);

      setLoading('cntt-file', true);

      try {
        const res = await fetch('/api/process-cntt', {
          method: 'POST',
          body: formData
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }

        const blob = await res.blob();
        downloadBlob(blob, 'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx');
        showAlert('alert-cntt-file', 'success', '✅ Đã tổng hợp thành công và tải file xuống!');
      } catch (err) {
        showAlert('alert-cntt-file', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('cntt-file', false);
      }
    }

    // Xử lý P.CNTT từ Google Sheets
    async function handleCNTTGSheet(e) {
      e.preventDefault();
      hideAlert('alert-cntt-gsheet');

      const url = document.getElementById('gsheet-url').value.trim();
      const sheetName = document.getElementById('gsheet-sheet-select').value;

      if (!url) {
        showAlert('alert-cntt-gsheet', 'error', '⚠️ Vui lòng nhập link Google Sheets!');
        return;
      }
      if (!sheetName) {
        showAlert('alert-cntt-gsheet', 'error', '⚠️ Vui lòng chọn Sheet ngày cần xử lý!');
        return;
      }

      const formData = new FormData();
      formData.append('url', url);
      formData.append('sheet_name', sheetName);

      setLoading('cntt-gsheet', true);

      try {
        const res = await fetch('/api/process-cntt', {
          method: 'POST',
          body: formData
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }

        const blob = await res.blob();
        downloadBlob(blob, 'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx');
        showAlert('alert-cntt-gsheet', 'success', '✅ Đã tổng hợp từ Google Sheets và tải xuống file thành công!');
      } catch (err) {
        showAlert('alert-cntt-gsheet', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('cntt-gsheet', false);
      }
    }

    function setLoading(prefix, isLoading) {
      const btn = document.getElementById('btn-' + prefix);
      const spin = document.getElementById('spin-' + prefix);
      const txt = document.getElementById('txt-' + prefix);
      if (btn) btn.disabled = isLoading;
      if (spin) spin.style.display = isLoading ? 'inline-block' : 'none';
      if (txt) {
        if (isLoading) {
          txt.textContent = 'Đang Xử Lý Dữ Liệu...';
        } else {
          txt.textContent = prefix === 'gd' ? '🚀 Xử Lý & Tải Xuống File Báo Cáo Giám Định' : '🚀 Tổng Hợp Báo Cáo P.CNTT (A4 Ngang)';
        }
      }
    }

    function downloadBlob(blob, filename) {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      a.remove();
    }
  </script>
</body>
</html>
"""

def handle_get_sheets():
    try:
        if "file" in request.files and request.files["file"].filename:
            f = request.files["file"]
            wb = load_workbook(f.stream, read_only=True)
            sheets = wb.sheetnames
        elif request.form.get("url"):
            url = request.form.get("url").strip()
            stream = download_google_sheet(url)
            wb = load_workbook(stream, read_only=True)
            sheets = wb.sheetnames
        else:
            return jsonify({"error": "Vui lòng cung cấp file hoặc URL Google Sheets."}), 400

        return jsonify({"sheets": sheets})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_process_giam_dinh():
    try:
        if "file" not in request.files or not request.files["file"].filename:
            return jsonify({"error": "Vui lòng tải lên file Excel."}), 400

        f = request.files["file"]
        date_range = request.form.get("date_range", "01-14")

        if date_range == "01-14":
            start_day, end_day = 1, 14
        elif date_range == "15-31":
            start_day, end_day = 15, 31
        else:
            start_day, end_day = 1, 31

        out_io = process_giam_dinh_web(f.stream, start_day, end_day)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_process_cntt():
    try:
        sheet_name = request.form.get("sheet_name")
        if not sheet_name:
            return jsonify({"error": "Vui lòng chọn Sheet cần xử lý."}), 400

        if "file" in request.files and request.files["file"].filename:
            stream = request.files["file"].stream
        elif request.form.get("url"):
            url = request.form.get("url").strip()
            stream = download_google_sheet(url)
        else:
            return jsonify({"error": "Vui lòng cung cấp file hoặc URL Google Sheets."}), 400

        out_io = process_cntt_web(stream, sheet_name)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/", methods=["GET", "POST"])
@app.route("/api", methods=["GET", "POST"])
@app.route("/api/index", methods=["GET", "POST"])
@app.route("/api/index.py", methods=["GET", "POST"])
def index_view():
    return render_template_string(HTML_PAGE)

@app.route("/api/get-sheets", methods=["POST"])
def get_sheets_route():
    return handle_get_sheets()

@app.route("/api/process-giam-dinh", methods=["POST"])
def process_giam_dinh_route():
    return handle_process_giam_dinh()

@app.route("/api/process-cntt", methods=["POST"])
def process_cntt_route():
    return handle_process_cntt()

@app.route("/<path:path>", methods=["GET", "POST"])
def catch_all(path):
    if "get-sheets" in path and request.method == "POST":
        return handle_get_sheets()
    if "process-giam-dinh" in path and request.method == "POST":
        return handle_process_giam_dinh()
    if "process-cntt" in path and request.method == "POST":
        return handle_process_cntt()
    return render_template_string(HTML_PAGE)

@app.errorhandler(404)
def error_404(e):
    return render_template_string(HTML_PAGE), 200

if __name__ == "__main__":
    app.run(debug=True, port=5000)
