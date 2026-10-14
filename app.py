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
  <title>Hệ Thống Tổng Hợp Báo Cáo Online - P.CNTT & BHYT</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --primary-light: #eff6ff;
      --secondary: #0f172a;
      --accent: #06b6d4;
      --bg-gradient: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
      --card-bg: rgba(30, 41, 59, 0.7);
      --card-border: rgba(255, 255, 255, 0.1);
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
      --success: #10b981;
      --error: #ef4444;
      --radius: 16px;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Plus Jakarta Sans', sans-serif;
    }

    body {
      min-height: 100vh;
      background: var(--bg-gradient);
      color: var(--text-main);
      padding: 30px 16px;
      display: flex;
      flex-direction: column;
      align-items: center;
    }

    .container {
      width: 100%;
      max-width: 900px;
    }

    .header {
      text-align: center;
      margin-bottom: 28px;
    }

    .header .badge {
      display: inline-block;
      background: rgba(37, 99, 235, 0.2);
      color: #60a5fa;
      border: 1px solid rgba(96, 165, 250, 0.3);
      padding: 6px 16px;
      border-radius: 999px;
      font-size: 0.82rem;
      font-weight: 600;
      margin-bottom: 12px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .header h1 {
      font-size: 2.1rem;
      font-weight: 800;
      letter-spacing: -0.5px;
      background: linear-gradient(to right, #ffffff, #cbd5e1);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      margin-bottom: 8px;
    }

    .header p {
      color: var(--text-muted);
      font-size: 0.95rem;
    }

    /* Tabs */
    .tabs-container {
      display: flex;
      background: rgba(15, 23, 42, 0.6);
      padding: 6px;
      border-radius: 14px;
      border: 1px solid var(--card-border);
      margin-bottom: 24px;
      gap: 6px;
    }

    .tab-btn {
      flex: 1;
      background: transparent;
      border: none;
      outline: none;
      color: var(--text-muted);
      padding: 12px 18px;
      font-size: 0.95rem;
      font-weight: 600;
      border-radius: 10px;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
    }

    .tab-btn.active {
      background: var(--primary);
      color: #ffffff;
      box-shadow: 0 4px 14px rgba(37, 99, 235, 0.4);
    }

    .card {
      background: var(--card-bg);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid var(--card-border);
      border-radius: var(--radius);
      padding: 30px;
      box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3);
      margin-bottom: 24px;
    }

    .tab-content {
      display: none;
    }

    .tab-content.active {
      display: block;
      animation: fadeIn 0.3s ease;
    }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .form-group {
      margin-bottom: 22px;
    }

    .form-group label {
      display: block;
      font-size: 0.9rem;
      font-weight: 600;
      margin-bottom: 8px;
      color: #e2e8f0;
    }

    /* Dropzone */
    .dropzone {
      border: 2px dashed rgba(255, 255, 255, 0.2);
      border-radius: 12px;
      padding: 32px 20px;
      text-align: center;
      cursor: pointer;
      transition: all 0.2s ease;
      background: rgba(15, 23, 42, 0.3);
    }

    .dropzone:hover, .dropzone.dragover {
      border-color: var(--primary);
      background: rgba(37, 99, 235, 0.08);
    }

    .dropzone .icon {
      font-size: 2.4rem;
      margin-bottom: 10px;
    }

    .dropzone .title {
      font-weight: 600;
      font-size: 1rem;
      margin-bottom: 4px;
    }

    .dropzone .desc {
      font-size: 0.82rem;
      color: var(--text-muted);
    }

    .file-chosen {
      display: none;
      margin-top: 12px;
      padding: 10px 14px;
      background: rgba(16, 185, 129, 0.1);
      border: 1px solid rgba(16, 185, 129, 0.3);
      border-radius: 8px;
      font-size: 0.88rem;
      color: #34d399;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    /* Radio Group */
    .radio-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
      gap: 10px;
    }

    .radio-card {
      border: 1px solid var(--card-border);
      background: rgba(15, 23, 42, 0.4);
      padding: 12px 16px;
      border-radius: 10px;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 0.9rem;
      font-weight: 500;
    }

    .radio-card:hover {
      border-color: rgba(255, 255, 255, 0.3);
    }

    .radio-card input {
      accent-color: var(--primary);
      width: 16px;
      height: 16px;
    }

    /* Inputs & Select */
    .text-input, .select-input {
      width: 100%;
      background: rgba(15, 23, 42, 0.5);
      border: 1px solid var(--card-border);
      color: #fff;
      padding: 12px 14px;
      border-radius: 10px;
      font-size: 0.92rem;
      outline: none;
      transition: border-color 0.2s ease;
    }

    .text-input:focus, .select-input:focus {
      border-color: var(--primary);
    }

    .select-input option {
      background: #0f172a;
      color: #fff;
    }

    .btn-quick {
      margin-top: 8px;
      background: rgba(37, 99, 235, 0.15);
      color: #60a5fa;
      border: 1px dashed rgba(96, 165, 250, 0.4);
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 0.8rem;
      cursor: pointer;
      font-weight: 600;
      transition: all 0.2s;
    }

    .btn-quick:hover {
      background: rgba(37, 99, 235, 0.3);
      color: #fff;
    }

    /* Submit Button */
    .btn-submit {
      width: 100%;
      background: var(--primary);
      color: #fff;
      border: none;
      padding: 14px;
      border-radius: 12px;
      font-size: 1rem;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35);
    }

    .btn-submit:hover:not(:disabled) {
      background: var(--primary-hover);
      transform: translateY(-1px);
    }

    .btn-submit:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    /* Alerts */
    .alert {
      padding: 14px 16px;
      border-radius: 10px;
      font-size: 0.9rem;
      margin-top: 18px;
      display: none;
      align-items: center;
      gap: 10px;
    }

    .alert.success {
      background: rgba(16, 185, 129, 0.12);
      border: 1px solid rgba(16, 185, 129, 0.3);
      color: #34d399;
    }

    .alert.error {
      background: rgba(239, 68, 68, 0.12);
      border: 1px solid rgba(239, 68, 68, 0.3);
      color: #f87171;
    }

    .spinner {
      border: 3px solid rgba(255, 255, 255, 0.2);
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

    .footer {
      text-align: center;
      color: var(--text-muted);
      font-size: 0.8rem;
      margin-top: auto;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="badge">🚀 Nền Tảng Trực Tuyến Vercel</div>
      <h1>HỆ THỐNG TỔNG HỢP BÁO CÁO</h1>
      <p>Tổng hợp dữ liệu tự động cho Phòng CNTT & Báo cáo Cổng Giám Định BHYT</p>
    </div>

    <div class="tabs-container">
      <button class="tab-btn active" onclick="switchTab('giamdinh')">
        📋 Cổng Giám Định BHYT
      </button>
      <button class="tab-btn" onclick="switchTab('cntt')">
        💻 Báo Cáo Phòng CNTT
      </button>
    </div>

    <!-- TAB 1: GIÁM ĐỊNH -->
    <div id="tab-giamdinh" class="tab-content active">
      <div class="card">
        <form id="form-giamdinh" onsubmit="handleGiamDinh(event)">
          <div class="form-group">
            <label>1️⃣ Tải lên file Excel Cổng Giám Định (.xlsx, .xlsm):</label>
            <div class="dropzone" id="dz-gd" onclick="document.getElementById('file-gd').click()">
              <div class="icon">📊</div>
              <div class="title">Bấm để chọn file hoặc kéo thả vào đây</div>
              <div class="desc">Chỉ chấp nhận file Excel (.xlsx, .xlsm) có sheet 'Người dùng'</div>
              <input type="file" id="file-gd" accept=".xlsx,.xlsm" style="display: none" onchange="onFileSelected(this, 'chosen-gd')">
            </div>
            <div id="chosen-gd" class="file-chosen" style="display: none;"></div>
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
                <span>Toàn bộ tháng (01 - 31)</span>
              </label>
            </div>
          </div>

          <button type="submit" id="btn-gd" class="btn-submit">
            <span class="spinner" id="spin-gd"></span>
            <span id="txt-gd">🚀 Bắt Đầu Xử Lý & Tải File Kết Quả</span>
          </button>

          <div id="alert-gd" class="alert"></div>
        </form>
      </div>
    </div>

    <!-- TAB 2: P.CNTT -->
    <div id="tab-cntt" class="tab-content">
      <div class="card">
        <form id="form-cntt" onsubmit="handleCNTT(event)">
          <div class="form-group">
            <label>1️⃣ Chọn nguồn dữ liệu Phòng CNTT:</label>
            <div class="radio-grid" style="margin-bottom: 12px;">
              <label class="radio-card">
                <input type="radio" name="cntt_source" value="gsheet" checked onchange="toggleCNTTSource('gsheet')">
                <span>🔗 Link Google Sheets (Tác vụ 6 & 8)</span>
              </label>
              <label class="radio-card">
                <input type="radio" name="cntt_source" value="file" onchange="toggleCNTTSource('file')">
                <span>📁 Tải File Excel Cục Bộ</span>
              </label>
            </div>

            <!-- GSheet Input -->
            <div id="cntt-gsheet-box">
              <input type="text" id="cntt-url" class="text-input" placeholder="https://docs.google.com/spreadsheets/d/..." value="https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139">
              <button type="button" class="btn-quick" onclick="useFixedLink()">
                ⚡ Dùng Link Google Sheets Cố Định (P.CNTT)
              </button>
              <button type="button" class="btn-quick" style="margin-left: 6px; background: rgba(16, 185, 129, 0.15); color: #34d399; border-color: rgba(52, 211, 153, 0.4);" onclick="fetchSheets()">
                🔄 Tải Danh Sách Sheet Ngày
              </button>
            </div>

            <!-- File Input -->
            <div id="cntt-file-box" style="display: none;">
              <div class="dropzone" onclick="document.getElementById('file-cntt').click()">
                <div class="icon">📁</div>
                <div class="title">Bấm để chọn file báo cáo P.CNTT (.xlsx, .xlsm)</div>
                <input type="file" id="file-cntt" accept=".xlsx,.xlsm" style="display: none" onchange="onCNTTFileSelected(this)">
              </div>
              <div id="chosen-cntt" class="file-chosen" style="display: none;"></div>
            </div>
          </div>

          <div class="form-group">
            <label>2️⃣ Chọn Sheet ngày cần tổng hợp:</label>
            <select id="cntt-sheet" class="select-input">
              <option value="">-- Bấm "Tải Danh Sách Sheet Ngày" hoặc Chọn File trước --</option>
            </select>
          </div>

          <button type="submit" id="btn-cntt" class="btn-submit">
            <span class="spinner" id="spin-cntt"></span>
            <span id="txt-cntt">🚀 Tổng Hợp Báo Cáo P.CNTT</span>
          </button>

          <div id="alert-cntt" class="alert"></div>
        </form>
      </div>
    </div>

    <div class="footer">
      Bệnh Viện Đa Khoa Bắc Ninh Số 2 • Phát triển phục vụ công tác P.CNTT & Giám Định BHYT
    </div>
  </div>

  <script>
    const FIXED_URL = "https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139";

    function switchTab(tab) {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      if (tab === 'giamdinh') {
        document.querySelectorAll('.tab-btn')[0].classList.add('active');
        document.getElementById('tab-giamdinh').classList.add('active');
      } else {
        document.querySelectorAll('.tab-btn')[1].classList.add('active');
        document.getElementById('tab-cntt').classList.add('active');
        // Auto fetch sheets for Google Sheet if empty
        const sel = document.getElementById('cntt-sheet');
        if (sel.options.length <= 1) {
          fetchSheets();
        }
      }
    }

    function toggleCNTTSource(type) {
      if (type === 'gsheet') {
        document.getElementById('cntt-gsheet-box').style.display = 'block';
        document.getElementById('cntt-file-box').style.display = 'none';
      } else {
        document.getElementById('cntt-gsheet-box').style.display = 'none';
        document.getElementById('cntt-file-box').style.display = 'block';
      }
    }

    function useFixedLink() {
      document.getElementById('cntt-url').value = FIXED_URL;
      fetchSheets();
    }

    function onFileSelected(input, targetId) {
      const target = document.getElementById(targetId);
      if (input.files && input.files[0]) {
        target.innerHTML = `<span>📄 <b>${input.files[0].name}</b> (${Math.round(input.files[0].size/1024)} KB)</span>`;
        target.style.display = 'flex';
      }
    }

    async function onCNTTFileSelected(input) {
      onFileSelected(input, 'chosen-cntt');
      if (input.files && input.files[0]) {
        const formData = new FormData();
        formData.append('file', input.files[0]);
        await loadSheetsFromFormData(formData);
      }
    }

    async function fetchSheets() {
      const url = document.getElementById('cntt-url').value.trim();
      if (!url) {
        alert('Vui lòng nhập đường link Google Sheets!');
        return;
      }
      const formData = new FormData();
      formData.append('url', url);
      await loadSheetsFromFormData(formData);
    }

    async function loadSheetsFromFormData(formData) {
      const sel = document.getElementById('cntt-sheet');
      sel.innerHTML = '<option value="">⏳ Đang tải danh sách các Sheet ngày...</option>';
      try {
        const res = await fetch('/api/get-sheets', {
          method: 'POST',
          body: formData
        });
        const data = await res.json();
        if (data.error) {
          throw new Error(data.error);
        }
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
        sel.innerHTML = '<option value="">❌ Không lấy được sheet. Vui lòng kiểm tra lại file/link!</option>';
        showAlert('alert-cntt', 'error', 'Lỗi tải Sheet: ' + err.message);
      }
    }

    function showAlert(id, type, msg) {
      const el = document.getElementById(id);
      el.className = 'alert ' + type;
      el.textContent = msg;
      el.style.display = 'flex';
    }

    function hideAlert(id) {
      document.getElementById(id).style.display = 'none';
    }

    // Xử lý Giám Định
    async function handleGiamDinh(e) {
      e.preventDefault();
      hideAlert('alert-gd');

      const fileInput = document.getElementById('file-gd');
      if (!fileInput.files || !fileInput.files[0]) {
        showAlert('alert-gd', 'error', '⚠️ Vui lòng chọn file Excel Cổng Giám Định trước khi thực hiện!');
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
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ khi xử lý dữ liệu.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }

        const blob = await res.blob();
        downloadBlob(blob, 'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx');
        showAlert('alert-gd', 'success', '✅ Đã tổng hợp và tải xuống file Excel báo cáo Giám Định thành công!');
      } catch (err) {
        showAlert('alert-gd', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('gd', false);
      }
    }

    // Xử lý CNTT
    async function handleCNTT(e) {
      e.preventDefault();
      hideAlert('alert-cntt');

      const source = document.querySelector('input[name="cntt_source"]:checked').value;
      const sheetName = document.getElementById('cntt-sheet').value;

      if (!sheetName) {
        showAlert('alert-cntt', 'error', '⚠️ Vui lòng chọn Sheet cần xử lý!');
        return;
      }

      const formData = new FormData();
      formData.append('sheet_name', sheetName);

      if (source === 'gsheet') {
        const url = document.getElementById('cntt-url').value.trim();
        if (!url) {
          showAlert('alert-cntt', 'error', '⚠️ Vui lòng nhập link Google Sheets!');
          return;
        }
        formData.append('url', url);
      } else {
        const fileInput = document.getElementById('file-cntt');
        if (!fileInput.files || !fileInput.files[0]) {
          showAlert('alert-cntt', 'error', '⚠️ Vui lòng chọn file Excel báo cáo công việc!');
          return;
        }
        formData.append('file', fileInput.files[0]);
      }

      setLoading('cntt', true);

      try {
        const res = await fetch('/api/process-cntt', {
          method: 'POST',
          body: formData
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ khi xử lý dữ liệu.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }

        const blob = await res.blob();
        downloadBlob(blob, 'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx');
        showAlert('alert-cntt', 'success', '✅ Đã tổng hợp và tải xuống file Báo cáo P.CNTT thành công!');
      } catch (err) {
        showAlert('alert-cntt', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('cntt', false);
      }
    }

    function setLoading(prefix, isLoading) {
      document.getElementById('btn-' + prefix).disabled = isLoading;
      document.getElementById('spin-' + prefix).style.display = isLoading ? 'inline-block' : 'none';
      document.getElementById('txt-' + prefix).textContent = isLoading ? 'Đang Xử Lý Dữ Liệu...' : (prefix === 'gd' ? '🚀 Bắt Đầu Xử Lý & Tải File Kết Quả' : '🚀 Tổng Hợp Báo Cáo P.CNTT');
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

@app.route("/")
def index():
    return render_template_string(HTML_PAGE)

@app.route("/api/get-sheets", methods=["POST"])
def get_sheets():
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

@app.route("/api/process-giam-dinh", methods=["POST"])
def api_process_giam_dinh():
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

@app.route("/api/process-cntt", methods=["POST"])
def api_process_cntt():
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

if __name__ == "__main__":
    app.run(debug=True, port=5000)
