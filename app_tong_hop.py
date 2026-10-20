import os
import re
import csv
import glob
import sys
import shutil
import copy
import unicodedata
import threading
import zipfile
import tempfile
import urllib.request
import http.client
import xml.etree.ElementTree as ET
from datetime import datetime

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Đảm bảo UTF-8 cho console
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ============================================================
# CẤU HÌNH & CHUẨN HÓA DÙNG CHUNG
# ============================================================
DEFAULT_ORG1 = "ỦY BAN NHÂN DÂN THÀNH PHỐ BẮC NINH"
DEFAULT_ORG2 = "SỞ Y TẾ"
DEFAULT_ORG3 = "BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2"
DEFAULT_PROVINCE = "Bắc Ninh"

def clean_text(value):
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)

def remove_accents(text):
    if not text:
        return ""
    text = str(text).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return text.lower()

def norm_name(value):
    text = clean_text(value)
    text = str(text).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-zA-Z0-9]", "", text).lower()

def norm_cccd(value):
    if not value:
        return ""
    digits = re.sub(r"\D", "", str(value))
    return digits.lstrip("0")

def format_cccd_output(value):
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 11:
        return "0" + digits
    return digits

def normalize_date(value):
    val = clean_text(value)
    if not val:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    match = re.search(r"(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})", val)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        if year < 100:
            year += 2000 if year < 50 else 1900
        try:
            return datetime(year, month, day).strftime("%d/%m/%Y")
        except ValueError:
            return val
    return val

def normalize_issue_place(value):
    val = clean_text(value).upper()
    if "CCSĐKQLCTVDLQGVDC" in val or "ĐKQL" in val or "DÂN CƯ" in val or "BỘ CÔNG AN" in val:
        return "CCSĐKQLCTVDLQGVDC"
    return "CCSQLHCVTTXH"

def clean_email(email_str):
    if not email_str:
        return ""
    email_clean = clean_text(email_str)
    if "@" in email_clean:
        parts = email_clean.split("@")
        return parts[0] + "@" + parts[1]
    return email_clean

def get_file_list(input_val, extensions=(".docx", ".doc")):
    files = []
    if isinstance(input_val, (list, tuple)):
        for item in input_val:
            files.extend(get_file_list(item, extensions))
        return files
    
    if not input_val:
        return []
    
    paths = [p.strip() for p in re.split(r"[;\n]", str(input_val)) if p.strip()]
    for p in paths:
        p = p.strip('"\'')
        if os.path.isfile(p):
            if p.lower().endswith(extensions) and not os.path.basename(p).startswith("~$"):
                files.append(p)
        elif os.path.isdir(p):
            for ext in extensions:
                files.extend(glob.glob(os.path.join(p, f"*{ext}")))
                files.extend(glob.glob(os.path.join(p, f"*{ext.upper()}")))
    
    seen = set()
    result = []
    for f in files:
        if f not in seen and not os.path.basename(f).startswith("~$"):
            seen.add(f)
            result.append(f)
    return result

def read_sso_excel(excel_input, log_func=print):
    excel_files = get_file_list(excel_input, extensions=(".xlsx", ".xls"))
    sso_map = {}
    
    for f in excel_files:
        filename = os.path.basename(f)
        if filename.startswith("~$") or filename in ["Ket_qua.xlsx", "Ket_qua_moi.xlsx", "Ket_qua_final.xlsx", "Template_import VGCA.xlsx", "DANH_SACH_TONG_HOP.xlsx"]:
            continue
        log_func(f"Đang đọc dữ liệu SSO từ file: {filename}")
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    continue
                header = [clean_text(c) for c in rows[0]]
                name_idx, email_idx, cccd_idx, unit_idx, pos_idx = -1, -1, -1, -1, -1
                for idx, col in enumerate(header):
                    c_low = remove_accents(col)
                    if ("ho va ten" in c_low or "ho ten" in c_low) and name_idx == -1:
                        name_idx = idx
                    elif ("email" in c_low or "thu dien tu" in c_low or "username" in c_low or "ten dang nhap" in c_low):
                        if email_idx == -1 or "email" in c_low:
                            email_idx = idx
                    elif ("cccd" in c_low or "cmt" in c_low or "cmnd" in c_low or "description" in c_low) and "ngay" not in c_low and "noi" not in c_low:
                        if cccd_idx == -1:
                            cccd_idx = idx
                    elif ("don vi" in c_low or "co quan" in c_low) and unit_idx == -1:
                        unit_idx = idx
                    elif "chuc vu" in c_low and pos_idx == -1:
                        pos_idx = idx

                if name_idx == -1: name_idx = 0
                if email_idx == -1: email_idx = 1
                if cccd_idx == -1 and len(header) >= 6: cccd_idx = 5

                for r in rows[1:]:
                    if not r:
                        continue
                    name_val = r[name_idx] if name_idx != -1 and name_idx < len(r) else ""
                    email_val = r[email_idx] if email_idx != -1 and email_idx < len(r) else ""
                    cccd_val = r[cccd_idx] if cccd_idx != -1 and cccd_idx < len(r) else ""
                    unit_val = r[unit_idx] if unit_idx != -1 and unit_idx < len(r) else ""
                    pos_val = r[pos_idx] if pos_idx != -1 and pos_idx < len(r) else ""

                    if name_val or email_val or cccd_val:
                        email_c = clean_email(email_val)
                        kn = norm_name(name_val)
                        kc = norm_cccd(cccd_val)
                        info = {
                            "email": email_c,
                            "unit": clean_text(unit_val),
                            "position": clean_text(pos_val)
                        }
                        if kn and kc: sso_map[(kn, kc)] = info
                        if kn and kn not in sso_map: sso_map[kn] = info
                        if kc and kc not in sso_map: sso_map[kc] = info
        except Exception as e:
            log_func(f"Lỗi đọc file Excel {filename}: {e}")

    return sso_map

def generate_mau_02_doc(output_dir, records, log_func=print):
    template_name_base = "Mẫu 02_CV cap CKS cho BVDK BN số 2_T08 2026_3_6"
    meipass = getattr(sys, '_MEIPASS', None)
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    candidate_paths = [
        os.path.join(output_dir, f"{template_name_base}.docx"),
        os.path.join(os.getcwd(), f"{template_name_base}.docx"),
        os.path.join(exe_dir, f"{template_name_base}.docx"),
    ]
    if meipass:
        candidate_paths.append(os.path.join(meipass, f"{template_name_base}.docx"))
    candidate_paths.extend([
        os.path.join(os.path.dirname(os.getcwd()), f"{template_name_base}.docx"),
        os.path.join(r"C:\Users\Admin\Desktop\Công việc\ĐƠN ĐỀ NGHỊ CẤP CTS CỤC CTS 2023-2025\ĐƠN ĐỀ NGHỊ CẤP CTS CỤC CTS 2023-2025\2026 T9", f"{template_name_base}.doc"),
        os.path.join(output_dir, f"{template_name_base}.doc"),
        os.path.join(os.getcwd(), f"{template_name_base}.doc"),
        os.path.join(exe_dir, f"{template_name_base}.doc")
    ])
    
    template_docx = None
    for cp in candidate_paths:
        if os.path.exists(cp):
            if cp.lower().endswith(".docx"):
                template_docx = cp
                break
            elif cp.lower().endswith(".doc"):
                target_docx = os.path.join(output_dir, f"{template_name_base}.docx")
                if not os.path.exists(target_docx):
                    try:
                        import win32com.client
                        word = win32com.client.Dispatch("Word.Application")
                        word.Visible = False
                        w_doc = word.Documents.Open(os.path.abspath(cp))
                        w_doc.SaveAs2(os.path.abspath(target_docx), FileFormat=16)
                        w_doc.Close()
                        word.Quit()
                    except Exception as ex:
                        log_func(f"Chuyển đổi file .doc mẫu bằng win32com không thành công: {ex}")
                if os.path.exists(target_docx):
                    template_docx = target_docx
                    break

    if not template_docx or not os.path.exists(template_docx):
        log_func("  [THÔNG BÁO] Không tìm thấy file mẫu Mẫu 02 (.docx). Hệ thống đang tự động khởi tạo văn bản Mẫu 02 chuẩn Nghị định 130...")
        try:
            doc = docx.Document()
            now = datetime.now()
            for style_name in ['Normal', 'Default Paragraph Font', 'Body Text']:
                if style_name in doc.styles:
                    doc.styles[style_name].font.name = 'Times New Roman'
            for section in doc.sections:
                section.top_margin = docx.shared.Inches(0.8)
                section.bottom_margin = docx.shared.Inches(0.8)
                section.left_margin = docx.shared.Inches(0.9)
                section.right_margin = docx.shared.Inches(0.6)

            table_top = doc.add_table(rows=1, cols=2)
            table_top.autofit = True
            c_left = table_top.cell(0, 0)
            c_right = table_top.cell(0, 1)

            p_left = c_left.paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_l1 = p_left.add_run("ỦY BAN NHÂN DÂN TỈNH BẮC NINH\n")
            r_l1.bold = True
            r_l1.font.size = Pt(11)
            r_l1.font.name = "Times New Roman"
            r_l2 = p_left.add_run("BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2\n")
            r_l2.bold = True
            r_l2.font.size = Pt(11)
            r_l2.font.name = "Times New Roman"
            r_l3 = p_left.add_run("Số:       /CV-BVĐK2")
            r_l3.font.size = Pt(11)
            r_l3.font.name = "Times New Roman"

            p_right = c_right.paragraphs[0]
            p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_r1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
            r_r1.bold = True
            r_r1.font.size = Pt(11)
            r_r1.font.name = "Times New Roman"
            r_r2 = p_right.add_run("Độc lập - Tự do - Hạnh phúc\n")
            r_r2.bold = True
            r_r2.font.size = Pt(11)
            r_r2.font.name = "Times New Roman"
            r_r3 = p_right.add_run(f"Bắc Ninh, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}")
            r_r3.italic = True
            r_r3.font.size = Pt(11)
            r_r3.font.name = "Times New Roman"

            doc.add_paragraph()
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run("CÔNG VĂN ĐỀ NGHỊ CẤP CHỨNG THƯ SỐ CHUYÊN DÙNG CHÍNH PHỦ\n")
            r_title.bold = True
            r_title.font.size = Pt(13)
            r_title.font.name = "Times New Roman"
            r_sub = p_title.add_run("(Mẫu 02 quy định ban hành kèm theo Nghị định số 130/2018/NĐ-CP)")
            r_sub.italic = True
            r_sub.font.size = Pt(10)
            r_sub.font.name = "Times New Roman"

            p_body = doc.add_paragraph()
            r_b1 = p_body.add_run("Kính gửi: Cục Chứng thực số và Bảo mật thông tin – Ban Cơ yếu Chính phủ.\n\n")
            r_b1.font.name = "Times New Roman"
            r_b1.font.size = Pt(11)
            r_b2 = p_body.add_run(f"Căn cứ nhu cầu sử dụng chứng thư số phục vụ ký số và xác thực điện tử trong công tác khám chữa bệnh và quản lý bệnh viện, Bệnh viện Đa khoa Bắc Ninh số 2 kính đề nghị cấp mới chứng thư số cho các cá nhân (tổng số: {len(records):02d} người) có danh sách chi tiết dưới đây:")
            r_b2.font.name = "Times New Roman"
            r_b2.font.size = Pt(11)

            table = doc.add_table(rows=1, cols=9)
            table.style = 'Table Grid'
            headers = ["STT", "Họ và tên", "Ngày sinh", "Số CCCD; Ngày cấp; Nơi cấp", "Thư điện tử công vụ", "Đơn vị công tác", "Tỉnh/TP", "Chức vụ", "Số ĐT"]
            for i, h in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = h
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].bold = True
                    p.runs[0].font.size = Pt(10)
                    p.runs[0].font.name = "Times New Roman"

            alignments = [
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.CENTER
            ]

            for idx, r_data in enumerate(records, 1):
                row = table.add_row()
                vals = [
                    str(idx), r_data[0], r_data[1], r_data[2],
                    r_data[3], r_data[4], r_data[5], r_data[6], r_data[7]
                ]
                for col_idx, text_val in enumerate(vals):
                    c = row.cells[col_idx]
                    c.text = str(text_val) if text_val is not None else ""
                    if c.paragraphs:
                        p = c.paragraphs[0]
                        p.alignment = alignments[col_idx]
                        if p.runs:
                            p.runs[0].font.size = Pt(10)
                            p.runs[0].font.name = "Times New Roman"

            doc.add_paragraph()
            p_sign = doc.add_paragraph()
            p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_sign1 = p_sign.add_run("GIÁM ĐỐC / THỦ TRƯỞNG ĐƠN VỊ\n")
            r_sign1.bold = True
            r_sign1.font.size = Pt(11)
            r_sign1.font.name = "Times New Roman"
            r_sign2 = p_sign.add_run("(Ký, ghi rõ họ tên và đóng dấu)")
            r_sign2.italic = True
            r_sign2.font.size = Pt(10)
            r_sign2.font.name = "Times New Roman"

            out_docx_name = f"{template_name_base}.docx"
            out_docx_path = os.path.join(output_dir, out_docx_name)
            doc.save(out_docx_path)
            log_func(f"  -> ĐÃ TẠO THÀNH CÔNG FILE CÔNG VÂN MẪU 02: {out_docx_name}")
            return out_docx_path
        except Exception as build_err:
            log_func(f"  ❌ Lỗi khi tự khởi tạo văn bản Mẫu 02: {build_err}")
            return None

    try:
        doc = docx.Document(template_docx)
        now = datetime.now()
        date_str = f"Bắc Ninh, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"

        # 1. Cập nhật ngày tháng tạo văn bản (Table 0, R1C1) - Giữ nguyên 100% định dạng nghiêng, cỡ chữ, khoảng cách
        cell_date = doc.tables[0].rows[1].cells[1]
        for p in cell_date.paragraphs:
            if "Bắc Ninh" in p.text:
                if p.runs:
                    p.runs[0].text = date_str
                    for r in p.runs[1:]:
                        r.text = ""

        # 2. Cập nhật số lượng bản ghi ở Mục 2 - Giữ nguyên 100% định dạng in đậm / chữ thường của mẫu
        records_count = len(records)
        for p in doc.paragraphs:
            if "Số lượng và danh" in p.text:
                found_colon = False
                count_runs = []
                for r in p.runs:
                    if ":" in r.text:
                        found_colon = True
                        continue
                    if found_colon:
                        if "," in r.text:
                            break
                        count_runs.append(r)
                if count_runs:
                    count_runs[0].text = f" {records_count:02d}"
                    for r in count_runs[1:]:
                        r.text = ""

        # 3. Điền bảng danh sách đăng ký (Table 1 - 11 cột) bằng cách nhân bản XML mẫu (giữ nguyên cỡ chữ, khoảng cách, viền, font)
        t_target = None
        for t in doc.tables:
            if len(t.columns) == 11:
                t_target = t
                break

        if t_target and len(t_target.rows) > 1:
            template_tr = t_target.rows[1]._tr
            parent = t_target._tbl

            # Xóa các dòng dữ liệu mẫu cũ từ dòng 1 trở đi
            for r in list(t_target.rows)[1:]:
                parent.remove(r._tr)

            for idx, r_data in enumerate(records, start=1):
                new_tr = copy.deepcopy(template_tr)
                new_row = docx.table._Row(new_tr, t_target)
                vals = [
                    idx,
                    r_data[0], # Họ tên
                    r_data[1], # Ngày sinh
                    r_data[2], # CCCD;Ngày cấp;Nơi cấp
                    r_data[3], # Email
                    r_data[4], # Đơn vị
                    r_data[5], # Tỉnh/TP
                    r_data[6], # Chức vụ
                    r_data[7], # SĐT
                    r_data[8], # Serial cũ
                    r_data[9]  # SIM PKI
                ]
                for c_idx, val in enumerate(vals):
                    cell = new_row.cells[c_idx]
                    val_str = str(val) if val is not None else ""
                    p = cell.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = val_str
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = val_str
                parent.append(new_tr)

        out_docx_name = f"{template_name_base}.docx"
        out_docx_path = os.path.join(output_dir, out_docx_name)
        doc.save(out_docx_path)

        log_func(f"  -> ĐÃ TẠO THÀNH CÔNG FILE CÔNG VÂN MẪU 02: {out_docx_name}")
        return out_docx_path
    except Exception as e:
        log_func(f"Lỗi khi tạo file Mẫu 02: {e}")
        return None

def generate_email_doc(output_dir, records, log_func=print):
    template_name_base = "Đề nghị cấp mới địa chỉ email công vụ"
    meipass = getattr(sys, '_MEIPASS', None)
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    candidate_paths = [
        os.path.join(output_dir, f"{template_name_base}.docx"),
        os.path.join(os.getcwd(), f"{template_name_base}.docx"),
        os.path.join(exe_dir, f"{template_name_base}.docx"),
    ]
    if meipass:
        candidate_paths.append(os.path.join(meipass, f"{template_name_base}.docx"))
    candidate_paths.extend([
        os.path.join(os.path.dirname(os.getcwd()), f"{template_name_base}.docx"),
        os.path.join(r"C:\Users\Admin\Desktop\Công việc\ĐƠN ĐỀ NGHỊ CẤP CTS CỤC CTS 2023-2025\ĐƠN ĐỀ NGHỊ CẤP CTS CỤC CTS 2023-2025\2026 T8", "27.08.2026. Đề nghị cấp mới địa chỉ email công vụ_TTBVSK_TTTM_HSTC_Trang_Nam_Thảo_Diện.docx"),
        os.path.join(output_dir, f"{template_name_base}.doc"),
        os.path.join(os.getcwd(), f"{template_name_base}.doc"),
        os.path.join(exe_dir, f"{template_name_base}.doc")
    ])
    
    template_docx = None
    for cp in candidate_paths:
        if os.path.exists(cp):
            if cp.lower().endswith(".docx"):
                template_docx = cp
                break
            elif cp.lower().endswith(".doc"):
                target_docx = os.path.join(output_dir, f"{template_name_base}.docx")
                if not os.path.exists(target_docx):
                    try:
                        import win32com.client
                        word = win32com.client.Dispatch("Word.Application")
                        word.Visible = False
                        w_doc = word.Documents.Open(os.path.abspath(cp))
                        w_doc.SaveAs2(os.path.abspath(target_docx), FileFormat=16)
                        w_doc.Close()
                        word.Quit()
                    except Exception as ex:
                        log_func(f"Chuyển đổi file .doc mẫu email bằng win32com không thành công: {ex}")
                if os.path.exists(target_docx):
                    template_docx = target_docx
                    break

    if not template_docx or not os.path.exists(template_docx):
        log_func("  [CẢNH BÁO] Không tìm thấy file mẫu Email (.docx)!")
        return None

    try:
        doc = docx.Document(template_docx)
        now = datetime.now()
        date_str = f"Bắc Ninh, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"

        # 1. Cập nhật ngày tháng văn bản (Table 0, R1C1) - Giữ nguyên 100% định dạng nghiêng & font
        cell_date = doc.tables[0].rows[1].cells[1]
        for p in cell_date.paragraphs:
            if "Bắc Ninh" in p.text:
                if p.runs:
                    p.runs[0].text = date_str
                    for r in p.runs[1:]:
                        r.text = ""

        # 2. Cập nhật số lượng bản ghi ở Mục 2 - Giữ nguyên 100% định dạng in đậm / chữ thường
        records_count = len(records)
        for p in doc.paragraphs:
            if "Số lượng và danh" in p.text:
                found_colon = False
                count_runs = []
                for r in p.runs:
                    if ":" in r.text:
                        found_colon = True
                        continue
                    if found_colon:
                        if "," in r.text:
                            break
                        count_runs.append(r)
                if count_runs:
                    count_runs[0].text = f" {records_count:02d}"
                    for r in count_runs[1:]:
                        r.text = ""

        # 3. Điền bảng danh sách đăng ký Email (Table 1 - 8 cột) nhân bản XML mẫu giữ nguyên cỡ chữ & viền
        t_target = None
        for t in doc.tables:
            if len(t.columns) == 8:
                t_target = t
                break

        if t_target and len(t_target.rows) > 1:
            template_tr = t_target.rows[1]._tr
            parent = t_target._tbl

            for r in list(t_target.rows)[1:]:
                parent.remove(r._tr)

            for idx, r_data in enumerate(records, start=1):
                new_tr = copy.deepcopy(template_tr)
                new_row = docx.table._Row(new_tr, t_target)
                vals = [
                    idx,
                    r_data[0], # Họ tên
                    r_data[1], # Ngày sinh
                    r_data[2], # Di động
                    r_data[3], # CCCD
                    r_data[4], # Đơn vị công tác
                    r_data[5], # Chức vụ
                    r_data[6]  # Ghi chú
                ]
                for c_idx, val in enumerate(vals):
                    cell = new_row.cells[c_idx]
                    val_str = str(val) if val is not None else ""
                    p = cell.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = val_str
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = val_str
                parent.append(new_tr)

        out_docx_name = f"{template_name_base}.docx"
        out_docx_path = os.path.join(output_dir, out_docx_name)
        doc.save(out_docx_path)

        log_func(f"  -> ĐÃ TẠO THÀNH CÔNG FILE CÔNG VÂN EMAIL: {out_docx_name}")
        return out_docx_path
    except Exception as e:
        log_func(f"Lỗi khi tạo file Công văn Email: {e}")
        return None

# ============================================================
# LOGIC TÁC VỤ 1: ĐỐI CHIẾU DỮ LIỆU & IMPORT VGCA (xu_ly_doi_chieu)
# ============================================================
def run_task_doi_chieu(word_input, excel_input, log_func=print, progress_func=lambda v: None):
    log_func("=== BẮT ĐẦU TÁC VỤ 1: ĐỐI CHIẾU DỮ LIỆU WORD & EXCEL SSO (XUẤT Ket_qua.xlsx) ===")
    
    word_files = get_file_list(word_input, extensions=(".docx", ".doc"))
    if not word_files:
        raise ValueError("Không tìm thấy file Word (.docx, .doc) nào để xử lý! Vui lòng chọn file Word.")

    sso_map = read_sso_excel(excel_input, log_func)
    if sso_map:
        log_func(f"  -> Đã nạp {len(sso_map)} bản ghi từ SSO Excel")
    else:
        log_func("  [CẢNH BÁO] Không nạp được dữ liệu SSO từ Excel - Cột email sẽ để trống!")

    people_list = []
    total_files = len(word_files)
    
    for i, f in enumerate(word_files, start=1):
        filename = os.path.basename(f)
        log_func(f"\nĐang xử lý file Word ({i}/{total_files}): {filename}")
        progress_func(int((i / max(1, total_files)) * 80))
        
        try:
            doc = docx.Document(f)
            for t_idx, table in enumerate(doc.tables):
                if not table.rows:
                    continue
                header = [clean_text(cell.text) for cell in table.rows[0].cells]
                name_idx, cccd_idx, dob_idx, issue_date_idx, issue_place_idx, phone_idx, pos_idx, unit_idx = [-1]*8
                
                for idx, col in enumerate(header):
                    c_no_acc = remove_accents(col)
                    if ("ho va ten" in c_no_acc or "ho ten" in c_no_acc) and "nguoi lap" not in c_no_acc and "truong" not in c_no_acc and "ky va" not in c_no_acc:
                        if name_idx == -1: name_idx = idx
                    elif ("cccd" in c_no_acc or "cmt" in c_no_acc or "cmnd" in c_no_acc or "can cuoc" in c_no_acc) and "noi cap" not in c_no_acc and "ngay cap" not in c_no_acc and "chup anh" not in c_no_acc and "anh" not in c_no_acc:
                        if cccd_idx == -1: cccd_idx = idx
                    elif "ngay sinh" in c_no_acc and dob_idx == -1:
                        dob_idx = idx
                    elif "ngay cap" in c_no_acc and issue_date_idx == -1:
                        issue_date_idx = idx
                    elif "noi cap" in c_no_acc and issue_place_idx == -1:
                        issue_place_idx = idx
                    elif ("sdt" in c_no_acc or "dien thoai" in c_no_acc) and phone_idx == -1:
                        phone_idx = idx
                    elif "chuc vu" in c_no_acc and pos_idx == -1:
                        pos_idx = idx
                    elif ("don vi" in c_no_acc or "co quan" in c_no_acc) and unit_idx == -1:
                        unit_idx = idx
                
                if name_idx != -1 and cccd_idx != -1:
                    log_func(f"  -> Tìm thấy bảng dữ liệu (Bảng {t_idx+1})")
                    for r in table.rows[1:]:
                        vals = [clean_text(cell.text) for cell in r.cells]
                        if len(vals) > max(name_idx, cccd_idx):
                            name_val = vals[name_idx]
                            cccd_val = vals[cccd_idx]
                            if name_val and cccd_val and "ho va ten" not in remove_accents(name_val):
                                kn = norm_name(name_val)
                                kc = norm_cccd(cccd_val)
                                sso_info = sso_map.get((kn, kc)) or sso_map.get(kc) or sso_map.get(kn) or {}
                                email = sso_info.get("email", "")
                                sso_unit = sso_info.get("unit", "")
                                sso_pos = sso_info.get("position", "")

                                dob = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                                issue_date = normalize_date(vals[issue_date_idx]) if issue_date_idx != -1 and issue_date_idx < len(vals) else ""
                                issue_place = normalize_issue_place(vals[issue_place_idx]) if issue_place_idx != -1 and issue_place_idx < len(vals) else "CCSQLHCVTTXH"
                                phone = vals[phone_idx] if phone_idx != -1 and phone_idx < len(vals) else ""
                                pos = vals[pos_idx] if pos_idx != -1 and pos_idx < len(vals) else (sso_pos if sso_pos else "Nhân viên")
                                unit = vals[unit_idx] if unit_idx != -1 and unit_idx < len(vals) else (sso_unit if sso_unit else DEFAULT_ORG3)

                                person = {
                                    "name": name_val, "dob": dob, "cccd": format_cccd_output(cccd_val),
                                    "issue_place": issue_place, "issue_date": issue_date, "email": email,
                                    "org1": DEFAULT_ORG1, "org2": DEFAULT_ORG2, "org3": unit if unit else DEFAULT_ORG3,
                                    "org4": "", "province": DEFAULT_PROVINCE, "position": pos, "phone": phone,
                                    "serial_old": "", "network": "", "is_transfer": "", "device_type": ""
                                }
                                people_list.append(person)
                                status_str = f"Email: {email}" if email else "CHƯA CÓ EMAIL SSO"
                                log_func(f"     + Lấy được: {name_val} - CCCD: {format_cccd_output(cccd_val)} -> {status_str}")
        except Exception as e:
            log_func(f"Lỗi đọc file Word {filename}: {e}")

    output_dir = os.path.dirname(word_files[0]) if word_files else os.getcwd()
    template_file = os.path.join(output_dir, "Template_import VGCA.xlsx")
    if not os.path.exists(template_file):
        template_file = os.path.join(output_dir, "Ket_qua.xlsx")

    if os.path.exists(template_file):
        try:
            wb = openpyxl.load_workbook(template_file)
        except Exception:
            wb = openpyxl.Workbook()
    else:
        wb = openpyxl.Workbook()

    ws = wb.active
    ws.cell(row=1, column=1, value="STT")
    ws.cell(row=1, column=2, value="Họ tên (bắt buộc)")
    ws.cell(row=1, column=3, value="Ngày sinh (bắt buộc, để dạng dd/MM/YYYY)")
    ws.cell(row=1, column=4, value="CCCC (Căn cước công dân, bắt buộc)")
    ws.cell(row=1, column=5, value="Nơi cấp (Nơi cấp CCCD; để một trong 2 nơi: CCSQLHCVTTXH hoặc CCSĐKQLCTVDLQGVDC)")
    ws.cell(row=1, column=6, value="Ngày cấp (bắt buộc, để dạng dd/MM/YYYY)")
    ws.cell(row=1, column=7, value="Địa chỉ thư điện tử công vụ (bắt buộc)")
    ws.cell(row=1, column=8, value="Tổ chức cấp 1 (bắt buộc)")
    ws.cell(row=1, column=9, value="Tổ chức cấp 2")
    ws.cell(row=1, column=10, value="Tổ chức cấp 3")
    ws.cell(row=1, column=11, value="Tổ chức cấp 4")
    ws.cell(row=1, column=12, value="Tỉnh thành phố (Bắt buộc)")
    ws.cell(row=1, column=13, value="Tỉnh/Thành phố")
    ws.cell(row=1, column=14, value="Chức vụ (Bắt buộc)")
    ws.cell(row=1, column=15, value="Điện thoại")
    ws.cell(row=1, column=16, value="Serial CTS cũ")
    ws.cell(row=1, column=17, value="Nhà mạng (bắt buộc nếu đăng ký SIM)")
    ws.cell(row=1, column=18, value="Là chuyển số (đánh dấu X nếu muốn chuyển số đang dùng sang SIM ký số)")
    ws.cell(row=1, column=19, value="Loại thiết bị( nếu không ghi sẽ là token; là một trong các loại sau: (TOKEN, SIM, SOFT, RSSP, HSM); RSSP là ký số tập trung")

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for idx, p in enumerate(people_list, start=1):
        row_num = idx + 1
        ws.cell(row=row_num, column=1, value=idx)
        ws.cell(row=row_num, column=2, value=p["name"])
        ws.cell(row=row_num, column=3, value=p["dob"])
        ws.cell(row=row_num, column=4, value=str(p["cccd"]))
        ws.cell(row=row_num, column=5, value=p["issue_place"])
        ws.cell(row=row_num, column=6, value=p["issue_date"])
        ws.cell(row=row_num, column=7, value=p["email"])
        ws.cell(row=row_num, column=8, value=p["org1"])
        ws.cell(row=row_num, column=9, value=p["org2"])
        ws.cell(row=row_num, column=10, value=p["org3"])
        ws.cell(row=row_num, column=11, value=p["org4"])
        ws.cell(row=row_num, column=12, value=p["province"])
        ws.cell(row=row_num, column=13, value=DEFAULT_PROVINCE)
        ws.cell(row=row_num, column=14, value=p["position"])
        ws.cell(row=row_num, column=15, value=p["phone"])
        ws.cell(row=row_num, column=16, value=p["serial_old"])
        ws.cell(row=row_num, column=17, value=p["network"])
        ws.cell(row=row_num, column=18, value=p["is_transfer"])
        ws.cell(row=row_num, column=19, value=p["device_type"])
        for col in range(1, 20):
            ws.cell(row=row_num, column=col).number_format = "@"

    output_file = os.path.join(output_dir, "Ket_qua.xlsx")
    saved = False
    for candidate in [output_file, os.path.join(output_dir, "Ket_qua_moi.xlsx"), os.path.join(output_dir, "Ket_qua_final.xlsx")]:
        try:
            wb.save(candidate)
            output_file = candidate
            saved = True
            break
        except PermissionError:
            continue

    if not saved:
        output_file = os.path.join(output_dir, f"Ket_qua_{int(datetime.now().timestamp())}.xlsx")
        wb.save(output_file)

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ XUẤT THÀNH CÔNG FILE KẾT QUẢ: {os.path.basename(output_file)}")
    log_func(f"Đường dẫn lưu file: {output_file}")
    log_func(f"Tổng số bản ghi: {len(people_list)}")
    log_func(f"==========================================")
    return output_file

# ============================================================
# LOGIC TÁC VỤ 7: TỔNG HỢP VGCA TỪ NHIỀU FILE WORD (XUẤT Ket_qua.xlsx)
# ============================================================
def run_task_vgca_word(word_input, log_func=print, progress_func=lambda v: None, output_file=None):
    log_func("=== BẮT ĐẦU TÁC VỤ 7: TỔNG HỢP VGCA TỪ CÁC FILE WORD (XUẤT Ket_qua.xlsx) ===")

    word_files = get_file_list(word_input, extensions=(".docx", ".doc"))
    if not word_files:
        raise ValueError("Không tìm thấy file Word (.docx, .doc) nào để xử lý! Vui lòng chọn file Word hoặc thư mục chứa file Word.")

    total_files = len(word_files)
    log_func(f"Tìm thấy tổng cộng {total_files} file Word cần xử lý.")
    people_list = []

    for i, target_file in enumerate(word_files, start=1):
        filename = os.path.basename(target_file)
        log_func(f"\nĐang đọc file Word ({i}/{total_files}): {filename}")
        progress_func(int((i / max(1, total_files)) * 80))

        actual_file = target_file
        if target_file.lower().endswith(".doc"):
            converted_docx = os.path.splitext(target_file)[0] + "_converted.docx"
            if not os.path.exists(converted_docx):
                try:
                    import win32com.client
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    w_doc = word.Documents.Open(os.path.abspath(target_file))
                    w_doc.SaveAs2(os.path.abspath(converted_docx), FileFormat=16)
                    w_doc.Close()
                    word.Quit()
                    actual_file = converted_docx
                    log_func(f"  -> Đã chuyển đổi file .doc sang .docx: {os.path.basename(actual_file)}")
                except Exception as e:
                    log_func(f"  [CẢNH BÁO] Không thể chuyển đổi tự động file .doc bằng win32com: {e}")
            else:
                actual_file = converted_docx

        try:
            doc = docx.Document(actual_file)
        except Exception as e:
            log_func(f"  ❌ Không thể mở file {filename}: {e}")
            continue

        file_people = []

        # 1. Quét các bảng trong tài liệu
        for t_idx, table in enumerate(doc.tables):
            if not table.rows:
                continue

            header_row_idx = -1
            name_idx, cccd_idx, dob_idx = -1, -1, -1
            issue_date_idx, issue_place_idx, email_idx = -1, -1, -1
            unit_idx, pos_idx, phone_idx, prov_idx = -1, -1, -1, -1
            serial_idx, network_idx, transfer_idx, device_idx = -1, -1, -1, -1

            for r_idx, row in enumerate(table.rows[:min(4, len(table.rows))]):
                header_texts = [clean_text(cell.text) for cell in row.cells]
                t_name, t_cccd, t_dob, t_idate, t_iplace, t_email = -1, -1, -1, -1, -1, -1
                t_unit, t_pos, t_phone, t_prov = -1, -1, -1, -1
                t_serial, t_network, t_transfer, t_dev = -1, -1, -1, -1

                for idx, col in enumerate(header_texts):
                    c_norm = remove_accents(col)
                    is_combo = (';' in col or ';' in c_norm) and any(k in c_norm for k in ['cccd', 'cmnd', 'cmt', 'can cuoc'])
                    if ("ho va ten" in c_norm or "ho ten" in c_norm) and "nguoi lap" not in c_norm and "truong" not in c_norm and "ky va" not in c_norm:
                        if t_name == -1: t_name = idx
                    elif is_combo:
                        if t_cccd == -1: t_cccd = idx
                    elif "ngay cap" in c_norm and "noi" not in c_norm:
                        if t_idate == -1: t_idate = idx
                    elif "noi cap" in c_norm:
                        if t_iplace == -1: t_iplace = idx
                    elif any(k in c_norm for k in ['cccd', 'cmt', 'cmnd', 'can cuoc']) and "chup anh" not in c_norm and "anh" not in c_norm and "ngay cap" not in c_norm and "noi cap" not in c_norm:
                        if t_cccd == -1: t_cccd = idx
                    elif "ngay sinh" in c_norm or "nam sinh" in c_norm or "dob" in c_norm:
                        if t_dob == -1: t_dob = idx
                    elif any(k in c_norm for k in ['email', 'thu dien tu', 'mail', 'hop thu']):
                        if t_email == -1: t_email = idx
                    elif any(k in c_norm for k in ['don vi', 'co quan', 'to chuc', 'khoa', 'phong']) and "tinh" not in c_norm:
                        if t_unit == -1: t_unit = idx
                    elif "chuc vu" in c_norm or "vi tri" in c_norm:
                        if t_pos == -1: t_pos = idx
                    elif any(k in c_norm for k in ['sdt', 'dien thoai', 'di dong', 'phone']):
                        if t_phone == -1: t_phone = idx
                    elif ("tinh" in c_norm or "thanh pho" in c_norm) and "don vi" not in c_norm and "gioi" not in c_norm:
                        if t_prov == -1: t_prov = idx
                    elif ("serial" in c_norm or "chung thu cu" in c_norm or "cts cu" in c_norm or "so hieu" in c_norm):
                        if t_serial == -1: t_serial = idx
                    elif ("nha mang" in c_norm or "mang" in c_norm):
                        if t_network == -1: t_network = idx
                    elif ("chuyen so" in c_norm):
                        if t_transfer == -1: t_transfer = idx
                    elif ("thiet bi" in c_norm or "sim" in c_norm or "pki" in c_norm):
                        if t_dev == -1: t_dev = idx

                if t_name != -1 and t_cccd != -1:
                    header_row_idx = r_idx
                    name_idx, cccd_idx, dob_idx = t_name, t_cccd, t_dob
                    issue_date_idx, issue_place_idx, email_idx = t_idate, t_iplace, t_email
                    unit_idx, pos_idx, phone_idx, prov_idx = t_unit, t_pos, t_phone, t_prov
                    serial_idx, network_idx, transfer_idx, device_idx = t_serial, t_network, t_transfer, t_dev
                    break

            if header_row_idx == -1:
                continue

            log_func(f"  -> Tìm thấy bảng dữ liệu (Bảng {t_idx+1}, dòng tiêu đề {header_row_idx+1})")

            for r_idx in range(header_row_idx + 1, len(table.rows)):
                row = table.rows[r_idx]
                vals = [clean_text(cell.text) for cell in row.cells]
                if len(vals) <= max(name_idx, cccd_idx):
                    continue

                name_val = vals[name_idx]
                cccd_raw = vals[cccd_idx]

                if not name_val or not cccd_raw:
                    continue
                if "ho va ten" in remove_accents(name_val) or "cong hoa xa hoi" in remove_accents(name_val):
                    continue

                dob_val = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                issue_date = normalize_date(vals[issue_date_idx]) if issue_date_idx != -1 and issue_date_idx < len(vals) else ""
                issue_place = normalize_issue_place(vals[issue_place_idx]) if issue_place_idx != -1 and issue_place_idx < len(vals) else ""

                cccd_clean = cccd_raw
                if ";" in cccd_raw:
                    parts = [clean_text(p) for p in cccd_raw.split(";")]
                    cccd_clean = parts[0]
                    if not issue_date and len(parts) > 1 and parts[1]:
                        issue_date = normalize_date(parts[1])
                    if (not issue_place or issue_place == "CCSQLHCVTTXH") and len(parts) > 2 and parts[2]:
                        issue_place = normalize_issue_place(parts[2])

                if not issue_place:
                    issue_place = "CCSQLHCVTTXH"

                email_val = clean_email(vals[email_idx]) if email_idx != -1 and email_idx < len(vals) else ""
                if not email_val:
                    for c_val in vals:
                        if "@" in c_val and "." in c_val:
                            m_mail = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", c_val)
                            if m_mail:
                                email_val = clean_email(m_mail.group(0))
                                break

                pos_val = vals[pos_idx] if pos_idx != -1 and pos_idx < len(vals) else "Nhân viên"
                phone_val = vals[phone_idx] if phone_idx != -1 and phone_idx < len(vals) else ""
                prov_val = vals[prov_idx] if prov_idx != -1 and prov_idx < len(vals) else DEFAULT_PROVINCE
                if not prov_val or prov_val.lower() in ["nam", "nu", "nữ", ""]:
                    prov_val = DEFAULT_PROVINCE
                serial_val = vals[serial_idx] if serial_idx != -1 and serial_idx < len(vals) else ""
                network_val = vals[network_idx] if network_idx != -1 and network_idx < len(vals) else ""
                transfer_val = vals[transfer_idx] if transfer_idx != -1 and transfer_idx < len(vals) else ""
                dev_val = vals[device_idx] if device_idx != -1 and device_idx < len(vals) else ""

                person = {
                    "name": name_val,
                    "dob": dob_val,
                    "cccd": format_cccd_output(cccd_clean),
                    "issue_place": issue_place,
                    "issue_date": issue_date,
                    "email": email_val,
                    "org1": DEFAULT_ORG1,
                    "org2": DEFAULT_ORG2,
                    "org3": DEFAULT_ORG3,
                    "org4": "",
                    "province": prov_val if prov_val else DEFAULT_PROVINCE,
                    "position": pos_val if pos_val else "Nhân viên",
                    "phone": phone_val,
                    "serial_old": serial_val,
                    "network": network_val,
                    "is_transfer": transfer_val,
                    "device_type": dev_val
                }
                file_people.append(person)
                status_str = f"Email: {email_val}" if email_val else "CHƯA CÓ EMAIL"
                log_func(f"     + Lấy được: {name_val} - CCCD: {format_cccd_output(cccd_clean)} -> {status_str}")

        # 2. Nếu không tìm thấy bảng trong file này, quét đoạn văn bản
        if not file_people:
            full_text = "\n".join(p.text for p in doc.paragraphs)
            m_name = re.search(r"(?:Họ và tên|Họ tên|Tên tôi là)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)
            m_cccd = re.search(r"(?:CCCD|Số CCCD|CMND|Số CMND)\s*[:：]\s*([0-9]{9,12})", full_text, re.IGNORECASE)
            m_dob = re.search(r"(?:Ngày sinh|Sinh ngày)\s*[:：]\s*([0-9\/\-\.]+)", full_text, re.IGNORECASE)
            m_mail = re.search(r"(?:Email|Thư điện tử|Hòm thư)\s*[:：]?\s*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", full_text, re.IGNORECASE)
            m_phone = re.search(r"(?:Điện thoại|SĐT|Số điện thoại|Di động)\s*[:：]\s*([0-9\.\s]+)", full_text, re.IGNORECASE)
            m_pos = re.search(r"(?:Chức vụ)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)
            m_unit = re.search(r"(?:Đơn vị|Khoa\/Phòng|Cơ quan)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)

            if m_name and m_cccd:
                name_v = clean_text(m_name.group(1))
                cccd_v = clean_text(m_cccd.group(1))
                dob_v = normalize_date(m_dob.group(1)) if m_dob else ""
                email_v = clean_email(m_mail.group(1)) if m_mail else ""
                phone_v = clean_text(m_phone.group(1)) if m_phone else ""
                pos_v = clean_text(m_pos.group(1)) if m_pos else "Nhân viên"

                file_people.append({
                    "name": name_v, "dob": dob_v, "cccd": format_cccd_output(cccd_v),
                    "issue_place": "CCSQLHCVTTXH", "issue_date": "", "email": email_v,
                    "org1": DEFAULT_ORG1, "org2": DEFAULT_ORG2,
                    "org3": DEFAULT_ORG3,
                    "org4": "",
                    "province": DEFAULT_PROVINCE, "position": pos_v, "phone": phone_v,
                    "serial_old": "", "network": "", "is_transfer": "", "device_type": ""
                })
                log_func(f"     + Lấy được từ đoạn văn: {name_v} - CCCD: {format_cccd_output(cccd_v)}")

        people_list.extend(file_people)

    if not people_list:
        raise ValueError("Không trích xuất được bản ghi nào từ các file Word đã chọn!\nVui lòng kiểm tra lại bảng danh sách hoặc nội dung trong các file Word.")

    progress_func(85)

    # 3. Xuất ra file Ket_qua.xlsx đúng 19 cột chuẩn VGCA y hệt Tác vụ 1
    output_dir = os.path.dirname(word_files[0]) if word_files else os.getcwd()
    template_file = os.path.join(output_dir, "Template_import VGCA.xlsx")
    if not os.path.exists(template_file):
        template_file = os.path.join(output_dir, "Ket_qua.xlsx")

    if os.path.exists(template_file):
        try:
            wb = openpyxl.load_workbook(template_file)
        except Exception:
            wb = openpyxl.Workbook()
    else:
        wb = openpyxl.Workbook()

    ws = wb.active
    ws.cell(row=1, column=1, value="STT")
    ws.cell(row=1, column=2, value="Họ tên (bắt buộc)")
    ws.cell(row=1, column=3, value="Ngày sinh (bắt buộc, để dạng dd/MM/YYYY)")
    ws.cell(row=1, column=4, value="CCCC (Căn cước công dân, bắt buộc)")
    ws.cell(row=1, column=5, value="Nơi cấp (Nơi cấp CCCD; để một trong 2 nơi: CCSQLHCVTTXH hoặc CCSĐKQLCTVDLQGVDC)")
    ws.cell(row=1, column=6, value="Ngày cấp (bắt buộc, để dạng dd/MM/YYYY)")
    ws.cell(row=1, column=7, value="Địa chỉ thư điện tử công vụ (bắt buộc)")
    ws.cell(row=1, column=8, value="Tổ chức cấp 1 (bắt buộc)")
    ws.cell(row=1, column=9, value="Tổ chức cấp 2")
    ws.cell(row=1, column=10, value="Tổ chức cấp 3")
    ws.cell(row=1, column=11, value="Tổ chức cấp 4")
    ws.cell(row=1, column=12, value="Tỉnh thành phố (Bắt buộc)")
    ws.cell(row=1, column=13, value="Tỉnh/Thành phố")
    ws.cell(row=1, column=14, value="Chức vụ (Bắt buộc)")
    ws.cell(row=1, column=15, value="Điện thoại")
    ws.cell(row=1, column=16, value="Serial CTS cũ")
    ws.cell(row=1, column=17, value="Nhà mạng (bắt buộc nếu đăng ký SIM)")
    ws.cell(row=1, column=18, value="Là chuyển số (đánh dấu X nếu muốn chuyển số đang dùng sang SIM ký số)")
    ws.cell(row=1, column=19, value="Loại thiết bị( nếu không ghi sẽ là token; là một trong các loại sau: (TOKEN, SIM, SOFT, RSSP, HSM); RSSP là ký số tập trung")

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for idx, p in enumerate(people_list, start=1):
        row_num = idx + 1
        ws.cell(row=row_num, column=1, value=idx)
        ws.cell(row=row_num, column=2, value=p["name"])
        ws.cell(row=row_num, column=3, value=p["dob"])
        ws.cell(row=row_num, column=4, value=str(p["cccd"]))
        ws.cell(row=row_num, column=5, value=p["issue_place"])
        ws.cell(row=row_num, column=6, value=p["issue_date"])
        ws.cell(row=row_num, column=7, value=p["email"])
        ws.cell(row=row_num, column=8, value=p["org1"])
        ws.cell(row=row_num, column=9, value=p["org2"])
        ws.cell(row=row_num, column=10, value=p["org3"])
        ws.cell(row=row_num, column=11, value=p["org4"])
        ws.cell(row=row_num, column=12, value=p["province"])
        ws.cell(row=row_num, column=13, value=DEFAULT_PROVINCE)
        ws.cell(row=row_num, column=14, value=p["position"])
        ws.cell(row=row_num, column=15, value=p["phone"])
        ws.cell(row=row_num, column=16, value=p["serial_old"])
        ws.cell(row=row_num, column=17, value=p["network"])
        ws.cell(row=row_num, column=18, value=p["is_transfer"])
        ws.cell(row=row_num, column=19, value=p["device_type"])
        for col in range(1, 20):
            ws.cell(row=row_num, column=col).number_format = "@"

    if not output_file:
        output_file = os.path.join(output_dir, "Ket_qua.xlsx")
    saved = False
    target_dir = os.path.dirname(output_file) or output_dir
    target_base = os.path.splitext(os.path.basename(output_file))[0]
    for candidate in [output_file, os.path.join(target_dir, f"{target_base}_moi.xlsx"), os.path.join(target_dir, f"{target_base}_final.xlsx")]:
        try:
            wb.save(candidate)
            output_file = candidate
            saved = True
            break
        except PermissionError:
            continue

    if not saved:
        output_file = os.path.join(target_dir, f"{target_base}_{int(datetime.now().timestamp())}.xlsx")
        wb.save(output_file)

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ TỔNG HỢP THÀNH CÔNG TỪ {total_files} FILE WORD VÀO: {os.path.basename(output_file)}")
    log_func(f"Đường dẫn lưu file: {output_file}")
    log_func(f"Tổng số bản ghi: {len(people_list)}")
    log_func(f"==========================================")
    return output_file

run_task_vgca_don_file = run_task_vgca_word

# ============================================================
# LOGIC TÁC VỤ 2: TỔNG HỢP XIN CẤP CKS (đọc Word + đối chiếu SSO Excel -> Xuất Mẫu 02)
# ============================================================
def run_task_cks(word_input, excel_input, log_func=print, progress_func=lambda v: None):
    log_func("=== BẮT ĐẦU TÁC VỤ 2: TỔNG HỢP XIN CẤP CKS (XUẤT CÔNG VÂN MẪU 02) ===")

    word_files = get_file_list(word_input, extensions=(".docx", ".doc"))
    if not word_files:
        raise ValueError("Không tìm thấy file Word (.docx, .doc) nào để xử lý! Vui lòng chọn file Word.")

    sso_map = read_sso_excel(excel_input, log_func)
    if sso_map:
        log_func(f"  -> Đã nạp {len(sso_map)} bản ghi từ SSO Excel")
    else:
        log_func("  [CẢNH BÁO] Không nạp được dữ liệu SSO từ Excel – Cột email sẽ để trống!")

    total_files = len(word_files)
    records = []

    for i, f in enumerate(word_files, start=1):
        filename = os.path.basename(f)
        log_func(f"\nĐang đọc file Word ({i}/{total_files}): {filename}")
        progress_func(int((i / max(1, total_files)) * 80))
        try:
            doc = docx.Document(f)
            for table in doc.tables:
                if not table.rows: continue
                header = [clean_text(cell.text) for cell in table.rows[0].cells]
                name_idx, dob_idx, phone_idx, cccd_idx, issue_date_idx, issue_place_idx, unit_idx, pos_idx = [-1]*8

                for idx, col in enumerate(header):
                    c_no_acc = remove_accents(col)
                    if ("ho va ten" in c_no_acc or "ho ten" in c_no_acc) and "nguoi lap" not in c_no_acc and "truong" not in c_no_acc: name_idx = idx
                    elif "ngay sinh" in c_no_acc and dob_idx == -1: dob_idx = idx
                    elif ("sdt" in c_no_acc or "dien thoai" in c_no_acc or "di dong" in c_no_acc) and phone_idx == -1: phone_idx = idx
                    elif ("cccd" in c_no_acc or "cmt" in c_no_acc or "cmnd" in c_no_acc or "can cuoc" in c_no_acc) and "noi cap" not in c_no_acc and "ngay cap" not in c_no_acc and "chup anh" not in c_no_acc and "anh" not in c_no_acc:
                        if cccd_idx == -1: cccd_idx = idx
                    elif "ngay cap" in c_no_acc and issue_date_idx == -1: issue_date_idx = idx
                    elif "noi cap" in c_no_acc and issue_place_idx == -1: issue_place_idx = idx
                    elif ("don vi" in c_no_acc or "co quan" in c_no_acc or "to chuc" in c_no_acc) and unit_idx == -1: unit_idx = idx
                    elif "chuc vu" in c_no_acc and pos_idx == -1: pos_idx = idx

                if name_idx != -1 and cccd_idx != -1:
                    for r in table.rows[1:]:
                        vals = [clean_text(cell.text) for cell in r.cells]
                        if len(vals) > max(name_idx, cccd_idx):
                            name_val = vals[name_idx]
                            cccd_val = vals[cccd_idx]
                            if name_val and cccd_val and "ho va ten" not in remove_accents(name_val):
                                dob_val     = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                                phone_val   = clean_text(vals[phone_idx]) if phone_idx != -1 and phone_idx < len(vals) else ""
                                cccd_fmt    = format_cccd_output(cccd_val)
                                issue_date  = normalize_date(vals[issue_date_idx]) if issue_date_idx != -1 and issue_date_idx < len(vals) else ""
                                issue_place = normalize_issue_place(vals[issue_place_idx]) if issue_place_idx != -1 and issue_place_idx < len(vals) else "CCSQLHCVTTXH"
                                
                                kn = norm_name(name_val)
                                kc = norm_cccd(cccd_val)
                                sso_info = sso_map.get((kn, kc)) or sso_map.get(kc) or sso_map.get(kn) or {}
                                email = sso_info.get("email", "")
                                sso_unit = sso_info.get("unit", "")
                                sso_pos = sso_info.get("position", "")

                                unit_val    = clean_text(vals[unit_idx]) if unit_idx != -1 and unit_idx < len(vals) else (sso_unit if sso_unit else DEFAULT_ORG3)
                                pos_val     = clean_text(vals[pos_idx]) if pos_idx != -1 and pos_idx < len(vals) else (sso_pos if sso_pos else "Nhân viên")

                                cccd_str = f"{cccd_fmt};{issue_date};{issue_place}".rstrip(";")

                                records.append([name_val, dob_val, cccd_str, email, unit_val, DEFAULT_PROVINCE, pos_val, phone_val, "", ""])
                                status = f"Email: {email}" if email else "CHƯA CÓ EMAIL SSO"
                                log_func(f"  + Lấy được: {name_val} - CCCD: {cccd_fmt} -> {status}")
        except Exception as e:
            log_func(f"Lỗi đọc file {filename}: {e}")

    output_dir = os.path.dirname(word_files[0]) if word_files else os.getcwd()
    
    # Chỉ xuất duy nhất 1 file Công văn Word Mẫu 02
    doc_result = generate_mau_02_doc(output_dir, records, log_func)

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ XUẤT THÀNH CÔNG FILE CÔNG VÂN MẪU 02: {os.path.basename(doc_result)}")
    log_func(f"Đường dẫn lưu file: {doc_result}")
    log_func(f"Tổng số bản ghi: {len(records)}")
    log_func(f"==========================================")
    return doc_result

# ============================================================
# LOGIC TÁC VỤ 3: CKS (MẪU CV02 CHỈ ĐỌC FILE WORD)
# ============================================================
def run_task_cks_word_only(word_input, log_func=print, progress_func=lambda v: None, output_file=None):
    log_func("=== BẮT ĐẦU TÁC VỤ 3: CKS (MẪU CV02 CHỈ ĐỌC FILE WORD) ===")

    word_files = get_file_list(word_input, extensions=(".docx", ".doc"))
    if not word_files:
        raise ValueError("Không tìm thấy file Word (.docx, .doc) nào để xử lý! Vui lòng chọn file Word.")

    total_files = len(word_files)
    records = []

    for i, target_file in enumerate(word_files, start=1):
        filename = os.path.basename(target_file)
        log_func(f"\nĐang đọc file Word ({i}/{total_files}): {filename}")
        progress_func(int((i / max(1, total_files)) * 80))

        actual_file = target_file
        if target_file.lower().endswith(".doc"):
            converted_docx = os.path.splitext(target_file)[0] + "_converted.docx"
            if not os.path.exists(converted_docx):
                try:
                    import win32com.client
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    w_doc = word.Documents.Open(os.path.abspath(target_file))
                    w_doc.SaveAs2(os.path.abspath(converted_docx), FileFormat=16)
                    w_doc.Close()
                    word.Quit()
                    actual_file = converted_docx
                    log_func(f"  -> Đã chuyển đổi file .doc sang .docx: {os.path.basename(actual_file)}")
                except Exception as e:
                    log_func(f"  [CẢNH BÁO] Không thể chuyển đổi tự động file .doc bằng win32com: {e}")
            else:
                actual_file = converted_docx

        try:
            doc = docx.Document(actual_file)
        except Exception as e:
            log_func(f"  ❌ Không thể mở file {filename}: {e}")
            continue

        file_records = []
        for t_idx, table in enumerate(doc.tables):
            if not table.rows: continue
            header_row_idx = -1
            name_idx, dob_idx, phone_idx, cccd_idx, issue_date_idx, issue_place_idx, unit_idx, pos_idx = [-1]*8
            serial_idx, dev_idx = -1, -1

            for r_idx, row in enumerate(table.rows[:min(4, len(table.rows))]):
                header_texts = [clean_text(cell.text) for cell in row.cells]
                t_name, t_dob, t_phone, t_cccd, t_idate, t_iplace, t_unit, t_pos = [-1]*8
                t_ser, t_dev = -1, -1

                for idx, col in enumerate(header_texts):
                    c_norm = remove_accents(col)
                    is_combo = (';' in col or ';' in c_norm) and any(k in c_norm for k in ['cccd', 'cmnd', 'cmt', 'can cuoc'])
                    if ("ho va ten" in c_norm or "ho ten" in c_norm) and "nguoi lap" not in c_norm and "truong" not in c_norm and "ky va" not in c_norm:
                        if t_name == -1: t_name = idx
                    elif is_combo:
                        if t_cccd == -1: t_cccd = idx
                    elif "ngay cap" in c_norm and "noi" not in c_norm:
                        if t_idate == -1: t_idate = idx
                    elif "noi cap" in c_norm:
                        if t_iplace == -1: t_iplace = idx
                    elif any(k in c_norm for k in ['cccd', 'cmt', 'cmnd', 'can cuoc']) and "chup anh" not in c_norm and "anh" not in c_norm and "ngay cap" not in c_norm and "noi cap" not in c_norm:
                        if t_cccd == -1: t_cccd = idx
                    elif "ngay sinh" in c_norm or "nam sinh" in c_norm or "dob" in c_norm:
                        if t_dob == -1: t_dob = idx
                    elif any(k in c_norm for k in ['sdt', 'dien thoai', 'di dong', 'phone']):
                        if t_phone == -1: t_phone = idx
                    elif any(k in c_norm for k in ['don vi', 'co quan', 'to chuc', 'khoa', 'phong']) and "tinh" not in c_norm:
                        if t_unit == -1: t_unit = idx
                    elif "chuc vu" in c_norm or "vi tri" in c_norm:
                        if t_pos == -1: t_pos = idx
                    elif "serial" in c_norm or "so hieu" in c_norm:
                        if t_ser == -1: t_ser = idx
                    elif "thiet bi" in c_norm or "sim" in c_norm:
                        if t_dev == -1: t_dev = idx

                if t_name != -1 and t_cccd != -1:
                    header_row_idx = r_idx
                    name_idx, dob_idx, phone_idx, cccd_idx = t_name, t_dob, t_phone, t_cccd
                    issue_date_idx, issue_place_idx, unit_idx, pos_idx = t_idate, t_iplace, t_unit, t_pos
                    serial_idx, dev_idx = t_ser, t_dev
                    break

            if header_row_idx == -1:
                continue

            for r in table.rows[header_row_idx + 1:]:
                vals = [clean_text(cell.text) for cell in r.cells]
                if len(vals) <= max(name_idx, cccd_idx): continue
                name_val = vals[name_idx]
                cccd_raw = vals[cccd_idx]
                if not name_val or not cccd_raw or "ho va ten" in remove_accents(name_val): continue

                dob_val = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                phone_val = clean_text(vals[phone_idx]) if phone_idx != -1 and phone_idx < len(vals) else ""
                issue_date = normalize_date(vals[issue_date_idx]) if issue_date_idx != -1 and issue_date_idx < len(vals) else ""
                issue_place = normalize_issue_place(vals[issue_place_idx]) if issue_place_idx != -1 and issue_place_idx < len(vals) else ""

                cccd_clean = cccd_raw
                if ";" in cccd_raw:
                    parts = [clean_text(p) for p in cccd_raw.split(";")]
                    cccd_clean = parts[0]
                    if not issue_date and len(parts) > 1 and parts[1]:
                        issue_date = normalize_date(parts[1])
                    if (not issue_place or issue_place == "CCSQLHCVTTXH") and len(parts) > 2 and parts[2]:
                        issue_place = normalize_issue_place(parts[2])

                if not issue_place:
                    issue_place = "CCSQLHCVTTXH"

                cccd_fmt = format_cccd_output(cccd_clean)
                cccd_str = f"{cccd_fmt};{issue_date};{issue_place}".rstrip(";")

                unit_val = clean_text(vals[unit_idx]) if unit_idx != -1 and unit_idx < len(vals) else DEFAULT_ORG3
                pos_val = clean_text(vals[pos_idx]) if pos_idx != -1 and pos_idx < len(vals) else "Nhân viên"
                ser_val = clean_text(vals[serial_idx]) if serial_idx != -1 and serial_idx < len(vals) else ""
                dev_val = clean_text(vals[dev_idx]) if dev_idx != -1 and dev_idx < len(vals) else ""

                email_val = ""
                for c_val in vals:
                    if "@" in c_val and "." in c_val:
                        m_mail = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", c_val)
                        if m_mail:
                            email_val = clean_email(m_mail.group(0))
                            break
                if not email_val:
                    kn = norm_name(name_val)
                    if kn:
                        email_val = f"{kn}@bacninh.gov.vn"

                file_records.append([name_val, dob_val, cccd_str, email_val, unit_val, DEFAULT_PROVINCE, pos_val, phone_val, ser_val, dev_val])
                log_func(f"  + Lấy được: {name_val} - CCCD: {cccd_fmt} -> Email: {email_val}")

        if not file_records:
            full_text = "\n".join(p.text for p in doc.paragraphs)
            m_name = re.search(r"(?:Họ và tên|Họ tên|Tên tôi là)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)
            m_cccd = re.search(r"(?:CCCD|Số CCCD|CMND|Số CMND)\s*[:：]\s*([0-9]{9,12})", full_text, re.IGNORECASE)
            m_dob = re.search(r"(?:Ngày sinh|Sinh ngày)\s*[:：]\s*([0-9\/\-\.]+)", full_text, re.IGNORECASE)
            m_mail = re.search(r"(?:Email|Thư điện tử|Hòm thư)\s*[:：]?\s*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", full_text, re.IGNORECASE)
            m_phone = re.search(r"(?:Điện thoại|SĐT|Số điện thoại|Di động)\s*[:：]\s*([0-9\.\s]+)", full_text, re.IGNORECASE)
            m_pos = re.search(r"(?:Chức vụ)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)
            m_unit = re.search(r"(?:Đơn vị|Khoa\/Phòng|Cơ quan)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)
            m_idate = re.search(r"(?:Ngày cấp)\s*[:：]\s*([0-9\/\-\.]+)", full_text, re.IGNORECASE)
            m_iplace = re.search(r"(?:Nơi cấp)\s*[:：]\s*([^\n\r,]+)", full_text, re.IGNORECASE)

            if m_name and m_cccd:
                name_v = clean_text(m_name.group(1))
                cccd_v = clean_text(m_cccd.group(1))
                dob_v = normalize_date(m_dob.group(1)) if m_dob else ""
                email_v = clean_email(m_mail.group(1)) if m_mail else ""
                phone_v = clean_text(m_phone.group(1)) if m_phone else ""
                pos_v = clean_text(m_pos.group(1)) if m_pos else "Nhân viên"
                unit_v = clean_text(m_unit.group(1)) if m_unit else DEFAULT_ORG3
                idate_v = normalize_date(m_idate.group(1)) if m_idate else ""
                iplace_v = normalize_issue_place(m_iplace.group(1)) if m_iplace else "CCSQLHCVTTXH"
                if not email_v:
                    kn = norm_name(name_v)
                    if kn:
                        email_v = f"{kn}@bacninh.gov.vn"

                cccd_fmt = format_cccd_output(cccd_v)
                cccd_str = f"{cccd_fmt};{idate_v};{iplace_v}".rstrip(";")
                file_records.append([name_v, dob_v, cccd_str, email_v, unit_v, DEFAULT_PROVINCE, pos_v, phone_v, "", ""])
                log_func(f"  + Lấy được (từ văn bản): {name_v} - CCCD: {cccd_fmt} -> Email: {email_v}")

        records.extend(file_records)

    output_dir = os.path.dirname(word_files[0]) if word_files else os.getcwd()
    doc_result = generate_mau_02_doc(output_dir, records, log_func)

    if doc_result and output_file and os.path.abspath(doc_result) != os.path.abspath(output_file):
        try:
            shutil.copy2(doc_result, output_file)
            doc_result = output_file
        except Exception as copy_err:
            log_func(f"  [CẢNH BÁO] Không thể copy sang {output_file}: {copy_err}")

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ XUẤT THÀNH CÔNG FILE CÔNG VÂN MẪU 02: {os.path.basename(doc_result) if doc_result else 'None'}")
    log_func(f"Đường dẫn lưu file: {doc_result}")
    log_func(f"Tổng số bản ghi: {len(records)}")
    log_func(f"==========================================")
    return doc_result

# ============================================================
# LOGIC TÁC VỤ 9: TỔNG HỢP XIN CẤP EMAIL CÔNG VỤ (Xuất Công văn Đề nghị)
# ============================================================
def run_task_email(word_input, log_func=print, progress_func=lambda v: None):
    log_func("=== BẮT ĐẦU TÁC VỤ 9: TỔNG HỢP XIN CẤP EMAIL CÔNG VỤ (XUẤT CÔNG VÂN ĐỀ NGHỊ) ===")
    word_files = get_file_list(word_input, extensions=(".docx", ".doc"))
    if not word_files:
        raise ValueError("Không tìm thấy file Word (.docx, .doc) nào để xử lý! Vui lòng chọn file Word.")
    
    total_files = len(word_files)
    records = []
    
    for i, f in enumerate(word_files, start=1):
        filename = os.path.basename(f)
        log_func(f"\nĐang đọc file Word ({i}/{total_files}): {filename}")
        progress_func(int((i / max(1, total_files)) * 80))
        try:
            doc = docx.Document(f)
            for table in doc.tables:
                if not table.rows: continue
                header = [clean_text(cell.text) for cell in table.rows[0].cells]
                name_idx, dob_idx, phone_idx, cccd_idx, unit_idx, pos_idx, note_idx = [-1]*7
                
                for idx, col in enumerate(header):
                    c_no_acc = remove_accents(col)
                    if ("ho va ten" in c_no_acc or "ho ten" in c_no_acc) and "nguoi lap" not in c_no_acc: name_idx = idx
                    elif "ngay sinh" in c_no_acc and dob_idx == -1: dob_idx = idx
                    elif ("sdt" in c_no_acc or "dien thoai" in c_no_acc or "di dong" in c_no_acc) and phone_idx == -1: phone_idx = idx
                    elif ("cccd" in c_no_acc or "cmt" in c_no_acc or "cmnd" in c_no_acc or "can cuoc" in c_no_acc) and "noi cap" not in c_no_acc and "ngay cap" not in c_no_acc and "chup anh" not in c_no_acc and "anh" not in c_no_acc:
                        if cccd_idx == -1: cccd_idx = idx
                    elif ("don vi" in c_no_acc or "co quan" in c_no_acc or "to chuc" in c_no_acc) and unit_idx == -1: unit_idx = idx
                    elif "chuc vu" in c_no_acc and pos_idx == -1: pos_idx = idx
                    elif "ghi chu" in c_no_acc and note_idx == -1: note_idx = idx

                if name_idx != -1 and cccd_idx != -1:
                    for r in table.rows[1:]:
                        vals = [clean_text(cell.text) for cell in r.cells]
                        if len(vals) > max(name_idx, cccd_idx):
                            name_val = vals[name_idx]
                            cccd_val = vals[cccd_idx]
                            if name_val and cccd_val and "họ và tên" not in name_val.lower():
                                dob_val = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                                phone_val = clean_text(vals[phone_idx]) if phone_idx != -1 and phone_idx < len(vals) else ""
                                cccd_val_fmt = format_cccd_output(cccd_val)
                                unit_val = clean_text(vals[unit_idx]) if unit_idx != -1 and unit_idx < len(vals) else DEFAULT_ORG3
                                pos_val = clean_text(vals[pos_idx]) if pos_idx != -1 and pos_idx < len(vals) else "Nhân viên"
                                note_val = clean_text(vals[note_idx]) if note_idx != -1 and note_idx < len(vals) else "Cấp mới"
                                if not note_val: note_val = "Cấp mới"
                                
                                records.append([name_val, dob_val, phone_val, cccd_val_fmt, unit_val, pos_val, note_val])
                                log_func(f"  + Lấy được Email: {name_val} - CCCD: {cccd_val_fmt} - Đơn vị: {unit_val}")
        except Exception as e:
            log_func(f"Lỗi đọc file {filename}: {e}")

    output_dir = os.path.dirname(word_files[0]) if word_files else os.getcwd()
    
    output_dir = os.path.dirname(word_files[0]) if word_files else os.getcwd()
    
    # Chỉ xuất duy nhất 1 file Công văn Word Đề nghị cấp mới địa chỉ email công vụ
    doc_result = generate_email_doc(output_dir, records, log_func)

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ XUẤT THÀNH CÔNG FILE CÔNG VÂN EMAIL: {os.path.basename(doc_result)}")
    log_func(f"Đường dẫn lưu file: {doc_result}")
    log_func(f"Tổng số bản ghi: {len(records)}")
    log_func(f"==========================================")
    return doc_result

def get_excel_sheet_names_fast(file_path):
    """Đọc danh sách Sheet siêu tốc trong 0.003 giây bằng cách đọc trực tiếp zipfile XML"""
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            with z.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                sheets = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
                names = [elem.attrib['name'] for elem in sheets if 'name' in elem.attrib]
                if names:
                    return names
    except Exception:
        pass
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []

# ============================================================
# LOGIC TÁC VỤ 4 & 6: BÁO CÁO TỔNG HỢP CÔNG VIỆC P.CNTT (CHUẨN FORMAT 08.09.2026)
# ============================================================
STANDARD_HEADERS_18 = [
    "STT",
    "KHOA, PHÒNG, TRUNG TÂM",
    "Xoá CLS, PTTT",
    "Sửa phí",
    "Huỷ CKS",
    "Huỷ B/án",
    "Sửa thông tin BN (NN; Địa chỉ, tuổi, CCCD,.v.v...)\nSửa công khám, sửa kiểu khám",
    "Sửa ngày giờ cls; pttt; giờ đọc kết quả; người TH\ncập nhật CLS, PTTT",
    "Sửa mã bệnh kèm theo; bệnh chính",
    "Hỗ trợ ký số PM, QLVBĐH",
    "Sửa giấy RV, sửa kết thúc ra viện, chuyển tuyến",
    "Đăng bài trên website",
    "Sửa cách dùng thuốc, ngày giờ thuốc",
    "Phân quyền USER, đổi nick user",
    "Lỗi không kết thúc được hồ sơ",
    "Sửa xoá giường; cập nhật giường, thêm giường",
    "Chuyển thuốc CP- DG",
    "SỬA LỖI KHÁC"
]

# Column widths matching EXACT sample file "08.09.2026 TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx"
# None = do not set (use Excel default ~8.43 chars), matching columns E,F,H,L-Q in sample
FIXED_COL_WIDTHS_18 = [
    (1,  5.0),          # A: STT
    (2,  25.0),         # B: KHOA, PHÒNG, TRUNG TÂM
    (3,  11.0),         # C: Xoá CLS, PTTT
    (4,  10.0),         # D: Sửa phí
    # E (5): Huỷ CKS  → DEFAULT (not set in sample)
    # F (6): Huỷ B/án → DEFAULT (not set in sample)
    (7,  18.140625),    # G: Sửa thông tin BN...
    # H (8): Sửa ngày giờ cls → DEFAULT (not set in sample)
    (9,  13.28515625),  # I: Sửa mã bệnh kèm theo...
    (10, 14.5703125),   # J: Hỗ trợ ký số PM...
    (11, 12.0),         # K: Sửa giấy RV...
    # L (12): Đăng bài trên website → DEFAULT
    # M (13): Sửa cách dùng thuốc  → DEFAULT
    # N (14): Phân quyền USER       → DEFAULT
    # O (15): Lỗi không kết thúc    → DEFAULT
    # P (16): Sửa xoá giường        → DEFAULT
    # Q (17): Chuyển thuốc CP-DG    → DEFAULT
    (18, 42.7109375),   # R: SỬA LỖI KHÁC
]

def map_software_header_to_index(header_name):
    norm_h = norm_name(header_name)
    if not norm_h: return None
    target_15 = STANDARD_HEADERS_18[2:17]
    for idx, th in enumerate(target_15):
        if norm_name(th) == norm_h: return idx
    if "xoacls" in norm_h: return 0
    if "suaphi" in norm_h or "epthe" in norm_h: return 1
    if "huycks" in norm_h: return 2
    if "huyban" in norm_h: return 3
    if "suathongtinbn" in norm_h or "suacongkham" in norm_h: return 4
    if "suangaygiocls" in norm_h or "suamaycls" in norm_h: return 5
    if "suamabenh" in norm_h: return 6
    if "kyso" in norm_h or "qlvbdh" in norm_h: return 7
    if "suagiayrv" in norm_h or "chuyentuyen" in norm_h: return 8
    if "dangbai" in norm_h or "website" in norm_h: return 9
    if "suacachdungthuoc" in norm_h or "ngaygiothuoc" in norm_h: return 10
    if "phanquyenuser" in norm_h or "doinick" in norm_h: return 11
    if "loikhongketthuc" in norm_h: return 12
    if "suaxoagiuong" in norm_h or "themgiuong" in norm_h: return 13
    if "chuyenthuoccp" in norm_h or "cpdg" in norm_h: return 14
    return None

def build_formatted_cntt_report(valid_departments, data_matrix, other_errors, output_file):
    from openpyxl.styles import PatternFill
    result_wb = Workbook()
    ws = result_wb.active
    ws.title = "BÁO CÁO TỔNG HỢP"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    # Dark gray header fill (matches sample: theme 0 with tint -0.35 ≈ #595959)
    header_fill = PatternFill(fill_type="solid", fgColor="595959")

    # 1. Row 1: Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=18)
    t_cell = ws.cell(1, 1, "TỔNG HỢP CÔNG VIỆC P.CNTT")
    t_cell.font = Font(name="Times New Roman", size=18, bold=True)
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42.0

    # 2. Row 3: Headers (with dark gray background like sample)
    header_row = 3
    data_start_row = 4

    for col_idx, h_text in enumerate(STANDARD_HEADERS_18, 1):
        cell = ws.cell(header_row, col_idx, h_text)
        cell.font = Font(name="Times New Roman", size=14, bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 112.5

    # 3. Data Rows
    for idx, dept in enumerate(valid_departments, 1):
        r_idx = data_start_row + idx - 1
        
        # Col 1 (A): STT
        c_stt = ws.cell(r_idx, 1, idx)
        c_stt.font = Font(name="Times New Roman", size=16)
        c_stt.border = border
        c_stt.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Col 2 (B): Khoa/Phòng
        c_dept = ws.cell(r_idx, 2, dept)
        c_dept.font = Font(name="Times New Roman", size=16)
        c_dept.border = border
        c_dept.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Col 3..17 (C..Q): 15 software error categories
        counts = data_matrix[dept]
        for c_offset in range(15):
            col = 3 + c_offset
            v = counts[c_offset]
            c_val = ws.cell(r_idx, col, int(v) if (v not in (None, "", 0)) else None)
            c_val.font = Font(name="Times New Roman", size=16)
            c_val.border = border
            c_val.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Col 18 (R): Sửa lỗi khác
        err_list = other_errors.get(dept, [])
        err_text = ", ".join(err_list) if err_list else None
        c_err = ws.cell(r_idx, 18, err_text)
        c_err.font = Font(name="Times New Roman", size=16)
        c_err.border = border
        c_err.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        line_count = 1
        if err_text:
            line_count = max(len(err_text.split('\n')), (len(err_text) // 35) + 1)
        ws.row_dimensions[r_idx].height = max(35.1, line_count * 26)

    # 4. Summary Row (TỔNG CỘNG)
    total_row = data_start_row + len(valid_departments)

    c_tot_stt = ws.cell(total_row, 1, None)
    c_tot_stt.border = border
    c_tot_stt.alignment = Alignment(horizontal="center", vertical="center")

    c_tot_label = ws.cell(total_row, 2, "TỔNG CỘNG")
    c_tot_label.font = Font(name="Times New Roman", size=14, bold=True)
    c_tot_label.border = border
    c_tot_label.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    category_sums = [0] * 15
    for dept in valid_departments:
        counts = data_matrix[dept]
        for i in range(15):
            val = counts[i]
            if isinstance(val, (int, float)):
                category_sums[i] += int(val)

    for c_offset in range(15):
        col = 3 + c_offset
        cell = ws.cell(total_row, col)
        cell.font = Font(name="Times New Roman", size=14, bold=True)
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        if c_offset == 9: # Column L: Đăng bài trên website (Excluded from summary)
            cell.value = None
        else:
            s_val = category_sums[c_offset]
            cell.value = s_val if s_val > 0 else None

    # Col 18 (R): Grand Total sum (excl col L)
    grand_total = sum(category_sums[i] for i in range(15) if i != 9)
    c_tot_r = ws.cell(total_row, 18, grand_total if grand_total > 0 else None)
    c_tot_r.font = Font(name="Times New Roman", size=14, bold=True)
    c_tot_r.border = border
    c_tot_r.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.row_dimensions[total_row].height = 35.1

    # 5. Set column widths – only set explicitly where sample defines them; others use Excel default
    for col_idx_or_letter, width in FIXED_COL_WIDTHS_18:
        col_letter = get_column_letter(col_idx_or_letter)
        ws.column_dimensions[col_letter].width = width

    # 6. Page Print setup – mirror exact settings from sample file:
    #    scale=46%, landscape, fitToPage=True (but NOT fitToWidth/fitToHeight = let Excel handle)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.page_setup.scale = 46  # 46% zoom – exact value from sample
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    ws.freeze_panes = "C4"

    try:
        result_wb.save(output_file)
    except PermissionError:
        output_file = os.path.join(os.path.dirname(output_file), f"{os.path.splitext(os.path.basename(output_file))[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        result_wb.save(output_file)

    return output_file


def run_task_cntt(excel_input, log_func=print, progress_func=lambda v: None):
    log_func("=== BẮT ĐẦU TÁC VỤ 4: BÁO CÁO TỔNG HỢP CÔNG VIỆC P.CNTT ===")
    
    if os.path.isdir(excel_input):
        folder = excel_input
        excel_files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$') and "TỔNG HỢP CÔNG VIỆC" not in f
        ]
        if not excel_files:
            raise ValueError(f"Không tìm thấy file Excel nào trong thư mục:\n{folder}")
        
        target_file = None
        for ef in excel_files:
            if "BÁO CÁO CÔNG TÁC SỬA CHỮA" in os.path.basename(ef).upper():
                target_file = ef
                break
        if not target_file:
            target_file = excel_files[0]
        actual_input_file = target_file
    else:
        actual_input_file = excel_input
        folder = os.path.dirname(excel_input)

    log_func(f"Quét nhanh danh sách Sheet từ file Excel: {os.path.basename(actual_input_file)}")
    progress_func(10)

    sheetnames = get_excel_sheet_names_fast(actual_input_file)
    if not sheetnames:
        raise ValueError(f"Không đọc được danh sách Sheet nào từ file: {actual_input_file}")

    if len(sheetnames) == 1:
        sheet_name = sheetnames[0]
    else:
        default_idx = len(sheetnames) - 1
        for idx in range(len(sheetnames) - 1, -1, -1):
            s_n = sheetnames[idx].strip()
            if not s_n.lower().startswith("sheet"):
                default_idx = idx
                break
        sheet_selected = {"val": sheetnames[default_idx]}
        if tk._default_root:
            evt = threading.Event()
            def ask_sheet_gui():
                sw = tk.Toplevel()
                sw.title("Chọn Sheet P.CNTT")
                sw.geometry("460x180")
                sw.resizable(False, False)
                ttk.Label(sw, text=f"CHỌN SHEET CẦN TỔNG HỢP ({len(sheetnames)} Sheets):", font=("Segoe UI", 10, "bold")).pack(pady=12)
                cb = ttk.Combobox(sw, values=sheetnames, state="readonly", width=42)
                cb.pack(pady=5)
                cb.current(default_idx if default_idx < len(sheetnames) else 0)
                def on_ok():
                    sheet_selected["val"] = cb.get()
                    sw.destroy()
                    evt.set()
                ttk.Button(sw, text="Đồng Ý & Tiếp Tục", command=on_ok, width=18).pack(pady=12)
                sw.grab_set()
                sw.protocol("WM_DELETE_WINDOW", lambda: (sw.destroy(), evt.set()))
            tk._default_root.after(0, ask_sheet_gui)
            evt.wait()
        sheet_name = sheet_selected["val"]

    log_func(f"  -> Đã chọn Sheet: '{sheet_name}'. Đang nạp và tổng hợp dữ liệu từ nhật ký cá nhân (dòng 7-96)...")
    progress_func(35)

    # Đọc KHÔNG data_only để lấy giá trị thực từ ô nhập tay (không bị ảnh hưởng bởi cache công thức)
    wb_source = load_workbook(actual_input_file, data_only=False)
    ws_source = wb_source[sheet_name]

    # Đọc header hàng 5 để map cột phần mềm → index trong STANDARD_HEADERS_18
    software_columns = list(range(5, 26))  # cột E–Y (index 5–25)
    sw_col_target_map = {}
    for col in software_columns:
        h = ws_source.cell(5, col).value or ws_source.cell(6, col).value
        sw_col_target_map[col] = map_software_header_to_index(clean_text(h))

    # Đọc danh sách khoa/phòng có thực tế nhập liệu (cột B từ dòng 7 đến dòng 96 hoặc dòng trống)
    departments_order = []
    data_matrix = {}
    other_errors = {}

    max_scan = min(ws_source.max_row, 200)  # an toàn, quét tối đa 200 dòng
    for row in range(7, max_scan + 1):
        dept_val = ws_source.cell(row, 2).value  # Cột B = KHOA/PHÒNG
        dept_str = clean_text(dept_val)
        if not dept_str:
            # Kiểm tra nếu cột A cũng trống → dòng trống hoàn toàn, dừng lại
            stt_val = ws_source.cell(row, 1).value
            if stt_val is None:
                continue  # có thể là dòng giữa chừng không có KHOA → bỏ qua
            continue
        if dept_str.upper() in ("KHOA, PHÒNG, TRUNG TÂM", "TỔNG CỘNG", "STT"):
            continue

        if dept_str not in data_matrix:
            departments_order.append(dept_str)
            data_matrix[dept_str] = [0] * 15
            other_errors[dept_str] = []

        # Đếm các loại phần mềm từ cột E–Y
        for col in software_columns:
            val = ws_source.cell(row, col).value
            if val in (None, "", 0):
                continue
            if isinstance(val, str):
                val_str = val.strip()
                if not val_str or val_str == "0":
                    continue
                try:
                    val_num = int(float(val_str))
                except (ValueError, TypeError):
                    val_num = 1
            elif isinstance(val, (int, float)):
                val_num = int(val)
                if val_num <= 0:
                    continue
            else:
                val_num = 1

            t_idx = sw_col_target_map[col]
            if t_idx is not None:
                data_matrix[dept_str][t_idx] += val_num
            else:
                col_h = clean_text(ws_source.cell(5, col).value or ws_source.cell(6, col).value or "")
                if col_h:
                    other_errors[dept_str].append(f"{col_h}: {val_num}")



# ============================================================
# HÀM HỖ TRỢ TÁC VỤ 6: TẢI DỮ LIỆU TỪ GOOGLE SHEETS
# ============================================================
def download_google_sheet(url, log_func=print):
    """
    Tải toàn bộ Google Sheets Workbook về dưới dạng file xlsx tạm thời.
    Luôn tải toàn bộ workbook (đủ tất cả sheet) để hộp thoại chọn sheet
    có thể hiển thị đầy đủ — giống như khi đọc file Excel cục bộ.
    Yêu cầu: Google Sheet phải được chia sẻ 'Đối tượng có được link có thể xem'.
    """
    # Trích xuất Spreadsheet ID từ URL
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9\-_]+)', url)
    if not match:
        raise ValueError(
            "URL không hợp lệ. Vui lòng dán đúng link Google Sheets.\n"
            "Ví dụ: https://docs.google.com/spreadsheets/d/1BxiM.../edit"
        )
    spreadsheet_id = match.group(1)

    # Luôn export TOÀN BỘ Workbook (không lọc theo gid)
    # Mục đích: để dialog chọn sheet hiển thị đầy đủ tất cả các sheet ngày
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    )

    log_func(f"  → Spreadsheet ID: {spreadsheet_id}")
    log_func(f"  → Đang tải toàn bộ workbook (tất cả sheet)...")

    try:
        req = urllib.request.Request(
            export_url,
            headers={"User-Agent": "Mozilla/5.0"}
        )
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
    except Exception as e:
        raise ConnectionError(
            f"Không tải được dữ liệu từ Google Sheets.\n"
            f"Hãy kiểm tra:\n"
            f"  1. File Google Sheets có được chia sẻ không?\n"
            f"     (Chia sẻ → Chọn 'Đối tượng có được link có thể xem')\n"
            f"  2. Máy tính có kết nối Internet không?\n"
            f"Chi tiết lỗi: {e}"
        )

    size_kb = len(content) // 1024
    log_func(f"  → Tải xuống thành công ({size_kb} KB)")

    # Lưu vào file tạm, tự xóa sau khi xử lý xong
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, prefix="gsheet_cntt_")
    tmp.write(content)
    tmp.close()
    return tmp.name


def run_task_cntt_tu_dong(excel_input, log_func=print, progress_func=lambda v: None, output_folder=None, task_label="Tác vụ 6"):
    log_func(f"=== BẮT ĐẦU {task_label.upper()}: TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG) ===")

    tmp_file = None  # file tạm khi đọc từ Google Sheets

    # ━━━ Xác định nguồn dữ liệu: Google Sheets URL hay file/thư mục cục bộ ━━━
    is_gsheet = isinstance(excel_input, str) and excel_input.strip().startswith("https://docs.google.com/spreadsheets")

    if is_gsheet:
        log_func(f"📶 Nguồn dữ liệu: Google Sheets")
        log_func(f"  URL: {excel_input.strip()[:80]}{'...' if len(excel_input) > 80 else ''}")
        progress_func(5)
        tmp_file = download_google_sheet(excel_input.strip(), log_func)
        actual_input_file = tmp_file
        # Thư mục xuất = output_folder hoặc thư mục hiện tại
        folder = output_folder if output_folder and os.path.isdir(output_folder) else os.getcwd()
        log_func(f"  → Thư mục lưu kết quả: {folder}")
        progress_func(25)
        source_label = "Google Sheets"
    elif os.path.isdir(excel_input):
        folder = excel_input
        excel_files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$') and "TỔNG HỢP CÔNG VIỆC" not in f
        ]
        if not excel_files:
            raise ValueError(f"Không tìm thấy file Excel nào trong thư mục:\n{folder}")
        target_file = None
        for ef in excel_files:
            if "BÁO CÁO CÔNG TÁC SỬA CHỮA" in os.path.basename(ef).upper():
                target_file = ef
                break
        if not target_file:
            target_file = excel_files[0]
        actual_input_file = target_file
        source_label = os.path.basename(actual_input_file)
    else:
        actual_input_file = excel_input
        folder = os.path.dirname(excel_input) or os.getcwd()
        source_label = os.path.basename(actual_input_file)
    if output_folder and os.path.isdir(output_folder):
        folder = output_folder

    log_func(f"Quét danh sách Sheet từ: {source_label}")
    progress_func(30)

    sheetnames = get_excel_sheet_names_fast(actual_input_file)
    if not sheetnames:
        raise ValueError(f"Không đọc được danh sách Sheet nào từ file: {actual_input_file}")

    log_func(f"  → Tìm thấy {len(sheetnames)} sheet: {', '.join(sheetnames)}")

    if len(sheetnames) == 1:
        sheet_name = sheetnames[0]
    else:
        default_idx = len(sheetnames) - 1
        for idx in range(len(sheetnames) - 1, -1, -1):
            s_n = sheetnames[idx].strip()
            if not s_n.lower().startswith("sheet"):
                default_idx = idx
                break
        sheet_selected = {"val": sheetnames[default_idx]}
        if tk._default_root:
            evt = threading.Event()
            def ask_sheet_gui():
                sw = tk.Toplevel()
                sw.title(f"Chọn Sheet P.CNTT ({task_label})")
                sw.geometry("460x180")
                sw.resizable(False, False)
                ttk.Label(sw, text=f"CHỌN SHEET CẦN TỔNG HỢP ({len(sheetnames)} Sheets):", font=("Segoe UI", 10, "bold")).pack(pady=12)
                cb = ttk.Combobox(sw, values=sheetnames, state="readonly", width=42)
                cb.pack(pady=5)
                cb.current(default_idx if default_idx < len(sheetnames) else 0)
                def on_ok():
                    sheet_selected["val"] = cb.get()
                    sw.destroy()
                    evt.set()
                ttk.Button(sw, text="Đồng Ý & Tiếp Tục", command=on_ok, width=18).pack(pady=12)
                sw.grab_set()
                sw.protocol("WM_DELETE_WINDOW", lambda: (sw.destroy(), evt.set()))
            tk._default_root.after(0, ask_sheet_gui)
            evt.wait()
        sheet_name = sheet_selected["val"]

    log_func(f"  -> Đã chọn Sheet: '{sheet_name}'. Đang tổng hợp dữ liệu (dòng 7 trở đi)...")
    progress_func(45)

    # Google Sheets xuất xlsx đã có giá trị tính sẵn → dùng data_only=True
    # File Excel cục bộ có thể bị cache 0 → dùng data_only=False để đọc formula rồi thiếu lấy giá trị trực tiếp
    wb_source = load_workbook(actual_input_file, data_only=True)
    ws_source = wb_source[sheet_name]

    software_columns = list(range(5, 26))  # Cột E–Y
    sw_col_target_map = {}
    for col in software_columns:
        h = ws_source.cell(5, col).value or ws_source.cell(6, col).value
        sw_col_target_map[col] = map_software_header_to_index(clean_text(h))

    departments_order = []
    data_matrix = {}
    other_errors = {}

    row = 7
    max_scan_row = min(ws_source.max_row, 200)
    consecutive_empty = 0
    while row <= max_scan_row:
        dept_val = ws_source.cell(row, 2).value  # Cột B = Khoa/Phòng
        dept_str = clean_text(dept_val)

        # Bỏ qua dòng tiêu đề hoặc tổng cộng
        if dept_str.upper() in ("KHOA, PHÒNG, TRUNG TÂM", "TỔNG CỘNG", "STT"):
            row += 1
            continue

        # Nếu cột B trống, kiểm tra xem dòng này có STT không
        if not dept_str:
            stt_val = ws_source.cell(row, 1).value
            if stt_val is None:
                consecutive_empty += 1
                if consecutive_empty >= 5:
                    log_func(f"  -> Gặp {consecutive_empty} dòng trống liên tiếp tại dòng {row}. Dừng quét.")
                    break
            row += 1
            continue

        consecutive_empty = 0

        if dept_str not in data_matrix:
            departments_order.append(dept_str)
            data_matrix[dept_str] = [0] * 15
            other_errors[dept_str] = []

        # Đếm loại phần mềm từ cột E–Y
        for col in software_columns:
            val = ws_source.cell(row, col).value
            if val in (None, ""):
                continue
            if isinstance(val, str):
                val_str = val.strip()
                if not val_str or val_str == "0":
                    continue
                try:
                    val_num = int(float(val_str))
                    if val_num <= 0:
                        continue
                except (ValueError, TypeError):
                    val_num = 1
            elif isinstance(val, (int, float)):
                val_num = int(val)
                if val_num <= 0:
                    continue
            else:
                val_num = 1

            t_idx = sw_col_target_map[col]
            if t_idx is not None:
                data_matrix[dept_str][t_idx] += val_num
            else:
                col_h = clean_text(ws_source.cell(5, col).value or ws_source.cell(6, col).value or "")
                if col_h:
                    other_errors[dept_str].append(f"{col_h}: {val_num}")

        # Cột AI (35) = Sửa chữa khác
        err_val = ws_source.cell(row, 35).value
        err_str = clean_text(err_val)
        if err_str and err_str not in ("0", ""):
            if err_str not in other_errors[dept_str]:
                other_errors[dept_str].append(err_str)

        row += 1

    valid_departments = [
        dept for dept in departments_order
        if any(v != 0 for v in data_matrix[dept]) or bool(other_errors[dept])
    ]

    log_func(f"  -> Tổng hợp xong. Số khoa có dữ liệu: {len(valid_departments)}")
    progress_func(75)

    output_file = os.path.join(folder, "TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx")
    saved_file = build_formatted_cntt_report(valid_departments, data_matrix, other_errors, output_file)

    try:
        os.unlink(tmp_file)
    except Exception:
        pass

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ XUẤT THÀNH CÔNG FILE P.CNTT TỰ ĐỘNG: {os.path.basename(saved_file)}")
    log_func(f"Đường dẫn lưu file: {saved_file}")
    log_func(f"Tổng số khoa có dữ liệu: {len(valid_departments)}")
    log_func(f"==========================================")
    return saved_file


# ============================================================
# CẤU HÌNH & HÀM XỬ LÝ TÁC VỤ 8: GOOGLE SHEETS CỐ ĐỊNH P.CNTT
# ============================================================
TASK8_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139"

def run_task_8(log_func=print, progress_func=lambda v: None, output_folder=None):
    return run_task_cntt_tu_dong(
        TASK8_GSHEET_URL,
        log_func=log_func,
        progress_func=progress_func,
        output_folder=output_folder,
        task_label="Tác vụ 8"
    )


# ============================================================
# LOGIC TÁC VỤ 5: BÁO CÁO CỔNG GIÁM ĐỊNH BẢO HIỂM (Xuất Excel A4 Portrait Fit 1 Page)
# ============================================================
ALLOWED_USERS = {
    "PHCN", "NGOẠI TH", "XNTT", "NHH", "TTTM", "NGOẠI CT", "DA LIỄU", "PTGMHS",
    "NỘI TH", "PHỤ SẢN", "MẮT", "TTBVSK", "DƯỢC", "THẬN", "TRUYỀN NHIỄM", "TCKT",
    "NGOẠI UB", "NGOẠI XẠ TRỊ", "RHM", "KB", "CXK", "CSGN", "NGOẠI TN", "HHLS",
    "TMH", "HSTC", "HTL", "NGOẠI TKLN", "NHI", "ĐÔNG Y", "CẤP CỨU", "LKTK"
}

def copy_style(source, target):
    if source.has_style:
        target._style = copy.copy(source._style)
    target.font = copy.copy(source.font)
    target.fill = copy.copy(source.fill)
    target.border = copy.copy(source.border)
    target.alignment = copy.copy(source.alignment)
    target.protection = copy.copy(source.protection)
    target.number_format = source.number_format

def center_cell(cell, wrap=False):
    alignment = copy.copy(cell.alignment)
    alignment.horizontal = "center"
    alignment.vertical = "center"
    alignment.wrap_text = wrap
    cell.alignment = alignment

def find_header_row(ws):
    for row in range(1, min(ws.max_row, 20) + 1):
        values = [remove_accents(clean_text(ws.cell(row, col).value)).upper() for col in range(1, ws.max_column + 1)]
        if "STT" in values and "NGUOI DUNG" in values:
            return row
    raise ValueError("Không tìm thấy dòng tiêu đề STT / NGƯỜI DÙNG.")

def find_columns(ws, header_row):
    user_col = None
    day_columns = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(header_row, col).value
        text = remove_accents(clean_text(val)).upper()
        if text == "NGUOI DUNG":
            user_col = col
        if val is not None:
            val_str = str(val).strip()
            if val_str.isdigit():
                day = int(val_str)
                if 1 <= day <= 31:
                    day_columns[day] = col
    if user_col is None:
        raise ValueError("Không tìm thấy cột NGƯỜI DÙNG.")
    if not day_columns:
        raise ValueError("Không tìm thấy các cột ngày (1-31).")
    return user_col, day_columns

def run_task_giam_dinh(excel_input, day_range_choice="7days", log_func=print, progress_func=lambda v: None, output_file=None):
    log_func("=== BẮT ĐẦU TÁC VỤ 5: BÁO CÁO CỔNG GIÁM ĐỊNH BẢO HIỂM ===")

    if os.path.isdir(excel_input):
        folder = excel_input
        excel_files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$') and "DANH SÁCH THEO DÕI GIÁM ĐỊNH" not in f
        ]
        if not excel_files:
            raise ValueError(f"Không tìm thấy file Excel nào trong thư mục:\n{folder}")
        target_file = None
        for ef in excel_files:
            if "GIÁM ĐỊNH" in os.path.basename(ef).upper():
                target_file = ef
                break
        if not target_file:
            target_file = excel_files[0]
        actual_input_file = target_file
    else:
        actual_input_file = excel_input
        folder = os.path.dirname(excel_input)

    log_func(f"Đang đọc file Excel: {os.path.basename(actual_input_file)}")
    progress_func(15)

    source_wb = load_workbook(actual_input_file)
    target_sheet = None
    for s_name in source_wb.sheetnames:
        norm = remove_accents(s_name).lower().strip()
        if "nguoi dung" in norm or norm in ("user", "users") or "cham cong" in norm:
            target_sheet = source_wb[s_name]
            break
    source = target_sheet if target_sheet else source_wb.active

    header_row = find_header_row(source)
    user_col, day_columns = find_columns(source, header_row)

    # Xác định ngày cuối có dữ liệu thực tế (cột ngày tồn tại trong file)
    available_days = sorted(day_columns.keys())  # danh sách ngày thực có trong sheet
    last_available_day = max(available_days) if available_days else 31

    # Lựa chọn khoảng ngày:
    choice = str(day_range_choice).strip().lower()
    today_day = datetime.now().day

    if "01-14" in choice or "1-14" in choice:
        # Lựa chọn: Ngày 01 → 14
        start_day, end_day = 1, 14
        range_label = "01 → 14"
    elif "full" in choice or "1-end" in choice or "01-end" in choice:
        # Lựa chọn: Full từ ngày 01 → ngày cuối có dữ liệu
        start_day = 1
        end_day = last_available_day
        range_label = f"01 → {end_day:02d} (FULL tháng)"
    elif "15-end" in choice or "15-31" in choice:
        # Lựa chọn: Ngày 15 → ngày cuối có dữ liệu
        start_day = 15
        end_day = last_available_day
        range_label = f"15 → {end_day:02d} (cuối có dữ liệu)"
    else:
        # Lựa chọn mặc định (7days): Trước và sau 7 ngày so với ngày hiện tại
        start_day = max(1, today_day - 7)
        end_day = min(last_available_day, today_day + 7)
        if start_day > end_day:
            start_day = max(1, last_available_day - 14)
            end_day = last_available_day
            range_label = f"{start_day:02d} → {end_day:02d} (Dữ liệu file kết thúc tại ngày {last_available_day:02d})"
        else:
            range_label = f"{start_day:02d} → {end_day:02d} (Trước & sau 7 ngày so với ngày hiện tại: {today_day:02d})"

    log_func(f"Khoảng ngày được chọn: {range_label}")
    log_func(f"Ngày cuối thực tế trong file: {last_available_day}")
    progress_func(25)

    # Lọc các ngày hợp lệ trong khoảng [start_day, end_day] mà CÓ trong file
    selected_days = [d for d in range(start_day, end_day + 1) if d in day_columns]
    if not selected_days:
        raise ValueError(f"Không có cột ngày nào từ {start_day:02d} đến {end_day:02d} trong sheet 'Người dùng'.")

    log_func(f"Các ngày được tổng hợp: {selected_days[0]:02d} đến {selected_days[-1]:02d} ({len(selected_days)} ngày)")

    source_columns = [1, 2] + [day_columns[d] for d in selected_days]
    result_last_col = len(source_columns)
    result_last_letter = get_column_letter(result_last_col)

    result_wb = Workbook()
    result = result_wb.active
    result.title = "Người dùng"

    for row in range(1, 6):
        for new_col, old_col in enumerate(source_columns, start=1):
            s_cell = source.cell(row, old_col)
            t_cell = result.cell(row, new_col, s_cell.value)
            copy_style(s_cell, t_cell)

    title = source["A1"].value
    for col in range(2, result_last_col + 1): result.cell(1, col).value = None
    result["A1"] = title
    result.merge_cells(start_row=1, start_column=1, end_row=1, end_column=result_last_col)
    center_cell(result["A1"], wrap=False)
    result.row_dimensions[1].height = 32

    note = source["A2"].value
    for col in range(2, result_last_col + 1): result.cell(2, col).value = None
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
        user_name = clean_text(source.cell(source_row, user_col).value).upper()
        if user_name not in ALLOWED_USERS: continue

        for new_col, old_col in enumerate(source_columns, start=1):
            s_cell = source.cell(source_row, old_col)
            t_cell = result.cell(output_row, new_col)
            t_cell.value = stt if new_col == 1 else s_cell.value
            copy_style(s_cell, t_cell)
        stt += 1
        output_row += 1

    result.page_setup.orientation = result.ORIENTATION_PORTRAIT
    result.page_setup.paperSize = result.PAPERSIZE_A4
    result.sheet_properties.pageSetUpPr.fitToPage = True
    result.page_setup.fitToWidth = 1
    result.page_setup.fitToHeight = 0
    result.page_margins.left = 0.25
    result.page_margins.right = 0.25
    result.page_margins.top = 0.4
    result.page_margins.bottom = 0.4

    for col in range(1, result_last_col + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(4, result.max_row + 1):
            val = result.cell(row, col).value
            if val is not None:
                for line in str(val).split('\n'):
                    max_len = max(max_len, len(line.strip()))
        if col == 1: width = max(max_len + 3, 6)
        elif col == 2: width = max(max_len + 4, 20)
        else: width = max(max_len + 2, 7)
        result.column_dimensions[col_letter].width = width

    for row in range(1, result.max_row + 1):
        for col in range(1, result_last_col + 1):
            cell = result.cell(row, col)
            is_bold = True if row in (1, 4, 5) or row == result.max_row else (cell.font.bold if cell.font else False)
            is_italic = True if row == 2 else (cell.font.italic if cell.font else False)
            cell.font = Font(name="Times New Roman", size=14, bold=is_bold, italic=is_italic)

    result.freeze_panes = "C6"
    result.auto_filter.ref = f"A4:{result_last_letter}{result.max_row}"

    if not output_file:
        output_file = os.path.join(folder, "DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx")
    target_dir = os.path.dirname(output_file) or folder
    target_base = os.path.splitext(os.path.basename(output_file))[0]
    try:
        result_wb.save(output_file)
    except PermissionError:
        output_file = os.path.join(target_dir, f"{target_base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        result_wb.save(output_file)

    progress_func(100)
    log_func(f"\n==========================================")
    log_func(f"ĐÃ XUẤT THÀNH CÔNG FILE GIÁM ĐỊNH BẢO HIỂM: {os.path.basename(output_file)}")
    log_func(f"Đường dẫn lưu file: {output_file}")
    log_func(f"Tổng số người dùng: {stt - 1}")
    log_func(f"==========================================")
    return output_file

# ============================================================
# GIAO DIỆN GUI TKINTER HỢP NHẤT
# ============================================================
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PHẦN MỀM TỔNG HỢP CẤP CKS, EMAIL, P.CNTT & CỔNG GIÁM ĐỊNH BẢO HIỂM")
        self.geometry("1240x840")
        self.minsize(1080, 720)
        
        style = ttk.Style(self)
        style.theme_use('clam')
        
        # 1. TOP HEADER BANNER
        header_frame = tk.Frame(self, bg="#0f172a", pady=10, padx=16)
        header_frame.pack(fill=tk.X)
        
        top_row = tk.Frame(header_frame, bg="#0f172a")
        top_row.pack(fill=tk.X)
        
        lbl_brand = tk.Label(
            top_row,
            text="🏥 BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2",
            font=("Segoe UI", 9, "bold"),
            fg="#60a5fa", bg="#0f172a"
        )
        lbl_brand.pack(side=tk.LEFT)
        
        lbl_status = tk.Label(
            top_row,
            text="🟢 Hệ thống sẵn sàng",
            font=("Segoe UI", 8),
            fg="#34d399", bg="#0f172a"
        )
        lbl_status.pack(side=tk.RIGHT)
        
        lbl_title = tk.Label(
            header_frame, 
            text="HỆ THỐNG TỔNG HỢP CẤP CKS, EMAIL, P.CNTT & GIÁM ĐỊNH BẢO HIỂM", 
            font=("Segoe UI", 13, "bold"), 
            fg="white", bg="#0f172a"
        )
        lbl_title.pack(anchor="w", pady=(3, 0))

        # 2. MAIN BODY (SIDEBAR BÊN TRÁI + KHUNG NỘI DUNG BÊN PHẢI)
        body_frame = tk.Frame(self, bg="#f1f5f9")
        body_frame.pack(fill=tk.BOTH, expand=True)

        # --- SIDEBAR (BÊN TRÁI) ---
        sidebar = tk.Frame(body_frame, bg="#0f172a", width=310)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)

        # Tiêu đề sidebar
        sb_title_box = tk.Frame(sidebar, bg="#1e293b", padx=14, pady=10)
        sb_title_box.pack(fill=tk.X, pady=(0, 6))
        tk.Label(
            sb_title_box,
            text="📑 DANH MỤC TÁC VỤ (8 CHỨC NĂNG)",
            font=("Segoe UI", 9, "bold"),
            fg="#93c5fd", bg="#1e293b"
        ).pack(anchor="w")
        tk.Label(
            sb_title_box,
            text="Nhấp chọn tác vụ bên dưới để thực thi",
            font=("Segoe UI", 8),
            fg="#94a3b8", bg="#1e293b"
        ).pack(anchor="w")

        # Container chứa các nút tác vụ
        sb_items_frame = tk.Frame(sidebar, bg="#0f172a")
        sb_items_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        # Footer sidebar
        sb_footer = tk.Frame(sidebar, bg="#0f172a", padx=12, pady=10)
        sb_footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(
            sb_footer,
            text="Phòng CNTT • BVĐK Bắc Ninh số 2\nPhiên bản tự động hóa v2.6",
            font=("Segoe UI", 8),
            fg="#64748b", bg="#0f172a", justify=tk.LEFT
        ).pack(anchor="w")

        # --- KHUNG NỘI DUNG BÊN PHẢI ---
        content_area = tk.Frame(body_frame, bg="#f8fafc")
        content_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Header thanh breadcrumb tác vụ đang chọn
        content_header = tk.Frame(content_area, bg="#ffffff", padx=16, pady=8, highlightthickness=1, highlightbackground="#e2e8f0")
        content_header.pack(fill=tk.X)
        self.lbl_task_badge = tk.Label(
            content_header,
            text="TÁC VỤ ĐANG CHỌN",
            font=("Segoe UI", 8, "bold"),
            bg="#dbeafe", fg="#1d4ed8",
            padx=8, pady=3
        )
        self.lbl_task_badge.pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_task_title = tk.Label(
            content_header,
            text="",
            font=("Segoe UI", 11, "bold"),
            fg="#0f172a", bg="#ffffff"
        )
        self.lbl_task_title.pack(side=tk.LEFT)

        # Khung chứa các trang nội dung
        self.page_container = tk.Frame(content_area, bg="#f8fafc")
        self.page_container.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self.pages = {}
        self.sidebar_buttons = {}
        self.current_task = 1

        self.tasks_meta = [
            (1, "📌", "1. VGCA (Word + Excel SSO)", "Đối chiếu cấp CKS Word & Excel SSO"),
            (2, "📝", "2. CKS (Mẫu 02)", "Xuất Công văn Word Mẫu 02 (Có SSO)"),
            (3, "📝", "3. CKS (Mẫu CV02 chỉ đọc file Word)", "Đọc toàn bộ file Word xuất Mẫu 02"),
            (4, "📊", "4. Báo Cáo P.CNTT", "Tổng hợp sửa chữa A4 ngang dòng 7"),
            (5, "🏥", "5. Giám Định BHYT", "Báo cáo cổng giám định A4 dọc"),
            (6, "🤖", "6. P.CNTT Tự Động", "Quét từ link Google Sheets bất kỳ"),
            (7, "📄", "7. VGCA (Từ File Word)", "Đọc nhiều file Word không cần SSO"),
            (8, "⚡", "8. P.CNTT (Sheet Cố Định)", "Tự động tải Google Sheets cố định"),
            (9, "📧", "9. Cấp Email Công Vụ", "Tổng hợp danh sách xin cấp Email"),
        ]

        # Khởi tạo 8 trang nội dung và 8 mục chọn ở Sidebar
        for task_id, icon, title, desc in self.tasks_meta:
            page = tk.Frame(self.page_container, bg="#f8fafc")
            self.pages[task_id] = page
            self.build_tab(page, task_type=task_id)
            self.create_sidebar_item(sb_items_frame, task_id, icon, title, desc)

        # Mặc định kích hoạt Tác vụ 1
        self.switch_task(1)

    def create_sidebar_item(self, parent, task_id, icon, title, desc):
        item_frame = tk.Frame(parent, bg="#1e293b", cursor="hand2", padx=6, pady=6)
        item_frame.pack(fill=tk.X, pady=3)

        # Vạch chỉ thị màu xanh khi active
        ind_bar = tk.Frame(item_frame, bg="#1e293b", width=4)
        ind_bar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        # Icon tác vụ
        lbl_icon = tk.Label(item_frame, text=icon, font=("Segoe UI Emoji", 12), bg="#1e293b", fg="#ffffff")
        lbl_icon.pack(side=tk.LEFT, padx=(0, 6))

        # Tiêu đề & mô tả
        text_box = tk.Frame(item_frame, bg="#1e293b")
        text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        lbl_t = tk.Label(text_box, text=title, font=("Segoe UI", 9, "bold"), fg="#e2e8f0", bg="#1e293b", anchor="w")
        lbl_t.pack(fill=tk.X)

        lbl_d = tk.Label(text_box, text=desc, font=("Segoe UI", 7), fg="#94a3b8", bg="#1e293b", anchor="w")
        lbl_d.pack(fill=tk.X)

        widgets = (item_frame, ind_bar, lbl_icon, text_box, lbl_t, lbl_d)

        def on_click(e=None):
            self.switch_task(task_id)

        for w in widgets:
            w.bind("<Button-1>", on_click)

        def on_enter(e=None):
            if self.current_task != task_id:
                for w in (item_frame, lbl_icon, text_box, lbl_t, lbl_d):
                    w.config(bg="#334155")

        def on_leave(e=None):
            if self.current_task != task_id:
                for w in (item_frame, lbl_icon, text_box, lbl_t, lbl_d):
                    w.config(bg="#1e293b")

        item_frame.bind("<Enter>", on_enter)
        item_frame.bind("<Leave>", on_leave)

        self.sidebar_buttons[task_id] = {
            "frame": item_frame,
            "bar": ind_bar,
            "icon": lbl_icon,
            "box": text_box,
            "title": lbl_t,
            "desc": lbl_d
        }

    def switch_task(self, task_id):
        self.current_task = task_id

        # Cập nhật giao diện Sidebar
        for tid, widgets in self.sidebar_buttons.items():
            if tid == task_id:
                widgets["frame"].config(bg="#2563eb")
                widgets["bar"].config(bg="#93c5fd")
                widgets["icon"].config(bg="#2563eb", fg="#ffffff")
                widgets["box"].config(bg="#2563eb")
                widgets["title"].config(bg="#2563eb", fg="#ffffff")
                widgets["desc"].config(bg="#2563eb", fg="#dbeafe")
            else:
                widgets["frame"].config(bg="#1e293b")
                widgets["bar"].config(bg="#1e293b")
                widgets["icon"].config(bg="#1e293b", fg="#ffffff")
                widgets["box"].config(bg="#1e293b")
                widgets["title"].config(bg="#1e293b", fg="#cbd5e1")
                widgets["desc"].config(bg="#1e293b", fg="#94a3b8")

        # Cập nhật tiêu đề Breadcrumb
        for tid, icon, title, desc in self.tasks_meta:
            if tid == task_id:
                self.lbl_task_title.config(text=f"{icon} {title}  ─  {desc}")
                break

        # Ẩn các trang khác, chỉ hiển thị trang được chọn
        for tid, page in self.pages.items():
            if tid == task_id:
                page.pack(fill=tk.BOTH, expand=True)
            else:
                page.pack_forget()

    def build_tab(self, tab, task_type):
        main_paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        guide_frame = ttk.LabelFrame(main_paned, text=" 📖 Hướng Dẫn Sử Dụng ", padding=10)
        main_paned.add(guide_frame, weight=1)

        guide_text = scrolledtext.ScrolledText(guide_frame, wrap=tk.WORD, font=("Segoe UI", 10), bg="#f8fafc", width=36)
        guide_text.pack(fill=tk.BOTH, expand=True)

        if task_type == 1:
            guide_content = """🎯 TÁC VỤ 1: ĐỐI CHIẾU WORD & EXCEL SSO (VGCA)

📌 Mục đích:
Tự động đọc thông tin từ các đơn đăng ký Word (.docx), ghép đôi với file tài khoản SSO Excel để lấy Email công vụ, tạo file Ket_qua.xlsx đúng 19 cột chuẩn VGCA.

📝 Các bước thực hiện:
1️⃣ Bước 1: Chọn các file Word (.docx) đề nghị cấp CKS (hoặc chọn thư mục chứa file Word).
2️⃣ Bước 2: Chọn file Excel SSO (.xlsx) chứa danh sách email công vụ.
3️⃣ Bước 3: Bấm nút '🚀 BẮT ĐẦU XỬ LÝ ĐỐI CHIẾU (VGCA)'.
4️⃣ Bước 4: Kiểm tra kết quả xuất ra file Ket_qua.xlsx.
"""
        elif task_type == 2:
            guide_content = """🎯 TÁC VỤ 2: TỔNG HỢP XIN CẤP CKS (XUẤT CÔNG VÂN MẪU 02)

📌 Mục đích:
Đọc các file Word (.docx) đề nghị cấp CKS, ghép với file Excel SSO để xuất ra file Công văn 'Mẫu 02_CV cap CKS cho BVDK BN số 2_T08 2026_3_6' giữ nguyên 100% thể thức văn bản & cỡ chữ.

📝 Các bước thực hiện:
1️⃣ Bước 1: Chọn các file Word (.docx) mẫu đơn CKS.
2️⃣ Bước 2: Chọn file Excel SSO (.xlsx) chứa tài khoản SSO.
3️⃣ Bước 3: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP CKS (MẪU 02)'.
4️⃣ Bước 4: Kiểm tra file Công văn Word 'Mẫu 02' được tự động tạo.
"""
        elif task_type == 3:
            guide_content = """🎯 TÁC VỤ 3: CKS (MẪU CV02 CHỈ ĐỌC FILE WORD)

📌 Mục đích:
Tổng hợp danh sách xin cấp Chữ ký số (CKS) xuất file Công văn Mẫu 02 (.docx) chỉ từ file Word đề nghị, độc lập hoàn toàn, KHÔNG CẦN file Excel SSO. Đọc toàn bộ dữ liệu ở file Word rồi điền thông tin để ra kết quả như Tác vụ 2.

📝 Các bước thực hiện:
1️⃣ Bước 1: Chọn các file Word (.docx, .doc) mẫu đơn CKS hoặc thư mục chứa file Word.
2️⃣ Bước 2: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP CKS (MẪU CV02 CHỈ ĐỌC FILE WORD)'.
3️⃣ Bước 3: Chọn vị trí lưu và nhận file Công văn Word 'Mẫu 02' hoàn chỉnh.
"""
        elif task_type == 9:
            guide_content = """🎯 TÁC VỤ 9: TỔNG HỢP DANH SÁCH XIN CẤP EMAIL CÔNG VỤ

📌 Mục đích:
Quét các file Word (.docx) để tổng hợp danh sách đề nghị cấp mới tài khoản Thư điện tử công vụ ra file Công văn 'Đề nghị cấp mới địa chỉ email công vụ'.

📝 Các bước thực hiện:
1️⃣ Bước 1: Chọn các file Word (.docx) xin cấp Email công vụ.
2️⃣ Bước 2: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP EMAIL'.
3️⃣ Bước 3: Kiểm tra file Công văn Word 'Đề nghị cấp mới địa chỉ email công vụ'.
"""
        elif task_type == 4:
            guide_content = """🎯 TÁC VỤ 4: BÁO CÁO TỔNG HỢP CÔNG VIỆC P.CNTT (EXCEL A4 NGANG)

📌 Mục đích:
Đọc file Excel báo cáo công tác sửa chữa hàng ngày của P.CNTT, trích xuất dữ liệu sửa chữa phần mềm & sửa lỗi khác theo từng Khoa/Phòng để xuất ra file 'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx' chuẩn khổ A4 nằm ngang, tự động giãn chiều cao dòng để căn kín chiều dọc khổ A4.

📝 Các bước thực hiện:
1️⃣ Bước 1: Chọn file Excel báo cáo P.CNTT (hoặc thư mục chứa file Excel).
2️⃣ Bước 2: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP P.CNTT'.
3️⃣ Bước 3: Chọn Sheet cần xử lý.
4️⃣ Bước 4: Kiểm tra file Excel 'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx'.
"""
        elif task_type == 5:
            guide_content = """🎯 TÁC VỤ 5: BÁO CÁO CỔNG GIÁM ĐỊNH BẢO HIỂM (EXCEL A4 DỌC)

📌 Mục đích:
Đọc file Excel báo cáo cổng giám định, lọc danh sách người dùng được phép, xuất ra file 'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx' chuẩn khổ A4 dọc, tự động autofit cột và co giãn vừa 1 trang ngang.

📝 Các bước thực hiện:
1️⃣ Bước 1: Chọn file Excel báo cáo Cổng Giám Định (.xlsx).
2️⃣ Bước 2: Chọn khoảng ngày cần tổng hợp (mặc định: Trước và sau 7 ngày so với ngày hiện tại):
   • [7days]  – Lấy trước và sau 7 ngày so với ngày hiện tại (Hôm nay ± 7 ngày)
   • [01-14]  – Lấy từ ngày 01 đến ngày 14 (nửa đầu tháng)
   • [15-end] – Lấy từ ngày 15 đến ngày cuối có dữ liệu (tự động phát hiện)
   • [full]   – Lấy toàn bộ từ ngày 01 đến ngày cuối có dữ liệu (cả tháng)
3️⃣ Bước 3: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP GIÁM ĐỊNH BẢO HIỂM'.
4️⃣ Bước 4: Kiểm tra file Excel 'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx'.

💡 Hệ thống tự động tính ngày trước và sau 7 ngày theo ngày hiện tại và phát hiện ngày có dữ liệu trong file!
"""
        elif task_type == 6:
            guide_content = """🎯 TÁC VỤ 6: TỔNG HỢP CÔNG VIỆC P.CNTT TỰ ĐỘNG (TỪ GOOGLE SHEETS)

📌 Mục đích:
Tự động đọc dữ liệu nhật ký sửa chữa từ Google Sheets (chia sẻ công khai), tổng hợp số lượt sửa chữa phần mềm theo từng Khoa/Phòng rồi xuất ra file 'TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx'.

📝 Các bước thực hiện:
1️⃣ Bước 1: Mở Google Sheets → Nút Chia sẻ → Chọn 'Đối tượng có được link có thể xem'.
2️⃣ Bước 2: Sao chép toàn bộ link Google Sheets, dán vào ô 'Link Google Sheets'.
3️⃣ Bước 3: Chọn thư mục lưu file kết quả (mặc định: thư mục chương trình).
4️⃣ Bước 4: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP' → Chọn Sheet cần xử lý.
5️⃣ Bước 5: Kiểm tra file 'TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx'.

⚠️ Lưu ý:
  • Google Sheets BẮT BUỘC chia sẻ 'Đối tượng có được link'
  • Máy tính phải có kết nối Internet
  • Hỗ trợ link có gid (chọn tab cụ thể)
"""
        elif task_type == 7:
            guide_content = """🎯 TÁC VỤ 7: TỔNG HỢP KẾT QUẢ VGCA TỪ CÁC FILE WORD (Ket_qua.xlsx)

📌 Mục đích:
Tương tự Tác vụ 1 (xuất file Ket_qua.xlsx đúng 19 cột chuẩn VGCA), nhưng KHÔNG CẦN file Excel SSO. Đọc trực tiếp từ một hoặc NHIỀU file Word (.docx, .doc) hoặc Thư mục chứa các file Word đã có đầy đủ thông tin (Họ tên, Ngày sinh, CCCD, Ngày cấp, Nơi cấp, Email công vụ, Đơn vị, Chức vụ, SĐT...).

📝 Các bước thực hiện:
1️⃣ Bước 1: Bấm '📄 Chọn các File Word' để chọn nhiều file Word (.docx / .doc), hoặc bấm '📁 Chọn Thư Mục' để quét toàn bộ file Word trong thư mục.
2️⃣ Bước 2: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP VGCA TỪ CÁC FILE WORD'.
3️⃣ Bước 3: Kiểm tra kết quả xuất ra file 'Ket_qua.xlsx' với 19 cột chuẩn VGCA tổng hợp từ tất cả các file.

💡 Ưu điểm nổi bật:
  • Hỗ trợ chọn nhiều file Word cùng lúc hoặc quét cả thư mục — tự động gộp tất cả bản ghi vào 1 file kết quả.
  • Độc lập hoàn toàn, không cần nạp thêm file Excel SSO.
  • Tự động tách thông tin 'Số CCCD;Ngày cấp;Nơi cấp' nếu gộp chung ô (chuẩn Công văn Mẫu 02) hoặc các cột bảng riêng lẻ.
  • Tự động nhận diện Email và chuẩn hóa toàn bộ 19 cột theo quy chuẩn Import VGCA.
"""
        else:
            guide_content = """🎯 TÁC VỤ 8: TỔNG HỢP CÔNG VIỆC P.CNTT (GOOGLE SHEETS CỐ ĐỊNH)

📌 Mục đích:
Tự động lấy dữ liệu từ đường link Google Sheets cố định của P.CNTT:
https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139

Thực hiện các công việc tương tự Tác vụ 6:
  • Tải toàn bộ bảng tính với tất cả các Sheet ngày.
  • Cho phép chọn Sheet ngày cần tổng hợp từ danh sách.
  • Quét dữ liệu từ dòng 7 trở đi (Khoa/Phòng tại cột B, phân loại lỗi phần mềm tại cột E–Y, sửa chữa khác tại cột AI).
  • Xuất file 'TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx' chuẩn A4 ngang, căn chỉnh fit all columns và tự động kéo giãn chiều cao dòng.

📝 Các bước thực hiện:
1️⃣ Bước 1: Đường link Google Sheets đã được tích hợp sẵn trong hệ thống (không cần nhập thủ công).
2️⃣ Bước 2: Chọn thư mục lưu file kết quả (mặc định: thư mục hiện tại của chương trình).
3️⃣ Bước 3: Bấm nút '🚀 BẮT ĐẦU TỔNG HỢP P.CNTT (GOOGLE SHEETS CỐ ĐỊNH)' → Chọn Sheet ngày cần tổng hợp.
4️⃣ Bước 4: Kiểm tra file Excel kết quả được tạo ra.

💡 Tiện ích nổi bật:
  • Hoàn toàn tự động, một chạm là chạy, không cần sao chép hay dán link thủ công.
  • Tự động kết nối và tải đầy đủ tất cả các sheet ngày từ Google Sheets.
"""
        guide_text.insert(tk.END, guide_content)
        guide_text.config(state=tk.DISABLED)

        action_frame = ttk.LabelFrame(main_paned, text=" ⚡ Thao Tác Thực Thi & Chọn Dữ Liệu ", padding=10)
        main_paned.add(action_frame, weight=2)

        select_box = ttk.LabelFrame(action_frame, text=" 📁 Chọn Dữ Liệu Đầu Vào ", padding=8)
        select_box.pack(fill=tk.X, pady=5)

        if task_type == 6:
            lbl_w_title = "1️⃣ Dán Link Google Sheets (phải chia sẻ 'Đối tượng có được link có thể xem'):"
        elif task_type == 8:
            lbl_w_title = "1️⃣ Link Google Sheets Cố Định (P.CNTT):"
        elif task_type == 4: lbl_w_title = "1️⃣ Chọn File Excel báo cáo P.CNTT (.xlsx, .xlsm) hoặc Thư Mục:"
        elif task_type == 5: lbl_w_title = "1️⃣ Chọn File Excel báo cáo Cổng Giám Định (.xlsx) hoặc Thư Mục:"
        elif task_type == 7: lbl_w_title = "1️⃣ Chọn Các File Word (.docx, .doc) hoặc Thư Mục chứa file Word:"
        else: lbl_w_title = "1️⃣ Chọn File(s) Word (.docx, .doc):"
        
        lbl_w = ttk.Label(select_box, text=lbl_w_title, font=("Segoe UI", 9, "bold"))
        lbl_w.pack(anchor=tk.W, padx=5, pady=(2, 0))

        word_frame = ttk.Frame(select_box)
        word_frame.pack(fill=tk.X, pady=2)

        txt_word = ttk.Entry(word_frame, font=("Segoe UI", 9))
        txt_word.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        cwd = os.getcwd()
        if task_type == 8:
            txt_word.insert(0, TASK8_GSHEET_URL)
            txt_word.config(state="readonly")
        elif task_type == 6:
            # Task 6: ô nhập URL Google Sheets – để trống hoặc placeholder
            txt_word.insert(0, "")
            txt_word.config(foreground="gray")
            placeholder = "https://docs.google.com/spreadsheets/d/..."
            txt_word.insert(0, placeholder)
            def on_focus_in(event):
                if txt_word.get() == placeholder:
                    txt_word.delete(0, tk.END)
                    txt_word.config(foreground="black")
            def on_focus_out(event):
                if not txt_word.get().strip():
                    txt_word.insert(0, placeholder)
                    txt_word.config(foreground="gray")
            txt_word.bind("<FocusIn>", on_focus_in)
            txt_word.bind("<FocusOut>", on_focus_out)
        elif task_type == 4:
            default_cntt_files = [f for f in glob.glob(os.path.join(cwd, "*.xlsx")) if "BÁO CÁO CÔNG TÁC SỬA CHỮA" in os.path.basename(f).upper()]
            txt_word.insert(0, default_cntt_files[0] if default_cntt_files else cwd)
        elif task_type == 5:
            default_giamdinh_files = [f for f in glob.glob(os.path.join(cwd, "*.xlsx")) if "GIÁM ĐỊNH" in os.path.basename(f).upper()]
            txt_word.insert(0, default_giamdinh_files[0] if default_giamdinh_files else cwd)
        else:
            default_words = glob.glob(os.path.join(cwd, "*.docx"))
            txt_word.insert(0, cwd if default_words else "")

        def browse_word_files():
            if task_type in (4, 5, 6):
                t_str = "Chọn file Excel báo cáo Giám Định (.xlsx)" if task_type == 5 else "Chọn file Excel báo cáo P.CNTT (.xlsx)"
                selected = filedialog.askopenfilename(
                    title=t_str,
                    filetypes=[("Excel Files", "*.xlsx;*.xlsm"), ("All Files", "*.*")]
                )
                if selected:
                    txt_word.delete(0, tk.END)
                    txt_word.insert(0, selected)
            else:
                title_word = "Chọn các file Word chứa thông tin (.docx, .doc)" if task_type == 7 else "Chọn các file Word đề nghị (.docx)"
                selected = filedialog.askopenfilenames(
                    title=title_word,
                    filetypes=[("Word Documents", "*.docx;*.doc"), ("All Files", "*.*")]
                )
                if selected:
                    txt_word.delete(0, tk.END)
                    txt_word.insert(0, "; ".join(selected))

        def browse_word_dir():
            title_dir = "Chọn thư mục chứa file Excel" if task_type in (4, 5, 6) else "Chọn thư mục chứa file Word"
            selected = filedialog.askdirectory(title=title_dir)
            if selected:
                txt_word.delete(0, tk.END)
                txt_word.insert(0, selected)

        if task_type == 8:
            def copy_fixed_url():
                try:
                    self.clipboard_clear()
                    self.clipboard_append(TASK8_GSHEET_URL)
                    messagebox.showinfo("Thông báo", "Đã sao chép đường link Google Sheets vào bộ nhớ tạm!")
                except Exception:
                    pass
            ttk.Button(word_frame, text="📋 Sao Chép Link", command=copy_fixed_url).pack(side=tk.LEFT, padx=2)
        elif task_type == 6:
            # Task 6: nút dán clipboard thay vì file picker
            def paste_from_clipboard():
                try:
                    clip = self.clipboard_get().strip()
                    txt_word.config(foreground="black")
                    txt_word.delete(0, tk.END)
                    txt_word.insert(0, clip)
                except Exception:
                    pass
            ttk.Button(word_frame, text="📋 Dán từ Clipboard", command=paste_from_clipboard).pack(side=tk.LEFT, padx=2)
        else:
            btn_w_text = "📊 Chọn File Excel" if task_type in (4, 5) else "📄 Chọn các File Word"
            ttk.Button(word_frame, text=btn_w_text, command=browse_word_files).pack(side=tk.LEFT, padx=2)
            ttk.Button(word_frame, text="📁 Chọn Thư Mục", command=browse_word_dir).pack(side=tk.LEFT, padx=2)

        # Task 6 & 8: thư mục lưu kết quả
        txt_output_folder = None
        if task_type in (6, 8):
            lbl_out = ttk.Label(select_box, text="2️⃣ Chọn Thư Mục Lưu File Kết Quả (mặc định: thư mục chương trình):", font=("Segoe UI", 9, "bold"))
            lbl_out.pack(anchor=tk.W, padx=5, pady=(6, 0))
            out_frame = ttk.Frame(select_box)
            out_frame.pack(fill=tk.X, pady=2)
            txt_output_folder = ttk.Entry(out_frame, font=("Segoe UI", 9))
            txt_output_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            txt_output_folder.insert(0, cwd)
            def browse_output_folder():
                sel = filedialog.askdirectory(title="Chọn thư mục lưu file kết quả")
                if sel:
                    txt_output_folder.delete(0, tk.END)
                    txt_output_folder.insert(0, sel)
            ttk.Button(out_frame, text="📂 Chọn Thư Mục", command=browse_output_folder).pack(side=tk.LEFT, padx=2)

        # 2. CHỌN FILE EXCEL SSO HOẶC KHOẢNG NGÀY GIÁM ĐỊNH
        txt_excel = None
        cb_day = None
        if task_type in (1, 2):
            lbl_e = ttk.Label(select_box, text="2️⃣ Chọn File Excel SSO (.xlsx, .xls):", font=("Segoe UI", 9, "bold"))
            lbl_e.pack(anchor=tk.W, padx=5, pady=(6, 0))

            excel_frame = ttk.Frame(select_box)
            excel_frame.pack(fill=tk.X, pady=2)

            txt_excel = ttk.Entry(excel_frame, font=("Segoe UI", 9))
            txt_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

            default_excels = [f for f in glob.glob(os.path.join(cwd, "*.xlsx")) if not os.path.basename(f).startswith("~$") and "Ket_qua" not in f and "TỔNG HỢP CÔNG VIỆC" not in f]
            txt_excel.insert(0, default_excels[0] if default_excels else cwd)

            def browse_excel_file():
                selected = filedialog.askopenfilename(
                    title="Chọn file Excel chứa tài khoản SSO (.xlsx)",
                    filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
                )
                if selected:
                    txt_excel.delete(0, tk.END)
                    txt_excel.insert(0, selected)

            def browse_excel_dir():
                selected = filedialog.askdirectory(title="Chọn thư mục chứa file Excel SSO")
                if selected:
                    txt_excel.delete(0, tk.END)
                    txt_excel.insert(0, selected)

            ttk.Button(excel_frame, text="📊 Chọn File Excel SSO", command=browse_excel_file).pack(side=tk.LEFT, padx=2)
            ttk.Button(excel_frame, text="📁 Chọn Thư Mục", command=browse_excel_dir).pack(side=tk.LEFT, padx=2)

        elif task_type == 5:
            lbl_d = ttk.Label(select_box, text="2️⃣ Chọn Khoảng Ngày Cần Tổng Hợp Giám Định:", font=("Segoe UI", 9, "bold"))
            lbl_d.pack(anchor=tk.W, padx=5, pady=(6, 0))

            day_frame = ttk.Frame(select_box)
            day_frame.pack(fill=tk.X, pady=2)

            today_day = datetime.now().day
            start_7 = max(1, today_day - 7)
            end_7 = min(31, today_day + 7)

            DAY_RANGE_OPTIONS = [
                f"7days   ─  Trước & sau 7 ngày (Ngày {start_7:02d} → {end_7:02d} theo ngày hiện tại: {today_day:02d})",
                "01-14   ─  Từ ngày 01 đến ngày 14",
                "15-end  ─  Từ ngày 15 đến ngày cuối có dữ liệu",
                "full    ─  Từ ngày 01 đến ngày cuối có dữ liệu (Full tháng)",
            ]
            cb_day = ttk.Combobox(
                day_frame,
                values=DAY_RANGE_OPTIONS,
                state="readonly",
                font=("Segoe UI", 9),
                width=55
            )
            cb_day.pack(side=tk.LEFT, padx=5)
            cb_day.current(0)   # Mặc định lựa chọn: Trước và sau 7 ngày so với ngày hiện tại

        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(action_frame, variable=progress_var, maximum=100)
        progress_bar.pack(fill=tk.X, pady=8, padx=5)

        log_box = scrolledtext.ScrolledText(action_frame, height=14, font=("Consolas", 10), bg="#1e293b", fg="#e2e8f0")
        log_box.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)

        def append_log(msg):
            log_box.insert(tk.END, msg + "\n")
            log_box.see(tk.END)

        def update_progress(val):
            progress_var.set(val)
            self.update_idletasks()

        def start_worker():
            if task_type == 8:
                w_val = TASK8_GSHEET_URL
            else:
                w_val = txt_word.get().strip()
            e_val = txt_excel.get().strip() if txt_excel else ""
            d_val = cb_day.get().split()[0] if cb_day else "7days"  # lấy phần code trước dấu ─

            if not w_val:
                err_msg = "Vui lòng nhập link Google Sheets!" if task_type == 6 else (
                    "Vui lòng chọn các file Word hoặc thư mục chứa file Word!" if task_type == 7
                    else ("Vui lòng chọn file Excel hoặc thư mục chứa file Excel!" if task_type in (4, 5)
                    else "Vui lòng chọn các file Word hoặc thư mục chứa file Word!")
                )
                messagebox.showerror("Lỗi", err_msg)
                return
            if task_type == 6 and not w_val.startswith("https://docs.google.com/spreadsheets"):
                messagebox.showerror(
                    "Lỗi URL",
                    "Link Google Sheets không hợp lệ!\n\n"
                    "Vui lòng dán đúng định dạng:\n"
                    "https://docs.google.com/spreadsheets/d/..."
                )
                return
            
            save_dest_file = None
            if task_type in (2, 3):
                init_dir = os.path.dirname(w_val) if os.path.isfile(w_val) else (w_val if os.path.isdir(w_val) else cwd)
                save_dest_file = filedialog.asksaveasfilename(
                    title="Chọn nơi lưu file Công văn Mẫu 02",
                    initialdir=init_dir,
                    initialfile="Mẫu 02_CV cap CKS cho BVDK BN số 2.docx",
                    defaultextension=".docx",
                    filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
                )
                if not save_dest_file:
                    return

            elif task_type == 5:
                init_dir = os.path.dirname(w_val) if os.path.isfile(w_val) else (w_val if os.path.isdir(w_val) else cwd)
                save_dest_file = filedialog.asksaveasfilename(
                    title="Chọn nơi lưu file Giám Định Bảo Hiểm",
                    initialdir=init_dir,
                    initialfile="DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx",
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
                )
                if not save_dest_file:
                    return

            elif task_type == 7:
                init_dir = os.path.dirname(w_val) if os.path.isfile(w_val) else (w_val if os.path.isdir(w_val) else cwd)
                save_dest_file = filedialog.asksaveasfilename(
                    title="Chọn nơi lưu file kết quả VGCA",
                    initialdir=init_dir,
                    initialfile="Ket_qua.xlsx",
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
                )
                if not save_dest_file:
                    return

            btn_run.config(state=tk.DISABLED)
            log_box.delete("1.0", tk.END)
            progress_var.set(0)

            def thread_target():
                try:
                    out_file = None
                    if task_type == 1:
                        out_file = run_task_doi_chieu(w_val, e_val, append_log, update_progress)
                    elif task_type == 2:
                        out_file = run_task_cks(w_val, e_val, append_log, update_progress)
                    elif task_type == 3:
                        out_file = run_task_cks_word_only(w_val, append_log, update_progress, output_file=save_dest_file)
                    elif task_type == 4:
                        out_file = run_task_cntt(w_val, append_log, update_progress)
                    elif task_type == 5:
                        out_file = run_task_giam_dinh(w_val, d_val, append_log, update_progress, output_file=save_dest_file)
                    elif task_type == 6:
                        out_folder_val = txt_output_folder.get().strip() if txt_output_folder else ""
                        out_file = run_task_cntt_tu_dong(w_val, append_log, update_progress, out_folder_val, task_label="Tác vụ 6")
                    elif task_type == 7:
                        out_file = run_task_vgca_word(w_val, append_log, update_progress, output_file=save_dest_file)
                    elif task_type == 8:
                        out_folder_val = txt_output_folder.get().strip() if txt_output_folder else ""
                        out_file = run_task_8(append_log, update_progress, out_folder_val)
                    elif task_type == 9:
                        out_file = run_task_email(w_val, append_log, update_progress)
                    
                    msg = "Tác vụ đã hoàn thành thành công!"
                    if out_file:
                        msg += f"\n\nFile kết quả đã được lưu tại:\n{out_file}"
                    messagebox.showinfo("Thành công", msg)
                except Exception as ex:
                    append_log(f"\n❌ LỖI HỆ THỐNG: {ex}")
                    messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {ex}")
                finally:
                    btn_run.config(state=tk.NORMAL)

            threading.Thread(target=thread_target, daemon=True).start()

        if task_type == 1: btn_text = "🚀 BẮT ĐẦU XỬ LÝ ĐỐI CHIẾU (VGCA)"
        elif task_type == 2: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP CKS (MẪU 02 - CÓ SSO)"
        elif task_type == 3: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP CKS (MẪU CV02 CHỈ ĐỌC FILE WORD)"
        elif task_type == 4: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP P.CNTT (EXCEL A4 NGANG)"
        elif task_type == 5: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP GIÁM ĐỊNH BẢO HIỂM (EXCEL A4 DỌC)"
        elif task_type == 6: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP P.CNTT TỰ ĐỘNG (DÒNG 7)"
        elif task_type == 7: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP VGCA TỪ CÁC FILE WORD (XUẤT Ket_qua.xlsx)"
        elif task_type == 8: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP P.CNTT (GOOGLE SHEETS CỐ ĐỊNH)"
        else: btn_text = "🚀 BẮT ĐẦU TỔNG HỢP EMAIL"

        btn_run = tk.Button(
            action_frame, text=btn_text, font=("Segoe UI", 11, "bold"), 
            bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white",
            pady=8, command=start_worker
        )
        btn_run.pack(fill=tk.X, pady=8, padx=5)

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
