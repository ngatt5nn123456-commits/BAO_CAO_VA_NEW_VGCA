import io
import os
import re
import csv
import copy
import json
import time
import zipfile
import warnings
import unicodedata
import urllib.request
import urllib.parse
import http.client
from datetime import datetime

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

import docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from flask import Flask, request, send_file, jsonify, render_template_string, make_response

warnings.filterwarnings("ignore", message="Workbook contains no default style")

# ============================================================
# CẤU HÌNH & CHUẨN HÓA DÙNG CHUNG
# ============================================================
DEFAULT_ORG1 = "ỦY BAN NHÂN DÂN THÀNH PHỐ BẮC NINH"
DEFAULT_ORG2 = "SỞ Y TẾ"
DEFAULT_ORG3 = "BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2"
DEFAULT_PROVINCE = "Bắc Ninh"

DEFAULT_FIXED_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139"
DEFAULT_ERROR_TRACKER_URL = "https://docs.google.com/spreadsheets/d/1VJ3_I4ehRDLSbm5w0CHAODtS8Y4fN3b-wsvo4p4ANVI/edit?gid=637670018#gid=637670018"
_DM_KHOA_CACHE = {}
_TRACKER_SHEETS_CACHE = {}
_LOCAL_STATUS_CACHE = {}
DEFAULT_TRACKER_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbydmlNGV_CDYGkMgRmInTi04zfjNpaP3qF1-uEVStDsyqZKxQ8xbNOpDvcsuDmT9m1N/exec"
TRACKER_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker_config.json")

def get_tracker_apps_script_url():
    env_url = os.environ.get("TRACKER_APPS_SCRIPT_URL", "").strip()
    if env_url:
        return env_url
    if os.path.exists(TRACKER_CONFIG_FILE):
        try:
            with open(TRACKER_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f).get("script_url", "").strip()
                if saved:
                    return saved
        except Exception:
            pass
    return DEFAULT_TRACKER_APPS_SCRIPT_URL

def set_tracker_apps_script_url(url):
    try:
        with open(TRACKER_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"script_url": url.strip()}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Error saving tracker config:", e)

ALLOWED_USERS = {
    "PHCN", "NGOẠI TH", "XNTT", "NHH", "TTTM", "NGOẠI CT", "DA LIỄU", "PTGMHS",
    "NỘI TH", "PHỤ SẢN", "MẮT", "TTBVSK", "DƯỢC", "THẬN", "TRUYỀN NHIỄM", "TCKT",
    "NGOẠI UB", "NGOẠI XẠ TRỊ", "RHM", "KB", "CXK", "CSGN", "NGOẠI TN", "HHLS",
    "TMH", "HSTC", "HTL", "NGOẠI TKLN", "NHI", "ĐÔNG Y", "CẤP CỨU", "LKTK"
}

def clean_text(value):
    if value is None:
        return ""
    return re.sub(r'\s+', ' ', str(value)).strip()

def remove_accents(text):
    if not text:
        return ""
    text = unicodedata.normalize('NFD', str(text))
    text = re.sub(r'[\u0300-\u036f]', '', text)
    return text.replace('đ', 'd').replace('Đ', 'D').lower().strip()

def norm_name(value):
    return re.sub(r'[^a-z0-9]', '', remove_accents(value))

def norm_cccd(value):
    digits = re.sub(r'\D', '', str(value or ''))
    return digits.zfill(12) if len(digits) == 12 else digits

def format_cccd_output(value):
    digits = re.sub(r'\D', '', str(value or ''))
    return digits.zfill(12) if len(digits) in [9, 12] else digits

def normalize_date(value):
    if not value:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    val_str = str(value).strip().replace('-', '/').replace('.', '/')
    match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', val_str)
    if match:
        d, m, y = match.groups()
        return f"{int(d):02d}/{int(m):02d}/{y}"
    match_rev = re.search(r'(\d{4})/(\d{1,2})/(\d{1,2})', val_str)
    if match_rev:
        y, m, d = match_rev.groups()
        return f"{int(d):02d}/{int(m):02d}/{y}"
    digits = re.sub(r'\D', '', val_str)
    if len(digits) == 8:
        return f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"
    return val_str

def normalize_issue_place(value):
    c_norm = remove_accents(value)
    if any(k in c_norm for k in ["cuc can cuoc", "cuc csqlhc", "ccsqlhcvttxh", "trat tu xa hoi", "bo cong an"]):
        return "CCSQLHCVTTXH"
    if any(k in c_norm for k in ["du lieu dan cu", "cu dan", "ccsdkqlctvdlqgvdc"]):
        return "CCSĐKQLCTVDLQGVDC"
    return "CCSQLHCVTTXH"

def clean_email(email_str):
    if not email_str:
        return ""
    match = re.search(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', str(email_str))
    return match.group(0).lower().strip() if match else str(email_str).strip()

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

# ============================================================
# TRÍCH XUẤT FILE WORD TỪ REQUEST
# ============================================================
def extract_word_documents_from_files(file_storage_list):
    """Đọc danh sách file Word từ request (hỗ trợ cả file .docx trực tiếp và file .zip chứa file .docx)."""
    docs = []
    for f in file_storage_list:
        if not f or not f.filename:
            continue
        fname = f.filename.lower()
        if fname.endswith(".docx"):
            stream = io.BytesIO(f.read())
            docs.append((f.filename, stream))
        elif fname.endswith(".zip"):
            zstream = io.BytesIO(f.read())
            try:
                with zipfile.ZipFile(zstream) as z:
                    for zname in z.namelist():
                        if zname.lower().endswith(".docx") and not os.path.basename(zname).startswith("~$"):
                            doc_bytes = z.read(zname)
                            docs.append((os.path.basename(zname), io.BytesIO(doc_bytes)))
            except Exception:
                pass
    return docs

def parse_people_from_word_stream(filename, doc_stream, sso_map=None):
    """Trích xuất dữ liệu người dùng từ bảng biểu và văn bản trong file Word."""
    people = []
    if sso_map is None:
        sso_map = {}
    try:
        doc = Document(doc_stream)
        full_doc_text = "\n".join(p.text for p in doc.paragraphs) + "\n" + "\n".join(
            " ".join(cell.text for cell in row.cells) for t in doc.tables for row in t.rows
        )
        norm_doc_full = remove_accents(full_doc_text).lower()
        doc_device_type = ""
        if "ky so tap trung" in norm_doc_full or "rssp" in norm_doc_full:
            doc_device_type = "RSSP"
        elif "sim pki" in norm_doc_full or "dang ky sim" in norm_doc_full:
            doc_device_type = "SIM"
        elif "token" in norm_doc_full:
            doc_device_type = "TOKEN"

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
                t_serial, t_network, t_transfer, t_dev, t_note = -1, -1, -1, -1, -1

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
                    elif any(k in c_norm for k in ['ghi chu', 'note']):
                        if t_note == -1: t_note = idx

                if t_name != -1 and t_cccd != -1:
                    header_row_idx = r_idx
                    name_idx, cccd_idx, dob_idx = t_name, t_cccd, t_dob
                    issue_date_idx, issue_place_idx, email_idx = t_idate, t_iplace, t_email
                    unit_idx, pos_idx, phone_idx, prov_idx = t_unit, t_pos, t_phone, t_prov
                    serial_idx, network_idx, transfer_idx, device_idx = t_serial, t_network, t_transfer, t_dev
                    note_idx = t_note
                    break

            if header_row_idx == -1:
                continue

            for r in table.rows[header_row_idx + 1:]:
                vals = [clean_text(cell.text) for cell in r.cells]
                if len(vals) <= max(name_idx, cccd_idx):
                    continue
                name_val = vals[name_idx]
                cccd_raw = vals[cccd_idx]
                if not name_val or not cccd_raw or "ho va ten" in remove_accents(name_val):
                    continue

                dob = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                issue_date = normalize_date(vals[issue_date_idx]) if issue_date_idx != -1 and issue_date_idx < len(vals) else ""
                issue_place = normalize_issue_place(vals[issue_place_idx]) if issue_place_idx != -1 and issue_place_idx < len(vals) else ""

                # Tách nếu cột gộp Số CCCD; Ngày cấp; Nơi cấp (chuẩn Mẫu 02)
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

                phone = vals[phone_idx] if phone_idx != -1 and phone_idx < len(vals) else ""
                unit = vals[unit_idx] if unit_idx != -1 and unit_idx < len(vals) else ""
                pos = vals[pos_idx] if pos_idx != -1 and pos_idx < len(vals) else ""
                note_val = vals[note_idx] if note_idx != -1 and note_idx < len(vals) else "Cấp mới"
                if not note_val: note_val = "Cấp mới"
                prov = vals[prov_idx] if prov_idx != -1 and prov_idx < len(vals) else DEFAULT_PROVINCE
                if not prov or prov.lower() in ["nam", "nu", "nữ", ""]:
                    prov = DEFAULT_PROVINCE
                serial_val = vals[serial_idx] if serial_idx != -1 and serial_idx < len(vals) else ""
                network_val = vals[network_idx] if network_idx != -1 and network_idx < len(vals) else ""
                transfer_val = vals[transfer_idx] if transfer_idx != -1 and transfer_idx < len(vals) else ""
                dev_val = vals[device_idx] if device_idx != -1 and device_idx < len(vals) else ""

                if not dev_val and doc_device_type:
                    dev_val = doc_device_type
                elif dev_val.upper() in ["X", "CÓ", "CO", "ĐĂNG KÝ", "DANG KY"] and device_idx != -1:
                    dev_val = "SIM" if "sim" in remove_accents(header_texts[device_idx]).lower() else (doc_device_type or "RSSP")

                kn = norm_name(name_val)
                kc = norm_cccd(cccd_clean)
                sso_info = sso_map.get((kn, kc)) or sso_map.get(kc) or sso_map.get(kn) or {}

                if not unit:
                    unit = sso_info.get("unit", DEFAULT_ORG3)
                if not pos:
                    pos = sso_info.get("position", "Nhân viên")

                # Chuẩn hóa org3
                org3_val = DEFAULT_ORG3
                if unit:
                    u_norm = remove_accents(unit).lower()
                    if any(k in u_norm for k in ["bvd-k", "benh vien", "bac ninh so 2", "khoa", "phong", "ttyt"]):
                        org3_val = DEFAULT_ORG3
                    else:
                        org3_val = unit

                email = sso_info.get("email", "")
                if not email and email_idx != -1 and email_idx < len(vals):
                    email = clean_email(vals[email_idx])
                if not email:
                    for c_val in vals:
                        if "@" in c_val and "." in c_val:
                            m_mail = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", c_val)
                            if m_mail:
                                email = clean_email(m_mail.group(0))
                                break
                if not email and kn:
                    # Gợi ý email công vụ nếu chưa có
                    email = f"{kn}@bacninh.gov.vn"

                people.append({
                    "name": name_val, "dob": dob, "cccd": format_cccd_output(cccd_clean),
                    "issue_place": issue_place, "issue_date": issue_date, "email": email,
                    "org1": DEFAULT_ORG1, "org2": DEFAULT_ORG2, "org3": org3_val,
                    "org4": "", "province": prov if prov else DEFAULT_PROVINCE,
                    "position": pos if pos else "Nhân viên", "phone": phone,
                    "unit": unit, "note": note_val,
                    "serial_old": serial_val, "network": network_val, "is_transfer": transfer_val, "device_type": dev_val
                })

        # Nếu không có bảng dữ liệu, quét các đoạn văn bản (đơn đề nghị cá nhân)
        if not people:
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

                kn = norm_name(name_v)
                kc = norm_cccd(cccd_v)
                sso_info = sso_map.get((kn, kc)) or sso_map.get(kc) or sso_map.get(kn) or {}
                if not email_v:
                    email_v = sso_info.get("email", "")
                if not email_v and kn:
                    email_v = f"{kn}@bacninh.gov.vn"
                if unit_v == DEFAULT_ORG3 and sso_info.get("unit"):
                    unit_v = sso_info.get("unit")
                if pos_v == "Nhân viên" and sso_info.get("position"):
                    pos_v = sso_info.get("position")

                org3_v = DEFAULT_ORG3
                if unit_v:
                    u_norm = remove_accents(unit_v).lower()
                    if any(k in u_norm for k in ["bvd-k", "benh vien", "bac ninh so 2", "khoa", "phong", "ttyt"]):
                        org3_v = DEFAULT_ORG3
                    else:
                        org3_v = unit_v

                people.append({
                    "name": name_v, "dob": dob_v, "cccd": format_cccd_output(cccd_v),
                    "issue_place": iplace_v, "issue_date": idate_v, "email": email_v,
                    "org1": DEFAULT_ORG1, "org2": DEFAULT_ORG2, "org3": org3_v,
                    "org4": "", "province": DEFAULT_PROVINCE, "position": pos_v, "phone": phone_v,
                    "unit": unit_v, "note": "Cấp mới",
                    "serial_old": "", "network": "", "is_transfer": "", "device_type": doc_device_type
                })
    except Exception as e:
        print(f"Lỗi đọc Word {filename}: {e}")
    return people

def read_sso_excel_stream(stream):
    """Đọc file Excel SSO để trích xuất mapping email công vụ."""
    sso_map = {}
    try:
        wb = openpyxl.load_workbook(stream, data_only=True)
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            header = [clean_text(c) for c in rows[0]]
            name_idx, email_idx, cccd_idx, unit_idx, pos_idx = -1, -1, -1, -1, -1
            for idx, col in enumerate(header):
                c_low = remove_accents(col)
                if ("ho va ten" in c_low or "ho ten" in c_low) and name_idx == -1: name_idx = idx
                elif ("email" in c_low or "thu dien tu" in c_low or "username" in c_low):
                    if email_idx == -1 or "email" in c_low: email_idx = idx
                elif ("cccd" in c_low or "cmt" in c_low or "cmnd" in c_low) and "ngay" not in c_low and "noi" not in c_low:
                    if cccd_idx == -1: cccd_idx = idx
                elif ("don vi" in c_low or "co quan" in c_low) and unit_idx == -1: unit_idx = idx
                elif "chuc vu" in c_low and pos_idx == -1: pos_idx = idx

            if name_idx == -1: name_idx = 0
            if email_idx == -1: email_idx = 1
            if cccd_idx == -1 and len(header) >= 6: cccd_idx = 5

            for r in rows[1:]:
                if not r: continue
                name_val = r[name_idx] if name_idx != -1 and name_idx < len(r) else ""
                email_val = r[email_idx] if email_idx != -1 and email_idx < len(r) else ""
                cccd_val = r[cccd_idx] if cccd_idx != -1 and cccd_idx < len(r) else ""
                unit_val = r[unit_idx] if unit_idx != -1 and unit_idx < len(r) else ""
                pos_val = r[pos_idx] if pos_idx != -1 and pos_idx < len(r) else ""

                if name_val or email_val or cccd_val:
                    email_c = clean_email(email_val)
                    kn = norm_name(name_val)
                    kc = norm_cccd(cccd_val)
                    info = {"email": email_c, "unit": clean_text(unit_val), "position": clean_text(pos_val)}
                    if kn and kc: sso_map[(kn, kc)] = info
                    if kn and kn not in sso_map: sso_map[kn] = info
                    if kc and kc not in sso_map: sso_map[kc] = info
    except Exception as e:
        print(f"Lỗi đọc SSO Excel: {e}")
    return sso_map

def build_vgca_19_col_excel(people_list):
    """Xuất file Ket_qua.xlsx đúng 19 cột quy chuẩn Import VGCA (Ban Cơ yếu Chính phủ)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Import_VGCA"

    columns = [
        "STT",
        "Họ tên (bắt buộc)",
        "Ngày sinh (bắt buộc, để dạng dd/MM/YYYY)",
        "CCCC (Căn cước công dân, bắt buộc)",
        "Nơi cấp (Nơi cấp CCCD; để một trong 2 nơi: CCSQLHCVTTXH hoặc CCSĐKQLCTVDLQGVDC)",
        "Ngày cấp (bắt buộc, để dạng dd/MM/YYYY)",
        "Địa chỉ thư điện tử công vụ (bắt buộc)",
        "Tổ chức cấp 1 (bắt buộc)",
        "Tổ chức cấp 2",
        "Tổ chức cấp 3",
        "Tổ chức cấp 4",
        "Tỉnh thành phố (Bắt buộc)",
        "Tỉnh/Thành phố",
        "Chức vụ (Bắt buộc)",
        "Điện thoại",
        "Serial CTS cũ",
        "Nhà mạng (bắt buộc nếu đăng ký SIM)",
        "Là chuyển số (đánh dấu X nếu muốn chuyển số đang dùng sang SIM ký số)",
        "Loại thiết bị( nếu không ghi sẽ là token; là một trong các loại sau: (TOKEN, SIM, SOFT, RSSP, HSM); RSSP là ký số tập trung"
    ]
    for col_idx, col_name in enumerate(columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)

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

    out_io = io.BytesIO()
    wb.save(out_io)
    out_io.seek(0)
    return out_io

def build_mau_02_docx(people_list):
    """Xuất file Công văn Mẫu 02 đề nghị cấp chứng thư số chuyên dùng Chính phủ theo đúng mẫu chính thức của BVĐK Bắc Ninh số 2 (Mẫu T10)."""
    template_name_base = "Mẫu 02_CV cap CKS cho BVDK BN số 2_T10 2026_3_6"
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidate_paths = [
        os.path.join(base_dir, "templates", f"{template_name_base}.docx"),
        os.path.join(os.getcwd(), "templates", f"{template_name_base}.docx"),
        os.path.join(base_dir, f"{template_name_base}.docx"),
        os.path.join(os.getcwd(), f"{template_name_base}.docx"),
        os.path.join(base_dir, "templates", "Mẫu 02_CV cap CKS cho BVDK BN số 2_T08 2026_3_6.docx"),
        os.path.join(os.getcwd(), "templates", "Mẫu 02_CV cap CKS cho BVDK BN số 2_T08 2026_3_6.docx"),
    ]

    template_path = None
    for cp in candidate_paths:
        if os.path.exists(cp):
            template_path = cp
            break

    doc = None
    if template_path:
        try:
            doc = Document(template_path)
        except Exception:
            doc = None

    now = datetime.now()
    date_str = f"Bắc Ninh, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"

    if doc is not None:
        # 1. Cập nhật ngày tháng ở Table 0 (R1C1)
        if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
            cell_date = doc.tables[0].rows[1].cells[1]
            for p in cell_date.paragraphs:
                if "Bắc Ninh" in p.text:
                    if p.runs:
                        p.runs[0].text = date_str
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = date_str

        # 2. Cập nhật số lượng ở Paragraph "2. Số lượng và danh sách đăng ký: ..."
        count_val = len(people_list)
        for p in doc.paragraphs:
            if "Số lượng và danh" in p.text or "danh sách đăng ký" in p.text:
                p.text = f"2. Số lượng và danh sách đăng ký: {count_val:02d}, gồm:"
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(11)

        # 3. Điền Table 1 (11 cột)
        t_target = None
        for t in doc.tables:
            if len(t.columns) == 11:
                t_target = t
                break
        if not t_target and len(doc.tables) >= 2:
            t_target = doc.tables[1]

        if t_target:
            template_tr = None
            if len(t_target.rows) > 1:
                template_tr = copy.deepcopy(t_target.rows[1]._tr)

            parent = t_target._tbl
            # Xóa các dòng mẫu cũ
            for r in list(t_target.rows)[1:]:
                parent.remove(r._tr)

            alignments = [
                WD_ALIGN_PARAGRAPH.CENTER, # STT
                WD_ALIGN_PARAGRAPH.LEFT,   # Họ tên
                WD_ALIGN_PARAGRAPH.CENTER, # Ngày sinh
                WD_ALIGN_PARAGRAPH.LEFT,   # CCCD;Ngày cấp;Nơi cấp
                WD_ALIGN_PARAGRAPH.LEFT,   # Email
                WD_ALIGN_PARAGRAPH.LEFT,   # Đơn vị công tác
                WD_ALIGN_PARAGRAPH.CENTER, # Tỉnh/TP
                WD_ALIGN_PARAGRAPH.CENTER, # Chức vụ
                WD_ALIGN_PARAGRAPH.CENTER, # SĐT
                WD_ALIGN_PARAGRAPH.CENTER, # Serial cũ
                WD_ALIGN_PARAGRAPH.CENTER  # SIM PKI
            ]

            for idx, p_info in enumerate(people_list, start=1):
                if isinstance(p_info, dict):
                    name = p_info.get("name", "")
                    dob = p_info.get("dob", "")
                    cccd_raw = p_info.get("cccd", "")
                    idate = p_info.get("issue_date", "")
                    iplace = p_info.get("issue_place", "")
                    cccd_col = f"{cccd_raw};{idate};{iplace}".rstrip(";")
                    email = p_info.get("email", "")
                    unit = p_info.get("org3", "") or p_info.get("unit", "")
                    prov = p_info.get("province", "Bắc Ninh")
                    pos = p_info.get("position", "")
                    phone = p_info.get("phone", "")
                    serial = p_info.get("serial_old", "")
                    sim = p_info.get("device_type", "") or p_info.get("sim", "")
                else:
                    name = p_info[0] if len(p_info) > 0 else ""
                    dob = p_info[1] if len(p_info) > 1 else ""
                    cccd_col = p_info[2] if len(p_info) > 2 else ""
                    email = p_info[3] if len(p_info) > 3 else ""
                    unit = p_info[4] if len(p_info) > 4 else ""
                    prov = p_info[5] if len(p_info) > 5 else "Bắc Ninh"
                    pos = p_info[6] if len(p_info) > 6 else ""
                    phone = p_info[7] if len(p_info) > 7 else ""
                    serial = p_info[8] if len(p_info) > 8 else ""
                    sim = p_info[9] if len(p_info) > 9 else ""

                row_vals = [
                    str(idx), name, dob, cccd_col, email, unit, prov, pos, phone, serial, sim
                ]

                if template_tr is not None:
                    new_tr = copy.deepcopy(template_tr)
                    new_row = docx.table._Row(new_tr, t_target)
                    for col_idx, text_val in enumerate(row_vals):
                        c = new_row.cells[col_idx]
                        p = c.paragraphs[0]
                        if p.runs:
                            p.runs[0].text = str(text_val)
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = str(text_val)
                        p.alignment = alignments[col_idx]
                    parent.append(new_tr)
                else:
                    new_row = t_target.add_row()
                    for col_idx, text_val in enumerate(row_vals):
                        c = new_row.cells[col_idx]
                        c.text = str(text_val)
                        p = c.paragraphs[0]
                        p.alignment = alignments[col_idx]
                        if p.runs:
                            p.runs[0].font.name = "Times New Roman"
                            p.runs[0].font.size = Pt(10)
    else:
        # Fallback tạo tài liệu chuẩn Mẫu 02 11 cột (chuẩn mẫu T10 mới nhất) nếu không có file template
        doc = Document()
        for style_name in ['Normal', 'Default Paragraph Font', 'Body Text']:
            if style_name in doc.styles:
                doc.styles[style_name].font.name = 'Times New Roman'
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.9)
            section.right_margin = Inches(0.6)

        t0 = doc.add_table(rows=2, cols=2)
        t0.autofit = True
        c00 = t0.cell(0, 0)
        p00 = c00.paragraphs[0]
        p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p00.add_run("SỞ Y TẾ TP BẮC NINH\nBỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2\n")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        c01 = t0.cell(0, 1)
        p01 = c01.paragraphs[0]
        p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        c10 = t0.cell(1, 0)
        p10 = c10.paragraphs[0]
        p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p10.add_run("Số: 10070/BVT-CNTT\nV/v đề nghị cấp chứng thư số cho cá nhân  (giải pháp ký số tập trung)")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

        c11 = t0.cell(1, 1)
        p11 = c11.paragraphs[0]
        p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p11.add_run(date_str)
        r.italic = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        doc.add_paragraph()
        p_kg = doc.add_paragraph()
        r = p_kg.add_run("Kính gửi: Cục Chứng thực số và Bảo mật thông tin, Ban Cơ yếu Chính phủ.")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        doc.add_paragraph()
        p1 = doc.add_paragraph()
        r = p1.add_run("1. Thông tin người tiếp nhận chứng thư số (thiết bị lưu khóa bí mật): \n")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r2 = p1.add_run("- Họ và tên: Nguyễn Lê Huyền Ly      Số CMND/CCCD/Hộ chiếu: 001191001662.     Ngày cấp: 17/05/2021       Nơi cấp: C06\n- Chức vụ: Nhân viên phòng CNTT. Số điện thoại di động: 0394022258. Địa chỉ thư điện tử công vụ: lynlh662@bacninh.gov.vn\n- Địa chỉ tiếp nhận: đường Nguyễn Quyền, phường Võ Cường, Tp. Bắc Ninh")
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(11)

        p2 = doc.add_paragraph()
        r = p2.add_run(f"2. Số lượng và danh sách đăng ký: {len(people_list):02d}, gồm:")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        t1 = doc.add_table(rows=1, cols=11)
        t1.style = 'Table Grid'
        headers = [
            "STT", "Họ và tên", "Ngày sinh", "Số CMND/ CCCD/Hộ chiếu;Ngày cấp; nơi cấp",
            "Địa chỉ thư điện tử công vụ", "Tên cơ quan-tổ chức công tác", "Tỉnh/Thành phố",
            "Chức vụ", "Số điện thoại di động", "Số hiệu chứng thư cũ (nếu có)", "SIM PKI"
        ]
        for col_i, h_text in enumerate(headers):
            cell = t1.cell(0, col_i)
            cell.text = h_text
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].bold = True
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(10)

        alignments = [
            WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER
        ]

        for idx, p_info in enumerate(people_list, start=1):
            if isinstance(p_info, dict):
                name = p_info.get("name", "")
                dob = p_info.get("dob", "")
                cccd_col = f"{p_info.get('cccd', '')};{p_info.get('issue_date', '')};{p_info.get('issue_place', '')}".rstrip(";")
                email = p_info.get("email", "")
                unit = p_info.get("org3", "") or p_info.get("unit", "")
                prov = p_info.get("province", "Bắc Ninh")
                pos = p_info.get("position", "")
                phone = p_info.get("phone", "")
                serial = p_info.get("serial_old", "")
                sim = p_info.get("device_type", "") or p_info.get("sim", "")
            else:
                name, dob, cccd_col, email, unit, prov, pos, phone, serial, sim = p_info[:10]

            row = t1.add_row()
            vals = [str(idx), name, dob, cccd_col, email, unit, prov, pos, phone, serial, sim]
            for col_i, text_v in enumerate(vals):
                c = row.cells[col_i]
                c.text = str(text_v)
                p = c.paragraphs[0]
                p.alignment = alignments[col_i]
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(10)

        doc.add_paragraph()
        t2 = doc.add_table(rows=1, cols=2)
        t2.autofit = True
        c20 = t2.cell(0, 0)
        p20 = c20.paragraphs[0]
        r = p20.add_run("Nơi nhận:\n- Như trên;\n- Lưu: VT, CNTT.")
        r.italic = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

        c21 = t2.cell(0, 1)
        p21 = c21.paragraphs[0]
        p21.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p21.add_run("KT. GIÁM ĐỐC BỆNH VIỆN\nPHÓ GIÁM ĐỐC\n(Ký, đóng dấu và ghi rõ họ tên)\n\n\n\nNguyễn Xuân Thủy")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)


    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io


def build_email_docx(people_list):
    """Xuất file Công văn Đề nghị cấp mới địa chỉ email công vụ chuẩn Sở TT&TT / BVĐK Bắc Ninh số 2."""
    template_name_base = "Đề nghị cấp mới địa chỉ email công vụ"
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidate_paths = [
        os.path.join(base_dir, "templates", f"{template_name_base}.docx"),
        os.path.join(os.getcwd(), "templates", f"{template_name_base}.docx"),
        os.path.join(base_dir, f"{template_name_base}.docx"),
        os.path.join(os.getcwd(), f"{template_name_base}.docx"),
    ]

    template_path = None
    for cp in candidate_paths:
        if os.path.exists(cp):
            template_path = cp
            break

    doc = None
    if template_path:
        try:
            doc = Document(template_path)
        except Exception:
            doc = None

    if doc is None:
        raise ValueError("Không tìm thấy file mẫu 'Đề nghị cấp mới địa chỉ email công vụ.docx'!")

    now = datetime.now()
    date_str = f"Bắc Ninh, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"

    # 1. Cập nhật ngày tháng văn bản (Table 0, R1C1)
    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        cell_date = doc.tables[0].rows[1].cells[1]
        for p in cell_date.paragraphs:
            if "Bắc Ninh" in p.text:
                if p.runs:
                    p.runs[0].text = date_str
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = date_str

    # 2. Cập nhật số lượng bản ghi ở Mục 2
    count_val = len(people_list)
    for p in doc.paragraphs:
        if "Số lượng và danh" in p.text or "danh sách đăng ký" in p.text:
            found_colon = False
            count_runs = []
            for r in p.runs:
                if ":" in r.text:
                    found_colon = True
                    continue
                if found_colon:
                    if "," in r.text or "gồm" in r.text:
                        break
                    count_runs.append(r)
            if count_runs:
                count_runs[0].text = f" {count_val:02d}"
                for r in count_runs[1:]:
                    r.text = ""
            else:
                p.text = f"2. Số lượng và danh sách đăng ký: {count_val:02d}, gồm:"

    # 3. Điền bảng danh sách đăng ký Email (Table 1 trong doc - 8 cột)
    t_target = None
    for t in doc.tables:
        if len(t.columns) == 8:
            t_target = t
            break
    if not t_target and len(doc.tables) >= 2:
        t_target = doc.tables[1]

    if t_target and len(t_target.rows) > 1:
        template_tr = copy.deepcopy(t_target.rows[1]._tr)
        parent = t_target._tbl

        for r in list(t_target.rows)[1:]:
            parent.remove(r._tr)

        alignments = [
            WD_ALIGN_PARAGRAPH.CENTER,  # 0: STT
            WD_ALIGN_PARAGRAPH.LEFT,    # 1: Họ và tên
            WD_ALIGN_PARAGRAPH.CENTER,  # 2: Ngày sinh
            WD_ALIGN_PARAGRAPH.CENTER,  # 3: Di động
            WD_ALIGN_PARAGRAPH.CENTER,  # 4: Số CCCD
            WD_ALIGN_PARAGRAPH.LEFT,    # 5: Đơn vị công tác
            WD_ALIGN_PARAGRAPH.CENTER,  # 6: Chức vụ
            WD_ALIGN_PARAGRAPH.CENTER   # 7: Ghi chú
        ]

        for idx, p_info in enumerate(people_list, start=1):
            name = p_info.get("name", "")
            dob = normalize_date(p_info.get("dob", ""))
            phone = clean_text(p_info.get("phone", ""))
            cccd = format_cccd_output(p_info.get("cccd", ""))
            
            raw_unit = clean_text(p_info.get("unit") or p_info.get("org3") or "")
            if raw_unit:
                u_low = raw_unit.lower()
                if any(k in u_low for k in ["bv", "bệnh viện", "benh vien", "bbvskcb"]):
                    unit_str = raw_unit
                else:
                    unit_str = f"{raw_unit} - BVĐK Bắc Ninh số 2"
            else:
                unit_str = "Bệnh viện Đa khoa Bắc Ninh số 2"

            pos = p_info.get("position", "Nhân viên")
            note = p_info.get("note", "Cấp mới")
            if not note:
                note = "Cấp mới"

            row_vals = [
                str(idx),
                name,
                dob,
                phone,
                cccd,
                unit_str,
                pos,
                note
            ]

            new_tr = copy.deepcopy(template_tr)
            new_row = docx.table._Row(new_tr, t_target)
            for col_idx, text_val in enumerate(row_vals):
                cell = new_row.cells[col_idx]
                p = cell.paragraphs[0]
                if p.runs:
                    p.runs[0].text = str(text_val)
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(12)
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = str(text_val)
                    if p.runs:
                        p.runs[0].font.name = "Times New Roman"
                        p.runs[0].font.size = Pt(12)
                p.alignment = alignments[col_idx]
            parent.append(new_tr)

    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io


def build_email_excel(people_list):
    """Xuất danh sách đăng ký mở hòm thư điện tử công vụ ra Excel."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Danh_sach_cap_email"

    ws.merge_cells("A1:G1")
    ws["A1"] = "DANH SÁCH ĐỀ NGHỊ CẤP MỚI ĐỊA CHỈ THƯ ĐIỆN TỬ CÔNG VỤ"
    ws["A1"].font = Font(name="Times New Roman", size=14, bold=True)
    center_cell(ws["A1"])
    ws.row_dimensions[1].height = 28

    headers = ["STT", "Họ và tên", "Ngày sinh", "Số điện thoại", "Số CCCD", "Khoa / Phòng / Đơn vị", "Chức vụ / Vị trí"]
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Times New Roman", size=11, bold=True)
        cell.border = border
        center_cell(cell, wrap=True)

    for idx, p in enumerate(people_list, 1):
        r = idx + 3
        raw_u = clean_text(p.get("unit") or p.get("org3") or "")
        if raw_u and not any(k in raw_u.lower() for k in ["bv", "bệnh viện", "benh vien", "bbvskcb"]):
            unit_disp = f"{raw_u} - BVĐK Bắc Ninh số 2"
        else:
            unit_disp = raw_u or "Bệnh viện Đa khoa Bắc Ninh số 2"

        row_vals = [idx, p["name"], p["dob"], p["phone"], p["cccd"], unit_disp, p["position"]]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            cell.font = Font(name="Times New Roman", size=11)
            cell.border = border
            if c_idx in [1, 3, 4, 5]:
                center_cell(cell)

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 17
    ws.column_dimensions["F"].width = 30
    ws.column_dimensions["G"].width = 20

    out_io = io.BytesIO()
    wb.save(out_io)
    out_io.seek(0)
    return out_io

# ============================================================
# LOGIC XỬ LÝ BÁO CÁO GIÁM ĐỊNH BHYT (A4 DỌC)
# ============================================================
def select_giam_dinh_sheet(wb):
    # 1. Ưu tiên sheet có tên chứa "người dùng" / "nguoi dung" / "chấm công" / "cham cong"
    for name in wb.sheetnames:
        norm = remove_accents(name).lower().strip()
        if "nguoi dung" in norm or norm in ("user", "users") or "cham cong" in norm:
            return wb[name]

    # 2. Thử tìm sheet nào chứa tiêu đề STT và NGƯỜI DÙNG
    for name in wb.sheetnames:
        ws = wb[name]
        try:
            r, s_c, u_c, days = find_giam_dinh_header_and_cols(ws)
            if u_c is not None and days:
                return ws
        except Exception:
            pass

    # 3. Sheet active hoặc sheet đầu tiên
    return wb.active or wb[wb.sheetnames[0]]

def find_giam_dinh_header_and_cols(ws):
    best_candidate = None
    max_days_found = 0

    max_rows = min(ws.max_row, 35)
    max_cols = min(ws.max_column, 65)

    for r in range(1, max_rows + 1):
        stt_col = None
        user_col = None
        day_cols = {}

        # Quét hàng r và hàng r+1 (hỗ trợ tiêu đề 2 tầng/merged headers)
        rows_to_check = [r]
        if r + 1 <= ws.max_row:
            rows_to_check.append(r + 1)

        for c in range(1, max_cols + 1):
            for row_idx in rows_to_check:
                raw_val = ws.cell(row_idx, c).value
                if raw_val is None:
                    continue
                txt = remove_accents(clean_text(raw_val)).upper()
                txt_flat = txt.replace(" ", "").replace("_", "")

                # Kiểm tra cột STT
                if stt_col is None and (txt in ("STT", "SO TT", "SOTT", "TT") or txt == "#"):
                    stt_col = c

                # Kiểm tra cột Người dùng
                if user_col is None and (
                    "NGUOIDUNG" in txt_flat or 
                    "TAIKHOAN" in txt_flat or 
                    txt in ("USER", "USERNAME", "HO VA TEN", "HO TEN", "NHAN VIEN", "CAN BO", "TEN NGUOI DUNG")
                ):
                    user_col = c

                # Kiểm tra cột Ngày (1..31)
                day = None
                if isinstance(raw_val, (int, float)):
                    int_v = int(raw_val)
                    if int_v == raw_val and 1 <= int_v <= 31:
                        day = int_v
                elif str(raw_val).strip().isdigit():
                    int_v = int(str(raw_val).strip())
                    if 1 <= int_v <= 31:
                        day = int_v
                elif re.match(r'^ngày\s*(\d+)$', remove_accents(str(raw_val)).lower().strip()):
                    day = int(re.match(r'^ngày\s*(\d+)$', remove_accents(str(raw_val)).lower().strip()).group(1))
                elif re.match(r'^(\d{1,2})/\d+', str(raw_val).strip()):
                    day = int(re.match(r'^(\d{1,2})/\d+', str(raw_val).strip()).group(1))

                if day is not None and 1 <= day <= 31 and day not in day_cols:
                    day_cols[day] = c

        # Đánh giá ứng viên
        if len(day_cols) >= 3 and len(day_cols) > max_days_found:
            max_days_found = len(day_cols)
            final_stt = stt_col if stt_col else 1
            final_user = user_col if user_col else (final_stt + 1)
            best_candidate = (r, final_stt, final_user, day_cols)

    if best_candidate:
        return best_candidate

    # Fallback: Quét bất kỳ dòng nào có STT hoặc NGUOI DUNG
    for r in range(1, max_rows + 1):
        vals = [remove_accents(clean_text(ws.cell(r, c).value)).upper() for c in range(1, max_cols + 1)]
        has_stt = any(v in ("STT", "SO TT", "SOTT", "TT") for v in vals)
        has_user = any("NGUOI DUNG" in v or "USER" in v for v in vals)
        if has_stt or has_user:
            u_c = None
            s_c = None
            d_cols = {}
            for c in range(1, max_cols + 1):
                raw = ws.cell(r, c).value
                txt = remove_accents(clean_text(raw)).upper()
                if txt in ("STT", "SO TT", "SOTT"):
                    s_c = c
                if "NGUOI DUNG" in txt or "USER" in txt:
                    u_c = c
                if raw is not None and str(raw).strip().isdigit():
                    v_int = int(str(raw).strip())
                    if 1 <= v_int <= 31:
                        d_cols[v_int] = c
                if r + 1 <= ws.max_row:
                    raw2 = ws.cell(r + 1, c).value
                    if raw2 is not None and str(raw2).strip().isdigit():
                        v_int2 = int(str(raw2).strip())
                        if 1 <= v_int2 <= 31 and v_int2 not in d_cols:
                            d_cols[v_int2] = c
            return r, (s_c or 1), (u_c or 2), d_cols

    sheet_names = ", ".join(ws.parent.sheetnames)
    raise Exception(f"Không tìm thấy dòng tiêu đề STT / NGƯỜI DÙNG trong sheet '{ws.title}'. Các sheet trong file gồm: {sheet_names}")

def process_giam_dinh_web(file_stream, start_day, end_day):
    wb_source = load_workbook(file_stream, data_only=True)
    source = select_giam_dinh_sheet(wb_source)

    header_row, stt_col, user_col, day_columns = find_giam_dinh_header_and_cols(source)

    available_days = sorted(day_columns.keys())
    if not available_days:
        raise Exception("Không tìm thấy các cột ngày (1-31) trong bảng dữ liệu.")

    last_available_day = max(available_days)
    if end_day > last_available_day:
        end_day = last_available_day

    selected_days = [d for d in range(start_day, end_day + 1) if d in day_columns]
    if not selected_days:
        selected_days = available_days

    source_columns = [stt_col, user_col] + [day_columns[d] for d in selected_days]
    result_last_col = len(source_columns)
    result_last_letter = get_column_letter(result_last_col)

    result_wb = Workbook()
    for s in result_wb._named_styles:
        if s.name == "Normal":
            s.font = Font(name="Times New Roman", size=14)
    result = result_wb.active
    result.title = "Người dùng"

    for row in range(1, 6):
        for new_col, old_col in enumerate(source_columns, start=1):
            source_cell = source.cell(row, old_col)
            target_cell = result.cell(row, new_col)
            target_cell.value = source_cell.value
            copy_style(source_cell, target_cell)
            if row in (4, 5):
                target_cell.font = Font(name="Times New Roman", size=14, bold=True)

    title = source["A1"].value
    for col in range(1, result_last_col + 1):
        c_cell = result.cell(1, col)
        if col > 1:
            c_cell.value = None
        c_cell.font = Font(name="Times New Roman", size=14, bold=True)
    result["A1"] = title
    result.merge_cells(start_row=1, start_column=1, end_row=1, end_column=result_last_col)
    center_cell(result["A1"], wrap=False)
    result.row_dimensions[1].height = 30

    note = source["A2"].value
    for col in range(1, result_last_col + 1):
        c_cell = result.cell(2, col)
        if col > 1:
            c_cell.value = None
        c_cell.font = Font(name="Times New Roman", size=14, italic=True)
    result["A2"] = note
    result.merge_cells(start_row=2, start_column=1, end_row=2, end_column=result_last_col)
    result.row_dimensions[2].height = 24

    result["A5"] = None
    result.merge_cells("A4:A5")
    center_cell(result["A4"], wrap=False)
    result["A4"].font = Font(name="Times New Roman", size=14, bold=True)

    result["B5"] = None
    result.merge_cells("B4:B5")
    center_cell(result["B4"], wrap=False)
    result["B4"].font = Font(name="Times New Roman", size=14, bold=True)

    result.row_dimensions[4].height = 28
    result.row_dimensions[5].height = 26

    output_row = 6
    stt = 1
    ALLOWED_USERS_NORM = {remove_accents(u).upper(): u for u in ALLOWED_USERS}

    for source_row in range(6, source.max_row + 1):
        raw_user = source.cell(source_row, user_col).value
        user_name = clean_text(raw_user).upper()
        user_norm = remove_accents(user_name)
        if user_name not in ALLOWED_USERS and user_norm not in ALLOWED_USERS_NORM:
            continue
        for new_col, old_col in enumerate(source_columns, start=1):
            source_cell = source.cell(source_row, old_col)
            target_cell = result.cell(output_row, new_col)
            target_cell.value = stt if new_col == 1 else source_cell.value
            copy_style(source_cell, target_cell)
            target_cell.font = Font(name="Times New Roman", size=14)
            if new_col == 1:
                center_cell(target_cell, wrap=False)
            elif new_col == 2:
                target_cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                center_cell(target_cell, wrap=False)
        result.row_dimensions[output_row].height = 25
        stt += 1
        output_row += 1

    total_row = output_row
    result.cell(total_row, 1, "")
    result.cell(total_row, 2, "Tổng cộng")

    for new_col in range(1, result_last_col + 1):
        cell = result.cell(total_row, new_col)
        cell.font = Font(name="Times New Roman", size=14, bold=True)
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    center_cell(result.cell(total_row, 2), wrap=False)
    result.row_dimensions[total_row].height = 28

    for new_col in range(3, result_last_col + 1):
        col_letter = get_column_letter(new_col)
        cell = result.cell(total_row, new_col)
        cell.value = f"=SUM({col_letter}6:{col_letter}{total_row - 1})"
        cell.font = Font(name="Times New Roman", size=14, bold=True)
        center_cell(cell, wrap=False)

    # ĐẢM BẢO TUYỆT ĐỐI TOÀN BỘ CÁC Ô ĐỀU LÀ FONT TIMES NEW ROMAN VÀ CỠ 14
    for r in range(1, total_row + 1):
        for c in range(1, result_last_col + 1):
            cell = result.cell(r, c)
            b = True if r in (1, 4, 5, total_row) else (cell.font.bold if cell.font else False)
            it = True if r == 2 else (cell.font.italic if cell.font else False)
            cell.font = Font(name="Times New Roman", size=14, bold=b, italic=it)

    result.column_dimensions["A"].width = 8
    result.column_dimensions["B"].width = 28
    for col in range(3, result_last_col + 1):
        result.column_dimensions[get_column_letter(col)].width = 9.5

    result.print_title_rows = "4:5"
    result.page_setup.orientation = result.ORIENTATION_PORTRAIT
    result.page_setup.paperSize = result.PAPERSIZE_A4
    result.sheet_properties.pageSetUpPr.fitToPage = True
    result.page_setup.fitToWidth = 1
    result.page_setup.fitToHeight = 0
    result.page_margins.left = 0.25
    result.page_margins.right = 0.25
    result.page_margins.top = 0.4
    result.page_margins.bottom = 0.4
    result.page_margins.header = 0.2
    result.page_margins.footer = 0.2

    out_io = io.BytesIO()
    result_wb.save(out_io)
    out_io.seek(0)
    return out_io

# ============================================================
# LOGIC XỬ LÝ BÁO CÁO P.CNTT & TẢI GOOGLE SHEETS
# ============================================================
def extract_spreadsheet_id(url):
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9\-_]+)', url)
    if not match:
        raise ValueError("URL không hợp lệ. Vui lòng cung cấp link Google Sheets hợp lệ.")
    return match.group(1)

def get_google_sheet_names(url):
    spreadsheet_id = extract_spreadsheet_id(url)
    html_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    req = urllib.request.Request(html_url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        sheets = re.findall(r'class="[^"]*docs-sheet-tab-caption[^"]*">([^<]+)<', html)
        if sheets:
            seen = set()
            clean_sheets = []
            for s in sheets:
                s_clean = s.strip()
                if s_clean and s_clean not in seen:
                    seen.add(s_clean)
                    clean_sheets.append(s_clean)
            if clean_sheets:
                return clean_sheets
    except Exception:
        pass

    stream = download_google_sheet(url)
    wb = load_workbook(stream, read_only=True)
    return wb.sheetnames

def download_google_sheet_single_tab(url, sheet_name):
    spreadsheet_id = extract_spreadsheet_id(url)
    encoded_sheet = urllib.parse.quote(sheet_name)
    csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"

    try:
        req = urllib.request.Request(csv_url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })
        with urllib.request.urlopen(req, timeout=20) as resp:
            raw_csv = resp.read().decode("utf-8", errors="ignore")

        if raw_csv and not raw_csv.startswith("<!DOCTYPE html>"):
            wb_temp = Workbook()
            ws_temp = wb_temp.active
            ws_temp.title = sheet_name

            reader = csv.reader(io.StringIO(raw_csv))
            for r_idx, row in enumerate(reader, 1):
                for c_idx, val in enumerate(row, 1):
                    if val != "":
                        try:
                            if "." in val:
                                val = float(val)
                            else:
                                val = int(val)
                        except ValueError:
                            pass
                        ws_temp.cell(r_idx, c_idx, val)

            temp_stream = io.BytesIO()
            wb_temp.save(temp_stream)
            temp_stream.seek(0)
            return temp_stream
    except Exception:
        pass

    return download_google_sheet(url)

def download_google_sheet(url):
    spreadsheet_id = extract_spreadsheet_id(url)
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

# ============================================================
# LOGIC BÁO CÁO TỔNG HỢP CÔNG VIỆC P.CNTT (CHUẨN 100% app_tong_hop.py)
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

FIXED_COL_WIDTHS_18 = [
    (1,  5.0),          # A: STT
    (2,  25.0),         # B: KHOA, PHÒNG, TRUNG TÂM
    (3,  11.0),         # C: Xoá CLS, PTTT
    (4,  10.0),         # D: Sửa phí
    (7,  18.140625),    # G: Sửa thông tin BN...
    (9,  13.28515625),  # I: Sửa mã bệnh kèm theo...
    (10, 14.5703125),   # J: Hỗ trợ ký số PM...
    (11, 12.0),         # K: Sửa giấy RV...
    (18, 42.7109375),   # R: SỬA LỖI KHÁC
]

def map_software_header_to_index(header_name):
    norm_h = norm_name(header_name)
    if not norm_h:
        return None
    target_15 = STANDARD_HEADERS_18[2:17]
    for idx, th in enumerate(target_15):
        if norm_name(th) == norm_h:
            return idx
    if "xoacls" in norm_h or "xoapttt" in norm_h: return 0
    if "suaphi" in norm_h or "epthe" in norm_h: return 1
    if "huycks" in norm_h: return 2
    if "huyban" in norm_h: return 3
    if "suathongtinbn" in norm_h or "suacongkham" in norm_h or "suakieukham" in norm_h: return 4
    if "suangaygiocls" in norm_h or "suamaycls" in norm_h or "capnhatcls" in norm_h or "giodocketqua" in norm_h: return 5
    if "suamabenh" in norm_h or "benhkemtheo" in norm_h: return 6
    if "kyso" in norm_h or "qlvbdh" in norm_h: return 7
    if "suagiayrv" in norm_h or "chuyentuyen" in norm_h or "ketthucravien" in norm_h: return 8
    if "dangbai" in norm_h or "website" in norm_h: return 9
    if "suacachdungthuoc" in norm_h or "ngaygiothuoc" in norm_h: return 10
    if "phanquyenuser" in norm_h or "doinick" in norm_h: return 11
    if "loikhongketthuc" in norm_h: return 12
    if "suaxoagiuong" in norm_h or "themgiuong" in norm_h or "capnhatgiuong" in norm_h: return 13
    if "chuyenthuoccp" in norm_h or "cpdg" in norm_h: return 14
    return None

def build_formatted_cntt_report(valid_departments, data_matrix, other_errors, output_file=None, sheet_name=None):
    result_wb = Workbook()
    ws = result_wb.active
    ws.title = "BÁO CÁO TỔNG HỢP"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    # Dark gray header fill (#595959)
    header_fill = PatternFill(fill_type="solid", fgColor="595959")

    # 1. Row 1: Title (Times New Roman cỡ 16 in đậm)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=18)
    title_text = f"TỔNG HỢP CÔNG VIỆC PHÒNG CNTT - NGÀY {sheet_name}" if (sheet_name and not str(sheet_name).lower().startswith("sheet")) else "TỔNG HỢP CÔNG VIỆC P.CNTT"
    for c in range(1, 19):
        cell_t = ws.cell(1, c)
        cell_t.font = Font(name="Times New Roman", size=16, bold=True)
        cell_t.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(1, 1, title_text)
    ws.row_dimensions[1].height = 42.0

    # 2. Row 3: Headers (Times New Roman cỡ 14 in đậm, chữ trắng, nền xám đậm)
    header_row = 3
    data_start_row = 4

    for col_idx, h_text in enumerate(STANDARD_HEADERS_18, 1):
        cell = ws.cell(header_row, col_idx, h_text)
        cell.font = Font(name="Times New Roman", size=14, bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 112.5

    # 3. Data Rows: Font Times New Roman, cỡ 14
    for idx, dept in enumerate(valid_departments, 1):
        r_idx = data_start_row + idx - 1

        # Col 1 (A): STT
        c_stt = ws.cell(r_idx, 1, idx)
        c_stt.font = Font(name="Times New Roman", size=14)
        c_stt.border = border
        c_stt.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Col 2 (B): Khoa/Phòng
        c_dept = ws.cell(r_idx, 2, dept)
        c_dept.font = Font(name="Times New Roman", size=14)
        c_dept.border = border
        c_dept.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Col 3..17 (C..Q): 15 danh mục phần mềm
        counts = data_matrix[dept]
        for c_offset in range(15):
            col = 3 + c_offset
            v = counts[c_offset]
            c_val = ws.cell(r_idx, col, int(v) if (v not in (None, "", 0)) else None)
            c_val.font = Font(name="Times New Roman", size=14)
            c_val.border = border
            c_val.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Col 18 (R): Sửa lỗi khác
        err_list = other_errors.get(dept, [])
        err_lines = []
        for e in err_list:
            e_str = str(e).strip()
            if not e_str:
                continue
            for sub_line in e_str.split("\n"):
                sub_line = sub_line.strip()
                if not sub_line:
                    continue
                nh_sub = norm_name(sub_line)
                # Loại bỏ hoàn toàn dòng tiêu đề hoặc mã số hồ sơ bị lẫn
                if "noidungdexuat" in nh_sub or "suachuavabosung" in nh_sub:
                    continue
                if re.search(r"\b\d{6,}\b", sub_line):
                    continue
                if sub_line.startswith("+") or sub_line.startswith("-"):
                    err_lines.append(sub_line)
                else:
                    err_lines.append(f"+ {sub_line}")
        err_text = "\n".join(err_lines) if err_lines else None

        c_err = ws.cell(r_idx, 18, err_text)
        c_err.font = Font(name="Times New Roman", size=14)
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

        if c_offset == 9:  # Column L: Đăng bài trên website (Excluded from summary)
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

    # 5. Set column widths
    for col_idx_or_letter, width in FIXED_COL_WIDTHS_18:
        col_letter = get_column_letter(col_idx_or_letter)
        ws.column_dimensions[col_letter].width = width

    # 6. Page Print setup: Khổ ngang A4, scale = 46%
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.page_setup.scale = 46
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    ws.freeze_panes = "C4"

    if output_file:
        if isinstance(output_file, str):
            result_wb.save(output_file)
            return output_file
        else:
            result_wb.save(output_file)
            output_file.seek(0)
            return output_file
    else:
        out_stream = io.BytesIO()
        result_wb.save(out_stream)
        out_stream.seek(0)
        return out_stream

def process_cntt_web(file_stream, sheet_name):
    wb_source = load_workbook(file_stream, data_only=True)
    if sheet_name in wb_source.sheetnames:
        ws_source = wb_source[sheet_name]
    else:
        target_name = None
        for s in wb_source.sheetnames:
            if s.strip().lower() == str(sheet_name).strip().lower():
                target_name = s
                break
        ws_source = wb_source[target_name] if target_name else wb_source.active

    # 1. Tự động phát hiện dòng tiêu đề cột
    header_row_source = None
    for r in range(1, 15):
        c2_val = clean_text(ws_source.cell(r, 2).value).upper()
        c1_val = clean_text(ws_source.cell(r, 1).value).upper()
        if "KHOA" in c2_val or "PHÒNG" in c2_val or "TRUNG TÂM" in c2_val or "STT" in c1_val:
            header_row_source = r
            break
    if not header_row_source:
        header_row_source = 4 if ws_source.max_row >= 4 else 1

    # 2. Tìm cột 'Sửa chữa khác' (bắt đầu tìm từ cột 5 để bỏ qua STT, Khoa/Phòng, Số hồ sơ, Số phiếu)
    other_col = None
    for c in range(5, ws_source.max_column + 1):
        h = clean_text(ws_source.cell(header_row_source, c).value)
        if not h and header_row_source > 1:
            h = clean_text(ws_source.cell(header_row_source - 1, c).value)
        nh = norm_name(h)
        if "suachuakhac" in nh or "sualoikhac" in nh or (nh.endswith("khac") and "xuat" not in nh):
            other_col = c
            break

    if not other_col:
        other_col = 35 if ws_source.max_column >= 30 else 25

    # 3. Chỉ quét các cột phần mềm từ cột 5 đến trước cột Sửa chữa khác (bỏ qua cột 1, 2, 3, 4)
    sw_col_target_map = {}
    for col in range(5, min(other_col, ws_source.max_column + 1)):
        h = clean_text(ws_source.cell(header_row_source, col).value)
        if not h and header_row_source > 1:
            h = clean_text(ws_source.cell(header_row_source - 1, col).value)
        nh = norm_name(h)
        # Bỏ qua các cột không phải phần mềm (số hồ sơ, số phiếu, cán bộ, ghi chú, phần cứng...)
        if any(skip_kw in nh for skip_kw in (
            "sohoso", "sophieu", "noidung", "dexuat", "canbo", "ghichu",
            "phancung", "domuc", "caiwin", "loimang", "suamaytinh", "nguon",
            "ocung", "suamayin", "loep", "cambien", "tinhtrang"
        )):
            continue
        t_idx = map_software_header_to_index(h)
        sw_col_target_map[col] = (t_idx, h)

    # 4. Quét các dòng dữ liệu (bắt đầu ngay sau header_row_source)
    departments_order = []
    data_matrix = {}
    other_errors = {}

    data_start_row = header_row_source + 1
    max_scan_row = min(ws_source.max_row, 350)
    consecutive_empty = 0

    for row in range(data_start_row, max_scan_row + 1):
        dept_val = ws_source.cell(row, 2).value
        dept_str = clean_text(dept_val)

        if not dept_str or dept_str.upper() in ("KHOA, PHÒNG, TRUNG TÂM", "TỔNG CỘNG", "STT"):
            stt_val = ws_source.cell(row, 1).value
            if stt_val is None:
                consecutive_empty += 1
                if consecutive_empty >= 8:
                    break
            continue

        consecutive_empty = 0

        if dept_str not in data_matrix:
            departments_order.append(dept_str)
            data_matrix[dept_str] = [0] * 15
            other_errors[dept_str] = []

        for col, (t_idx, col_name) in sw_col_target_map.items():
            val = ws_source.cell(row, col).value
            if val in (None, ""):
                continue
            try:
                val_num = int(float(str(val).strip()))
                if val_num <= 0:
                    continue
            except (ValueError, TypeError):
                val_num = 1

            # Bỏ qua các số lớn > 50 (chỉ có thể là số hồ sơ hoặc mã bệnh nhân nhầm lẫn)
            if val_num > 50:
                continue

            if t_idx is not None:
                data_matrix[dept_str][t_idx] += val_num
            else:
                if col_name and not any(k in norm_name(col_name) for k in ("noidung", "dexuat", "hoso", "phieu")):
                    other_errors[dept_str].append(f"{col_name}: {val_num}")

        # Đọc cột Sửa chữa khác
        err_val = ws_source.cell(row, other_col).value
        if err_val is not None:
            err_str = str(err_val).strip()
            if err_str and err_str not in ("0", "None"):
                nh_err = norm_name(err_str)
                # Loại bỏ hoàn toàn dòng tiêu đề hoặc mã số hồ sơ
                if not any(k in nh_err for k in ("noidungdexuat", "suachuavabosung")):
                    if not re.search(r"^\d{6,}$", err_str):
                        if err_str not in other_errors[dept_str]:
                            other_errors[dept_str].append(err_str)

    valid_departments = [
        dept for dept in departments_order
        if any(v != 0 for v in data_matrix[dept]) or bool(other_errors[dept])
    ]

    out_io = io.BytesIO()
    build_formatted_cntt_report(valid_departments, data_matrix, other_errors, output_file=out_io, sheet_name=sheet_name)
    return out_io

# ============================================================
# FLASK WEB APPLICATION (TƯƠNG THÍCH HOÀN TOÀN VERCEL)
# ============================================================
app = Flask(__name__)
application = app
handler = app

HTML_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="0">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Hệ Thống Tổng Hợp CKS, Email, P.CNTT & Giám Định BHYT - BVĐK Bắc Ninh Số 2</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --primary-light: #eff6ff;
      --bg-dark: #0f172a;
      --sidebar-bg: #0f172a;
      --sidebar-item: #1e293b;
      --sidebar-hover: #334155;
      --sidebar-active: #2563eb;
      --content-bg: #f8fafc;
      --border: #e2e8f0;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --success: #10b981;
      --error: #ef4444;
      --radius: 12px;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', sans-serif; }
    body { background-color: var(--content-bg); color: var(--text-main); display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

    /* TOP HEADER */
    .header {
      background: var(--bg-dark);
      color: white;
      padding: 12px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 1px solid #1e293b;
      flex-shrink: 0;
    }
    .header .brand { font-size: 0.85rem; font-weight: 700; color: #60a5fa; letter-spacing: 0.5px; }
    .header .title { font-size: 1.15rem; font-weight: 800; color: white; margin-top: 2px; }
    .header .status {
      display: inline-flex; align-items: center; gap: 6px;
      background: #064e3b; color: #34d399; font-size: 0.8rem; font-weight: 600;
      padding: 5px 12px; border-radius: 9999px;
    }
    .status-dot { width: 8px; height: 8px; background: #34d399; border-radius: 50%; animation: pulse 2s infinite; }
    @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.4; } 100% { opacity: 1; } }

    /* LAYOUT: SIDEBAR + CONTENT */
    .app-container { display: flex; flex: 1; overflow: hidden; }

    /* SIDEBAR */
    .sidebar {
      width: 340px;
      background: #090d16;
      display: flex;
      flex-direction: column;
      flex-shrink: 0;
      border-right: 1px solid #1e293b;
      overflow-y: auto;
    }
    .sidebar-header {
      padding: 14px 16px 10px;
      border-bottom: 1px solid #1e293b;
      background: #0b1120;
    }
    .sidebar-title { font-size: 0.76rem; font-weight: 800; color: #93c5fd; text-transform: uppercase; letter-spacing: 0.8px; }
    .sidebar-subtitle { font-size: 0.72rem; color: #94a3b8; margin-top: 3px; }

    .sidebar-scrollable {
      padding: 10px 10px;
      display: flex;
      flex-direction: column;
      gap: 12px;
      flex: 1;
    }

    /* NHÓM TÁC VỤ CARD BOX */
    .task-group-card {
      background: #0f172a;
      border: 1px solid #1e293b;
      border-radius: 10px;
      padding: 8px 6px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
    }
    .task-group-card.card-cks {
      border-top: 3px solid #38bdf8;
      background: linear-gradient(180deg, rgba(56, 189, 248, 0.08) 0%, rgba(15, 23, 42, 0.95) 100%);
    }
    .task-group-card.card-bhyt {
      border-top: 3px solid #c084fc;
      background: linear-gradient(180deg, rgba(192, 132, 252, 0.08) 0%, rgba(15, 23, 42, 0.95) 100%);
    }

    .group-header-box {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 4px 6px 8px 6px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
      margin-bottom: 6px;
    }
    .group-header-title {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 0.73rem;
      font-weight: 800;
      letter-spacing: 0.6px;
      text-transform: uppercase;
    }
    .card-cks .group-header-title { color: #38bdf8; }
    .card-bhyt .group-header-title { color: #c084fc; }

    .group-header-badge {
      font-size: 0.62rem;
      font-weight: 700;
      padding: 2px 7px;
      border-radius: 999px;
    }
    .card-cks .group-header-badge {
      background: rgba(56, 189, 248, 0.18);
      color: #7dd3fc;
      border: 1px solid rgba(56, 189, 248, 0.35);
    }
    .card-bhyt .group-header-badge {
      background: rgba(192, 132, 252, 0.18);
      color: #e9d5ff;
      border: 1px solid rgba(192, 132, 252, 0.35);
    }

    .sidebar-menu { list-style: none; padding: 0; margin: 0; display: flex; flex-direction: column; gap: 4px; }
    .sidebar-menu li {
      display: flex; align-items: center; gap: 10px;
      padding: 8px 10px;
      border-radius: 7px;
      background: rgba(30, 41, 59, 0.45);
      border: 1px solid rgba(255, 255, 255, 0.04);
      cursor: pointer;
      transition: all 0.2s ease;
      position: relative;
    }
    .sidebar-menu li:hover {
      background: rgba(51, 65, 85, 0.7);
      transform: translateX(2px);
      border-color: rgba(255, 255, 255, 0.1);
    }
    .sidebar-menu li.active {
      background: #2563eb;
      box-shadow: 0 3px 10px rgba(37, 99, 235, 0.4);
      border-color: #3b82f6;
    }
    .sidebar-menu li .nav-icon { font-size: 1.15rem; flex-shrink: 0; }
    .sidebar-menu li .nav-text { display: flex; flex-direction: column; overflow: hidden; }
    .sidebar-menu li .nav-title { font-size: 0.84rem; font-weight: 700; color: #e2e8f0; }
    .sidebar-menu li.active .nav-title { color: #ffffff; }
    .sidebar-menu li .nav-desc { font-size: 0.70rem; color: #94a3b8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .sidebar-menu li.active .nav-desc { color: #dbeafe; }
    .sidebar-menu li .nav-lock {
      margin-left: auto;
      font-size: 0.78rem;
      opacity: 0.65;
      flex-shrink: 0;
      transition: all 0.2s ease;
    }
    .sidebar-menu li:hover .nav-lock { opacity: 1; transform: scale(1.15); }
    .sidebar-menu li.active .nav-lock { opacity: 0.95; }
    @keyframes authShake {
      0%, 100% { transform: translateX(0); }
      20%, 60% { transform: translateX(-6px); }
      40%, 80% { transform: translateX(6px); }
    }
    .auth-shake { animation: authShake 0.35s ease; }

    .sidebar-footer {
      padding: 12px 16px;
      border-top: 1px solid #1e293b;
      font-size: 0.74rem;
      color: #64748b;
      line-height: 1.4;
      background: #0b1120;
    }

    /* RIGHT CONTENT AREA */
    .content-area {
      flex: 1;
      display: flex;
      flex-direction: column;
      background: #f8fafc;
      overflow-y: auto;
    }
    .breadcrumb-bar {
      background: white;
      padding: 12px 24px;
      border-bottom: 1px solid var(--border);
      display: flex;
      align-items: center;
      gap: 12px;
      flex-shrink: 0;
    }
    .badge {
      background: #dbeafe; color: #1d4ed8;
      font-size: 0.75rem; font-weight: 700;
      padding: 4px 10px; border-radius: 6px;
      text-transform: uppercase;
    }
    .breadcrumb-title { font-size: 1rem; font-weight: 700; color: #0f172a; }

    /* TASK PAGES */
    .page { display: none; padding: 24px; }
    .page.active { display: block; }

    .two-col-grid {
      display: grid;
      grid-template-columns: 1fr 1.35fr;
      gap: 24px;
      align-items: start;
    }

    /* CARD STYLING */
    .card {
      background: white;
      border-radius: var(--radius);
      border: 1px solid var(--border);
      padding: 24px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .card-header {
      display: flex; align-items: center; gap: 10px;
      margin-bottom: 16px; padding-bottom: 12px;
      border-bottom: 1px solid var(--border);
    }
    .card-header h2 { font-size: 1.1rem; font-weight: 700; color: #0f172a; }
    .card-header .icon { font-size: 1.3rem; }

    /* GUIDE BOX */
    .guide-box {
      font-size: 0.88rem;
      line-height: 1.6;
      color: #334155;
    }
    .guide-box h3 { font-size: 0.95rem; font-weight: 700; color: #1e293b; margin: 12px 0 6px; }
    .guide-box p { margin-bottom: 8px; }
    .guide-box ul, .guide-box ol { padding-left: 20px; margin-bottom: 12px; }
    .guide-box li { margin-bottom: 6px; }
    .guide-box .highlight {
      background: #eff6ff; border-left: 3px solid #2563eb;
      padding: 8px 12px; border-radius: 4px; margin: 10px 0; font-size: 0.84rem;
    }

    /* FORM CONTROLS */
    .form-group { margin-bottom: 18px; }
    .form-label { display: block; font-size: 0.88rem; font-weight: 700; margin-bottom: 8px; color: #1e293b; }
    .form-desc { font-size: 0.78rem; color: var(--text-muted); margin-top: 4px; }

    .file-dropzone {
      border: 2px dashed #cbd5e1;
      border-radius: 10px;
      padding: 24px 16px;
      text-align: center;
      cursor: pointer;
      transition: all 0.2s ease;
      background: #f8fafc;
    }
    .file-dropzone:hover { border-color: var(--primary); background: #f0f7ff; }
    .file-dropzone .drop-icon { font-size: 2rem; margin-bottom: 6px; }
    .file-dropzone .drop-text { font-size: 0.9rem; font-weight: 600; color: #334155; }
    .file-dropzone .drop-sub { font-size: 0.75rem; color: #64748b; margin-top: 2px; }
    .file-dropzone input { display: none; }

    .file-chosen-badge {
      margin-top: 8px;
      padding: 8px 12px;
      background: #e0f2fe;
      color: #0369a1;
      border-radius: 6px;
      font-size: 0.8rem;
      font-weight: 600;
      display: none;
      word-break: break-all;
    }

    .form-input {
      width: 100%;
      padding: 10px 14px;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      font-size: 0.9rem;
      outline: none;
      transition: border-color 0.2s;
    }
    .form-input:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15); }

    .radio-group { display: flex; flex-direction: column; gap: 8px; }
    .radio-card {
      display: flex; align-items: center; gap: 10px;
      padding: 10px 14px; border: 1px solid #cbd5e1; border-radius: 8px;
      cursor: pointer; transition: all 0.2s;
    }
    .radio-card:hover { border-color: var(--primary); background: #f0f7ff; }
    .radio-card input[type="radio"] { accent-color: var(--primary); }
    .radio-title { font-size: 0.88rem; font-weight: 600; }
    .radio-desc { font-size: 0.75rem; color: var(--text-muted); }

    .btn {
      width: 100%;
      padding: 12px 20px;
      border: none;
      border-radius: 8px;
      font-size: 0.95rem;
      font-weight: 700;
      color: white;
      background: var(--primary);
      cursor: pointer;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 8px;
      transition: all 0.2s;
      box-shadow: 0 2px 8px rgba(37, 99, 235, 0.25);
    }
    .btn:hover { background: var(--primary-hover); transform: translateY(-1px); }
    .btn:disabled { background: #94a3b8; cursor: not-allowed; transform: none; box-shadow: none; }

    .btn-secondary {
      background: #f1f5f9; color: #1e293b; border: 1px solid #cbd5e1;
      box-shadow: none; width: auto; padding: 8px 14px; font-size: 0.85rem;
    }
    .btn-secondary:hover { background: #e2e8f0; color: black; }

    .alert {
      padding: 12px 16px; border-radius: 8px; font-size: 0.85rem;
      margin-top: 14px; display: none; align-items: center; gap: 10px;
      line-height: 1.4;
    }
    .alert.success { background: #dcfce7; color: #166534; border: 1px solid #86efac; }
    .alert.error { background: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }

    .spinner {
      width: 18px; height: 18px; border: 2px solid #ffffff;
      border-bottom-color: transparent; border-radius: 50%;
      display: none; animation: rotation 1s linear infinite;
    }
    @keyframes rotation { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

    /* TÁC VỤ 10: THỐNG KÊ LỖI XML BHYT */
    .tracker-card {
      background: white; border-radius: var(--radius); border: 1px solid var(--border);
      padding: 22px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 20px;
    }
    .tracker-header-box {
      display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;
      margin-bottom: 18px; padding-bottom: 14px; border-bottom: 1px solid var(--border);
    }
    .tracker-title-box h2 { font-size: 1.15rem; font-weight: 800; color: #0f172a; display: flex; align-items: center; gap: 8px; }
    .tracker-title-box p { font-size: 0.82rem; color: var(--text-muted); margin-top: 3px; }

    .tracker-filter-bar {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px;
      align-items: flex-end; background: #f8fafc; padding: 16px; border-radius: 10px; border: 1px solid #e2e8f0;
    }
    .tracker-filter-item { display: flex; flex-direction: column; gap: 5px; }
    .tracker-filter-item label { font-size: 0.76rem; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.4px; }
    .tracker-select, .tracker-input {
      padding: 8px 12px; font-size: 0.85rem; border-radius: 7px; border: 1px solid #cbd5e1;
      background: white; color: #0f172a; outline: none; transition: border-color 0.15s; width: 100%;
    }
    .tracker-select:focus, .tracker-input:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12); }

    .tracker-stats-grid {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin-bottom: 22px;
    }
    .stat-card {
      background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
      border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; position: relative; overflow: hidden;
      box-shadow: 0 1px 3px rgba(0,0,0,0.04); transition: transform 0.2s, box-shadow 0.2s;
    }
    .stat-card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.06); }
    .stat-card::before {
      content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
    }
    .stat-card.c-total::before { background: linear-gradient(90deg, #ef4444, #f97316); }
    .stat-card.c-unique::before { background: linear-gradient(90deg, #3b82f6, #60a5fa); }
    .stat-card.c-khoa::before { background: linear-gradient(90deg, #8b5cf6, #a78bfa); }
    .stat-card.c-phong::before { background: linear-gradient(90deg, #10b981, #34d399); }
    .stat-card.c-xml::before { background: linear-gradient(90deg, #f59e0b, #fbbf24); }
    .stat-card.c-ok::before { background: linear-gradient(90deg, #059669, #10b981); }

    .status-badge {
      display: inline-flex; align-items: center; justify-content: center; gap: 4px; padding: 3px 8px; border-radius: 6px;
      font-size: 0.74rem; font-weight: 700; white-space: nowrap; transition: all 0.2s;
    }
    .badge-ok { background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }
    .badge-pending { background: #f1f5f9; color: #64748b; border: 1px solid #e2e8f0; }
    .badge-saving { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
    .tracker-status-check {
      width: 17px; height: 17px; cursor: pointer; accent-color: #10b981; margin: 0;
    }

    .stat-card-title { font-size: 0.74rem; font-weight: 700; color: #64748b; text-transform: uppercase; margin-bottom: 6px; }
    .stat-card-value { font-size: 1.65rem; font-weight: 800; color: #0f172a; line-height: 1.1; }
    .stat-card-sub { font-size: 0.72rem; color: #94a3b8; margin-top: 4px; }

    .tracker-cols-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 22px; }

    .ranking-box {
      border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; max-height: 320px; overflow-y: auto;
    }
    .ranking-item {
      display: flex; align-items: center; justify-content: space-between; padding: 10px 14px;
      border-bottom: 1px solid #f1f5f9; cursor: pointer; transition: background 0.15s; font-size: 0.85rem;
    }
    .ranking-item:hover { background: #f0fdf4; }
    .ranking-item:last-child { border-bottom: none; }
    .ranking-name { font-weight: 600; color: #1e293b; display: flex; align-items: center; gap: 8px; flex: 1; }
    .ranking-count-badge {
      background: #fee2e2; color: #dc2626; font-weight: 800; padding: 2px 8px; border-radius: 999px; font-size: 0.75rem;
    }
    .ranking-bar-bg { width: 90px; height: 6px; background: #e2e8f0; border-radius: 3px; overflow: hidden; margin-left: 12px; }
    .ranking-bar-fill { height: 100%; background: #3b82f6; border-radius: 3px; }

    .tracker-table-wrap {
      border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; max-height: 540px; overflow-y: auto; background: white;
    }
    .tracker-table { width: 100%; border-collapse: collapse; font-size: 0.84rem; text-align: left; }
    .tracker-table th {
      background: #f8fafc; color: #475569; font-weight: 700; font-size: 0.77rem; text-transform: uppercase;
      padding: 10px 14px; border-bottom: 1px solid #cbd5e1; position: sticky; top: 0; z-index: 5;
    }
    .tracker-table td { padding: 9px 14px; border-bottom: 1px solid #f1f5f9; vertical-align: top; color: #1e293b; }
    .tracker-table tr:hover td { background: #f8fafc; }

    .badge-xml {
      display: inline-block; padding: 2px 7px; border-radius: 5px; font-size: 0.72rem; font-weight: 800;
      letter-spacing: 0.3px; text-transform: uppercase; white-space: nowrap;
    }
    .badge-xml1 { background: #ede9fe; color: #6d28d9; border: 1px solid #ddd6fe; }
    .badge-xml2 { background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }
    .badge-xml3 { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
    .badge-xml4 { background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; }
    .badge-xml5 { background: #e0e7ff; color: #3730a3; border: 1px solid #c7d2fe; }
    .badge-xml13 { background: #ffe4e6; color: #be123c; border: 1px solid #fecdd3; }
    .badge-xml-default { background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; }

    .btn-copy-hs {
      background: none; border: 1px dashed #cbd5e1; border-radius: 4px; padding: 1px 5px; font-size: 0.72rem;
      cursor: pointer; color: #475569; margin-left: 6px; transition: all 0.15s;
    }
    .btn-copy-hs:hover { background: #eff6ff; color: #2563eb; border-color: #93c5fd; }

    @media (max-width: 992px) {
      .app-container { flex-direction: column; }
      .sidebar { width: 100%; height: auto; border-right: none; border-bottom: 1px solid #1e293b; }
      .two-col-grid { grid-template-columns: 1fr; }
      .tracker-cols-2 { grid-template-columns: 1fr; }
    }
  

    /* ============================================================
       BC SỬA CNTT (DASHBOARD TỔNG HỢP SHEET) STYLES
       ============================================================ */
    .cntt-header-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 16px;
      padding: 12px 18px;
      background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(30, 41, 59, 0.95));
      border: 1px solid rgba(59, 130, 246, 0.25);
      border-radius: 12px;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
    }
    .cntt-header-left {
      display: flex;
      align-items: center;
      gap: 14px;
    }
    .cntt-btn-back {
      padding: 7px 14px;
      background: rgba(30, 41, 59, 0.8);
      border: 1px solid rgba(148, 163, 184, 0.3);
      color: #94a3b8;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s;
    }
    .cntt-btn-back:hover {
      background: rgba(51, 65, 85, 0.9);
      color: #f1f5f9;
      border-color: #38bdf8;
    }
    .cntt-header-title {
      font-size: 20px;
      font-weight: 800;
      color: #ffffff;
      letter-spacing: -0.3px;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .cntt-sync-badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      font-weight: 600;
      color: #10b981;
      background: rgba(16, 185, 129, 0.15);
      border: 1px solid rgba(16, 185, 129, 0.3);
      padding: 3px 10px;
      border-radius: 20px;
    }
    .cntt-header-actions {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .cntt-btn-action {
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 600;
      border-radius: 8px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      border: 1px solid transparent;
      transition: all 0.2s;
    }
    .cntt-btn-sync {
      background: linear-gradient(135deg, #0284c7, #2563eb);
      color: white;
      box-shadow: 0 2px 8px rgba(37, 99, 235, 0.35);
    }
    .cntt-btn-sync:hover {
      background: linear-gradient(135deg, #0369a1, #1d4ed8);
      transform: translateY(-1px);
    }
    .cntt-btn-excel {
      background: rgba(15, 23, 42, 0.8);
      border-color: #0284c7;
      color: #38bdf8;
    }
    .cntt-btn-excel:hover {
      background: #0284c7;
      color: white;
    }
    .cntt-btn-link {
      background: transparent;
      border-color: rgba(148, 163, 184, 0.3);
      color: #94a3b8;
    }
    .cntt-btn-link:hover {
      color: #e2e8f0;
      border-color: #94a3b8;
    }

    /* Sub Tabs Nav */
    .cntt-subnav {
      display: flex;
      align-items: center;
      gap: 6px;
      margin-bottom: 16px;
      border-bottom: 1px solid rgba(51, 65, 85, 0.7);
      padding-bottom: 4px;
      overflow-x: auto;
    }
    .cntt-tab-item {
      padding: 8px 16px;
      font-size: 13px;
      font-weight: 600;
      color: #94a3b8;
      border-radius: 8px 8px 0 0;
      cursor: pointer;
      border: 1px solid transparent;
      border-bottom: none;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      transition: all 0.2s;
      white-space: nowrap;
    }
    .cntt-tab-item:hover {
      color: #f1f5f9;
      background: rgba(30, 41, 59, 0.5);
    }
    .cntt-tab-item.active {
      color: #38bdf8;
      background: rgba(30, 41, 59, 0.95);
      border-color: rgba(59, 130, 246, 0.4);
      border-bottom: 2px solid #38bdf8;
    }

    /* Suggestions & Month selectors */
    .cntt-filter-card {
      background: rgba(15, 23, 42, 0.85);
      border: 1px solid rgba(51, 65, 85, 0.6);
      border-radius: 12px;
      padding: 14px 18px;
      margin-bottom: 16px;
    }
    .cntt-quick-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 12px;
      padding-bottom: 10px;
      border-bottom: 1px solid rgba(51, 65, 85, 0.4);
    }
    .cntt-quick-left {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 8px;
    }
    .cntt-quick-pill {
      padding: 5px 12px;
      background: rgba(30, 41, 59, 0.9);
      border: 1px solid rgba(148, 163, 184, 0.25);
      color: #e2e8f0;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      transition: all 0.2s;
    }
    .cntt-quick-pill:hover {
      background: #0284c7;
      border-color: #38bdf8;
      color: white;
    }
    .cntt-quick-right {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .cntt-btn-small {
      padding: 4px 10px;
      font-size: 11px;
      font-weight: 600;
      border-radius: 6px;
      cursor: pointer;
      border: 1px solid rgba(148, 163, 184, 0.3);
      background: rgba(30, 41, 59, 0.6);
      color: #cbd5e1;
    }
    .cntt-btn-small:hover {
      background: #334155;
      color: white;
    }
    .cntt-btn-small.btn-success {
      color: #34d399;
      border-color: rgba(52, 211, 153, 0.3);
    }
    .cntt-btn-small.btn-danger {
      color: #f87171;
      border-color: rgba(248, 113, 113, 0.3);
    }

    /* Month Tabs */
    .cntt-month-tabs {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 12px;
    }
    .cntt-month-pill {
      padding: 5px 14px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      background: rgba(30, 41, 59, 0.7);
      border: 1px solid rgba(148, 163, 184, 0.2);
      color: #94a3b8;
      transition: all 0.2s;
    }
    .cntt-month-pill:hover {
      background: rgba(51, 65, 85, 0.9);
      color: #f1f5f9;
    }
    .cntt-month-pill.active {
      background: #0284c7;
      color: white;
      border-color: #38bdf8;
      box-shadow: 0 2px 6px rgba(2, 132, 199, 0.4);
    }

    /* Sheet Grid */
    .cntt-sheets-grid {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      max-height: 180px;
      overflow-y: auto;
      padding: 4px;
    }
    .cntt-sheet-chip {
      padding: 6px 12px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      background: rgba(30, 41, 59, 0.85);
      border: 1px solid rgba(71, 85, 105, 0.6);
      color: #cbd5e1;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.15s ease;
      user-select: none;
    }
    .cntt-sheet-chip input[type="checkbox"] {
      accent-color: #0284c7;
      cursor: pointer;
    }
    .cntt-sheet-chip:hover {
      border-color: #38bdf8;
      color: white;
    }
    .cntt-sheet-chip.selected {
      background: linear-gradient(135deg, #0284c7, #1d4ed8);
      border-color: #60a5fa;
      color: #ffffff;
      box-shadow: 0 2px 8px rgba(37, 99, 235, 0.35);
    }

    /* Filter & Search Bar */
    .cntt-action-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 18px;
      padding: 10px 16px;
      background: rgba(15, 23, 42, 0.7);
      border: 1px solid rgba(51, 65, 85, 0.5);
      border-radius: 10px;
    }
    .cntt-action-inputs {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
      flex: 1;
    }
    .cntt-select-cb {
      padding: 8px 14px;
      background: #1e293b;
      border: 1px solid #475569;
      color: #f1f5f9;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      min-width: 220px;
    }
    .cntt-search-input {
      padding: 8px 14px;
      background: #1e293b;
      border: 1px solid #475569;
      color: #f1f5f9;
      border-radius: 8px;
      font-size: 13px;
      flex: 1;
      min-width: 200px;
    }
    .cntt-search-input:focus, .cntt-select-cb:focus {
      outline: none;
      border-color: #38bdf8;
      box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.2);
    }
    .cntt-success-pill {
      font-size: 13px;
      font-weight: 700;
      color: #10b981;
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }

    /* 4 KPI Cards */
    .cntt-kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 14px;
      margin-bottom: 22px;
    }
    .cntt-kpi-card {
      background: linear-gradient(135deg, rgba(30, 41, 59, 0.95), rgba(15, 23, 42, 0.95));
      border: 1px solid rgba(51, 65, 85, 0.8);
      border-radius: 12px;
      padding: 16px 20px;
      position: relative;
      overflow: hidden;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }
    .cntt-kpi-card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0; height: 3px;
    }
    .cntt-kpi-card.kpi-blue::before { background: linear-gradient(90deg, #38bdf8, #0284c7); }
    .cntt-kpi-card.kpi-purple::before { background: linear-gradient(90deg, #a855f7, #6366f1); }
    .cntt-kpi-card.kpi-emerald::before { background: linear-gradient(90deg, #34d399, #059669); }
    .cntt-kpi-card.kpi-amber::before { background: linear-gradient(90deg, #fbbf24, #d97706); }
    .cntt-kpi-label {
      font-size: 11px;
      font-weight: 700;
      color: #94a3b8;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      display: flex;
      align-items: center;
      gap: 6px;
      margin-bottom: 6px;
    }
    .cntt-kpi-value {
      font-size: 32px;
      font-weight: 900;
      color: #ffffff;
      line-height: 1.1;
      margin-bottom: 4px;
    }
    .cntt-kpi-sub {
      font-size: 12px;
      color: #94a3b8;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    /* Staff Aggregation Cards */
    .cntt-staff-card {
      background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(30, 41, 59, 0.9));
      border: 1px solid rgba(59, 130, 246, 0.3);
      border-radius: 14px;
      padding: 20px 22px;
      margin-bottom: 20px;
      box-shadow: 0 6px 20px rgba(0, 0, 0, 0.35);
    }
    .cntt-staff-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
      padding-bottom: 14px;
      border-bottom: 1px solid rgba(51, 65, 85, 0.6);
      margin-bottom: 16px;
    }
    .cntt-staff-left {
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }
    .cntt-avatar {
      width: 42px;
      height: 42px;
      border-radius: 50%;
      background: linear-gradient(135deg, #0284c7, #2563eb);
      color: white;
      font-size: 18px;
      font-weight: 800;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 2px 10px rgba(37, 99, 235, 0.4);
    }
    .cntt-staff-name {
      font-size: 18px;
      font-weight: 800;
      color: #ffffff;
    }
    .cntt-tag-badge {
      padding: 3px 10px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 700;
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }
    .cntt-tag-blue { background: rgba(2, 132, 199, 0.2); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.3); }
    .cntt-tag-slate { background: rgba(51, 65, 85, 0.5); color: #cbd5e1; border: 1px solid rgba(148, 163, 184, 0.25); }
    .cntt-tag-cyan { background: rgba(6, 182, 212, 0.2); color: #22d3ee; border: 1px solid rgba(34, 211, 238, 0.3); }
    .cntt-tag-amber { background: rgba(217, 119, 6, 0.2); color: #fbbf24; border: 1px solid rgba(251, 191, 36, 0.3); }
    
    .cntt-btn-copy {
      padding: 7px 14px;
      background: rgba(30, 41, 59, 0.85);
      border: 1px solid rgba(148, 163, 184, 0.3);
      color: #e2e8f0;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s;
    }
    .cntt-btn-copy:hover {
      background: #0284c7;
      color: white;
      border-color: #38bdf8;
    }

    /* Categories Pills */
    .cntt-sec-title {
      font-size: 13px;
      font-weight: 800;
      color: #e2e8f0;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .cntt-sec-title.accent-amber {
      color: #fbbf24;
      margin-top: 18px;
    }
    .cntt-cats-container {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 14px;
    }
    .cntt-cat-pill {
      background: rgba(30, 41, 59, 0.7);
      border: 1px solid rgba(71, 85, 105, 0.5);
      border-radius: 8px;
      padding: 6px 12px;
      font-size: 12px;
      color: #e2e8f0;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }
    .cntt-cat-count {
      background: #0284c7;
      color: #ffffff;
      font-weight: 800;
      font-size: 11px;
      padding: 2px 7px;
      border-radius: 6px;
    }

    /* Other items list */
    .cntt-other-list {
      background: rgba(15, 23, 42, 0.6);
      border: 1px solid rgba(51, 65, 85, 0.5);
      border-radius: 10px;
      padding: 10px 14px;
      max-height: 260px;
      overflow-y: auto;
    }
    .cntt-other-row {
      padding: 8px 10px;
      border-bottom: 1px solid rgba(51, 65, 85, 0.35);
      font-size: 13px;
      color: #cbd5e1;
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 8px;
    }
    .cntt-other-row:last-child {
      border-bottom: none;
    }
    .cntt-other-idx {
      font-weight: 700;
      color: #fbbf24;
      min-width: 22px;
    }
    .badge-sheet {
      background: rgba(2, 132, 199, 0.2);
      color: #38bdf8;
      border: 1px solid rgba(56, 189, 248, 0.3);
      padding: 2px 6px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 700;
    }
    .badge-dept {
      background: rgba(168, 85, 247, 0.2);
      color: #c084fc;
      border: 1px solid rgba(192, 132, 252, 0.3);
      padding: 2px 6px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 700;
    }
    .badge-meta {
      color: #94a3b8;
      font-size: 11px;
    }
    .badge-hs {
      color: #38bdf8;
      font-size: 11px;
      font-weight: 600;
    }

    /* Toast Notification */
    .cntt-toast-popup {
      position: fixed;
      bottom: 24px;
      right: 24px;
      background: rgba(15, 23, 42, 0.95);
      border: 1px solid #10b981;
      border-radius: 10px;
      padding: 12px 18px;
      display: flex;
      align-items: center;
      gap: 12px;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.5);
      z-index: 99999;
      transform: translateY(120%);
      opacity: 0;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .cntt-toast-popup.show {
      transform: translateY(0);
      opacity: 1;
    }
    .cntt-toast-icon {
      width: 28px;
      height: 28px;
      border-radius: 50%;
      background: #10b981;
      color: white;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 16px;
      font-weight: 800;
    }
    .cntt-toast-title {
      font-size: 13px;
      font-weight: 800;
      color: #ffffff;
    }
    .cntt-toast-msg {
      font-size: 12px;
      color: #94a3b8;
    }

</style>
</head>
<body>

  <!-- TOP HEADER -->
  <div class="header">
    <div>
      <div class="brand">🏥 BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2</div>
      <div class="title">HỆ THỐNG TỔNG HỢP CẤP CKS, EMAIL, P.CNTT & GIÁM ĐỊNH BẢO HIỂM</div>
    </div>
    <div class="status">
      <span class="status-dot"></span>
      <span>Hệ thống trực tuyến sẵn sàng</span>
    </div>
  </div>

  <!-- MAIN BODY: SIDEBAR + CONTENT -->
  <div class="app-container">

    <!-- SIDEBAR ĐƯỢC CHIA THÀNH 2 NHÓM -->
    <div class="sidebar">
      <div class="sidebar-header">
        <div class="sidebar-title">📑 DANH MỤC TÁC VỤ</div>
        <div class="sidebar-subtitle">Hệ thống phân thành 2 nhóm chuyên biệt</div>
      </div>

      <div class="sidebar-scrollable">
        <!-- NHÓM 1: CHỨNG THƯ SỐ + EMAIL CÔNG VỤ -->
        <div class="task-group-card card-cks">
          <div class="group-header-box">
            <div class="group-header-title">
              <span>🔐 NHÓM CKS & MAIL CÔNG VỤ</span>
            </div>
            <span class="group-header-badge">5 tác vụ</span>
          </div>
          <ul class="sidebar-menu">
            <li data-task="1" onclick="switchTask(1)">
              <div class="nav-icon">🎯</div>
              <div class="nav-text">
                <span class="nav-title">VGCA (Word + Excel SSO)</span>
                <span class="nav-desc">Đối chiếu cấp CKS Word & Excel SSO</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
            <li data-task="2" onclick="switchTask(2)">
              <div class="nav-icon">📝</div>
              <div class="nav-text">
                <span class="nav-title">CKS (Công văn Mẫu 02)</span>
                <span class="nav-desc">Xuất Công văn Mẫu 02 (Có Excel SSO)</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
            <li data-task="3" onclick="switchTask(3)">
              <div class="nav-icon">📝</div>
              <div class="nav-text">
                <span class="nav-title">CKS (Mẫu CV02 chỉ đọc file Word)</span>
                <span class="nav-desc">Đọc toàn bộ file Word xuất Mẫu 02</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
            <li data-task="7" onclick="switchTask(7)">
              <div class="nav-icon">📄</div>
              <div class="nav-text">
                <span class="nav-title">VGCA (Từ File Word)</span>
                <span class="nav-desc">Đọc nhiều file Word không cần SSO</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
            <li data-task="9" onclick="switchTask(9)">
              <div class="nav-icon">📧</div>
              <div class="nav-text">
                <span class="nav-title">Cấp Email Công Vụ</span>
                <span class="nav-desc">Tổng hợp danh sách xin cấp Email</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
          </ul>
        </div>

        <!-- NHÓM 2: BÁO CÁO VÀ GIÁM ĐỊNH BHYT -->
        <div class="task-group-card card-bhyt">
          <div class="group-header-box">
            <div class="group-header-title">
              <span>📊 NHÓM BÁO CÁO & GIÁM ĐỊNH BHYT</span>
            </div>
            <span class="group-header-badge">4 tác vụ</span>
          </div>
          <ul class="sidebar-menu">
            <li data-task="5" onclick="switchTask(5)">
              <div class="nav-icon">🏥</div>
              <div class="nav-text">
                <span class="nav-title">Check Log MDINSIGHT</span>
                <span class="nav-desc">Lọc người dùng & xuất Excel A4</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
            <li data-task="11" onclick="switchTask(11)">
              <div class="nav-icon">🛠</div>
              <div class="nav-text">
                <span class="nav-title">BC Sửa CNTT</span>
                <span class="nav-desc">Tổng hợp ca tiếp nhận & báo cáo cán bộ</span>
              </div>
            </li>
            <li data-task="8" onclick="switchTask(8)">
              <div class="nav-icon">⚡</div>
              <div class="nav-text">
                <span class="nav-title">P.CNTT (Sheet Cố Định)</span>
                <span class="nav-desc">Tự động tải Google Sheets của viện</span>
              </div>
              <span class="nav-lock" title="Chức năng yêu cầu mật khẩu (8889)">🔒</span>
            </li>
            <li class="active" data-task="10" onclick="switchTask(10)">
              <div class="nav-icon">📊</div>
              <div class="nav-text">
                <span class="nav-title">Thống Kê Lỗi XML BHYT</span>
                <span class="nav-desc">Google Sheets • Lọc Khoa/Phòng & Lỗi</span>
              </div>
            </li>
          </ul>
        </div>
      </div>

      <div class="sidebar-footer">
        <b>Phòng CNTT • BVĐK Bắc Ninh số 2</b><br>
        Phiên bản Web v2.8 • Phân Quyền Bảo Mật
        <div id="auth-status-bar" style="margin-top: 8px; padding: 6px 10px; border-radius: 6px; font-size: 0.72rem; display: flex; align-items: center; justify-content: space-between; background: rgba(30, 41, 59, 0.7); border: 1px solid rgba(255, 255, 255, 0.08);">
          <span id="auth-status-text">🔒 7 tác vụ có khóa (8889)</span>
          <button type="button" id="btn-relock" onclick="lockApp()" style="display:none; background: #ef4444; color: #fff; border: none; border-radius: 4px; padding: 2px 7px; font-size: 0.7rem; cursor: pointer; font-weight:600;">Khóa lại</button>
        </div>
      </div>
    </div>

    <!-- CONTENT AREA -->
    <div class="content-area">
      <div class="breadcrumb-bar">
        <span class="badge" id="task-badge">LỖI XML BHYT</span>
        <span class="breadcrumb-title" id="task-breadcrumb">📊 Thống Kê & Lọc Lỗi XML BHYT (Google Sheets) ─ Đọc & Tra Cứu Lỗi Đẩy Cổng BHYT</span>
      </div>

      <!-- ==================== TÁC VỤ 1 ==================== -->
      <div class="page" id="page-task-1">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 1</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Tự động đọc thông tin từ các đơn đăng ký Word (.docx), ghép đôi với file tài khoản SSO Excel để lấy Email công vụ, tạo file <b>Ket_qua.xlsx</b> đúng 19 cột chuẩn VGCA (Ban Cơ yếu Chính phủ).</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên các file Word (.docx) đề nghị cấp CKS (có thể chọn nhiều file cùng lúc hoặc nén thành file .zip).</li>
                <li><b>Bước 2:</b> Tải lên file Excel SSO (.xlsx) chứa danh bạ email công vụ của đơn vị.</li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Bắt Đầu Đối Chiếu & Xuất Ket_qua.xlsx"</b>.</li>
                <li><b>Bước 4:</b> Trình duyệt sẽ tự động tải file <b>Ket_qua.xlsx</b> chuẩn 19 cột về máy tính.</li>
              </ol>
              <div class="highlight">
                💡 <b>Mẹo hay:</b> Hệ thống tự động so khớp theo cả CCCD 12 số lẫn Họ và tên không dấu để đảm bảo độ chính xác 100%.
              </div>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask1(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Chọn các file Word (.docx) hoặc file .zip:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t1-word').click()">
                  <div class="drop-icon">📄</div>
                  <div class="drop-text">Nhấp để chọn hoặc kéo thả các file Word vào đây</div>
                  <div class="drop-sub">Hỗ trợ chọn nhiều file .docx hoặc 1 file .zip chứa các đơn Word</div>
                  <input type="file" id="file-t1-word" multiple accept=".docx,.zip" onchange="previewFiles(this, 'chosen-t1-word')">
                </div>
                <div class="file-chosen-badge" id="chosen-t1-word"></div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Chọn file Excel SSO (.xlsx, .xls):</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t1-sso').click()">
                  <div class="drop-icon">📊</div>
                  <div class="drop-text">Nhấp để chọn file Excel SSO</div>
                  <div class="drop-sub">Chứa cột Họ tên, Email, CCCD</div>
                  <input type="file" id="file-t1-sso" accept=".xlsx,.xls" onchange="previewFiles(this, 'chosen-t1-sso')">
                </div>
                <div class="file-chosen-badge" id="chosen-t1-sso"></div>
              </div>

              <button type="submit" class="btn" id="btn-t1">
                <span class="spinner" id="spin-t1"></span>
                <span id="txt-t1">🚀 Bắt Đầu Đối Chiếu & Xuất Ket_qua.xlsx</span>
              </button>
              <div class="alert" id="alert-t1"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 2 ==================== -->
      <div class="page" id="page-task-2">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 2</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Đọc các file Word (.docx) đề nghị cấp CKS, ghép với file Excel SSO để xuất ra file Công văn <b>'Mẫu 02_CV cap CKS cho BVDK BN số 2'</b> (.docx) chuẩn thể thức văn bản hành chính.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên các file Word mẫu đơn CKS (.docx hoặc .zip).</li>
                <li><b>Bước 2:</b> Tải lên file Excel SSO chứa tài khoản email (tùy chọn).</li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Bắt Đầu Tổng Hợp CKS (Mẫu 02)"</b>.</li>
                <li><b>Bước 4:</b> Nhận file Công văn Word (.docx) hoàn chỉnh sẵn sàng in và trình ký.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask2(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Chọn các file Word (.docx) đơn CKS:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t2-word').click()">
                  <div class="drop-icon">📄</div>
                  <div class="drop-text">Nhấp để chọn các file Word (.docx)</div>
                  <input type="file" id="file-t2-word" multiple accept=".docx,.zip" onchange="previewFiles(this, 'chosen-t2-word')">
                </div>
                <div class="file-chosen-badge" id="chosen-t2-word"></div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Chọn file Excel SSO (.xlsx, tùy chọn):</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t2-sso').click()">
                  <div class="drop-icon">📊</div>
                  <div class="drop-text">Nhấp để chọn file Excel SSO (nếu có)</div>
                  <input type="file" id="file-t2-sso" accept=".xlsx,.xls" onchange="previewFiles(this, 'chosen-t2-sso')">
                </div>
                <div class="file-chosen-badge" id="chosen-t2-sso"></div>
              </div>

              <button type="submit" class="btn" id="btn-t2">
                <span class="spinner" id="spin-t2"></span>
                <span id="txt-t2">🚀 Bắt Đầu Tổng Hợp CKS (Mẫu 02)</span>
              </button>
              <div class="alert" id="alert-t2"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 3 ==================== -->
      <div class="page" id="page-task-3">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 3</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Tổng hợp danh sách xin cấp Chữ ký số (CKS) xuất file <b>Công văn Mẫu 02</b> (.docx) <b>chỉ từ file Word</b> đề nghị, độc lập hoàn toàn, <b>KHÔNG CẦN file Excel SSO</b>.</p>
              <p>Hệ thống tự động đọc toàn bộ thông tin từ file Word (Họ tên, Ngày sinh, CCCD, Ngày cấp, Nơi cấp, Email công vụ, Đơn vị, Chức vụ, SĐT...) và tự động điền vào Công văn Mẫu 02 chuẩn thể thức văn bản hành chính tương tự Tác vụ 2.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên một hoặc nhiều file Word (.docx) hoặc file nén (.zip) chứa các file Word đề nghị.</li>
                <li><b>Bước 2:</b> Nhấn nút <b>"🚀 Bắt Đầu Tổng Hợp CKS Mẫu 02 (Chỉ Từ File Word)"</b>.</li>
                <li><b>Bước 3:</b> Chọn vị trí lưu và nhận ngay file Công văn Word Mẫu 02 (.docx) hoàn chỉnh sẵn sàng in và trình ký.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask3(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Chọn các file Word (.docx) đề nghị cấp CKS:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t3-word').click()">
                  <div class="drop-icon">📄</div>
                  <div class="drop-text">Nhấp để chọn một hoặc nhiều file Word (.docx / .zip)</div>
                  <input type="file" id="file-t3-word" multiple accept=".docx,.zip" onchange="previewFiles(this, 'chosen-t3-word')">
                </div>
                <div class="file-chosen-badge" id="chosen-t3-word"></div>
              </div>

              <button type="submit" class="btn" id="btn-t3">
                <span class="spinner" id="spin-t3"></span>
                <span id="txt-t3">🚀 Bắt Đầu Tổng Hợp CKS (Mẫu CV02 - Chỉ File Word)</span>
              </button>
              <div class="alert" id="alert-t3"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 4 ==================== -->
      <div class="page" id="page-task-4">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 4</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Đọc file Excel báo cáo công tác sửa chữa của P.CNTT từ máy tính, trích xuất dữ liệu sửa chữa phần mềm & sửa lỗi khác theo từng Khoa/Phòng để xuất ra file <b>'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx'</b> chuẩn khổ A4 nằm ngang.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên file Excel báo cáo P.CNTT (.xlsx, .xlsm).</li>
                <li><b>Bước 2:</b> Chọn Sheet ngày cần xử lý từ danh sách tự động nạp.</li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Tổng Hợp Báo Cáo P.CNTT"</b>.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask4(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Tải lên file Excel báo cáo P.CNTT (.xlsx):</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t4').click()">
                  <div class="drop-icon">📊</div>
                  <div class="drop-text">Nhấp để chọn file Excel</div>
                  <input type="file" id="file-t4" accept=".xlsx,.xlsm" onchange="onTask4FileSelected(this)">
                </div>
                <div class="file-chosen-badge" id="chosen-t4"></div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Chọn Sheet ngày cần tổng hợp:</label>
                <select id="select-t4-sheet" class="form-input">
                  <option value="">-- Vui lòng chọn file Excel trước --</option>
                </select>
              </div>

              <button type="submit" class="btn" id="btn-t4">
                <span class="spinner" id="spin-t4"></span>
                <span id="txt-t4">🚀 Tổng Hợp Báo Cáo P.CNTT (A4 Ngang)</span>
              </button>
              <div class="alert" id="alert-t4"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 5 ==================== -->
      <div class="page" id="page-task-5">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Check Log MDINSIGHT</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Đọc file Excel báo cáo Check Log MDINSIGHT, lọc danh sách người dùng được phép, xuất ra file <b>'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx'</b> chuẩn khổ A4 dọc, tự động autofit cột và co giãn vừa 1 trang ngang.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên file Excel Check Log MDINSIGHT.</li>
                <li><b>Bước 2:</b> Chọn khoảng ngày cần tổng hợp:
                  <ul>
                    <li><b>Từ ngày 01 đến 14:</b> Lấy nửa đầu tháng.</li>
                    <li><b>Từ ngày 15 đến cuối tháng:</b> Lấy nửa cuối tháng.</li>
                    <li><b>Toàn bộ tháng (01 đến hết tháng):</b> Lấy trọn vẹn dữ liệu.</li>
                  </ul>
                </li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Xử Lý & Xuất Check Log MDINSIGHT"</b>.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask5(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Tải lên file Excel Check Log MDINSIGHT:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t5').click()">
                  <div class="drop-icon">🏥</div>
                  <div class="drop-text">Nhấp để chọn file Excel Check Log MDINSIGHT</div>
                  <input type="file" id="file-t5" accept=".xlsx,.xls" onchange="previewFiles(this, 'chosen-t5')">
                </div>
                <div class="file-chosen-badge" id="chosen-t5"></div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Chọn phạm vi ngày cần lấy báo cáo:</label>
                <div class="radio-group">
                  <label class="radio-card">
                    <input type="radio" name="t5_date_range" value="01-14" checked>
                    <div>
                      <div class="radio-title">Lựa chọn 1: Từ ngày 01 đến ngày 14</div>
                      <div class="radio-desc">Lấy dữ liệu giám định nửa đầu tháng</div>
                    </div>
                  </label>
                  <label class="radio-card">
                    <input type="radio" name="t5_date_range" value="15-31">
                    <div>
                      <div class="radio-title">Lựa chọn 2: Từ ngày 15 đến ngày cuối cùng có dữ liệu</div>
                      <div class="radio-desc">Lấy dữ liệu giám định nửa cuối tháng</div>
                    </div>
                  </label>
                  <label class="radio-card">
                    <input type="radio" name="t5_date_range" value="full">
                    <div>
                      <div class="radio-title">Lựa chọn 3: Lấy full từ ngày 01 đến ngày cuối cùng của tháng</div>
                      <div class="radio-desc">Tổng hợp toàn bộ các ngày có dữ liệu trong tháng</div>
                    </div>
                  </label>
                </div>
              </div>

              <button type="submit" class="btn" id="btn-t5">
                <span class="spinner" id="spin-t5"></span>
                <span id="txt-t5">🚀 Xử Lý & Xuất Check Log MDINSIGHT (A4 Dọc)</span>
              </button>
              <div class="alert" id="alert-t5"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 6 ==================== -->
      <div class="page" id="page-task-6">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 6</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Tự động đọc dữ liệu nhật ký sửa chữa từ Google Sheets (bất kỳ link chia sẻ nào), tổng hợp số lượt sửa chữa phần mềm theo từng Khoa/Phòng rồi xuất ra file <b>'TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx'</b>.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Dán đường link Google Sheets (phải được bật 'Đối tượng có link có thể xem').</li>
                <li><b>Bước 2:</b> Nhấn nút <b>"⚡ Nạp Danh Sách Sheet Ngày"</b> (hệ thống quét siêu tốc trong 1 giây).</li>
                <li><b>Bước 3:</b> Chọn Sheet ngày mong muốn.</li>
                <li><b>Bước 4:</b> Nhấn nút <b>"🚀 Tổng Hợp Từ Google Sheets"</b>.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask6(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Dán đường link Google Sheets:</label>
                <input type="text" id="url-t6" class="form-input" placeholder="https://docs.google.com/spreadsheets/d/..." value="">
                <div style="margin-top: 8px;">
                  <button type="button" class="btn btn-secondary" onclick="fetchGSheetsForTask(6)">⚡ Nạp Danh Sách Sheet Ngày</button>
                </div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Chọn Sheet ngày cần xử lý:</label>
                <select id="select-t6-sheet" class="form-input">
                  <option value="">-- Vui lòng nhấn 'Nạp Danh Sách Sheet Ngày' trước --</option>
                </select>
              </div>

              <button type="submit" class="btn" id="btn-t6">
                <span class="spinner" id="spin-t6"></span>
                <span id="txt-t6">🚀 Tổng Hợp Báo Cáo P.CNTT (A4 Ngang)</span>
              </button>
              <div class="alert" id="alert-t6"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 7 ==================== -->
      <div class="page" id="page-task-7">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 7</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Tổng hợp VGCA trực tiếp từ các file Word (.docx) <b>KHÔNG CẦN FILE SSO EXCEL</b>. Tự động đọc thông tin từ các đơn Word và xuất ra file <b>Ket_qua.xlsx</b> với đầy đủ 19 cột chuẩn VGCA.</p>
              <h3>💡 Ưu điểm nổi bật:</h3>
              <ul>
                <li>Độc lập hoàn toàn, không cần chuẩn bị file SSO Excel.</li>
                <li>Hỗ trợ tải lên nhiều file Word cùng lúc hoặc file nén .zip.</li>
                <li>Tự động nhận diện Email trong đơn hoặc tạo email gợi ý.</li>
              </ul>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Chọn các file Word (.docx) đề nghị cấp CKS.</li>
                <li><b>Bước 2:</b> Nhấn nút <b>"🚀 Bắt Đầu Tổng Hợp VGCA (Từ Word)"</b>.</li>
                <li><b>Bước 3:</b> Nhận file <b>Ket_qua.xlsx</b> 19 cột.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask7(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Chọn các file Word (.docx) hoặc file .zip:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t7-word').click()">
                  <div class="drop-icon">📄</div>
                  <div class="drop-text">Nhấp để chọn các file Word (.docx)</div>
                  <div class="drop-sub">Không cần file SSO – Xử lý tự động</div>
                  <input type="file" id="file-t7-word" multiple accept=".docx,.zip" onchange="previewFiles(this, 'chosen-t7-word')">
                </div>
                <div class="file-chosen-badge" id="chosen-t7-word"></div>
              </div>

              <button type="submit" class="btn" id="btn-t7">
                <span class="spinner" id="spin-t7"></span>
                <span id="txt-t7">🚀 Bắt Đầu Tổng Hợp VGCA (Xuất Ket_qua.xlsx)</span>
              </button>
              <div class="alert" id="alert-t7"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 8 ==================== -->
      <div class="page" id="page-task-8">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 8</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Tự động lấy dữ liệu từ đường link Google Sheets <b>cố định của Phòng CNTT Bệnh viện</b>, quét tất cả sheet ngày và xuất báo cáo chuẩn A4 ngang.</p>
              <h3>💡 Tiện ích đặc biệt:</h3>
              <p>Đường link nội bộ đã được cấu hình sẵn trong hệ thống, chỉ cần 1 chạm là nạp toàn bộ danh sách ngày mà không cần copy/paste link thủ công.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Nhấn nút <b>"⚡ Nạp Danh Sách Sheet Ngày"</b>.</li>
                <li><b>Bước 2:</b> Chọn sheet ngày cần tổng hợp (hệ thống tự chọn ngày mới nhất).</li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Tổng Hợp Báo Cáo P.CNTT"</b> để tải file Excel về.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask8(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Đường link Google Sheets của viện (Đã nạp sẵn):</label>
                <input type="text" id="url-t8" class="form-input" value="https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=800962139#gid=800962139" readonly style="background:#f1f5f9; color:#475569;">
                <div style="margin-top: 8px;">
                  <button type="button" class="btn btn-secondary" onclick="fetchGSheetsForTask(8)">⚡ Nạp Danh Sách Sheet Ngày</button>
                </div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Chọn Sheet ngày cần xử lý:</label>
                <select id="select-t8-sheet" class="form-input">
                  <option value="">-- Vui lòng nhấn 'Nạp Danh Sách Sheet Ngày' trước --</option>
                </select>
              </div>

              <button type="submit" class="btn" id="btn-t8">
                <span class="spinner" id="spin-t8"></span>
                <span id="txt-t8">🚀 Tổng Hợp Từ Google Sheets (A4 Ngang)</span>
              </button>
              <div class="alert" id="alert-t8"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 9 ==================== -->
      <div class="page" id="page-task-9">
        <div class="two-col-grid">
          <div class="card">
            <div class="card-header">
              <span class="icon">📖</span>
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 9</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Quét các file Word (.docx) để tổng hợp danh sách đề nghị cấp mới tài khoản Thư điện tử công vụ ra file Công văn và bảng Excel chuẩn Sở TT&TT Bắc Ninh.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên các file Word (.docx) xin cấp Email công vụ.</li>
                <li><b>Bước 2:</b> Chọn định dạng mong muốn: File Excel bảng danh sách hoặc File Word công văn.</li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Bắt Đầu Tổng Hợp Danh Sách Email"</b>.</li>
              </ol>
            </div>
          </div>

          <div class="card">
            <div class="card-header">
              <span class="icon">⚡</span>
              <h2>Thao Tác Thực Thi & Chọn Dữ Liệu</h2>
            </div>
            <form onsubmit="handleTask9(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Chọn các file Word (.docx) xin cấp Email:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t9-word').click()">
                  <div class="drop-icon">📧</div>
                  <div class="drop-text">Nhấp để chọn các file Word (.docx)</div>
                  <input type="file" id="file-t9-word" multiple accept=".docx,.zip" onchange="previewFiles(this, 'chosen-t9-word')">
                </div>
                <div class="file-chosen-badge" id="chosen-t9-word"></div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Định dạng xuất kết quả:</label>
                <div class="radio-group">
                  <label class="radio-card">
                    <input type="radio" name="email_format_9" value="excel" checked>
                    <div>
                      <div class="radio-title">Bảng tính Excel (.xlsx) chuẩn Sở TT&TT</div>
                      <div class="radio-desc">Gồm 7 cột: STT, Họ tên, Ngày sinh, SĐT, CCCD, Đơn vị, Chức vụ</div>
                    </div>
                  </label>
                  <label class="radio-card">
                    <input type="radio" name="email_format_9" value="word">
                    <div>
                      <div class="radio-title">Văn bản Công văn Word (.docx) hoàn chỉnh</div>
                      <div class="radio-desc">Có đầy đủ tiêu ngữ, căn cứ và phần ký tên</div>
                    </div>
                  </label>
                </div>
              </div>

              <button type="submit" class="btn" id="btn-t9">
                <span class="spinner" id="spin-t9"></span>
                <span id="txt-t9">🚀 Bắt Đầu Tổng Hợp Danh Sách Email</span>
              </button>
              <div class="alert" id="alert-t9"></div>
            </form>
          </div>
        </div>
      </div>

      <!-- ==================== TÁC VỤ 10: THỐNG KÊ LỖI XML BHYT ==================== -->
      <div class="page active" id="page-task-10">
        <div class="tracker-card">
          <div class="tracker-header-box">
            <div class="tracker-title-box">
              <h2>📊 Bảng Điều Khiển Thống Kê & Lọc Lỗi XML Đẩy Cổng Giám Định BHYT</h2>
              <p>Chỉ đọc dữ liệu từ Google Sheets: <a href="https://docs.google.com/spreadsheets/d/1VJ3_I4ehRDLSbm5w0CHAODtS8Y4fN3b-wsvo4p4ANVI/edit?gid=637670018#gid=637670018" target="_blank" style="color:#2563eb; text-decoration:underline; font-weight:600;">Xem file gốc trên Google Drive ↗</a></p>
            </div>
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
              <button type="button" class="btn btn-secondary" onclick="openAppsScriptModal()" id="btn-apps-script-cfg" style="background:linear-gradient(135deg, #059669 0%, #047857 100%); color:#ffffff; border:none; padding:8px 16px; font-size:0.85rem;" title="Hệ thống đã tự động kết nối Google Sheet của viện">
                ⚡ Đã Kết Nối Google Sheet (Sẵn Sàng ✅)
              </button>
              <button type="button" class="btn btn-secondary" onclick="loadTrackerSheets(true)" title="Cập nhật lại danh sách Sheet">
                🔄 Làm Mới Sheets
              </button>
              <button type="button" class="btn" onclick="exportTrackerExcel()" id="btn-tracker-export" style="width:auto; padding:8px 16px; font-size:0.85rem; background:linear-gradient(135deg, #059669 0%, #047857 100%);">
                📥 Xuất Báo Cáo Excel
              </button>
            </div>
          </div>

          <!-- THANH CÔNG CỤ LỌC -->
          <div class="tracker-filter-bar">
            <div class="tracker-filter-item">
              <label>📅 Chọn Sheet Ngày:</label>
              <select id="tracker-sheet-select" class="tracker-select" onchange="onTrackerSheetChange()">
                <option value="">⏳ Đang nạp danh sách sheet...</option>
              </select>
            </div>
            <div class="tracker-filter-item">
              <label>🏥 Lọc Theo Khoa:</label>
              <select id="tracker-khoa-select" class="tracker-select" onchange="applyTrackerFilters()">
                <option value="ALL">-- Tất Cả Khoa --</option>
              </select>
            </div>
            <div class="tracker-filter-item">
              <label>🩺 Lọc Phòng Khám:</label>
              <select id="tracker-phong-select" class="tracker-select" onchange="applyTrackerFilters()">
                <option value="ALL">-- Tất Cả Phòng --</option>
              </select>
            </div>
            <div class="tracker-filter-item">
              <label>📑 Lọc Bảng XML:</label>
              <select id="tracker-xml-select" class="tracker-select" onchange="applyTrackerFilters()">
                <option value="ALL">-- Tất Cả Bảng XML --</option>
              </select>
            </div>
            <div class="tracker-filter-item">
              <label>✅ Lọc Trạng Thái:</label>
              <select id="tracker-status-select" class="tracker-select" onchange="applyTrackerFilters()">
                <option value="ALL">-- Tất Cả Trạng Thái --</option>
                <option value="PENDING">⏳ Chưa Xử Lý</option>
                <option value="OK">✅ Đã Xử Lý (OK)</option>
              </select>
            </div>
            <div class="tracker-filter-item">
              <label>🔍 Tìm Kiếm (Mã HS, Lỗi):</label>
              <input type="text" id="tracker-search-input" class="tracker-input" placeholder="Nhập mã hồ sơ hoặc từ khóa lỗi..." oninput="applyTrackerFilters()">
            </div>
          </div>
        </div>

        <!-- 6 THẺ THỐNG KÊ TỔNG QUAN -->
        <div class="tracker-stats-grid">
          <div class="stat-card c-total">
            <div class="stat-card-title">⚠️ Tổng Số Lỗi</div>
            <div class="stat-card-value" id="stat-total-errors">0</div>
            <div class="stat-card-sub" id="stat-sub-errors">Đang tải...</div>
          </div>
          <div class="stat-card c-unique">
            <div class="stat-card-title">📑 Hồ Sơ Bị Lỗi</div>
            <div class="stat-card-value" id="stat-unique-hs">0</div>
            <div class="stat-card-sub">Hồ sơ khác nhau</div>
          </div>
          <div class="stat-card c-ok">
            <div class="stat-card-title">✅ Đã Xử Lý (OK)</div>
            <div class="stat-card-value" id="stat-total-ok">0</div>
            <div class="stat-card-sub" id="stat-sub-ok">0% hoàn thành</div>
          </div>
          <div class="stat-card c-khoa">
            <div class="stat-card-title">🏥 Khoa Vi Phạm</div>
            <div class="stat-card-value" id="stat-total-khoa">0</div>
            <div class="stat-card-sub">Có hồ sơ phát sinh lỗi</div>
          </div>
          <div class="stat-card c-phong">
            <div class="stat-card-title">🩺 Phòng Khám Lỗi</div>
            <div class="stat-card-value" id="stat-total-phong">0</div>
            <div class="stat-card-sub">Phòng ngoại trú ghi nhận</div>
          </div>
          <div class="stat-card c-xml">
            <div class="stat-card-title">📑 XML Lỗi Nhiều Nhất</div>
            <div class="stat-card-value" id="stat-top-xml">-</div>
            <div class="stat-card-sub" id="stat-top-xml-sub">Tỷ lệ lỗi cao nhất</div>
          </div>
        </div>

        <!-- 2 KHỐI THỐNG KÊ CHI TIẾT: KHOA & BẢNG XML -->
        <div class="tracker-cols-2">
          <!-- Cột 1: Xếp hạng Khoa -->
          <div class="tracker-card" style="margin-bottom:0;">
            <div class="tracker-title-box" style="margin-bottom:12px;">
              <h2>🏥 Xếp Hạng Lỗi Theo Khoa <span style="font-size:0.75rem; color:#64748b; font-weight:normal;">(Nhấp vào khoa để lọc nhanh)</span></h2>
            </div>
            <div class="ranking-box" id="tracker-khoa-ranking">
              <div style="padding:16px; text-align:center; color:#94a3b8;">Đang nạp dữ liệu thống kê...</div>
            </div>
          </div>

          <!-- Cột 2: Thống kê XML & Top lỗi -->
          <div class="tracker-card" style="margin-bottom:0;">
            <div class="tracker-title-box" style="margin-bottom:12px;">
              <h2>📑 Cơ Cấu Bảng XML & Nhóm Lỗi Thường Gặp</h2>
            </div>
            <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:14px;" id="tracker-xml-chips">
            </div>
            <div style="font-size:0.76rem; font-weight:700; color:#475569; text-transform:uppercase; margin-bottom:8px;">Top Nhóm Lỗi Phổ Biến:</div>
            <div class="ranking-box" style="max-height:220px;" id="tracker-top-errors-box">
              <div style="padding:16px; text-align:center; color:#94a3b8;">Đang nạp dữ liệu...</div>
            </div>
          </div>
        </div>

        <!-- BẢNG CHI TIẾT HỒ SƠ LỖI -->
        <div class="tracker-card" style="margin-top:20px;">
          <div class="tracker-header-box" style="margin-bottom:12px; padding-bottom:10px;">
            <div class="tracker-title-box">
              <h2>📋 Danh Sách Chi Tiết Các Bản Ghi Lỗi</h2>
              <p id="tracker-table-summary">Hiển thị 0 / 0 hồ sơ lỗi</p>
            </div>
            <div id="tracker-active-filter-badge" style="display:none; font-size:0.8rem; background:#eff6ff; color:#1d4ed8; padding:4px 10px; border-radius:6px; border:1px solid #bfdbfe;">
            </div>
          </div>

          <div class="tracker-table-wrap">
            <table class="tracker-table" id="tracker-records-table">
              <thead>
                <tr>
                  <th style="width: 45px; text-align:center;">STT</th>
                  <th style="width: 130px;">Mã Hồ Sơ</th>
                  <th style="width: 140px;">Ngày Giờ Đẩy</th>
                  <th style="width: 80px; text-align:center;">Bảng XML</th>
                  <th style="width: 190px;">Khoa / Bộ Phận</th>
                  <th style="width: 190px;">Phòng Khám</th>
                  <th>Nội Dung Chi Tiết Lỗi</th>
                  <th style="width: 135px; text-align:center;">Trạng Thái Xử Lý</th>
                </tr>
              </thead>
              <tbody id="tracker-records-tbody">
                <tr>
                  <td colspan="8" style="text-align:center; padding:30px; color:#94a3b8;">
                    ⏳ Đang tải dữ liệu hồ sơ lỗi...
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>


      <!-- MODAL BẢO VỆ MẬT KHẨU CHO CÁC CHỨC NĂNG -->
      
      <!-- ==========================================
           TÁC VỤ 11: BC SỬA CNTT (DASHBOARD TỔNG HỢP SHEET)
           ========================================== -->
      <div class="page" id="page-task-11">
        <!-- Top Bar -->
        <div class="cntt-header-bar">
          <div class="cntt-header-left">
            <button class="cntt-btn-back" onclick="switchTask(10)">
              ‹ Quay lại
            </button>
            <div class="cntt-header-title">
              🛠 BC Sửa CNTT
              <span class="cntt-sync-badge" id="cntt-sync-badge">
                <span style="color:#10b981">●</span> Đã đồng bộ (<span id="cntt-sync-clock">11:07</span>)
              </span>
            </div>
          </div>
          <div class="cntt-header-actions">
            <button class="cntt-btn-action cntt-btn-sync" onclick="refreshCnttLive()">
              🔄 Đồng Bộ
            </button>
            <button class="cntt-btn-action cntt-btn-excel" onclick="exportCnttExcel()">
              📄 Xuất Excel
            </button>
            <a href="https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=1894058788#gid=1894058788" target="_blank" class="cntt-btn-action cntt-btn-link">
              ↗ Mở Google Sheets
            </a>
          </div>
        </div>

        <!-- Sub Tabs Nav -->
        <div class="cntt-subnav">
          <div class="cntt-tab-item" onclick="switchCnttSubTab('diem-ca')">📋 Điểm Ca</div>
          <div class="cntt-tab-item" onclick="switchCnttSubTab('google-sheet')">📂 Google Sheet</div>
          <div class="cntt-tab-item" onclick="switchCnttSubTab('thong-ke')">📊 Thống Kê</div>
          <div class="cntt-tab-item" onclick="switchCnttSubTab('danh-sach')">📑 Danh Sách (<span id="cntt-tab-rec-count">16</span>)</div>
          <div class="cntt-tab-item active" onclick="switchCnttSubTab('tong-hop-sheet')">📑 Tổng Hợp Sheet</div>
        </div>

        <!-- TAB CONTENT: TỔNG HỢP SHEET -->
        <div id="cntt-subtab-tong-hop-sheet">
          <!-- Filter Card: Suggestions & Month Grid -->
          <div class="cntt-filter-card">
            <!-- Quick Suggestions -->
            <div class="cntt-quick-row">
              <div class="cntt-quick-left">
                <span style="font-size:12px; font-weight:700; color:#94a3b8; margin-right:4px;">⚡ Gợi ý:</span>
                <span class="cntt-quick-pill" onclick="selectCnttShortcut('today')">⚡ Hôm nay (29.10)</span>
                <span class="cntt-quick-pill" onclick="selectCnttShortcut('week')">⚡ Tuần này</span>
                <span class="cntt-quick-pill" onclick="selectCnttShortcut('m10')">⚡ Tháng 10</span>
                <span class="cntt-quick-pill" onclick="selectCnttShortcut('m9')">⚡ Tháng 9</span>
              </div>
              <div class="cntt-quick-right">
                <button class="cntt-btn-small" onclick="toggleSheetGridCollapse()">Thu gọn ▲</button>
                <button class="cntt-btn-small btn-success" onclick="selectAllCurrentMonthSheets()">✓ Chọn tất cả</button>
                <button class="cntt-btn-small btn-danger" onclick="clearSelectedSheets()">✗ Bỏ chọn</button>
              </div>
            </div>

            <!-- Month Selection Tabs -->
            <div class="cntt-month-tabs">
              <span class="cntt-month-pill active" onclick="switchCnttMonth('m10', this)">★ Tháng 10 (Hiện tại)</span>
              <span class="cntt-month-pill" onclick="switchCnttMonth('m9', this)">Tháng 9</span>
              <span class="cntt-month-pill" onclick="switchCnttMonth('m8', this)">Tháng 8</span>
              <span class="cntt-month-pill" onclick="switchCnttMonth('m7', this)">Tháng 7</span>
              <span class="cntt-month-pill" onclick="switchCnttMonth('all', this)">Tất cả (67 sheet)</span>
            </div>

            <!-- Sheets Checkbox Grid -->
            <div class="cntt-sheets-grid" id="cntt-sheets-grid">
              <!-- Dynamically populated -->
            </div>
          </div>

          <!-- Action & Search Bar -->
          <div class="cntt-action-bar">
            <div class="cntt-action-inputs">
              <div style="display:flex; align-items:center; gap:8px;">
                <label style="font-size:12px; font-weight:700; color:#94a3b8;">Lọc cán bộ:</label>
                <select class="cntt-select-cb" id="cntt-filter-staff" onchange="renderCnttDashboard()">
                  <option value="">-- Tất cả cán bộ (7) --</option>
                  <option value="Chu Thị Dương">Chu Thị Dương</option>
                  <option value="Nguyễn Lê Huyền Ly">Nguyễn Lê Huyền Ly</option>
                  <option value="Đỗ Văn Hiệp">Đỗ Văn Hiệp</option>
                  <option value="Nguyễn Thị Quyên">Nguyễn Thị Quyên</option>
                  <option value="Nguyễn Thị Huyền">Nguyễn Thị Huyền</option>
                  <option value="Vũ Phương">Vũ Phương</option>
                  <option value="Chưa gán">Chưa gán</option>
                </select>
              </div>
              <input type="text" id="cntt-search-keyword" class="cntt-search-input" placeholder="🔍 Tìm kiếm công việc, khoa, nội dung sửa..." oninput="renderCnttDashboard()">
            </div>
            <div class="cntt-success-pill" id="cntt-status-msg">
              ✓ Đã tổng hợp thành công 128 ca từ 3 sheet!
            </div>
          </div>

          <!-- 4 KPI Metrics Grid -->
          <div class="cntt-kpi-grid">
            <div class="cntt-kpi-card kpi-blue">
              <div class="cntt-kpi-label">📖 SHEET ĐÃ CHỌN</div>
              <div class="cntt-kpi-value" id="cntt-kpi-sheets">3</div>
              <div class="cntt-kpi-sub" id="cntt-kpi-sheets-sub">29.10, 28.10, 27.10</div>
            </div>
            <div class="cntt-kpi-card kpi-purple">
              <div class="cntt-kpi-label">🛠 TỔNG SỐ CA TIẾP NHẬN</div>
              <div class="cntt-kpi-value" id="cntt-kpi-cases">128</div>
              <div class="cntt-kpi-sub" id="cntt-kpi-cases-sub">Bao gồm 36 khoa/phòng</div>
            </div>
            <div class="cntt-kpi-card kpi-emerald">
              <div class="cntt-kpi-label">👥 CÁN BỘ THỰC HIỆN</div>
              <div class="cntt-kpi-value" id="cntt-kpi-staff">7</div>
              <div class="cntt-kpi-sub">Cán bộ tham gia xử lý</div>
            </div>
            <div class="cntt-kpi-card kpi-amber">
              <div class="cntt-kpi-label">📝 NỘI DUNG SỬA CHỮA KHÁC</div>
              <div class="cntt-kpi-value" id="cntt-kpi-other">51</div>
              <div class="cntt-kpi-sub">Nội dung ghi chi tiết</div>
            </div>
          </div>

          <!-- Staff Cards Container -->
          <div id="cntt-staff-cards-list">
            <!-- Dynamically populated -->
          </div>
        </div>

        <!-- TAB CONTENT: DANH SÁCH -->
        <div id="cntt-subtab-danh-sach" style="display:none;">
          <div class="card" style="padding:16px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
              <h3 style="color:#ffffff; margin:0;">📑 Danh Sách Chi Tiết Công Việc (<span id="cntt-table-total">128</span> ca)</h3>
              <input type="text" id="cntt-table-search" placeholder="🔍 Lọc trong bảng..." oninput="filterCnttTable()" style="padding:6px 12px; background:#1e293b; border:1px solid #475569; color:white; border-radius:6px; font-size:12px;">
            </div>
            <div style="max-height:550px; overflow-y:auto;">
              <table class="error-table" id="cntt-raw-table">
                <thead>
                  <tr>
                    <th>STT</th>
                    <th>Sheet</th>
                    <th>Khoa / Phòng</th>
                    <th>Số HS</th>
                    <th>Cán Bộ Xử Lý</th>
                    <th>Người Y/C</th>
                    <th>Nội Dung Sửa Chữa Khác</th>
                    <th>Các Mục Sửa Khác</th>
                  </tr>
                </thead>
                <tbody id="cntt-table-body">
                  <!-- Populated via JS -->
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- TAB CONTENT: THỐNG KÊ -->
        <div id="cntt-subtab-thong-ke" style="display:none;">
          <div class="tracker-charts-grid">
            <div class="tracker-chart-card">
              <div class="tracker-chart-header">
                <h3>🏥 Xếp Hạng Khoa/Phòng Yêu Cầu Nhiều Nhất</h3>
              </div>
              <div id="cntt-stat-dept-list" style="padding:10px 0;"></div>
            </div>
            <div class="tracker-chart-card">
              <div class="tracker-chart-header">
                <h3>📊 Cơ Cấu Các Nhóm Sửa Chữa</h3>
              </div>
              <div id="cntt-stat-cat-list" style="padding:10px 0;"></div>
            </div>
          </div>
        </div>

        <!-- TAB CONTENT: GOOGLE SHEET -->
        <div id="cntt-subtab-google-sheet" style="display:none;">
          <div class="card" style="padding:16px; text-align:center;">
            <h3 style="color:#ffffff; margin-bottom:10px;">📂 Trực Tiếp Xem & Thao Tác Trên Google Sheets</h3>
            <p style="color:#94a3b8; font-size:13px; margin-bottom:14px;">Bấm nút bên dưới để mở file Google Sheets gốc của P.CNTT trong tab mới:</p>
            <a href="https://docs.google.com/spreadsheets/d/1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA/edit?gid=1894058788#gid=1894058788" target="_blank" class="btn btn-primary" style="display:inline-flex; align-items:center; gap:8px;">
              ↗ Mở File Google Sheets Gốc (P.CNTT)
            </a>
          </div>
        </div>

        <!-- TAB CONTENT: ĐIỂM CA -->
        <div id="cntt-subtab-diem-ca" style="display:none;">
          <div class="card" style="padding:16px;">
            <h3 style="color:#ffffff; margin-bottom:12px;">📋 Tổng Hợp Điểm Ca Theo Ngày</h3>
            <div id="cntt-diem-ca-content"></div>
          </div>
        </div>

        <!-- Floating Toast -->
        <div class="cntt-toast-popup" id="cntt-toast">
          <div class="cntt-toast-icon">✓</div>
          <div>
            <div class="cntt-toast-title" id="cntt-toast-title">THÀNH CÔNG</div>
            <div class="cntt-toast-msg" id="cntt-toast-msg">Đã tổng hợp 128 ca công việc từ 3 sheet!</div>
          </div>
        </div>
      </div>

      <div id="auth-modal" style="display:none; position:fixed; inset:0; background:rgba(15,23,42,0.75); backdrop-filter:blur(5px); z-index:999999; align-items:center; justify-content:center; padding:20px;">
        <div style="background:#ffffff; border-radius:14px; max-width:440px; width:100%; padding:26px; box-shadow:0 25px 50px -12px rgba(0,0,0,0.35); border:1px solid #e2e8f0;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
            <div style="display:flex; align-items:center; gap:10px;">
              <div style="width:38px; height:38px; border-radius:10px; background:#eff6ff; color:#2563eb; display:flex; align-items:center; justify-content:center; font-size:1.25rem;">🔒</div>
              <h3 style="margin:0; font-size:1.1rem; color:#0f172a; font-weight:700;">Khóa Bảo Vệ Chức Năng</h3>
            </div>
            <button type="button" onclick="closeAuthModal()" style="background:none; border:none; font-size:1.4rem; cursor:pointer; color:#94a3b8; line-height:1;">&times;</button>
          </div>
          
          <p style="font-size:0.86rem; color:#475569; line-height:1.5; margin-top:0; margin-bottom:16px;">
            Chức năng <b id="auth-task-name" style="color:#1d4ed8;"></b> được phân quyền bảo mật. Vui lòng nhập mật khẩu để mở khóa truy cập:
          </p>

          <div style="margin-bottom:14px; position:relative;">
            <input type="password" id="auth-password-input" placeholder="Nhập mật khẩu truy cập (mã PIN)..." 
                   style="width:100%; padding:10px 42px 10px 14px; border:1.5px solid #cbd5e1; border-radius:8px; font-size:0.95rem; outline:none; box-sizing:border-box; transition:border-color 0.2s;"
                   onkeydown="if(event.key==='Enter') submitAuthPassword()"
                   oninput="document.getElementById('auth-error-msg').style.display='none'; this.style.borderColor='#3b82f6';">
            <button type="button" onclick="toggleAuthPasswordVisibility()" style="position:absolute; right:10px; top:50%; transform:translateY(-50%); background:none; border:none; cursor:pointer; font-size:1rem; color:#64748b;" title="Ẩn/Hiện mật khẩu">👁️</button>
          </div>

          <div id="auth-error-msg" style="display:none; color:#dc2626; font-size:0.82rem; font-weight:600; margin-bottom:14px; padding:8px 12px; background:#fef2f2; border-radius:6px; border:1px solid #fee2e2;"></div>

          <div style="display:flex; justify-content:flex-end; gap:10px;">
            <button type="button" class="btn btn-secondary" onclick="closeAuthModal()" style="width:auto; padding:9px 18px; font-size:0.86rem;">Hủy Bỏ</button>
            <button type="button" class="btn" onclick="submitAuthPassword()" id="btn-submit-auth" style="width:auto; padding:9px 22px; font-size:0.86rem; background:linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);">
              🔓 Mở Khóa
            </button>
          </div>
        </div>
      </div>

      <!-- MODAL CẤU HÌNH APPS SCRIPT -->
      <div id="apps-script-modal" style="display:none; position:fixed; inset:0; background:rgba(15,23,42,0.65); backdrop-filter:blur(4px); z-index:9999; align-items:center; justify-content:center; padding:20px;">
        <div style="background:#ffffff; border-radius:14px; max-width:680px; width:100%; max-height:90vh; overflow-y:auto; padding:24px; box-shadow:0 20px 25px -5px rgba(0,0,0,0.2);">
          <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #e2e8f0; padding-bottom:14px; margin-bottom:16px;">
            <h3 style="margin:0; font-size:1.15rem; color:#1e293b; display:flex; align-items:center; gap:8px;">
              <span>⚡ Cấu Hình Tự Động Cập Nhật Vào Google Sheet</span>
            </h3>
            <button type="button" onclick="closeAppsScriptModal()" style="background:none; border:none; font-size:1.4rem; cursor:pointer; color:#64748b; line-height:1;">&times;</button>
          </div>
          
          <div style="font-size:0.88rem; color:#475569; line-height:1.6; margin-bottom:16px;">
            <p style="margin-top:0;">Để phần mềm có quyền ghi <b>"OK"</b> vào cột <b>Trạng thái xử lý</b> trực tuyến trên Google Sheet khi bạn tích chọn, vui lòng thiết lập Apps Script một lần duy nhất theo các bước sau:</p>
            
            <ol style="padding-left:20px; margin-bottom:14px;">
              <li style="margin-bottom:6px;">Mở file: <a href="https://docs.google.com/spreadsheets/d/1VJ3_I4ehRDLSbm5w0CHAODtS8Y4fN3b-wsvo4p4ANVI/edit?gid=637670018#gid=637670018" target="_blank" style="color:#2563eb; font-weight:600; text-decoration:underline;">Google Sheets Lỗi XML BHYT ↗</a></li>
              <li style="margin-bottom:6px;">Vào menu <b>Tiện ích mở rộng (Extensions)</b> ➔ <b>Apps Script</b>.</li>
              <li style="margin-bottom:6px;">Xóa hết mã cũ trong cửa sổ soạn thảo, bấm nút <b>📋 Sao Chép Mã</b> bên dưới rồi <b>Dán (Ctrl + V)</b> vào và bấm <b>Lưu (Ctrl + S)</b>.</li>
              <li style="margin-bottom:6px;">Bấm nút <b>Triển khai (Deploy)</b> ➔ <b>Triển khai mới (New deployment)</b> ➔ Bấm bánh răng ⚙️ chọn <b>Ứng dụng web (Web app)</b>:
                <ul style="margin-top:4px; font-size:0.82rem; color:#b91c1c; font-weight:600;">
                  <li>Thực thi dưới dạng (Execute as): <b>Tôi (Me)</b></li>
                  <li>Người có quyền truy cập (Who has access): <span style="background:#fef08a; color:#854d0e; padding:1px 4px; border-radius:3px;">Bắt buộc chọn: Bất kỳ ai (Anyone)</span></li>
                </ul>
              </li>
              <li>Bấm <b>Triển khai (Deploy)</b> (cấp quyền Google nếu có yêu cầu) ➔ Sao chép <b>URL ứng dụng web</b> (link có đuôi <code>/exec</code>) dán vào ô bên dưới.</li>
            </ol>

            <div style="position:relative; margin-bottom:16px;">
              <div style="display:flex; justify-content:space-between; align-items:center; background:#1e293b; color:#94a3b8; padding:6px 12px; border-top-left-radius:8px; border-top-right-radius:8px; font-size:0.78rem;">
                <span>MÃ GOOGLE APPS SCRIPT (Chuẩn tự động điền OK vào Cột M)</span>
                <button type="button" onclick="copyAppsScriptCode()" id="btn-copy-script" style="background:#3b82f6; color:#fff; border:none; border-radius:4px; padding:3px 8px; cursor:pointer; font-size:0.75rem;">📋 Sao Chép Mã</button>
              </div>
              <textarea id="apps-script-code-box" readonly style="width:100%; height:160px; font-family:monospace; font-size:0.78rem; background:#0f172a; color:#f8fafc; border:none; border-bottom-left-radius:8px; border-bottom-right-radius:8px; padding:10px; resize:vertical; outline:none; box-sizing:border-box;">function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }
function handleRequest(e) {
  try {
    var p = e.parameter || {};
    if (p.action === "ping") {
      return ContentService.createTextOutput(JSON.stringify({success:true, ping:"pong"})).setMimeType(ContentService.MimeType.JSON);
    }
    var sheetName = (p.sheet || "").trim();
    var hs = (p.hs || "").trim();
    var row = parseInt(p.row || 0);
    var status = p.status !== undefined ? p.status : "OK";
    var SPREADSHEET_ID = "1VJ3_I4ehRDLSbm5w0CHAODtS8Y4fN3b-wsvo4p4ANVI";
    var ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(_) {}
    if (!ss) {
      try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); } catch(err) {
        return ContentService.createTextOutput(JSON.stringify({success:false, error:"Không thể mở Spreadsheet: " + err.toString()})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    var sheet = null;
    if (sheetName) {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet && sheetName.indexOf("Ngày ") !== -1) sheet = ss.getSheetByName(sheetName.replace("Ngày ", "").trim());
      if (!sheet && sheetName.indexOf("Ngày") === -1) sheet = ss.getSheetByName("Ngày " + sheetName);
    }
    if (!sheet) sheet = ss.getActiveSheet();
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({success:false, error:"Sheet not found: " + sheetName})).setMimeType(ContentService.MimeType.JSON);

    var lastRow = sheet.getLastRow();
    var targetRow = -1;
    if (row >= 2 && row <= lastRow) {
      if (hs) {
        var cVal = String(sheet.getRange(row, 2).getValue()).split(".")[0].trim();
        if (cVal === hs) targetRow = row;
      } else targetRow = row;
    }
    if (targetRow <= 0 && hs) {
      var hsVals = sheet.getRange(1, 2, Math.max(lastRow, 2), 1).getValues();
      for (var i = 1; i < hsVals.length; i++) {
        if (String(hsVals[i][0]).split(".")[0].trim() === hs) { targetRow = i + 1; break; }
      }
    }
    if (targetRow <= 1) return ContentService.createTextOutput(JSON.stringify({success:false, error:"Không tìm thấy dòng cho HS: " + hs})).setMimeType(ContentService.MimeType.JSON);

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 20)).getValues()[0];
    var colIdx = -1;
    for (var j = 0; j < headers.length; j++) {
      var h = String(headers[j]).trim().toLowerCase();
      if (h.indexOf("trạng thái") !== -1 || h.indexOf("kết quả xử lý") !== -1) { colIdx = j + 1; break; }
    }
    if (colIdx === -1) { colIdx = 13; sheet.getRange(1, colIdx).setValue("Trạng thái xử lý"); }
    sheet.getRange(targetRow, colIdx).setValue(status);
    return ContentService.createTextOutput(JSON.stringify({success:true, sheet:sheet.getName(), row:targetRow, col:colIdx, status:status, hs:hs})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}</textarea>
            </div>

            <div style="margin-bottom:12px;">
              <label style="display:block; font-size:0.8rem; font-weight:700; color:#334155; margin-bottom:6px;">Dán Link URL Web App (kết thúc bằng /exec):</label>
              <div style="display:flex; gap:8px;">
                <input type="text" id="apps-script-url-input" class="tracker-input" placeholder="https://script.google.com/macros/s/AKfycb.../exec" style="flex:1; font-family:monospace; font-size:0.82rem;">
                <button type="button" class="btn btn-secondary" onclick="testAppsScriptUrl()" id="btn-test-script" style="width:auto; padding:8px 14px; font-size:0.82rem; white-space:nowrap; background:#0ea5e9; color:#fff; border:none;" title="Kiểm tra xem link Web App đã hoạt động và có quyền ghi Google Sheet hay chưa">
                  🔍 Thử Kết Nối
                </button>
              </div>
              <div id="test-script-result" style="display:none; margin-top:8px; padding:10px; border-radius:6px; font-size:0.82rem; line-height:1.4;"></div>
              <p style="font-size:0.75rem; color:#64748b; margin-top:4px; margin-bottom:0;">
                ⚠️ <i>Lưu ý: Không dán link trình duyệt có đuôi <b>/edit</b> hoặc link file Google Sheets! Link chuẩn phải có đuôi <b>/exec</b>.</i>
              </p>
            </div>

            <div style="display:flex; justify-content:flex-end; gap:10px;">
              <button type="button" class="btn btn-secondary" onclick="closeAppsScriptModal()" style="width:auto; padding:8px 16px;">Đóng</button>
              <button type="button" class="btn" onclick="saveAppsScriptUrl()" style="width:auto; padding:8px 20px; background:linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);">💾 Lưu Cấu Hình</button>
            </div>
          </div>
        </div>
      </div>
        </div>
      </div>

    </div>
  </div>

  <script>
    const taskMeta = {
      1: { badge: 'VGCA', title: '🎯 VGCA (Word + Excel SSO)  ─  Đối chiếu cấp CKS Word & Excel SSO' },
      2: { badge: 'MẪU 02', title: '📝 CKS (Công văn Mẫu 02)  ─  Xuất Công văn Word Mẫu 02 chuẩn' },
      3: { badge: 'MẪU 02 (WORD)', title: '📝 CKS (Mẫu CV02 chỉ đọc file Word)  ─  Đọc toàn bộ file Word xuất Mẫu 02' },
      5: { badge: 'CHECK LOG MDINSIGHT', title: '🏥 Check Log MDINSIGHT  ─  Lọc người dùng & xuất Excel A4 dọc' },
      7: { badge: 'VGCA (WORD)', title: '📄 VGCA (Từ File Word)  ─  Đọc nhiều file Word không cần SSO' },
      8: { badge: 'P.CNTT', title: '⚡ P.CNTT (Sheet Cố Định)  ─  Tự động tải Google Sheets của viện' },
      9: { badge: 'EMAIL CÔNG VỤ', title: '📧 Cấp Email Công Vụ  ─  Tổng hợp danh sách xin cấp Email' },
      10: { badge: 'LỖI XML BHYT', title: '📊 Thống Kê & Lọc Lỗi XML BHYT (Google Sheets) ─ Đọc & Tra Cứu Lỗi Đẩy Cổng BHYT' },
      11: { badge: 'BC SỬA CNTT', title: '🛠 BC Sửa CNTT ─ Tổng Hợp Sheet & Báo Cáo Cán Bộ P.CNTT' }
    };

    const UNLOCK_PASSWORD = "8889";
    let pendingTaskId = null;

    function isAppUnlocked() {
      return sessionStorage.getItem('app_unlocked_8889') === 'true' || localStorage.getItem('app_unlocked_8889') === 'true';
    }

    function updateAuthUI() {
      const unlocked = isAppUnlocked();
      document.querySelectorAll('.sidebar-menu li[data-task]').forEach(li => {
        const id = parseInt(li.getAttribute('data-task'));
        const lockEl = li.querySelector('.nav-lock');
        if (lockEl) {
          if (id === 10) {
            lockEl.style.display = 'none';
          } else {
            lockEl.textContent = unlocked ? '🔓' : '🔒';
            lockEl.title = unlocked ? 'Đã mở khóa' : 'Cần mật khẩu (8889)';
          }
        }
      });
      const statusText = document.getElementById('auth-status-text');
      const btnRelock = document.getElementById('btn-relock');
      if (statusText) {
        statusText.innerHTML = unlocked ? '🔓 Đã mở khóa đầy đủ' : '🔒 7 tác vụ có khóa (8889)';
        statusText.style.color = unlocked ? '#34d399' : '#94a3b8';
      }
      if (btnRelock) {
        btnRelock.style.display = unlocked ? 'inline-block' : 'none';
      }
    }

    function switchTask(taskId) {
      taskId = parseInt(taskId);
      // Task 10 (Thống Kê Lỗi XML) và Task 11 (BC Sửa CNTT) là 2 dashboard Google Sheets mở trực tiếp
      if (taskId === 10 || taskId === 11) {
        executeSwitchTask(taskId);
        return;
      }
      // Các tác vụ quản trị, xuất file (1, 2, 3, 5, 7, 8, 9) yêu cầu mật khẩu 8889 mỗi lần chọn
      pendingTaskId = taskId;
      openAuthModal(taskId);
    }

    function executeSwitchTask(taskId) {
      document.querySelectorAll('.sidebar-menu li[data-task]').forEach((el) => {
        const id = parseInt(el.getAttribute('data-task'));
        el.classList.toggle('active', id === taskId);
      });
      document.querySelectorAll('.page').forEach(el => el.classList.remove('active'));
      const activePage = document.getElementById('page-task-' + taskId);
      if (activePage) activePage.classList.add('active');

      const meta = taskMeta[taskId];
      if (meta) {
        document.getElementById('task-badge').textContent = meta.badge;
        document.getElementById('task-breadcrumb').textContent = meta.title;
      }
      if (taskId === 10) {
        initErrorTracker();
      }
      if (taskId === 11) {
        initCnttDashboard();
      }
    }

    function openAuthModal(taskId) {
      const modal = document.getElementById('auth-modal');
      const input = document.getElementById('auth-password-input');
      const nameEl = document.getElementById('auth-task-name');
      const errEl = document.getElementById('auth-error-msg');
      if (nameEl) {
        const meta = taskMeta[taskId];
        nameEl.textContent = meta ? meta.badge : ('Tác vụ ' + taskId);
      }
      if (errEl) errEl.style.display = 'none';
      if (input) {
        input.value = '';
        input.style.borderColor = '#cbd5e1';
      }
      if (modal) modal.style.display = 'flex';
      setTimeout(() => { if (input) input.focus(); }, 120);
    }

    function closeAuthModal() {
      const modal = document.getElementById('auth-modal');
      if (modal) modal.style.display = 'none';
      pendingTaskId = null;
    }

    function toggleAuthPasswordVisibility() {
      const input = document.getElementById('auth-password-input');
      if (!input) return;
      input.type = (input.type === 'password') ? 'text' : 'password';
    }

    function submitAuthPassword() {
      const input = document.getElementById('auth-password-input');
      const errEl = document.getElementById('auth-error-msg');
      if (!input) return;
      const val = input.value.trim();
      if (val === UNLOCK_PASSWORD) {
        sessionStorage.setItem('app_unlocked_8889', 'true');
        localStorage.setItem('app_unlocked_8889', 'true');
        document.cookie = "app_auth_pass=" + UNLOCK_PASSWORD + "; path=/; max-age=864000; SameSite=Lax";
        updateAuthUI();
        closeAuthModal();
        const target = pendingTaskId || 1;
        pendingTaskId = null;
        executeSwitchTask(target);
        if (target === 11) {
          initCnttDashboard();
        }
        showMiniToast('🔓 Mở khóa thành công! Bạn có thể sử dụng tất cả các chức năng.');
      } else {
        if (errEl) {
          errEl.textContent = '❌ Mật khẩu không chính xác! Vui lòng thử lại.';
          errEl.style.display = 'block';
        }
        input.classList.remove('auth-shake');
        void input.offsetWidth;
        input.classList.add('auth-shake');
        input.style.borderColor = '#ef4444';
        input.select();
      }
    }

    function lockApp() {
      sessionStorage.removeItem('app_unlocked_8889');
      localStorage.removeItem('app_unlocked_8889');
      document.cookie = "app_auth_pass=; path=/; max-age=0; SameSite=Lax";
      updateAuthUI();
      executeSwitchTask(10);
      showMiniToast('🔒 Đã khóa lại các chức năng. Chuyển về màn hình Thống Kê Lỗi XML BHYT.');
    }

    function previewFiles(input, badgeId) {
      const badge = document.getElementById(badgeId);
      if (!badge) return;
      if (input.files && input.files.length > 0) {
        if (input.files.length === 1) {
          badge.innerHTML = `📄 <b>${input.files[0].name}</b> (${Math.round(input.files[0].size/1024)} KB)`;
        } else {
          badge.innerHTML = `📁 <b>Đã chọn ${input.files.length} files</b>`;
        }
        badge.style.display = 'block';
      } else {
        badge.style.display = 'none';
      }
    }

    function showAlert(id, type, msg) {
      const el = document.getElementById(id);
      if (!el) return;
      el.className = 'alert ' + type;
      el.textContent = msg;
      el.style.display = 'flex';
    }

    function hideAlert(id) {
      const el = document.getElementById(id);
      if (el) el.style.display = 'none';
    }

    function setLoading(prefix, isLoading, normalTxt) {
      const btn = document.getElementById('btn-' + prefix);
      const spin = document.getElementById('spin-' + prefix);
      const txt = document.getElementById('txt-' + prefix);
      if (btn) btn.disabled = isLoading;
      if (spin) spin.style.display = isLoading ? 'inline-block' : 'none';
      if (txt) txt.textContent = isLoading ? 'Đang Xử Lý Dữ Liệu...' : normalTxt;
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

    async function promptSaveFile(suggestedName) {
      if ('showSaveFilePicker' in window) {
        try {
          const parts = suggestedName.split('.');
          const ext = parts.pop().toLowerCase();
          let mimeType = 'application/octet-stream';
          let description = 'Tệp tin';
          if (ext === 'xlsx') {
            mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            description = 'Tệp Excel (*.xlsx)';
          } else if (ext === 'docx') {
            mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            description = 'Tệp Word (*.docx)';
          }
          const handle = await window.showSaveFilePicker({
            suggestedName: suggestedName,
            types: [{
              description: description,
              accept: { [mimeType]: ['.' + ext] }
            }]
          });
          return { handle: handle, canceled: false };
        } catch (err) {
          if (err.name === 'AbortError') {
            return { handle: null, canceled: true };
          }
          console.warn('showSaveFilePicker không khả dụng hoặc bị từ chối:', err);
        }
      }
      return { handle: null, canceled: false };
    }

    async function writeBlobToHandleOrDownload(blob, fileHandle, defaultFilename) {
      if (fileHandle) {
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        return fileHandle.name || defaultFilename;
      } else {
        downloadBlob(blob, defaultFilename);
        return defaultFilename;
      }
    }

    // Nạp Sheet cho Task 4 từ file Excel
    async function onTask4FileSelected(input) {
      previewFiles(input, 'chosen-t4');
      if (input.files && input.files[0]) {
        const formData = new FormData();
        formData.append('file', input.files[0]);
        await loadSheetsIntoSelect(formData, 'select-t4-sheet', 'alert-t4');
      }
    }

    // Nạp Sheet cho Task 6 & 8 từ Google Sheets URL
    async function fetchGSheetsForTask(taskId) {
      const inputId = taskId === 8 ? 'url-t8' : 'url-t6';
      const selectId = taskId === 8 ? 'select-t8-sheet' : 'select-t6-sheet';
      const alertId = taskId === 8 ? 'alert-t8' : 'alert-t6';
      const url = document.getElementById(inputId).value.trim();

      if (!url) {
        showAlert(alertId, 'error', '⚠️ Vui lòng nhập link Google Sheets!');
        return;
      }
      const formData = new FormData();
      formData.append('url', url);
      await loadSheetsIntoSelect(formData, selectId, alertId);
    }

    async function loadSheetsIntoSelect(formData, selectId, alertId) {
      const sel = document.getElementById(selectId);
      sel.innerHTML = '<option value="">⏳ Đang nạp danh sách các sheet ngày...</option>';
      hideAlert(alertId);
      try {
        const res = await fetch('/api/get-sheets?action=get-sheets', {
          method: 'POST',
          body: formData
        });
        let data;
        try {
          data = await res.json();
        } catch (jsonErr) {
          throw new Error('Máy chủ phản hồi không đúng định dạng (Mã HTTP ' + res.status + '). Vui lòng thử lại!');
        }
        if (!res.ok || data.error) {
          throw new Error(data.error || ('Lỗi tải sheet: ' + res.status));
        }

        sel.innerHTML = '';
        if (!data.sheets || data.sheets.length === 0) {
          throw new Error('Không tìm thấy sheet nào trong tài liệu này!');
        }
        data.sheets.forEach((s, idx) => {
          const opt = document.createElement('option');
          opt.value = s;
          opt.textContent = `📅 Sheet: ${s}`;
          if (idx === data.sheets.length - 1) {
            opt.selected = true;
          }
          sel.appendChild(opt);
        });
        showAlert(alertId, 'success', `✅ Đã nạp thành công ${data.sheets.length} sheet ngày! Hãy chọn sheet cần xuất báo cáo.`);
      } catch (err) {
        sel.innerHTML = '<option value="">❌ Không đọc được sheet. Vui lòng kiểm tra lại!</option>';
        showAlert(alertId, 'error', 'Lỗi: ' + err.message);
      }
    }

    // XỬ LÝ TÁC VỤ 1
    async function handleTask1(e) {
      e.preventDefault();
      hideAlert('alert-t1');
      const wInput = document.getElementById('file-t1-word');
      const sInput = document.getElementById('file-t1-sso');
      if (!wInput.files || wInput.files.length === 0) {
        showAlert('alert-t1', 'error', '⚠️ Vui lòng chọn ít nhất 1 file Word (.docx) hoặc .zip!');
        return;
      }

      const picker = await promptSaveFile('Ket_qua.xlsx');
      if (picker.canceled) return;

      const formData = new FormData();
      for (let i = 0; i < wInput.files.length; i++) {
        formData.append('word_files', wInput.files[i]);
      }
      if (sInput.files && sInput.files[0]) {
        formData.append('sso_file', sInput.files[0]);
      }

      setLoading('t1', true);
      try {
        const res = await fetch('/api/task-1?action=task-1', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'Ket_qua.xlsx');
        if (picker.handle) {
          showAlert('alert-t1', 'success', `✅ Đã đối chiếu thành công và lưu file "${savedName}" tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t1', 'success', `✅ Đã đối chiếu thành công và tải xuống file "${savedName}" chuẩn 19 cột!`);
        }
      } catch (err) {
        showAlert('alert-t1', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t1', false, '🚀 Bắt Đầu Đối Chiếu & Xuất Ket_qua.xlsx');
      }
    }

    // XỬ LÝ TÁC VỤ 2
    async function handleTask2(e) {
      e.preventDefault();
      hideAlert('alert-t2');
      const wInput = document.getElementById('file-t2-word');
      const sInput = document.getElementById('file-t2-sso');
      if (!wInput.files || wInput.files.length === 0) {
        showAlert('alert-t2', 'error', '⚠️ Vui lòng chọn ít nhất 1 file Word (.docx)!');
        return;
      }

      const defaultMau02Name = 'Mẫu 02_CV cap CKS cho BVDK BN số 2_T10 2026_3_6.docx';
      const picker = await promptSaveFile(defaultMau02Name);
      if (picker.canceled) return;

      const formData = new FormData();
      for (let i = 0; i < wInput.files.length; i++) {
        formData.append('word_files', wInput.files[i]);
      }
      if (sInput.files && sInput.files[0]) {
        formData.append('sso_file', sInput.files[0]);
      }

      setLoading('t2', true);
      try {
        const res = await fetch('/api/task-2?action=task-2', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, defaultMau02Name);
        if (picker.handle) {
          showAlert('alert-t2', 'success', `✅ Đã xuất Công văn Mẫu 02 và lưu file "${savedName}" tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t2', 'success', `✅ Đã xuất Công văn Mẫu 02 (.docx) thành công!`);
        }
      } catch (err) {
        showAlert('alert-t2', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t2', false, '🚀 Bắt Đầu Tổng Hợp CKS (Mẫu 02)');
      }
    }

    // XỬ LÝ TÁC VỤ 3 (CKS MẪU CV02 CHỈ ĐỌC FILE WORD)
    async function handleTask3(e) {
      e.preventDefault();
      hideAlert('alert-t3');
      const wInput = document.getElementById('file-t3-word');
      if (!wInput.files || wInput.files.length === 0) {
        showAlert('alert-t3', 'error', '⚠️ Vui lòng chọn ít nhất 1 file Word (.docx / .zip)!');
        return;
      }

      const defaultName = 'Mẫu 02_CV cap CKS cho BVDK BN số 2_T10 2026_3_6.docx';
      const picker = await promptSaveFile(defaultName);
      if (picker.canceled) return;

      const formData = new FormData();
      for (let i = 0; i < wInput.files.length; i++) {
        formData.append('word_files', wInput.files[i]);
      }

      setLoading('t3', true);
      try {
        const res = await fetch('/api/task-3?action=task-3', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, defaultName);
        if (picker.handle) {
          showAlert('alert-t3', 'success', `✅ Đã xuất Công văn Mẫu 02 và lưu file "${savedName}" tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t3', 'success', `✅ Đã xuất Công văn Mẫu 02 (.docx) thành công!`);
        }
      } catch (err) {
        showAlert('alert-t3', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t3', false, '🚀 Bắt Đầu Tổng Hợp CKS (Mẫu CV02 - Chỉ File Word)');
      }
    }

    // XỬ LÝ TÁC VỤ 9 (CẤP EMAIL CÔNG VỤ)
    async function handleTask9(e) {
      e.preventDefault();
      hideAlert('alert-t9');
      const wInput = document.getElementById('file-t9-word');
      if (!wInput.files || wInput.files.length === 0) {
        showAlert('alert-t9', 'error', '⚠️ Vui lòng chọn các file Word xin cấp Email!');
        return;
      }
      const fmt = document.querySelector('input[name="email_format_9"]:checked').value;
      let defaultName = fmt === 'word' ? 'Đề nghị cấp mới địa chỉ email công vụ.docx' : 'Danh_sach_cap_email_cong_vu.xlsx';
      if (wInput.files.length === 1) {
        const fnameStr = wInput.files[0].name;
        const lastDot = fnameStr.lastIndexOf('.');
        const rawFname = lastDot !== -1 ? fnameStr.substring(0, lastDot) : fnameStr;
        const cleanSub = rawFname.replace(/^(Cấp mới và cấp lại CKS|Đơn đề nghị cấp CKS|Đề nghị cấp CKS)[ \t_–—-]*/i, '').trim();
        if (cleanSub) {
          defaultName = fmt === 'word' ? `Đề nghị cấp mới địa chỉ email công vụ_${cleanSub}.docx` : `Danh_sach_cap_email_cong_vu_${cleanSub}.xlsx`;
        }
      }

      const picker = await promptSaveFile(defaultName);
      if (picker.canceled) return;

      const formData = new FormData();
      for (let i = 0; i < wInput.files.length; i++) {
        formData.append('word_files', wInput.files[i]);
      }
      formData.append('format', fmt);

      setLoading('t9', true);
      try {
        const res = await fetch('/api/task-9?action=task-9', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }

        let finalFilename = defaultName;
        const xFilename = res.headers.get('X-Filename');
        if (xFilename) {
          try { finalFilename = decodeURIComponent(xFilename); } catch (_) {}
        }

        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, finalFilename);
        if (picker.handle) {
          showAlert('alert-t9', 'success', `✅ Đã lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t9', 'success', `✅ Đã tổng hợp và tải xuống file "${savedName}" thành công!`);
        }
      } catch (err) {
        showAlert('alert-t9', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t9', false, '🚀 Bắt Đầu Tổng Hợp Danh Sách Email');
      }
    }

    // XỬ LÝ TÁC VỤ 4
    async function handleTask4(e) {
      e.preventDefault();
      hideAlert('alert-t4');
      const fInput = document.getElementById('file-t4');
      const sheet = document.getElementById('select-t4-sheet').value;
      if (!fInput.files || !fInput.files[0]) {
        showAlert('alert-t4', 'error', '⚠️ Vui lòng chọn file Excel báo cáo P.CNTT!');
        return;
      }
      if (!sheet) {
        showAlert('alert-t4', 'error', '⚠️ Vui lòng chọn Sheet ngày cần tổng hợp!');
        return;
      }

      const picker = await promptSaveFile('TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx');
      if (picker.canceled) return;

      const formData = new FormData();
      formData.append('file', fInput.files[0]);
      formData.append('sheet_name', sheet);

      setLoading('t4', true);
      try {
        const res = await fetch('/api/task-4?action=task-4', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx');
        if (picker.handle) {
          showAlert('alert-t4', 'success', `✅ Đã lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t4', 'success', `✅ Đã tổng hợp thành công và tải file A4 ngang!`);
        }
      } catch (err) {
        showAlert('alert-t4', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t4', false, '🚀 Tổng Hợp Báo Cáo P.CNTT (A4 Ngang)');
      }
    }

    // XỬ LÝ TÁC VỤ 5
    async function handleTask5(e) {
      e.preventDefault();
      hideAlert('alert-t5');
      const fInput = document.getElementById('file-t5');
      if (!fInput.files || !fInput.files[0]) {
        showAlert('alert-t5', 'error', '⚠️ Vui lòng chọn file Excel Check Log MDINSIGHT!');
        return;
      }

      // Mở hộp thoại chọn vị trí và tên file lưu trên máy tính
      const picker = await promptSaveFile('DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx');
      if (picker.canceled) {
        return;
      }

      const range = document.querySelector('input[name="t5_date_range"]:checked').value;
      const formData = new FormData();
      formData.append('file', fInput.files[0]);
      formData.append('date_range', range);

      setLoading('t5', true);
      try {
        const res = await fetch('/api/task-5?action=task-5', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx');
        if (picker.handle) {
          showAlert('alert-t5', 'success', `✅ Đã lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t5', 'success', `✅ Đã xuất và tải xuống file "${savedName}" A4 dọc thành công!`);
        }
      } catch (err) {
        showAlert('alert-t5', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t5', false, '🚀 Xử Lý & Xuất Check Log MDINSIGHT (A4 Dọc)');
      }
    }

    // XỬ LÝ TÁC VỤ 6
    async function handleTask6(e) {
      e.preventDefault();
      hideAlert('alert-t6');
      const url = document.getElementById('url-t6').value.trim();
      const sheet = document.getElementById('select-t6-sheet').value;
      if (!url) {
        showAlert('alert-t6', 'error', '⚠️ Vui lòng nhập link Google Sheets!');
        return;
      }
      if (!sheet) {
        showAlert('alert-t6', 'error', '⚠️ Vui lòng nhấn "Nạp Danh Sách Sheet Ngày" và chọn sheet cần xử lý!');
        return;
      }

      const picker = await promptSaveFile('TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx');
      if (picker.canceled) return;

      const formData = new FormData();
      formData.append('url', url);
      formData.append('sheet_name', sheet);

      setLoading('t6', true);
      try {
        const res = await fetch('/api/task-6?action=task-6', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx');
        if (picker.handle) {
          showAlert('alert-t6', 'success', `✅ Đã lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t6', 'success', `✅ Đã tổng hợp từ Google Sheets và tải xuống file thành công!`);
        }
      } catch (err) {
        showAlert('alert-t6', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t6', false, '🚀 Tổng Hợp Báo Cáo P.CNTT (A4 Ngang)');
      }
    }

    // XỬ LÝ TÁC VỤ 7
    async function handleTask7(e) {
      e.preventDefault();
      hideAlert('alert-t7');
      const wInput = document.getElementById('file-t7-word');
      if (!wInput.files || wInput.files.length === 0) {
        showAlert('alert-t7', 'error', '⚠️ Vui lòng chọn các file Word (.docx) để tổng hợp!');
        return;
      }

      // Mở hộp thoại chọn vị trí và tên file lưu trên máy tính
      const picker = await promptSaveFile('Ket_qua.xlsx');
      if (picker.canceled) {
        return;
      }

      const formData = new FormData();
      for (let i = 0; i < wInput.files.length; i++) {
        formData.append('word_files', wInput.files[i]);
      }

      setLoading('t7', true);
      try {
        const res = await fetch('/api/task-7?action=task-7', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'Ket_qua.xlsx');
        if (picker.handle) {
          showAlert('alert-t7', 'success', `✅ Đã tổng hợp từ Word và lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t7', 'success', `✅ Đã tổng hợp từ Word và tải xuống file "${savedName}" chuẩn 19 cột thành công!`);
        }
      } catch (err) {
        showAlert('alert-t7', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t7', false, '🚀 Bắt Đầu Tổng Hợp VGCA (Xuất Ket_qua.xlsx)');
      }
    }

    // XỬ LÝ TÁC VỤ 8
    async function handleTask8(e) {
      e.preventDefault();
      hideAlert('alert-t8');
      const sheet = document.getElementById('select-t8-sheet').value;
      if (!sheet) {
        showAlert('alert-t8', 'error', '⚠️ Vui lòng nhấn "Nạp Danh Sách Sheet Ngày" và chọn sheet cần xử lý!');
        return;
      }

      const picker = await promptSaveFile('TỔNG HỢP CÔNG VIỆC P.CNTT (CỐ ĐỊNH).xlsx');
      if (picker.canceled) return;

      const formData = new FormData();
      formData.append('url', document.getElementById('url-t8').value);
      formData.append('sheet_name', sheet);

      setLoading('t8', true);
      try {
        const res = await fetch('/api/task-8?action=task-8', { method: 'POST', body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: 'Lỗi máy chủ.' }));
          throw new Error(errData.error || 'Xử lý thất bại.');
        }
        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'TỔNG HỢP CÔNG VIỆC P.CNTT (CỐ ĐỊNH).xlsx');
        if (picker.handle) {
          showAlert('alert-t8', 'success', `✅ Đã lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t8', 'success', `✅ Đã tổng hợp từ Google Sheets của viện và tải xuống file thành công!`);
        }
      } catch (err) {
        showAlert('alert-t8', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t8', false, '🚀 Tổng Hợp Từ Google Sheets (A4 Ngang)');
      }
    }

    // ============================================================
    // LOGIC TÁC VỤ 10: THỐNG KÊ & LỌC LỖI XML BHYT
    // ============================================================
    const DEFAULT_TRACKER_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbydmlNGV_CDYGkMgRmInTi04zfjNpaP3qF1-uEVStDsyqZKxQ8xbNOpDvcsuDmT9m1N/exec';
    let trackerInitialized = false;
    let currentTrackerSheet = '';
    let trackerAllRecords = [];
    let trackerFilteredRecords = [];

    async function initErrorTracker() {
      if (trackerInitialized) return;
      trackerInitialized = true;
      // Dọn sạch mọi cache trạng thái cũ trong LocalStorage để luôn đọc chuẩn từ Google Sheet
      try {
        Object.keys(localStorage).forEach(k => {
          if (k.startsWith('tracker_status_')) localStorage.removeItem(k);
        });
      } catch (_) {}
      const stored = localStorage.getItem('tracker_apps_script_url') || '';
      if (!stored || stored.includes('/edit') || stored.includes('vercel.app')) {
        localStorage.setItem('tracker_apps_script_url', DEFAULT_TRACKER_APPS_SCRIPT_URL);
      }
      try {
        const res = await fetch('/api/error-tracker/config?action=error-tracker-config&_t=' + Date.now());
        if (res.ok) {
          const cfg = await res.json();
          if (cfg.script_url && !cfg.script_url.includes('/edit')) {
            localStorage.setItem('tracker_apps_script_url', cfg.script_url);
          }
        }
      } catch (_) {}
      updateAppsScriptBtnStatus();
      await loadTrackerSheets();
    }

    async function loadTrackerSheets(forceReload = false) {
      const select = document.getElementById('tracker-sheet-select');
      select.innerHTML = '<option value="">⏳ Đang nạp danh sách sheet ngày...</option>';
      try {
        const res = await fetch('/api/error-tracker/sheets' + (forceReload ? '?action=error-tracker-sheets&reload=1' : '?action=error-tracker-sheets'));
        if (!res.ok) throw new Error('Không thể tải danh sách sheets.');
        const data = await res.json();
        if (!data.sheets || data.sheets.length === 0) {
          select.innerHTML = '<option value="">Không tìm thấy sheet nào</option>';
          return;
        }

        select.innerHTML = '';
        data.sheets.forEach((s) => {
          const opt = document.createElement('option');
          opt.value = s;
          opt.textContent = 'Ngày ' + s;
          if (s === data.default_sheet) opt.selected = true;
          select.appendChild(opt);
        });

        currentTrackerSheet = select.value;
        await loadTrackerSheetData(currentTrackerSheet);
      } catch (err) {
        select.innerHTML = '<option value="">❌ Lỗi tải sheet</option>';
        alert('Lỗi tải danh sách sheets: ' + err.message);
      }
    }

    async function onTrackerSheetChange() {
      const select = document.getElementById('tracker-sheet-select');
      const sheet = select.value;
      if (!sheet) return;
      currentTrackerSheet = sheet;
      await loadTrackerSheetData(sheet);
    }

    async function loadTrackerSheetData(sheetName) {
      const tbody = document.getElementById('tracker-records-tbody');
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:30px; color:#2563eb;">⏳ Đang tải và phân tích dữ liệu ngày ' + sheetName + '...</td></tr>';

      try {
        const res = await fetch('/api/error-tracker/data?action=error-tracker-data&sheet=' + encodeURIComponent(sheetName) + '&_t=' + Date.now());
        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.error || ('Lỗi máy chủ ' + res.status));
        }
        const data = await res.json();
        trackerAllRecords = data.records || [];

        // Nạp dropdown Khoa
        const kSel = document.getElementById('tracker-khoa-select');
        kSel.innerHTML = '<option value="ALL">-- Tất Cả Khoa (' + (data.total_khoa || 0) + ') --</option>';
        (data.khoa_list || []).forEach(k => {
          const opt = document.createElement('option');
          opt.value = k.name;
          opt.textContent = `${k.name} (${k.count})`;
          kSel.appendChild(opt);
        });

        // Nạp dropdown Phòng
        const pSel = document.getElementById('tracker-phong-select');
        pSel.innerHTML = '<option value="ALL">-- Tất Cả Phòng (' + (data.total_phong || 0) + ') --</option>';
        (data.phong_list || []).forEach(p => {
          const opt = document.createElement('option');
          opt.value = p.name;
          opt.textContent = `${p.name} (${p.count})`;
          pSel.appendChild(opt);
        });

        // Nạp dropdown XML
        const xSel = document.getElementById('tracker-xml-select');
        xSel.innerHTML = '<option value="ALL">-- Tất Cả Bảng XML --</option>';
        (data.xml_list || []).forEach(x => {
          const opt = document.createElement('option');
          opt.value = x.name;
          opt.textContent = `${x.name} (${x.count})`;
          xSel.appendChild(opt);
        });

        applyTrackerFilters();
      } catch (err) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:30px; color:#dc2626;">❌ Không thể tải dữ liệu: ${err.message}</td></tr>`;
      }
    }

    function applyTrackerFilters() {
      const kVal = document.getElementById('tracker-khoa-select').value;
      const pVal = document.getElementById('tracker-phong-select').value;
      const xVal = document.getElementById('tracker-xml-select').value;
      const statusSelect = document.getElementById('tracker-status-select');
      const statusVal = statusSelect ? statusSelect.value : 'ALL';
      const search = document.getElementById('tracker-search-input').value.trim().toLowerCase();

      let filtered = trackerAllRecords.filter(r => {
        if (kVal !== 'ALL' && r.khoa !== kVal) return false;
        if (pVal !== 'ALL' && r.phong !== pVal) return false;
        if (xVal !== 'ALL' && r.xml !== xVal) return false;
        if (statusVal === 'PENDING' && r.is_ok) return false;
        if (statusVal === 'OK' && !r.is_ok) return false;
        if (search) {
          const matchHs = r.hs && r.hs.toLowerCase().includes(search);
          const matchLoi = r.loi && r.loi.toLowerCase().includes(search);
          const matchKhoa = r.khoa && r.khoa.toLowerCase().includes(search);
          const matchPhong = r.phong && r.phong.toLowerCase().includes(search);
          if (!matchHs && !matchLoi && !matchKhoa && !matchPhong) return false;
        }
        return true;
      });

      trackerFilteredRecords = filtered;
      renderTrackerStats(filtered);
      renderTrackerTable(filtered);
    }

    function renderTrackerStats(records) {
      const totalErrors = records.length;
      const uniqueHs = new Set(records.map(r => r.hs).filter(Boolean)).size;
      const khoaSet = new Set(records.map(r => r.khoa).filter(Boolean));
      const phongSet = new Set(records.map(r => r.phong).filter(Boolean));
      const totalOk = records.filter(r => r.is_ok).length;
      const pctOk = totalErrors > 0 ? Math.round((totalOk / totalErrors) * 100) : 0;

      document.getElementById('stat-total-errors').textContent = totalErrors.toLocaleString();
      document.getElementById('stat-sub-errors').textContent = `Ngày ${currentTrackerSheet}`;
      document.getElementById('stat-unique-hs').textContent = uniqueHs.toLocaleString();
      document.getElementById('stat-total-khoa').textContent = khoaSet.size;
      document.getElementById('stat-total-phong').textContent = phongSet.size;

      const statOkVal = document.getElementById('stat-total-ok');
      if (statOkVal) statOkVal.textContent = totalOk.toLocaleString();
      const statOkSub = document.getElementById('stat-sub-ok');
      if (statOkSub) statOkSub.textContent = `${pctOk}% hoàn thành`;

      // XML stats
      const xmlCounts = {};
      records.forEach(r => {
        if (r.xml) xmlCounts[r.xml] = (xmlCounts[r.xml] || 0) + 1;
      });
      const sortedXml = Object.entries(xmlCounts).sort((a,b) => b[1] - a[1]);
      if (sortedXml.length > 0) {
        const top = sortedXml[0];
        const pct = totalErrors > 0 ? Math.round((top[1] / totalErrors) * 100) : 0;
        document.getElementById('stat-top-xml').textContent = top[0];
        document.getElementById('stat-top-xml-sub').textContent = `${top[1]} lỗi (${pct}%)`;
      } else {
        document.getElementById('stat-top-xml').textContent = '-';
        document.getElementById('stat-top-xml-sub').textContent = 'Chưa có lỗi';
      }

      // XML Chips
      const chipsBox = document.getElementById('tracker-xml-chips');
      if (sortedXml.length > 0) {
        chipsBox.innerHTML = sortedXml.map(([name, count]) => {
          const badgeClass = getXmlBadgeClass(name);
          return `<span class="badge-xml ${badgeClass}" style="cursor:pointer; font-size:0.8rem; padding:4px 9px;" onclick="filterByXmlQuick('${name}')">${name}: <b>${count}</b></span>`;
        }).join('');
      } else {
        chipsBox.innerHTML = '<span style="color:#94a3b8; font-size:0.82rem;">Không có dữ liệu</span>';
      }

      // Khoa ranking
      const khoaCounts = {};
      records.forEach(r => {
        const k = r.khoa || 'Chưa xác định';
        khoaCounts[k] = (khoaCounts[k] || 0) + 1;
      });
      const sortedKhoa = Object.entries(khoaCounts).sort((a,b) => b[1] - a[1]);
      const rankingBox = document.getElementById('tracker-khoa-ranking');
      if (sortedKhoa.length > 0) {
        const maxK = sortedKhoa[0][1];
        rankingBox.innerHTML = sortedKhoa.map(([name, count], idx) => {
          const pct = Math.round((count / maxK) * 100);
          return `
            <div class="ranking-item" data-khoa="${encodeURIComponent(name)}" onclick="filterByKhoaQuick(decodeURIComponent(this.getAttribute('data-khoa')))" title="Nhấp để lọc ${name}">
              <div class="ranking-name">
                <span style="color:#94a3b8; font-size:0.75rem; width:20px;">#${idx+1}</span>
                <span>${name}</span>
              </div>
              <div style="display:flex; align-items:center;">
                <span class="ranking-count-badge">${count} lỗi</span>
                <div class="ranking-bar-bg"><div class="ranking-bar-fill" style="width:${pct}%;"></div></div>
              </div>
            </div>`;
        }).join('');
      } else {
        rankingBox.innerHTML = '<div style="padding:16px; text-align:center; color:#94a3b8;">Không có dữ liệu phù hợp bộ lọc</div>';
      }

      // Top errors patterns
      const loiCounts = {};
      records.forEach(r => {
        const l = r.loi || '';
        const firstLine = (l.split(String.fromCharCode(10))[0] || '').replace(String.fromCharCode(13), '').trim();
        let pat = firstLine.length > 70 ? firstLine.substring(0, 70) + '...' : firstLine;
        loiCounts[pat] = (loiCounts[pat] || 0) + 1;
      });
      const sortedLoi = Object.entries(loiCounts).sort((a,b) => b[1] - a[1]).slice(0, 8);
      const topErrorsBox = document.getElementById('tracker-top-errors-box');
      if (sortedLoi.length > 0) {
        topErrorsBox.innerHTML = sortedLoi.map(([rule, count]) => `
          <div class="ranking-item" style="cursor:default;">
            <div class="ranking-name" style="font-size:0.8rem; line-height:1.3;">
              <span style="color:#ef4444; font-weight:700;">•</span>
              <span>${rule}</span>
            </div>
            <span class="ranking-count-badge" style="background:#fee2e2; color:#b91c1c;">${count}</span>
          </div>`).join('');
      } else {
        topErrorsBox.innerHTML = '<div style="padding:16px; text-align:center; color:#94a3b8;">Không có dữ liệu</div>';
      }
    }

    function renderTrackerTable(records) {
      const tbody = document.getElementById('tracker-records-tbody');
      const summary = document.getElementById('tracker-table-summary');
      summary.textContent = `Hiển thị ${records.length} / ${trackerAllRecords.length} hồ sơ lỗi (Ngày ${currentTrackerSheet})`;

      if (!records || records.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align:center; padding:30px; color:#94a3b8;">🔍 Không tìm thấy bản ghi nào khớp với điều kiện lọc hiện tại.</td></tr>';
        return;
      }

      tbody.innerHTML = records.map((r, i) => {
        const badgeClass = getXmlBadgeClass(r.xml);
        const escapedLoi = (r.loi || '').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(new RegExp(String.fromCharCode(10), 'g'), '<br>');
        return `
          <tr>
            <td style="text-align:center; font-weight:700; color:#64748b;">${i + 1}</td>
            <td>
              <span style="font-family:monospace; font-weight:700; color:#1e293b; font-size:0.88rem;">${r.hs || '-'}</span>
              ${r.hs ? `<button type="button" class="btn-copy-hs" onclick="copyHsToClipboard('${r.hs}')" title="Copy mã HS">📋</button>` : ''}
            </td>
            <td style="color:#475569; font-size:0.8rem; white-space:nowrap;">${r.ngay || '-'}</td>
            <td style="text-align:center;"><span class="badge-xml ${badgeClass}">${r.xml || '-'}</span></td>
            <td><b style="color:#0f172a;">${r.khoa || '-'}</b></td>
            <td style="color:#334155;">${r.phong || '-'}</td>
            <td style="line-height:1.4; color:#991b1b; font-size:0.82rem;">${escapedLoi || '-'}</td>
            <td style="text-align:center; vertical-align:middle;">
              <label style="display:inline-flex; align-items:center; justify-content:center; gap:6px; cursor:pointer; user-select:none; margin:0;" title="${r.is_ok ? 'Đã xử lý (Nhấp để hủy OK)' : 'Chưa xử lý (Nhấp để đánh dấu OK)'}">
                <input type="checkbox" class="tracker-status-check" 
                  data-hs="${r.hs}" 
                  data-row="${r.sheet_row}"
                  ${r.is_ok ? 'checked' : ''} 
                  onchange="onTrackerStatusToggle(this, '${r.hs}', ${r.sheet_row})">
                <span class="status-badge ${r.is_ok ? 'badge-ok' : 'badge-pending'}" id="status-badge-${r.sheet_row}">
                  ${r.is_ok ? '✓ OK' : 'Chưa xử lý'}
                </span>
              </label>
            </td>
          </tr>`;
      }).join('');
    }

    async function onTrackerStatusToggle(checkbox, hs, sheetRow) {
      const isChecked = checkbox.checked;
      const statusVal = isChecked ? 'OK' : '';
      const badge = document.getElementById('status-badge-' + sheetRow);

      // 1. Phản hồi giao diện tức thì
      if (badge) {
        badge.className = isChecked ? 'status-badge badge-ok' : 'status-badge badge-pending';
        badge.textContent = isChecked ? '✓ OK' : 'Chưa xử lý';
      }

      // 2. Cập nhật mảng bản ghi trong bộ nhớ
      const rec = trackerAllRecords.find(r => r.sheet_row === sheetRow || r.hs === hs);
      if (rec) {
        rec.status = statusVal;
        rec.is_ok = isChecked;
      }

      // 3. Cập nhật lại các thẻ thống kê
      renderTrackerStats(trackerFilteredRecords);

      // 5. Gửi ngầm lên server để lưu vào hệ thống và đồng bộ sang Google Sheet nếu có Webhook
      let scriptUrl = localStorage.getItem('tracker_apps_script_url') || DEFAULT_TRACKER_APPS_SCRIPT_URL;
      if (scriptUrl.includes('/edit') || scriptUrl.includes('vercel.app')) {
        scriptUrl = DEFAULT_TRACKER_APPS_SCRIPT_URL;
        localStorage.setItem('tracker_apps_script_url', DEFAULT_TRACKER_APPS_SCRIPT_URL);
      }

      try {
        const res = await fetch('/api/error-tracker/update-status?action=error-tracker-update-status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sheet: currentTrackerSheet,
            hs: hs,
            row: sheetRow,
            status: statusVal,
            script_url: scriptUrl
          })
        });

        const data = await res.json().catch(() => ({}));

        if (data && data.synced_to_sheet) {
          showMiniToast(`✅ Đã tự động điền 'OK' vào Google Sheet (Dòng ${sheetRow}, HS ${hs})!`);
        } else {
          const detail = (data && (data.error_detail || data.error)) ? data.error_detail || data.error : 'Chưa thiết lập Webhook';
          showMiniToast(`⚠️ Đã lưu trên Web! Google Sheet chưa cập nhật: ${detail}.`);
        }
      } catch (err) {
        showMiniToast(`⚠️ Đã lưu trạng thái trên Web! (Lỗi gửi: ${err.message})`);
      }
    }

    function updateAppsScriptBtnStatus() {
      let url = localStorage.getItem('tracker_apps_script_url') || DEFAULT_TRACKER_APPS_SCRIPT_URL;
      if (url.includes('/edit') || url.includes('vercel.app')) {
        url = DEFAULT_TRACKER_APPS_SCRIPT_URL;
        localStorage.setItem('tracker_apps_script_url', DEFAULT_TRACKER_APPS_SCRIPT_URL);
      }
      const btn = document.getElementById('btn-apps-script-cfg');
      if (!btn) return;
      btn.innerHTML = '⚡ Đã Kết Nối Google Sheet (Sẵn Sàng ✅)';
      btn.style.background = 'linear-gradient(135deg, #059669 0%, #047857 100%)';
    }

    function openAppsScriptModal() {
      const modal = document.getElementById('apps-script-modal');
      if (modal) {
        let currentUrl = localStorage.getItem('tracker_apps_script_url') || DEFAULT_TRACKER_APPS_SCRIPT_URL;
        if (currentUrl.includes('/edit') || currentUrl.includes('vercel.app')) {
          currentUrl = DEFAULT_TRACKER_APPS_SCRIPT_URL;
          localStorage.setItem('tracker_apps_script_url', DEFAULT_TRACKER_APPS_SCRIPT_URL);
        }
        const input = document.getElementById('apps-script-url-input');
        if (input) input.value = currentUrl;
        const resBox = document.getElementById('test-script-result');
        if (resBox) {
          resBox.style.display = 'block';
          resBox.style.background = '#ecfdf5';
          resBox.style.color = '#065f46';
          resBox.style.border = '1px solid #a7f3d0';
          resBox.innerHTML = '<b>✅ ĐÃ TỰ ĐỘNG KẾT NỐI MẶC ĐỊNH:</b><br>Hệ thống đã tự động tích hợp link Web App của viện. Bạn và mọi người chỉ cần vào dùng là sẽ tự động cập nhật vào Google Sheet!';
        }
        modal.style.display = 'flex';
      }
    }

    function closeAppsScriptModal() {
      const modal = document.getElementById('apps-script-modal');
      if (modal) modal.style.display = 'none';
    }

    function copyAppsScriptCode() {
      const codeBox = document.getElementById('apps-script-code-box');
      if (codeBox) {
        codeBox.select();
        navigator.clipboard.writeText(codeBox.value).then(() => {
          const btn = document.getElementById('btn-copy-script');
          if (btn) {
            const orig = btn.textContent;
            btn.textContent = '✅ Đã Chép!';
            setTimeout(() => { btn.textContent = orig; }, 2000);
          }
        });
      }
    }

    async function testAppsScriptUrl() {
      const input = document.getElementById('apps-script-url-input');
      const resBox = document.getElementById('test-script-result');
      const btn = document.getElementById('btn-test-script');
      if (!input || !resBox || !btn) return;
      const url = input.value.trim();
      if (!url) {
        resBox.style.display = 'block';
        resBox.style.background = '#fef2f2';
        resBox.style.color = '#991b1b';
        resBox.style.border = '1px solid #fecaca';
        resBox.innerHTML = '⚠️ Vui lòng dán link Web App vào ô trước khi bấm thử!';
        return;
      }
      if (url.includes('/edit') || url.includes('docs.google.com/spreadsheets')) {
        resBox.style.display = 'block';
        resBox.style.background = '#fef2f2';
        resBox.style.color = '#991b1b';
        resBox.style.border = '1px solid #fecaca';
        resBox.innerHTML = '❌ <b>Đường link không đúng:</b> Bạn đang dán link chỉnh sửa (/edit) hoặc link Google Sheet! Link hợp lệ bắt buộc phải lấy từ nút <b>Triển khai ➔ Triển khai mới ➔ Ứng dụng web</b> và kết thúc bằng <code>/exec</code>.';
        return;
      }
      btn.disabled = true;
      btn.textContent = '⏳ Đang kiểm tra...';
      resBox.style.display = 'block';
      resBox.style.background = '#f0f9ff';
      resBox.style.color = '#0369a1';
      resBox.style.border = '1px solid #bae6fd';
      resBox.innerHTML = '⏳ Đang gửi yêu cầu ping tới Google Apps Script...';
      try {
        const resp = await fetch('/api/error-tracker/test-connection?action=error-tracker-test-connection', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ script_url: url })
        });
        const data = await resp.json().catch(() => ({}));
        const isSuccess = data && (data.success === true || (data.error && data.error.includes('"success":true')) || (data.error && data.error.includes('"ping":"pong"')));
        if (isSuccess) {
          resBox.style.background = '#ecfdf5';
          resBox.style.color = '#065f46';
          resBox.style.border = '1px solid #a7f3d0';
          resBox.innerHTML = '<b>✅ KẾT NỐI THÀNH CÔNG 100%!</b><br>Google Apps Script đã sẵn sàng ghi dữ liệu vào Google Sheet. Hãy bấm nút <b>"💾 Lưu Cấu Hình"</b> bên dưới để áp dụng!';
        } else {
          resBox.style.background = '#fef2f2';
          resBox.style.color = '#991b1b';
          resBox.style.border = '1px solid #fecaca';
          resBox.innerHTML = `<b>❌ KẾT NỐI THẤT BẠI:</b><br>${data.error || 'Google Apps Script không phản hồi hợp lệ.'}`;
        }
      } catch (err) {
        resBox.style.background = '#fef2f2';
        resBox.style.color = '#991b1b';
        resBox.style.border = '1px solid #fecaca';
        resBox.innerHTML = `<b>❌ Lỗi kết nối:</b> ${err.message}`;
      } finally {
        btn.disabled = false;
        btn.textContent = '🔍 Thử Kết Nối';
      }
    }

    async function saveAppsScriptUrl() {
      const input = document.getElementById('apps-script-url-input');
      if (!input) return;
      const val = input.value.trim();
      if (val && !val.startsWith('http')) {
        alert('Vui lòng nhập đúng đường link Web App (bắt đầu bằng https://)!');
        return;
      }
      if (val.includes('/edit') || val.includes('docs.google.com/spreadsheets')) {
        if (!confirm('Cảnh báo: Bạn đang dán link trình duyệt /edit hoặc link Google Sheet chứ không phải link Web App /exec! Bạn có chắc chắn muốn lưu không?')) {
          return;
        }
      }
      localStorage.setItem('tracker_apps_script_url', val);
      try {
        await fetch('/api/error-tracker/config?action=error-tracker-config', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ script_url: val })
        });
      } catch (_) {}
      updateAppsScriptBtnStatus();
      closeAppsScriptModal();
      showMiniToast('✅ Đã lưu cấu hình kết nối Google Sheet thành công!');
    }

    function showMiniToast(msg) {
      let toast = document.getElementById('tracker-mini-toast');
      if (!toast) {
        toast = document.createElement('div');
        toast.id = 'tracker-mini-toast';
        toast.style.cssText = 'position:fixed; bottom:24px; right:24px; color:#ffffff; padding:12px 20px; border-radius:8px; font-size:0.88rem; font-weight:600; box-shadow:0 10px 15px -3px rgba(0,0,0,0.15); z-index:99999; display:none; transition:opacity 0.3s; max-width:440px; line-height:1.4;';
        document.body.appendChild(toast);
      }
      toast.textContent = msg;
      if (msg.startsWith('⚠️')) {
        toast.style.background = '#b45309';
      } else {
        toast.style.background = '#065f46';
      }
      toast.style.display = 'block';
      toast.style.opacity = '1';
      clearTimeout(toast._timer);
      toast._timer = setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => { toast.style.display = 'none'; }, 300);
      }, 5000);
    }

    function getXmlBadgeClass(xml) {
      if (!xml) return 'badge-xml-default';
      const x = xml.toUpperCase().trim();
      if (x === 'XML1') return 'badge-xml1';
      if (x === 'XML2') return 'badge-xml2';
      if (x === 'XML3') return 'badge-xml3';
      if (x === 'XML4') return 'badge-xml4';
      if (x === 'XML5') return 'badge-xml5';
      if (x === 'XML13') return 'badge-xml13';
      return 'badge-xml-default';
    }

    function filterByKhoaQuick(khoaName) {
      const kSel = document.getElementById('tracker-khoa-select');
      kSel.value = khoaName;
      applyTrackerFilters();
      kSel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function filterByXmlQuick(xmlName) {
      const xSel = document.getElementById('tracker-xml-select');
      xSel.value = xmlName;
      applyTrackerFilters();
      xSel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function copyHsToClipboard(hs) {
      if (!hs) return;
      navigator.clipboard.writeText(hs).then(() => {
        alert('Đã copy mã hồ sơ: ' + hs);
      }).catch(() => {
        prompt('Copy mã hồ sơ:', hs);
      });
    }

    async function exportTrackerExcel() {
      if (!trackerFilteredRecords || trackerFilteredRecords.length === 0) {
        alert('Không có dữ liệu nào để xuất!');
        return;
      }

      const defaultFileName = `Bao_cao_loi_XML_${currentTrackerSheet.replace(/[^a-zA-Z0-9.-]/g, '_')}.xlsx`;
      const picker = await promptSaveFile(defaultFileName);
      if (picker.canceled) return;

      const btn = document.getElementById('btn-tracker-export');
      const oldTxt = btn.innerHTML;
      btn.innerHTML = '⏳ Đang Xuất File...';
      btn.disabled = true;

      try {
        const payload = {
          sheet: currentTrackerSheet,
          records: trackerFilteredRecords,
          filter_khoa: document.getElementById('tracker-khoa-select').value,
          filter_phong: document.getElementById('tracker-phong-select').value,
          filter_xml: document.getElementById('tracker-xml-select').value
        };

        const res = await fetch('/api/error-tracker/export?action=error-tracker-export', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.error || 'Lỗi xuất file Excel.');
        }

        const blob = await res.blob();
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, defaultFileName);
        alert(`✅ Đã lưu file "${savedName}" thành công!`);
      } catch (err) {
        alert('❌ Lỗi xuất file: ' + err.message);
      } finally {
        btn.innerHTML = oldTxt;
        btn.disabled = false;
      }
    }

      window.addEventListener('DOMContentLoaded', () => {
      updateAuthUI();
      executeSwitchTask(10);
    });
  

    /* ============================================================
       BC SỬA CNTT JAVASCRIPT CONTROLLER
       ============================================================ */
    const CNTT_ALL_MONTHS = {
      m10: [
        '29.10', '28.10', '27.10', '26.10', '23-24-25.10', '22.10', '21.10', 
        '20.10', '19.10', '16.10-18.10', '15.10', '14.10', '13.10', '12.10', 
        '9.10-10.10-11.10', '8.10', '7.10', '6.10', '05.10', '3-4.10', '2.10', '1.10'
      ],
      m9: [
        '30.9', '29.9', '28.9', '25-26-27.9', '24.9', '23.9', '22.9', '21.9',
        '18-19-20.9', '17.9', '16.9', '15.9', '14.9', '11-12-13.09', '10.9', '9.9', '8.9', '7.9', '4.9', '3.9', '31.8-1.9-2.9'
      ],
      m8: [
        '28-29-30.8', '27.8', '26.8', '25.8', '24.8', '21-22-23.8', '20.8', '19.8',
        '18.8', '17.8', '14-15-16.8', '15.8', '13.8', '12.8', '11.8', '10.8', '7-8-9.8', '6.8', '5.8', '4.8', '3.8'
      ],
      m7: ['31.7-2.8', '30.7', '29.7']
    };

    let cnttSelectedMonth = 'm10';
    let cnttSelectedSheets = new Set(['29.10', '28.10', '27.10']);
    let cnttRawRecords = [];
    let cnttAggregatedData = null;
    let cnttIsGridCollapsed = false;

    // Preloaded initial records from Python
    const CNTT_INITIAL_DATA = {};

    function initCnttDashboard() {
      renderCnttMonthPills();
      renderCnttSheetGrid();
      
      // Fetch data cleanly from API
      fetchOrAggregateCnttSheets(false);
      
      // Show sync time
      const now = new Date();
      const h = String(now.getHours()).padStart(2, '0');
      const m = String(now.getMinutes()).padStart(2, '0');
      const clockEl = document.getElementById('cntt-sync-clock');
      if (clockEl) clockEl.textContent = `${h}:${m}`;
    }

    function renderCnttMonthPills() {
      // Setup click handlers for month buttons
    }

    function switchCnttMonth(monthKey, el) {
      cnttSelectedMonth = monthKey;
      document.querySelectorAll('.cntt-month-pill').forEach(p => p.classList.remove('active'));
      if (el) el.classList.add('active');
      renderCnttSheetGrid();
    }

    function renderCnttSheetGrid() {
      const grid = document.getElementById('cntt-sheets-grid');
      if (!grid) return;
      grid.innerHTML = '';

      let list = [];
      if (cnttSelectedMonth === 'all') {
        list = [...CNTT_ALL_MONTHS.m10, ...CNTT_ALL_MONTHS.m9, ...CNTT_ALL_MONTHS.m8, ...CNTT_ALL_MONTHS.m7];
      } else {
        list = CNTT_ALL_MONTHS[cnttSelectedMonth] || CNTT_ALL_MONTHS.m10;
      }

      list.forEach(sheetName => {
        const chip = document.createElement('label');
        const isSelected = cnttSelectedSheets.has(sheetName);
        chip.className = 'cntt-sheet-chip' + (isSelected ? ' selected' : '');
        chip.innerHTML = `
          <input type="checkbox" value="${sheetName}" ${isSelected ? 'checked' : ''} onchange="toggleCnttSheet('${sheetName}', this)">
          ${sheetName}
        `;
        grid.appendChild(chip);
      });
    }

    async function toggleCnttSheet(sheetName, checkbox) {
      if (checkbox.checked) {
        cnttSelectedSheets.add(sheetName);
      } else {
        cnttSelectedSheets.delete(sheetName);
      }
      renderCnttSheetGrid();
      await fetchOrAggregateCnttSheets();
    }

    function selectAllCurrentMonthSheets() {
      const list = CNTT_ALL_MONTHS[cnttSelectedMonth] || CNTT_ALL_MONTHS.m10;
      list.forEach(s => cnttSelectedSheets.add(s));
      renderCnttSheetGrid();
      fetchOrAggregateCnttSheets();
    }

    function clearSelectedSheets() {
      cnttSelectedSheets.clear();
      renderCnttSheetGrid();
      fetchOrAggregateCnttSheets();
    }

    function selectCnttShortcut(type) {
      cnttSelectedSheets.clear();
      if (type === 'today') {
        cnttSelectedSheets.add('29.10');
      } else if (type === 'week') {
        ['29.10', '28.10', '27.10', '26.10', '23-24-25.10'].forEach(s => cnttSelectedSheets.add(s));
      } else if (type === 'm10') {
        CNTT_ALL_MONTHS.m10.forEach(s => cnttSelectedSheets.add(s));
      } else if (type === 'm9') {
        CNTT_ALL_MONTHS.m9.forEach(s => cnttSelectedSheets.add(s));
      }
      renderCnttSheetGrid();
      fetchOrAggregateCnttSheets();
    }

    function toggleSheetGridCollapse() {
      const grid = document.getElementById('cntt-sheets-grid');
      cnttIsGridCollapsed = !cnttIsGridCollapsed;
      if (grid) {
        grid.style.display = cnttIsGridCollapsed ? 'none' : 'flex';
      }
    }

    async function fetchOrAggregateCnttSheets(forceSync = false) {
      const sheetsArr = Array.from(cnttSelectedSheets);
      if (sheetsArr.length === 0) {
        cnttRawRecords = [];
        recomputeCnttAggregation();
        renderCnttDashboard();
        return;
      }

      // Check if all selected sheets are already in local cache
      const missingSheets = sheetsArr.filter(s => !CNTT_INITIAL_DATA[s]);
      if (missingSheets.length === 0 && !forceSync) {
        cnttRawRecords = [];
        sheetsArr.forEach(s => {
          if (CNTT_INITIAL_DATA[s]) cnttRawRecords = cnttRawRecords.concat(CNTT_INITIAL_DATA[s]);
        });
        recomputeCnttAggregation();
        renderCnttDashboard();
        showCnttToast('✓ THÀNH CÔNG', `Đã tổng hợp ${cnttRawRecords.length} ca từ ${sheetsArr.length} sheet!`);
        return;
      }

      // Fetch from API
      try {
        const url = `/api/cntt-summary?action=cntt-summary&sheets=${encodeURIComponent(sheetsArr.join(','))}${forceSync ? '&sync=1' : ''}`;
        const resp = await fetch(url, { headers: { 'X-App-Password': '8889' } });
        if (resp.ok) {
          const res = await resp.json();
          if (res.records) {
            cnttRawRecords = res.records;
            recomputeCnttAggregation();
            renderCnttDashboard();
            showCnttToast('✓ THÀNH CÔNG', `Đã tổng hợp ${cnttRawRecords.length} ca từ ${sheetsArr.length} sheet!`);
          }
        }
      } catch (err) {
        console.error('Lỗi tải dữ liệu CNTT:', err);
      }
    }

    function recomputeCnttAggregation() {
      const staffMap = {};
      const depts = new Set();
      let otherCount = 0;

      cnttRawRecords.forEach(r => {
        const cb = r.can_bo || 'Chưa gán';
        if (!staffMap[cb]) {
          staffMap[cb] = {
            name: cb,
            cases: 0,
            sheets: new Set(),
            departments: new Set(),
            other_count: 0,
            categories: {},
            other_items: []
          };
        }
        const st = staffMap[cb];
        st.cases++;
        st.sheets.add(r.sheet);
        if (r.khoa) {
          st.departments.add(r.khoa);
          depts.add(r.khoa);
        }
        if (r.sua_khac) {
          st.other_count++;
          otherCount++;
          st.other_items.push({
            sheet: r.sheet,
            khoa: r.khoa,
            content: r.sua_khac,
            yc: r.yeu_cau,
            hs: r.hs
          });
        }
        if (r.categories) {
          Object.entries(r.categories).forEach(([k, v]) => {
            st.categories[k] = (st.categories[k] || 0) + v;
          });
        }
      });

      const staffList = Object.values(staffMap).map(st => {
        const sortedCats = Object.entries(st.categories).sort((a, b) => b[1] - a[1]);
        return {
          ...st,
          sheets_count: st.sheets.size,
          sheets_list: Array.from(st.sheets).sort().reverse(),
          departments_count: st.departments.size,
          categories: sortedCats.map(([name, count]) => ({ name, count }))
        };
      });
      staffList.sort((a, b) => b.cases - a.cases);

      cnttAggregatedData = {
        total_cases: cnttRawRecords.length,
        selected_sheets_count: cnttSelectedSheets.size,
        departments_count: depts.size,
        staff_count: staffList.length,
        other_count: otherCount,
        staff_list: staffList
      };
    }

    function renderCnttDashboard() {
      if (!cnttAggregatedData) recomputeCnttAggregation();
      const d = cnttAggregatedData;

      // Update KPI cards
      const kpiSheets = document.getElementById('cntt-kpi-sheets');
      const kpiSheetsSub = document.getElementById('cntt-kpi-sheets-sub');
      const kpiCases = document.getElementById('cntt-kpi-cases');
      const kpiCasesSub = document.getElementById('cntt-kpi-cases-sub');
      const kpiStaff = document.getElementById('cntt-kpi-staff');
      const kpiOther = document.getElementById('cntt-kpi-other');
      const statusMsg = document.getElementById('cntt-status-msg');
      const tabRecCount = document.getElementById('cntt-tab-rec-count');

      const sheetsListText = Array.from(cnttSelectedSheets).join(', ') || 'Chưa chọn';
      if (kpiSheets) kpiSheets.textContent = d.selected_sheets_count;
      if (kpiSheetsSub) kpiSheetsSub.textContent = sheetsListText;
      if (kpiCases) kpiCases.textContent = d.total_cases;
      if (kpiCasesSub) kpiCasesSub.textContent = `Bao gồm ${d.departments_count} khoa/phòng`;
      if (kpiStaff) kpiStaff.textContent = d.staff_count;
      if (kpiOther) kpiOther.textContent = d.other_count;
      if (statusMsg) statusMsg.textContent = `✓ Đã tổng hợp thành công ${d.total_cases} ca từ ${d.selected_sheets_count} sheet!`;
      if (tabRecCount) tabRecCount.textContent = (cnttRawRecords.filter(r => r.sheet === '29.10').length) || d.total_cases;

      // Filters
      const filterCb = (document.getElementById('cntt-filter-staff')?.value || '').trim();
      const searchKw = (document.getElementById('cntt-search-keyword')?.value || '').trim().toLowerCase();

      // Render Staff Cards
      const container = document.getElementById('cntt-staff-cards-list');
      if (!container) return;
      container.innerHTML = '';

      d.staff_list.forEach(st => {
        if (filterCb && st.name !== filterCb) return;

        // Filter other items by keyword
        const filteredOtherItems = st.other_items.filter(it => {
          if (!searchKw) return true;
          return it.content.toLowerCase().includes(searchKw) || 
                 it.khoa.toLowerCase().includes(searchKw) || 
                 (it.hs && it.hs.includes(searchKw));
        });

        // Filter categories by keyword
        const filteredCats = st.categories.filter(c => {
          if (!searchKw) return true;
          return c.name.toLowerCase().includes(searchKw);
        });

        if (searchKw && filteredOtherItems.length === 0 && filteredCats.length === 0 && !st.name.toLowerCase().includes(searchKw)) {
          return;
        }

        const initial = st.name.charAt(0).toUpperCase();
        const card = document.createElement('div');
        card.className = 'cntt-staff-card';

        // Badges
        const sheetsBadgeText = `${st.sheets_count} sheet (${st.sheets_list.join(', ')})`;

        // Build HTML
        card.innerHTML = `
          <div class="cntt-staff-header">
            <div class="cntt-staff-left">
              <div class="cntt-avatar">${initial}</div>
              <div class="cntt-staff-name">${st.name}</div>
              <span class="cntt-tag-badge cntt-tag-blue">${st.cases} ca</span>
              <span class="cntt-tag-badge cntt-tag-slate">${sheetsBadgeText}</span>
              <span class="cntt-tag-badge cntt-tag-cyan">${st.departments_count} khoa/phòng</span>
              <span class="cntt-tag-badge cntt-tag-amber">${st.other_count} sửa chữa khác</span>
            </div>
            <button class="cntt-btn-copy" onclick="copyStaffReport('${st.name}')">
              📋 Copy Báo Cáo Cán Bộ
            </button>
          </div>

          <!-- Section Categories -->
          <div class="cntt-sec-title">📌 SỐ LƯỢNG CÁC MỤC ĐÃ THỰC HIỆN:</div>
          <div class="cntt-cats-container">
            ${filteredCats.map(c => `
              <div class="cntt-cat-pill">
                <span>${c.name}</span>
                <span class="cntt-cat-count">${c.count}</span>
              </div>
            `).join('')}
          </div>

          <!-- Section Other Items -->
          ${filteredOtherItems.length > 0 ? `
            <div class="cntt-sec-title accent-amber">🛠 Chi Tiết Sửa Chữa Khác (${filteredOtherItems.length} nội dung cụ thể):</div>
            <div class="cntt-other-list">
              ${filteredOtherItems.map((it, idx) => `
                <div class="cntt-other-row">
                  <span class="cntt-other-idx">${idx + 1}.</span>
                  <span class="badge-sheet">Sheet ${it.sheet}</span>
                  <span class="badge-dept">${it.khoa}</span>
                  <span style="color:#ffffff; font-weight:600;">${it.content}</span>
                  ${it.yc ? `<span class="badge-meta">Y/C: ${it.yc}</span>` : ''}
                  ${it.hs ? `<span class="badge-hs">HS: ${it.hs}</span>` : ''}
                </div>
              `).join('')}
            </div>
          ` : ''}
        `;
        container.appendChild(card);
      });

      renderCnttRawTable();
      renderCnttStats();
    }

    function copyStaffReport(staffName) {
      if (!cnttAggregatedData) return;
      const st = cnttAggregatedData.staff_list.find(s => s.name === staffName);
      if (!st) return;

      const lines = [
        `BÁO CÁO CÔNG VIỆC CÁN BỘ: ${st.name.toUpperCase()}`,
        `- Sheets tổng hợp: ${st.sheets_list.join(', ')}`,
        `- Tổng số ca tiếp nhận: ${st.cases} ca (${st.departments_count} khoa/phòng)`,
        `- Sửa chữa khác: ${st.other_count} ca`,
        '',
        '📌 NỘI DUNG CÁC MỤC ĐÃ THỰC HIỆN:'
      ];

      st.categories.forEach((c, idx) => {
        lines.push(`${idx + 1}. ${c.name}: ${c.count}`);
      });

      if (st.other_items.length > 0) {
        lines.push('');
        lines.push(`🛠 CHI TIẾT SỬA CHỮA KHÁC (${st.other_items.length} nội dung):`);
        st.other_items.forEach((it, idx) => {
          const metaYc = it.yc ? ` | Y/C: ${it.yc}` : '';
          const metaHs = it.hs ? ` | HS: ${it.hs}` : '';
          lines.push(`${idx + 1}. Sheet ${it.sheet} | ${it.khoa} | ${it.content}${metaYc}${metaHs}`);
        });
      }

      const reportText = lines.join(String.fromCharCode(10));
      navigator.clipboard.writeText(reportText).then(() => {
        showCnttToast('✓ ĐÃ SAO CHÉP', `Đã sao chép báo cáo của cán bộ ${st.name} vào bộ nhớ tạm!`);
      }).catch(() => {
        showCnttToast('✓ THÔNG BÁO', `Báo cáo: ${st.cases} ca của ${st.name}`);
      });
    }

    function renderCnttRawTable() {
      const tbody = document.getElementById('cntt-table-body');
      const totalEl = document.getElementById('cntt-table-total');
      if (!tbody) return;
      tbody.innerHTML = '';
      if (totalEl) totalEl.textContent = cnttRawRecords.length;

      cnttRawRecords.slice(0, 150).forEach((r, idx) => {
        const catSummary = Object.entries(r.categories || {}).map(([k, v]) => `${k} (${v})`).join(', ');
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${idx + 1}</td>
          <td><b>${r.sheet}</b></td>
          <td><span class="badge-dept">${r.khoa}</span></td>
          <td><code>${r.hs || '-'}</code></td>
          <td><b>${r.can_bo}</b></td>
          <td>${r.yeu_cau || '-'}</td>
          <td style="color:#fbbf24;">${r.sua_khac || '-'}</td>
          <td style="font-size:11px; color:#94a3b8;">${catSummary || '-'}</td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderCnttStats() {
      const deptContainer = document.getElementById('cntt-stat-dept-list');
      const catContainer = document.getElementById('cntt-stat-cat-list');
      if (!deptContainer || !catContainer) return;

      const deptMap = {};
      const catMap = {};
      cnttRawRecords.forEach(r => {
        if (r.khoa) deptMap[r.khoa] = (deptMap[r.khoa] || 0) + 1;
        if (r.categories) {
          Object.entries(r.categories).forEach(([k, v]) => {
            catMap[k] = (catMap[k] || 0) + v;
          });
        }
      });

      const sortedDepts = Object.entries(deptMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
      deptContainer.innerHTML = sortedDepts.map(([name, count], i) => `
        <div style="display:flex; justify-content:space-between; padding:6px 10px; border-bottom:1px solid #334155; font-size:12px;">
          <span>#${i + 1}. <b>${name}</b></span>
          <span class="badge" style="background:#0284c7; color:white;">${count} ca</span>
        </div>
      `).join('');

      const sortedCats = Object.entries(catMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
      catContainer.innerHTML = sortedCats.map(([name, count], i) => `
        <div style="display:flex; justify-content:space-between; padding:6px 10px; border-bottom:1px solid #334155; font-size:12px;">
          <span>#${i + 1}. ${name}</span>
          <span class="badge" style="background:#a855f7; color:white;">${count}</span>
        </div>
      `).join('');
    }

    function switchCnttSubTab(tabId) {
      document.querySelectorAll('.cntt-subnav .cntt-tab-item').forEach(el => {
        el.classList.remove('active');
      });
      event.currentTarget.classList.add('active');

      const tabs = ['tong-hop-sheet', 'danh-sach', 'thong-ke', 'google-sheet', 'diem-ca'];
      tabs.forEach(t => {
        const pane = document.getElementById(`cntt-subtab-${t}`);
        if (pane) pane.style.display = (t === tabId) ? 'block' : 'none';
      });
    }

    async function refreshCnttLive() {
      showCnttToast('⏳ ĐỒNG BỘ', 'Đang nạp dữ liệu mới nhất từ Google Sheets...');
      await fetchOrAggregateCnttSheets(true);
      const now = new Date();
      const h = String(now.getHours()).padStart(2, '0');
      const m = String(now.getMinutes()).padStart(2, '0');
      const clockEl = document.getElementById('cntt-sync-clock');
      if (clockEl) clockEl.textContent = `${h}:${m}`;
    }

    function exportCnttExcel() {
      // Trigger download for current active sheets
      const sheets = Array.from(cnttSelectedSheets).join(',');
      window.open(`/api/process-cntt?sheet=${encodeURIComponent(sheets)}`, '_blank');
      showCnttToast('📄 XUẤT EXCEL', 'Đang tạo file Excel Báo Cáo P.CNTT...');
    }

    function showCnttToast(title, msg) {
      const toast = document.getElementById('cntt-toast');
      const tTitle = document.getElementById('cntt-toast-title');
      const tMsg = document.getElementById('cntt-toast-msg');
      if (!toast) return;

      if (tTitle) tTitle.textContent = title;
      if (tMsg) tMsg.textContent = msg;

      toast.classList.add('show');
      setTimeout(() => {
        toast.classList.remove('show');
      }, 3500);
    }

</script>
</body>
</html>
"""

# ============================================================
# API HANDLERS CHO CẢ 8 TÁC VỤ
# ============================================================
def handle_get_sheets():
    try:
        if "file" in request.files and request.files["file"].filename:
            f = request.files["file"]
            wb = load_workbook(f.stream, read_only=True)
            sheets = wb.sheetnames
        elif request.form.get("url"):
            url = request.form.get("url").strip()
            sheets = get_google_sheet_names(url)
        else:
            return jsonify({"error": "Vui lòng cung cấp file hoặc URL Google Sheets."}), 400
        return jsonify({"sheets": sheets})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_1():
    try:
        word_files = request.files.getlist("word_files")
        doc_list = extract_word_documents_from_files(word_files)
        if not doc_list:
            return jsonify({"error": "Không tìm thấy file Word (.docx) hợp lệ nào để xử lý!"}), 400

        sso_map = {}
        if "sso_file" in request.files and request.files["sso_file"].filename:
            sso_map = read_sso_excel_stream(request.files["sso_file"].stream)

        all_people = []
        for fname, dstream in doc_list:
            people = parse_people_from_word_stream(fname, dstream, sso_map)
            all_people.extend(people)

        if not all_people:
            return jsonify({"error": "Không trích xuất được bản ghi nào từ các file Word đã chọn. Vui lòng kiểm tra lại cấu trúc bảng trong file!"}), 400

        out_io = build_vgca_19_col_excel(all_people)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="Ket_qua_VGCA.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_2():
    try:
        word_files = request.files.getlist("word_files")
        doc_list = extract_word_documents_from_files(word_files)
        if not doc_list:
            return jsonify({"error": "Không tìm thấy file Word (.docx) hợp lệ nào!"}), 400

        sso_map = {}
        if "sso_file" in request.files and request.files["sso_file"].filename:
            sso_map = read_sso_excel_stream(request.files["sso_file"].stream)

        all_people = []
        for fname, dstream in doc_list:
            people = parse_people_from_word_stream(fname, dstream, sso_map)
            all_people.extend(people)

        if not all_people:
            return jsonify({"error": "Không trích xuất được bản ghi nào từ các file Word!"}), 400

        out_io = build_mau_02_docx(all_people)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="Mẫu 02_CV cap CKS cho BVDK BN số 2_T10 2026_3_6.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_3():
    try:
        word_files = request.files.getlist("word_files")
        doc_list = extract_word_documents_from_files(word_files)
        if not doc_list:
            return jsonify({"error": "Không tìm thấy file Word (.docx) hợp lệ nào!"}), 400

        all_people = []
        for fname, dstream in doc_list:
            people = parse_people_from_word_stream(fname, dstream, sso_map=None)
            all_people.extend(people)

        if not all_people:
            return jsonify({"error": "Không trích xuất được bản ghi nào từ các file Word!"}), 400

        out_io = build_mau_02_docx(all_people)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="Mẫu 02_CV cap CKS cho BVDK BN số 2_T10 2026_3_6.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_9():
    try:
        word_files = request.files.getlist("word_files")
        doc_list = extract_word_documents_from_files(word_files)
        if not doc_list:
            return jsonify({"error": "Không tìm thấy file Word (.docx) hợp lệ nào!"}), 400

        all_people = []
        for fname, dstream in doc_list:
            people = parse_people_from_word_stream(fname, dstream)
            all_people.extend(people)

        if not all_people:
            return jsonify({"error": "Không trích xuất được bản ghi nào từ các file Word!"}), 400

        fmt = request.form.get("format", "word")
        if fmt == "word":
            out_io = build_email_docx(all_people)
            if len(all_people) == 1:
                p_name = all_people[0].get("name", "").strip()
                out_name = f"Đề nghị cấp mới địa chỉ email công vụ_{p_name}.docx"
            else:
                out_name = "Đề nghị cấp mới địa chỉ email công vụ.docx"
            
            quoted_name = urllib.parse.quote(out_name)
            resp = make_response(send_file(
                out_io,
                as_attachment=True,
                download_name=out_name,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ))
            resp.headers["Content-Disposition"] = f"attachment; filename=\"{out_name.encode('ascii', 'ignore').decode()}\"; filename*=UTF-8''{quoted_name}"
            resp.headers["X-Filename"] = quoted_name
            return resp
        else:
            out_io = build_email_excel(all_people)
            if len(all_people) == 1:
                p_name = all_people[0].get("name", "").strip()
                out_name = f"Danh_sach_cap_email_cong_vu_{p_name}.xlsx"
            else:
                out_name = "Danh_sach_cap_email_cong_vu.xlsx"

            quoted_name = urllib.parse.quote(out_name)
            resp = make_response(send_file(
                out_io,
                as_attachment=True,
                download_name=out_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ))
            resp.headers["Content-Disposition"] = f"attachment; filename=\"{out_name.encode('ascii', 'ignore').decode()}\"; filename*=UTF-8''{quoted_name}"
            resp.headers["X-Filename"] = quoted_name
            return resp
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_4():
    try:
        sheet_name = request.form.get("sheet_name")
        if not sheet_name:
            return jsonify({"error": "Vui lòng chọn Sheet cần xử lý."}), 400
        if "file" not in request.files or not request.files["file"].filename:
            return jsonify({"error": "Vui lòng tải lên file Excel báo cáo P.CNTT."}), 400

        out_io = process_cntt_web(request.files["file"].stream, sheet_name)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_5():
    try:
        if "file" not in request.files or not request.files["file"].filename:
            return jsonify({"error": "Vui lòng tải lên file Excel Check Log MDINSIGHT."}), 400

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

def handle_task_6():
    try:
        url = request.form.get("url")
        sheet_name = request.form.get("sheet_name")
        if not url:
            return jsonify({"error": "Vui lòng nhập link Google Sheets."}), 400
        if not sheet_name:
            return jsonify({"error": "Vui lòng chọn Sheet ngày cần xử lý."}), 400

        try:
            stream = download_google_sheet_single_tab(url.strip(), sheet_name)
        except Exception:
            stream = download_google_sheet(url.strip())
        out_io = process_cntt_web(stream, sheet_name)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_7():
    try:
        word_files = request.files.getlist("word_files")
        doc_list = extract_word_documents_from_files(word_files)
        if not doc_list:
            return jsonify({"error": "Không tìm thấy file Word (.docx) hợp lệ nào!"}), 400

        all_people = []
        for fname, dstream in doc_list:
            people = parse_people_from_word_stream(fname, dstream)
            all_people.extend(people)

        if not all_people:
            return jsonify({"error": "Không trích xuất được bản ghi nào từ các file Word!"}), 400

        out_io = build_vgca_19_col_excel(all_people)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="Ket_qua.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_task_8():
    try:
        sheet_name = request.form.get("sheet_name")
        if not sheet_name:
            return jsonify({"error": "Vui lòng chọn Sheet ngày cần xử lý."}), 400

        url = request.form.get("url") or DEFAULT_FIXED_GSHEET_URL
        try:
            stream = download_google_sheet_single_tab(url.strip(), sheet_name)
        except Exception:
            stream = download_google_sheet(url.strip())
        out_io = process_cntt_web(stream, sheet_name)
        return send_file(
            out_io,
            as_attachment=True,
            download_name="TỔNG HỢP CÔNG VIỆC P.CNTT (CỐ ĐỊNH).xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ============================================================
# LOGIC TÁC VỤ 10: THỐNG KÊ & LỌC LỖI XML BHYT (GOOGLE SHEETS)
# ============================================================
def get_dm_khoa_mapping(url=None):
    url = url or DEFAULT_ERROR_TRACKER_URL
    sid = extract_spreadsheet_id(url)
    if sid in _DM_KHOA_CACHE:
        return _DM_KHOA_CACHE[sid]

    encoded = urllib.parse.quote('Danh mục khoa phòng')
    csv_url = f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet={encoded}"
    mapping = {}
    try:
        req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=12) as resp:
            content = resp.read().decode('utf-8', errors='ignore')
        reader = csv.reader(io.StringIO(content))
        rows = list(reader)
        for r in rows[1:]:
            if len(r) >= 3 and r[1].strip() and r[2].strip():
                code = r[1].strip().upper()
                name = r[2].strip()
                mapping[code] = name
        _DM_KHOA_CACHE[sid] = mapping
    except Exception as e:
        print("Error get_dm_khoa_mapping:", e)
    return mapping

def get_error_tracker_sheets(url=None):
    url = url or DEFAULT_ERROR_TRACKER_URL
    sid = extract_spreadsheet_id(url)
    if sid in _TRACKER_SHEETS_CACHE:
        return _TRACKER_SHEETS_CACHE[sid]

    html_url = f"https://docs.google.com/spreadsheets/d/{sid}/edit"
    req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    clean_sheets = []
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        sheets = re.findall(r'class="[^"]*docs-sheet-tab-caption[^"]*">([^<]+)<', html)
        seen = set()
        for s in sheets:
            s_clean = s.strip()
            if s_clean and s_clean not in seen and s_clean != 'Danh mục khoa phòng' and not s_clean.startswith('Copy'):
                seen.add(s_clean)
                clean_sheets.append(s_clean)
    except Exception as e:
        print("Error get_error_tracker_sheets:", e)

    clean_sheets.reverse()
    if clean_sheets:
        _TRACKER_SHEETS_CACHE[sid] = clean_sheets
    return clean_sheets

def parse_error_tracker_sheet(url, sheet_name):
    url = url or DEFAULT_ERROR_TRACKER_URL
    sid = extract_spreadsheet_id(url)
    dm_map = get_dm_khoa_mapping(url)

    encoded = urllib.parse.quote(sheet_name)
    now_ts = int(time.time() * 1000)
    csv_url = f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet={encoded}&_t={now_ts}"

    req = urllib.request.Request(csv_url, headers={
        'User-Agent': 'Mozilla/5.0',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
    })
    with urllib.request.urlopen(req, timeout=15) as resp:
        content = resp.read().decode('utf-8', errors='ignore')
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    if not rows:
        return {"records": [], "stats": {}, "khoa_list": [], "phong_list": [], "xml_list": []}

    header = rows[0]
    hs_col = -1
    ngay_col = -1
    xml_col = -1
    loi_col = -1
    status_col = -1
    phong_cols = []
    khoa_cols = []

    for idx, c in enumerate(header):
        c_clean = c.strip().upper()
        if c_clean == 'HS': hs_col = idx
        elif ('NGÀY' in c_clean or 'NGAY' in c_clean) and ngay_col == -1: ngay_col = idx
        elif 'XML' in c_clean and xml_col == -1: xml_col = idx
        elif c_clean in ('LỖI', 'LOI'): loi_col = idx
        elif ('LỖI' in c_clean or 'LOI' in c_clean) and 'TRÙNG' not in c_clean and 'TRUNG' not in c_clean and loi_col == -1:
            loi_col = idx
        elif 'PHÒNG' in c_clean or 'PHONG' in c_clean:
            phong_cols.append(idx)
        elif 'KHOA' in c_clean:
            khoa_cols.append(idx)
        elif any(k in c_clean for k in ['TRẠNG THÁI', 'TRANG THAI', 'KẾT QUẢ XỬ LÝ', 'KET QUA XU LY']) and status_col == -1:
            status_col = idx

    if status_col == -1:
        for idx, c in enumerate(header):
            c_u = c.strip().upper()
            if 'XỬ LÝ' in c_u or 'XU LY' in c_u or c_u in ('KẾT QUẢ', 'KET QUA'):
                status_col = idx
                break

    if status_col == -1 and len(header) > 12:
        status_col = 12

    if hs_col == -1: hs_col = 1
    if ngay_col == -1: ngay_col = 2
    if xml_col == -1: xml_col = 3
    if loi_col == -1: loi_col = 4

    records = []
    for r_idx, r in enumerate(rows[1:], start=1):
        if not any(r): continue
        hs = r[hs_col].strip() if hs_col < len(r) else ''
        if hs.endswith('.0'): hs = hs[:-2]
        loi = r[loi_col].strip() if loi_col < len(r) else ''
        if not hs and not loi: continue

        ngay = r[ngay_col].strip() if ngay_col < len(r) else ''
        xml = r[xml_col].strip() if xml_col < len(r) else ''
        if xml.upper().startswith('XML'):
            xml = xml.upper()

        phong = ''
        for p_i in phong_cols:
            if p_i < len(r) and r[p_i].strip():
                phong = r[p_i].strip()
                break

        khoa_raw = ''
        for k_i in khoa_cols:
            if k_i < len(r) and r[k_i].strip():
                khoa_raw = r[k_i].strip()
                break

        khoa_norm = khoa_raw.upper()
        if khoa_norm in dm_map:
            khoa = f"{dm_map[khoa_norm]} ({khoa_raw})"
        else:
            khoa = khoa_raw if khoa_raw else "Chưa xác định"

        sheet_row = r_idx + 1
        status_val = r[status_col].strip() if status_col != -1 and status_col < len(r) else ''
        is_ok = status_val.upper() in ('OK', 'XONG', 'ĐÃ XỬ LÝ', 'DA XU LY', 'TRUE', '1')

        records.append({
            'stt': r_idx,
            'sheet_row': sheet_row,
            'hs': hs,
            'ngay': ngay,
            'xml': xml,
            'loi': loi,
            'phong': phong,
            'khoa': khoa,
            'khoa_raw': khoa_raw,
            'status': status_val,
            'is_ok': is_ok
        })

    from collections import Counter
    total_records = len(records)
    unique_hs = len(set(r['hs'] for r in records if r['hs']))
    total_ok = sum(1 for r in records if r.get('is_ok'))
    total_pending = total_records - total_ok

    khoa_counter = Counter(r['khoa'] for r in records if r['khoa'])
    phong_counter = Counter(r['phong'] for r in records if r['phong'])
    xml_counter = Counter(r['xml'] for r in records if r['xml'])

    error_patterns = []
    for r in records:
        l = r['loi']
        first_line = l.split('\n')[0].strip()
        m_rule = re.match(r'([A-Z0-9_]+(?:\s+[0-9.]+)?(?:[^:\d]+)?)', first_line)
        pattern = m_rule.group(1).strip() if m_rule else first_line[:60]
        if len(pattern) > 75: pattern = pattern[:75] + "..."
        error_patterns.append(pattern)
    loi_counter = Counter(error_patterns)

    khoa_list = [{"name": k, "count": v} for k, v in khoa_counter.most_common()]
    phong_list = [{"name": k, "count": v} for k, v in phong_counter.most_common()]
    xml_list = [{"name": k, "count": v} for k, v in xml_counter.most_common()]
    top_errors = [{"rule": k, "count": v} for k, v in loi_counter.most_common(10)]

    return {
        "sheet": sheet_name,
        "total_records": total_records,
        "unique_hs": unique_hs,
        "total_ok": total_ok,
        "total_pending": total_pending,
        "total_khoa": len(khoa_list),
        "total_phong": len(phong_list),
        "records": records,
        "khoa_list": khoa_list,
        "phong_list": phong_list,
        "xml_list": xml_list,
        "top_errors": top_errors
    }

def build_error_tracker_excel(records, sheet_name, filter_desc=""):
    wb = Workbook()
    ws = wb.active
    ws.title = f"Lỗi XML {sheet_name}"[:31]

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_view.showGridLines = True

    ws.merge_cells("A1:H1")
    ws["A1"] = "BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2 - PHÒNG CÔNG NGHỆ THÔNG TIN"
    ws["A1"].font = Font(name="Times New Roman", size=11, italic=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A2:H2")
    ws["A2"] = f"BÁO CÁO THỐNG KÊ LỖI XML ĐẨY CỔNG GIÁM ĐỊNH BHYT (NGÀY {sheet_name})"
    ws["A2"].font = Font(name="Times New Roman", size=14, bold=True, color="1E3A8A")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 28

    ws.merge_cells("A3:H3")
    desc = f"Ngày dữ liệu: {sheet_name} • Tổng số bản ghi lỗi: {len(records)}"
    if filter_desc:
        desc += f" • Lọc: {filter_desc}"
    ws["A3"] = desc
    ws["A3"].font = Font(name="Times New Roman", size=11, italic=True, color="4B5563")
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 20

    headers = ["STT", "Mã Hồ Sơ", "Ngày Giờ Đẩy", "Bảng XML", "Khoa / Bộ Phận", "Phòng Khám", "Nội Dung Lỗi Chi Tiết", "Trạng Thái Xử Lý"]
    header_row = 4
    ws.row_dimensions[header_row].height = 26

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = Font(name="Times New Roman", size=12, bold=True, color="0F172A")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        cell.fill = header_fill

    for r_idx, r in enumerate(records, start=header_row + 1):
        ws.row_dimensions[r_idx].height = 24
        st_val = r.get("status", "")
        if not st_val:
            st_val = "OK" if r.get("is_ok") else "Chưa xử lý"
        vals = [
            r_idx - header_row,
            r.get("hs", ""),
            r.get("ngay", ""),
            r.get("xml", ""),
            r.get("khoa", ""),
            r.get("phong", ""),
            r.get("loi", ""),
            st_val
        ]
        for c_idx, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.font = Font(name="Times New Roman", size=11)
            cell.border = thin_border
            if c_idx in (1, 4):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (2, 3):
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "@"
            elif c_idx in (5, 6):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 8:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if str(v).upper() == "OK":
                    cell.font = Font(name="Times New Roman", size=11, bold=True, color="047857")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 30
    ws.column_dimensions["G"].width = 65
    ws.column_dimensions["H"].width = 18

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def handle_error_tracker_sheets():
    try:
        url = request.args.get("url") or request.form.get("url") or DEFAULT_ERROR_TRACKER_URL
        if request.args.get("reload") == "1":
            sid = extract_spreadsheet_id(url)
            _TRACKER_SHEETS_CACHE.pop(sid, None)
            _DM_KHOA_CACHE.pop(sid, None)
        sheets = get_error_tracker_sheets(url)
        default_sheet = sheets[0] if sheets else ""
        return jsonify({
            "sheets": sheets,
            "default_sheet": default_sheet,
            "total_sheets": len(sheets)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_error_tracker_data():
    try:
        url = request.args.get("url") or request.form.get("url") or DEFAULT_ERROR_TRACKER_URL
        sheet_name = request.args.get("sheet") or request.form.get("sheet")
        if not sheet_name:
            sheets = get_error_tracker_sheets(url)
            sheet_name = sheets[0] if sheets else ""
        if not sheet_name:
            return jsonify({"error": "Không tìm thấy sheet nào để xử lý."}), 400

        data = parse_error_tracker_sheet(url, sheet_name)
        resp = jsonify(data)
        resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handle_error_tracker_export():
    try:
        body = request.get_json(silent=True) or request.form
        sheet_name = body.get("sheet") or "Bao_cao"
        records = body.get("records")

        if not records:
            url = body.get("url") or DEFAULT_ERROR_TRACKER_URL
            data = parse_error_tracker_sheet(url, sheet_name)
            records = data.get("records", [])

            filter_khoa = body.get("filter_khoa", "ALL")
            filter_phong = body.get("filter_phong", "ALL")
            filter_xml = body.get("filter_xml", "ALL")

            filtered = []
            for r in records:
                if filter_khoa != "ALL" and r.get("khoa") != filter_khoa: continue
                if filter_phong != "ALL" and r.get("phong") != filter_phong: continue
                if filter_xml != "ALL" and r.get("xml") != filter_xml: continue
                filtered.append(r)
            records = filtered

        f_desc_parts = []
        if body.get("filter_khoa") and body.get("filter_khoa") != "ALL":
            f_desc_parts.append(f"Khoa: {body.get('filter_khoa')}")
        if body.get("filter_phong") and body.get("filter_phong") != "ALL":
            f_desc_parts.append(f"Phòng: {body.get('filter_phong')}")
        if body.get("filter_xml") and body.get("filter_xml") != "ALL":
            f_desc_parts.append(f"XML: {body.get('filter_xml')}")
        filter_desc = " • ".join(f_desc_parts)

        out_io = build_error_tracker_excel(records, sheet_name, filter_desc)
        safe_name = re.sub(r'[^a-zA-Z0-9.-]', '_', sheet_name)
        return send_file(
            out_io,
            as_attachment=True,
            download_name=f"Bao_cao_loi_XML_{safe_name}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def parse_apps_script_response(content):
    if not content:
        return {}
    content = content.strip()
    if content.startswith(")]}'"):
        content = content[4:].strip()
    try:
        data = json.loads(content)
        if isinstance(data, str):
            try:
                data = json.loads(data)
            except Exception:
                pass
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    clean = content.replace(" ", "").replace("\r", "").replace("\n", "")
    if '"success":true' in clean or '"ping":"pong"' in clean:
        return {"success": True, "ping": "pong"}
    return {}

def handle_error_tracker_test_connection():
    try:
        data = {}
        if request.is_json:
            data = request.get_json(silent=True) or {}
        else:
            data = request.form.to_dict() or request.args.to_dict()

        script_url = data.get("script_url", "").strip() or get_tracker_apps_script_url()
        if not script_url:
            return jsonify({
                "success": False,
                "error": "Chưa có URL Apps Script. Vui lòng dán link Web App vào ô."
            })

        if not script_url.startswith("http"):
            return jsonify({
                "success": False,
                "error": "Đường link không hợp lệ (phải bắt đầu bằng https://)."
            })

        if "/edit" in script_url or "docs.google.com/spreadsheets" in script_url:
            return jsonify({
                "success": False,
                "error": "Bạn đang dán link trình duyệt (/edit) hoặc Google Sheet! Hãy bấm nút 'Triển khai (Deploy)' ➔ 'Triển khai mới' ➔ 'Ứng dụng web' và sao chép link kết thúc bằng '/exec'."
            })

        # Ping test to Apps Script
        test_url = f"{script_url}{'&' if '?' in script_url else '?'}action=ping"
        req = urllib.request.Request(test_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        with urllib.request.urlopen(req, timeout=12) as resp:
            content = resp.read().decode('utf-8', errors='ignore')

        res_json = parse_apps_script_response(content)
        if res_json.get("success") is True or "ping" in res_json:
            return jsonify({
                "success": True,
                "message": "Kết nối thành công 100%! Google Sheet đã sẵn sàng tự động điền."
            })
        elif res_json.get("error"):
            return jsonify({
                "success": False,
                "error": res_json.get("error")
            })

        if "<!DOCTYPE html>" in content or "<html" in content:
            if "accounts.google.com" in content or "Sign in" in content or "Đăng nhập" in content:
                return jsonify({
                    "success": False,
                    "error": "Google chặn quyền truy cập (yêu cầu đăng nhập). Khi Triển khai (Deploy), tại mục 'Người có quyền truy cập (Who has access)' bạn bắt buộc phải chọn 'Bất kỳ ai (Anyone)'."
                })
            return jsonify({
                "success": False,
                "error": "Google trả về trang đăng nhập HTML thay vì JSON. Vui lòng kiểm tra lại quyền truy cập 'Bất kỳ ai (Anyone)'."
            })

        return jsonify({
            "success": False,
            "error": f"Phản hồi không nhận diện được: {content[:120]}"
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "error": f"Không thể kết nối tới Google Apps Script: {str(e)}"
        })

def handle_error_tracker_update_status():
    try:
        data = {}
        if request.is_json:
            data = request.get_json(silent=True) or {}
        else:
            data = request.form.to_dict() or request.args.to_dict()

        sheet_name = data.get("sheet", "").strip()
        hs = data.get("hs", "").strip()
        row = data.get("row")
        status = data.get("status", "OK").strip()
        script_url = data.get("script_url", "").strip() or get_tracker_apps_script_url()

        # Luôn ghi nhận vào bộ nhớ cache hệ thống để giao diện cập nhật ngay lập tức
        if sheet_name and hs:
            _LOCAL_STATUS_CACHE[(sheet_name, hs)] = status

        if not script_url:
            return jsonify({
                "success": True,
                "synced_to_sheet": False,
                "need_config": True,
                "status": status,
                "error_detail": "Chưa kết nối Webhook Google Apps Script",
                "message": f"Đã lưu trạng thái '{status or 'Chưa xử lý'}' vào hệ thống."
            })

        if not script_url.startswith("http"):
            return jsonify({
                "success": True,
                "synced_to_sheet": False,
                "status": status,
                "error_detail": "Link Web App không đúng định dạng (phải bắt đầu bằng https://)"
            })

        if "/edit" in script_url or "docs.google.com/spreadsheets" in script_url:
            return jsonify({
                "success": True,
                "synced_to_sheet": False,
                "status": status,
                "error_detail": "Link đã lưu là link trang web/Google Sheet, không phải link Web App /exec! Bấm nút Cấu hình để chỉnh lại."
            })

        # Gửi cập nhật sang Google Apps Script Web App
        params = urllib.parse.urlencode({
            "sheet": sheet_name,
            "hs": hs,
            "row": row if row else "",
            "status": status
        })
        sep = "&" if "?" in script_url else "?"
        full_url = f"{script_url}{sep}{params}"

        req = urllib.request.Request(full_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        with urllib.request.urlopen(req, timeout=15) as resp:
            content = resp.read().decode('utf-8', errors='ignore')

        res_json = parse_apps_script_response(content)
        if res_json.get("success") is True:
            res_json["synced_to_sheet"] = True
            return jsonify(res_json)
        elif res_json.get("error"):
            return jsonify({
                "success": True,
                "synced_to_sheet": False,
                "status": status,
                "error_detail": f"Google Apps Script báo: {res_json.get('error')}"
            })

        if "<!DOCTYPE html>" in content or "<html" in content:
            if "accounts.google.com" in content or "Sign in" in content or "Đăng nhập" in content:
                detail = "Google chặn quyền truy cập (cần chọn 'Bất kỳ ai / Anyone' khi Triển khai)"
            elif not script_url.endswith("/exec"):
                detail = "Link Script phải kết thúc bằng '/exec', không được dùng link /edit"
            else:
                detail = "Google yêu cầu xác thực hoặc đăng nhập"
        else:
            detail = f"Phản hồi từ Google không đúng JSON: {content[:80]}"
        return jsonify({
            "success": True,
            "synced_to_sheet": False,
            "status": status,
            "error_detail": detail
        })

    except Exception as e:
        return jsonify({
            "success": True, 
            "synced_to_sheet": False, 
            "status": status,
            "error_detail": f"Không thể kết nối đến Web App: {str(e)}", 
            "message": "Đã lưu nội bộ trên hệ thống."
        })

def handle_error_tracker_config():
    try:
        if request.method == "POST":
            data = request.get_json(silent=True) or request.form.to_dict()
            script_url = data.get("script_url", "").strip()
            set_tracker_apps_script_url(script_url)
            return jsonify({"success": True, "script_url": script_url})
        else:
            return jsonify({"script_url": get_tracker_apps_script_url()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ============================================================
# MASTER DISPATCH ROUTER (TƯƠNG THÍCH MỌI CẤU HÌNH REWRITE VERCEL)
# ============================================================

# ============================================================
# LOGIC BC SỬA CNTT (TỔNG HỢP SHEET & BÁO CÁO CÁN BỘ)
# ============================================================
CNTT_SPREADSHEET_ID = "1_I7qmDx7mtIOcCLkEkDz_PCaC8KPd_Po-pqEh7krdPA"

CNTT_STAFF_MAP = {
    "dương": "Chu Thị Dương",
    "chu thị dương": "Chu Thị Dương",
    "ly": "Nguyễn Lê Huyền Ly",
    "huyền ly": "Nguyễn Lê Huyền Ly",
    "nguyễn lê huyền ly": "Nguyễn Lê Huyền Ly",
    "hiệp": "Đỗ Văn Hiệp",
    "đỗ văn hiệp": "Đỗ Văn Hiệp",
    "quyen": "Nguyễn Thị Quyên",
    "quyên": "Nguyễn Thị Quyên",
    "nguyễn thị quyên": "Nguyễn Thị Quyên",
    "huyền": "Nguyễn Thị Huyền",
    "nguyễn thị huyền": "Nguyễn Thị Huyền",
    "v phương": "Vũ Phương",
    "vũ phương": "Vũ Phương",
    "vũ thị phương": "Vũ Phương",
    "nhân": "Nguyễn Trọng Nhân",
    "nguyễn trọng nhân": "Nguyễn Trọng Nhân"
}

def normalize_cntt_staff(raw):
    if not raw or not str(raw).strip():
        return "Chưa gán"
    clean = str(raw).strip().lower()
    return CNTT_STAFF_MAP.get(clean, str(raw).strip())

CNTT_CAT_COLUMNS = [
    (4, "Xoá CLS, PTTT"),
    (5, "Ép thẻ BH"),
    (6, "Sửa phí, mã QR sang TM và ngược lại"),
    (7, "Huỷ CKS"),
    (8, "Huỷ B/án"),
    (9, "Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám"),
    (10, "Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS"),
    (11, "Sửa mã bệnh kèm theo, bệnh chính"),
    (12, "Hỗ trợ ký số PM, QLVBĐH"),
    (13, "Sửa giấy RV, kết thúc ra viện, chuyển tuyến"),
    (14, "Đăng bài trên website"),
    (15, "Sửa cách dùng thuốc, ngày giờ thuốc, vật tư"),
    (16, "Phân quyền USER, đổi nick user"),
    (17, "Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)"),
    (18, "Sửa ngày giờ, sửa xoá phiếu xuất khác"),
    (19, "Sửa, xoá ảnh trên PACS"),
    (20, "Mở HS ký số bổ sung"),
    (21, "Sửa xoá giường, cập nhật giường, thêm giường"),
    (22, "Chuyển thuốc CP-DG"),
    (23, "Sửa lỗi xuất toán tự động"),
    (24, "VIMES Xử Lý"),
    (25, "Đổ mực"),
    (26, "Cài Win"),
    (27, "Sửa lỗi mạng"),
    (28, "Sửa máy tính (Main, chip)"),
    (29, "Nguồn"),
    (30, "Ổ cứng"),
    (31, "Sửa máy in (Vỏ lụa)"),
    (32, "Lô ép"),
    (33, "Cảm biến"),
]

INITIAL_CNTT_RECORDS_CACHE = {"29.10": [{"sheet": "29.10", "stt": "1", "khoa": "NGOẠI ƯB", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hồng", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1, "Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "29.10", "stt": "2", "khoa": "TMH", "hs": "26228514", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hoa", "sua_khac": "Sửa mức hưởng BH", "categories": {}}, {"sheet": "29.10", "stt": "3", "khoa": "KB", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Sơn PK UB 219", "sua_khac": "", "categories": {"Hỗ trợ ký số PM, QLVBĐH": 1}}, {"sheet": "29.10", "stt": "4", "khoa": "NHH", "hs": "26221694", "phieu": "8043555", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Thùy", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "29.10", "stt": "5", "khoa": "NGOẠI XT", "hs": "26209471", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Huệ", "sua_khac": "", "categories": {"Huỷ CKS": 1}}, {"sheet": "29.10", "stt": "6", "khoa": "MẮT", "hs": "26228870", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Thu", "sua_khac": "Sửa mức hưởng BH", "categories": {}}, {"sheet": "29.10", "stt": "7", "khoa": "NGOẠI TKLN", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Huyền", "sua_khac": "Nhập BH cùng chi trả", "categories": {"Sửa xoá giường, cập nhật giường, thêm giường": 7}}, {"sheet": "29.10", "stt": "8", "khoa": "PHỤ SẢN", "hs": "", "phieu": "TL0109518", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Hiền", "sua_khac": "", "categories": {"Huỷ CKS": 1}}, {"sheet": "29.10", "stt": "9", "khoa": "DA LIỄU", "hs": "26229792", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Thủy", "sua_khac": "Sủa mức hưởng BH", "categories": {}}, {"sheet": "29.10", "stt": "10", "khoa": "TRUYỀN NHIỄM", "hs": "26225431", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "Chuyển trạng thái phiếu", "categories": {}}, {"sheet": "29.10", "stt": "11", "khoa": "LKTK", "hs": "26227036", "phieu": "804444", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Loan", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "29.10", "stt": "12", "khoa": "NHI", "hs": "26224147", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Thu Hiền", "sua_khac": "", "categories": {}}, {"sheet": "29.10", "stt": "13", "khoa": "TTTM", "hs": "26225561", "phieu": "8028816", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Nga", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "29.10", "stt": "14", "khoa": "NGOẠI TH", "hs": "26223088", "phieu": "8019001", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Dung", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "29.10", "stt": "15", "khoa": "KHTH", "hs": "26229981", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Kim Nội", "sua_khac": "", "categories": {"Huỷ B/án": 1}}, {"sheet": "29.10", "stt": "16", "khoa": "KHTH", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Kim Nội", "sua_khac": "sửa tên CSKCBBĐ 30122", "categories": {}}], "28.10": [{"sheet": "28.10", "stt": "1", "khoa": "THẬN", "hs": "26225416", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Quy", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "28.10", "stt": "2", "khoa": "KHTH", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Kim Nội", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 4}}, {"sheet": "28.10", "stt": "3", "khoa": "KB", "hs": "26228486", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Nhung PK N3", "sua_khac": "", "categories": {"Huỷ B/án": 1}}, {"sheet": "28.10", "stt": "4", "khoa": "THẬN", "hs": "26204563", "phieu": "2572041", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Sơn", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "5", "khoa": "MẮT", "hs": "26228137", "phieu": "8039270", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Hằng", "sua_khac": "Thêm diễn giải", "categories": {}}, {"sheet": "28.10", "stt": "6", "khoa": "NGOẠI CT", "hs": "26228161", "phieu": "8039422", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Chinh", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "28.10", "stt": "7", "khoa": "NGOẠI TKLN", "hs": "26226732", "phieu": "8033865", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Huyền Trang", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "28.10", "stt": "8", "khoa": "CTXH", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Nga", "sua_khac": "", "categories": {"Đăng bài trên website": 2}}, {"sheet": "28.10", "stt": "9", "khoa": "TTTM", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Máy trực", "sua_khac": "Kết nối in qua Lan", "categories": {}}, {"sheet": "28.10", "stt": "10", "khoa": "KHTH", "hs": "26228722", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Kim Nội", "sua_khac": "", "categories": {"Ép thẻ BH": 1}}, {"sheet": "28.10", "stt": "11", "khoa": "DƯỢC", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bích Hạnh", "sua_khac": "Hạ thuốc treo kho chính", "categories": {}}, {"sheet": "28.10", "stt": "12", "khoa": "LKTK", "hs": "26222801", "phieu": "8027542", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Vân", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "28.10", "stt": "13", "khoa": "TRUYỀN NHIỄM", "hs": "26221141", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "Sửa giờ kết thúc công khám ngoại TH", "categories": {}}, {"sheet": "28.10", "stt": "14", "khoa": "NGOẠI TKLN", "hs": "26220268", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Huyền", "sua_khac": "", "categories": {"Sửa xoá giường, cập nhật giường, thêm giường": 1}}, {"sheet": "28.10", "stt": "15", "khoa": "NHH", "hs": "26222373", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Liễu", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "28.10", "stt": "16", "khoa": "NỘI TH", "hs": "26228205", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Huế", "sua_khac": "", "categories": {"Huỷ CKS": 1}}, {"sheet": "28.10", "stt": "17", "khoa": "NGOẠI TN", "hs": "26223185", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hương", "sua_khac": "Sửa nơi ĐK KCB BĐ", "categories": {}}, {"sheet": "28.10", "stt": "18", "khoa": "NỘI TH", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Mai", "sua_khac": "Sửa lỗi không in được P866", "categories": {}}, {"sheet": "28.10", "stt": "19", "khoa": "ĐIỀU DƯỠNG", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Kiều", "sua_khac": "Kết nối máy in cho 2 máy tính", "categories": {}}, {"sheet": "28.10", "stt": "20", "khoa": "DƯỢC", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "E Thảo Quầy thuốc BH", "sua_khac": "Lỗi đơn thuốc trả về O ( 02 HS)", "categories": {}}, {"sheet": "28.10", "stt": "21", "khoa": "NGOẠI TN", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Sâm", "sua_khac": "Sửa lỗi máy in bị in liên tục", "categories": {}}, {"sheet": "28.10", "stt": "22", "khoa": "NHH", "hs": "26218168", "phieu": "2572984", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Phương", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "23", "khoa": "HSTC", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Loan", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 2}}, {"sheet": "28.10", "stt": "24", "khoa": "NGOẠI CT", "hs": "26217179", "phieu": "2556472", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "25", "khoa": "TRUYỀN NHIỄM", "hs": "26228936", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "28.10", "stt": "26", "khoa": "NHH", "hs": "26219481", "phieu": "2571206", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Thuận", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "27", "khoa": "NGOẠI CT", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Thảnh", "sua_khac": "Sửa nơi KCB BĐ 22370, 93078", "categories": {}}, {"sheet": "28.10", "stt": "28", "khoa": "NGOẠI ƯB", "hs": "26222162", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hồng", "sua_khac": "Sửa mức hưởng BH", "categories": {}}, {"sheet": "28.10", "stt": "29", "khoa": "KB", "hs": "26228417", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Huyền PK N5", "sua_khac": "", "categories": {"Sửa mã bệnh kèm theo, bệnh chính": 1}}, {"sheet": "28.10", "stt": "30", "khoa": "NGOẠI TKLN", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Huyền", "sua_khac": "Sửa mức hưởng BH; Sửa giờ khám bệnh", "categories": {}}, {"sheet": "28.10", "stt": "31", "khoa": "HTL", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Xâm", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 5}}, {"sheet": "28.10", "stt": "32", "khoa": "TCKT", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hoa", "sua_khac": "Sửa ngày phiếu tạm ứng sang ngày 27/10", "categories": {}}, {"sheet": "28.10", "stt": "33", "khoa": "CC115", "hs": "26228212", "phieu": "8039654", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bảo Anh", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "34", "khoa": "NHI", "hs": "26225394", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Loan", "sua_khac": "", "categories": {"Mở HS ký số bổ sung": 1}}, {"sheet": "28.10", "stt": "35", "khoa": "VT", "hs": "", "phieu": "(645', '657', '658', '356', 383 ,335\n'274', '659', '648', '671',\n '672', '727', '729','250',\n '710', '285', '728', '142', \n'448', '25', '60', '730', \n'746','284', '536', '1128', \n'748', '59', '712', '275', \n'157', '270','287', '446', \n'447', '731', '732', '243',\n '244', '450','47', '49', '48')", "can_bo": "Nguyễn Thị Quyên", "cb_raw": "Quyên", "yeu_cau": "Thái", "sua_khac": "ẩn hoạt động máy", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 43}}, {"sheet": "28.10", "stt": "36", "khoa": "NHI", "hs": "26225395", "phieu": "8028283", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bích", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "37", "khoa": "CĐHA", "hs": "26229136", "phieu": "8043108", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Sen", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "38", "khoa": "CC115", "hs": "", "phieu": "8043459\n8043461\n8043614", "can_bo": "Vũ Phương", "cb_raw": "V Phương", "yeu_cau": "Sơn\nHướng\nChinh", "sua_khac": "Sửa mạch huyết áp", "categories": {"Xoá CLS, PTTT": 3}}, {"sheet": "28.10", "stt": "39", "khoa": "HSTC", "hs": "26229146", "phieu": "2573795", "can_bo": "Vũ Phương", "cb_raw": "V Phương", "yeu_cau": "Thái", "sua_khac": "Cập nhật TT sang S", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "40", "khoa": "NGOẠI TN", "hs": "26225247", "phieu": "2571469", "can_bo": "Vũ Phương", "cb_raw": "V Phương", "yeu_cau": "Khải", "sua_khac": "Cập nhật TT sang S", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "28.10", "stt": "41", "khoa": "PTGMHS", "hs": "26229204", "phieu": "8043541\n8043549\n8043717,8043719,8043720,8043729,8043730", "can_bo": "Vũ Phương", "cb_raw": "V Phương", "yeu_cau": "Dư, Khánh Linh", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 5, "Sửa ngày giờ, sửa xoá phiếu xuất khác": 2}}], "27.10": [{"sheet": "27.10", "stt": "1", "khoa": "KB", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Luyên", "sua_khac": "Sửa địa chỉ TP BN", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 40}}, {"sheet": "27.10", "stt": "2", "khoa": "RHM", "hs": "26225404", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "E An", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "3", "khoa": "PHCN", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Nguyệt", "sua_khac": "Sửa mức hưởng BH ( 02 HS)", "categories": {}}, {"sheet": "27.10", "stt": "4", "khoa": "KHTH", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Kim Nội", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 2}}, {"sheet": "27.10", "stt": "5", "khoa": "NGOẠI TH", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Dương Nga", "sua_khac": "Bỏ xác nhận phiếu TLT0227181", "categories": {}}, {"sheet": "27.10", "stt": "6", "khoa": "LKTK", "hs": "26223173", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Quỳnh", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "7", "khoa": "TRUYỀN NHIỄM", "hs": "26216584", "phieu": "15102819; 15104940", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "27.10", "stt": "8", "khoa": "TTTM", "hs": "", "phieu": "7999113; 7978582", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Phúc", "sua_khac": "Thêm mã máy vào CTM ( 02 HS)", "categories": {}}, {"sheet": "27.10", "stt": "9", "khoa": "TTTM", "hs": "26223944", "phieu": "8033549", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Linh", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "27.10", "stt": "10", "khoa": "TRUYỀN NHIỄM", "hs": "26219362", "phieu": "8003942", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "27.10", "stt": "11", "khoa": "NHI", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hiền", "sua_khac": "Sửa phí công khám", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1, "Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "27.10", "stt": "12", "khoa": "NGOẠI TKLN", "hs": "26224191", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Bảo", "sua_khac": "", "categories": {"Huỷ CKS": 1}}, {"sheet": "27.10", "stt": "13", "khoa": "CXK", "hs": "26214158", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "E Thảo", "sua_khac": "Thêm ngày hẹn khám lại", "categories": {}}, {"sheet": "27.10", "stt": "14", "khoa": "NHH", "hs": "26219530", "phieu": "2569009", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Linh", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "27.10", "stt": "15", "khoa": "NGOẠI CT", "hs": "26049982", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Chinh", "sua_khac": "", "categories": {"Huỷ CKS": 1}}, {"sheet": "27.10", "stt": "16", "khoa": "NỘI TH", "hs": "26223944", "phieu": "8034907", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Đoàn Thảo", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "27.10", "stt": "17", "khoa": "CĐHA", "hs": "26223019", "phieu": "8021771", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Dung", "sua_khac": "thêm mã máy", "categories": {}}, {"sheet": "27.10", "stt": "18", "khoa": "KB", "hs": "26227352", "phieu": "8036310", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Nhung PK N3", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "27.10", "stt": "19", "khoa": "KHTH", "hs": "26225307", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Kim Nội", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "20", "khoa": "NGOẠI TN", "hs": "26227593", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Cuộc", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "27.10", "stt": "21", "khoa": "NGOẠI TKLN", "hs": "26222375", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Huyền", "sua_khac": "", "categories": {"Sửa xoá giường, cập nhật giường, thêm giường": 1}}, {"sheet": "27.10", "stt": "22", "khoa": "NHH", "hs": "26224992", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Duyên", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "27.10", "stt": "23", "khoa": "TMH", "hs": "26225643", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hoa", "sua_khac": "", "categories": {"Huỷ B/án": 1}}, {"sheet": "27.10", "stt": "24", "khoa": "TRUYỀN NHIỄM", "hs": "26216846", "phieu": "8018954", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "27.10", "stt": "25", "khoa": "CĐT", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "C Dương", "sua_khac": "Hỗ trợ phần mềm WPS", "categories": {}}, {"sheet": "27.10", "stt": "26", "khoa": "HSTC", "hs": "26219281", "phieu": "6608562, 6612828, 6612829, 6612830", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hiền", "sua_khac": "", "categories": {"Huỷ CKS": 4}}, {"sheet": "27.10", "stt": "27", "khoa": "KB", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Hạnh PK Nội tiết", "sua_khac": "", "categories": {"Sửa cách dùng thuốc, ngày giờ thuốc, vật tư": 2}}, {"sheet": "27.10", "stt": "28", "khoa": "HSTC", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Loan", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 2}}, {"sheet": "27.10", "stt": "29", "khoa": "TRUYỀN NHIỄM", "hs": "26216846", "phieu": "15104937", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Vui", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "27.10", "stt": "30", "khoa": "KB", "hs": "26227560", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Huệ PK CXK", "sua_khac": "Sửa mức hưởng BH", "categories": {}}, {"sheet": "27.10", "stt": "31", "khoa": "NGOẠI XT", "hs": "26218981", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Huệ", "sua_khac": "Thêm ngày hẹn khám lại", "categories": {}}, {"sheet": "27.10", "stt": "32", "khoa": "PTGMHS", "hs": "26223148", "phieu": "2564613", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Hảo", "sua_khac": "Chuyển trạng thái PT từ S về T", "categories": {}}, {"sheet": "27.10", "stt": "33", "khoa": "TDCN", "hs": "", "phieu": "8035656; 8037315; 8037493; 8037514; 8037853", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Hách", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 5}}, {"sheet": "27.10", "stt": "34", "khoa": "KB", "hs": "26227976", "phieu": "8038638", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Thanh PK TMH", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "27.10", "stt": "35", "khoa": "THẬN", "hs": "26209627", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Quy", "sua_khac": "Kiểm tra lỗi MDinsight", "categories": {}}, {"sheet": "27.10", "stt": "36", "khoa": "ĐIỀU DƯỠNG", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Kiều", "sua_khac": "Lỗi kẹt giấy, kết nối máy in", "categories": {}}, {"sheet": "27.10", "stt": "37", "khoa": "TRUYỀN NHIỄM", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "", "sua_khac": "Kéo dây mạng cho tòa B2, B3", "categories": {}}, {"sheet": "27.10", "stt": "38", "khoa": "KB", "hs": "26228051", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Khuyên", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "39", "khoa": "QLCL", "hs": "", "phieu": "", "can_bo": "Chưa gán", "cb_raw": "", "yeu_cau": "", "sua_khac": "Biên bản góp ý", "categories": {}}, {"sheet": "27.10", "stt": "40", "khoa": "PHỤ SẢN", "hs": "26226630", "phieu": "8033367", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Hiền", "sua_khac": "", "categories": {"Xoá CLS, PTTT": 1}}, {"sheet": "27.10", "stt": "41", "khoa": "KB", "hs": "26226307", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Thủy PK LKTK", "sua_khac": "", "categories": {"Sửa cách dùng thuốc, ngày giờ thuốc, vật tư": 1}}, {"sheet": "27.10", "stt": "42", "khoa": "TCKT", "hs": "26217377", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Vân", "sua_khac": "BN để quá hạn thanh toán k đẩy đc cổng yêu cầu sửa ngày kết thúc sang ngày 27/10/2026", "categories": {}}, {"sheet": "27.10", "stt": "43", "khoa": "NHI", "hs": "26226312", "phieu": "8034470", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Hiền", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "27.10", "stt": "44", "khoa": "HSTC", "hs": "26213465", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Loan", "sua_khac": "Sửa giờ kết thúc hố sơ", "categories": {}}, {"sheet": "27.10", "stt": "45", "khoa": "KB", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Quyền PK TM", "sua_khac": "", "categories": {"Sửa cách dùng thuốc, ngày giờ thuốc, vật tư": 2}}, {"sheet": "27.10", "stt": "46", "khoa": "TCKT", "hs": "26223747", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "A Trường", "sua_khac": "", "categories": {"Sửa phí, mã QR sang TM và ngược lại": 1}}, {"sheet": "27.10", "stt": "48", "khoa": "PHỤ SẢN", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Hiền", "sua_khac": "Hủy CKS số ra viện, giấy chứng nhận PTTT", "categories": {"Huỷ CKS": 2}}, {"sheet": "27.10", "stt": "49", "khoa": "NHI", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "C Hiền", "sua_khac": "Tách chứng từ ( 02 HS)", "categories": {}}, {"sheet": "27.10", "stt": "50", "khoa": "NGOẠI TN", "hs": "26213939", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Sâm", "sua_khac": "", "categories": {"Huỷ CKS": 1, "Sửa giấy RV, kết thúc ra viện, chuyển tuyến": 1}}, {"sheet": "27.10", "stt": "52", "khoa": "HTL", "hs": "", "phieu": "", "can_bo": "Nguyễn Lê Huyền Ly", "cb_raw": "Ly", "yeu_cau": "Xâm", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 2}}, {"sheet": "27.10", "stt": "53", "khoa": "CC115", "hs": "26220933", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hoa", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "56", "khoa": "NHH", "hs": "", "phieu": "2480147; 2478660", "can_bo": "Nguyễn Thị Quyên", "cb_raw": "Quyên", "yeu_cau": "Duyên", "sua_khac": "kiểm tra người thực hiện PTTT", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 2}}, {"sheet": "27.10", "stt": "57", "khoa": "KB", "hs": "", "phieu": "", "can_bo": "Nguyễn Thị Quyên", "cb_raw": "Quyên", "yeu_cau": "Sang", "sua_khac": "sửa 1.14\nmã bệnh thành L40.1", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 2}}, {"sheet": "27.10", "stt": "58", "khoa": "LKTK", "hs": "26227954", "phieu": "", "can_bo": "Nguyễn Thị Quyên", "cb_raw": "Quyên", "yeu_cau": "Nơi", "sua_khac": "Ép thẻ cls", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "59", "khoa": "BAN BVSK", "hs": "", "phieu": "", "can_bo": "Nguyễn Thị Quyên", "cb_raw": "Quyên", "yeu_cau": "Liên", "sua_khac": "", "categories": {"Sửa xoá giường, cập nhật giường, thêm giường": 1}}, {"sheet": "27.10", "stt": "60", "khoa": "PTGMHS", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Toại", "sua_khac": "Sửa lỗi chữ ký số USB không ký được", "categories": {}}, {"sheet": "27.10", "stt": "61", "khoa": "TDCN", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hồng Liên", "sua_khac": "", "categories": {"Phân quyền USER, đổi nick user": 1}}, {"sheet": "27.10", "stt": "62", "khoa": "TMH", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hoa", "sua_khac": "Sửa mức hưởng BH ( 02 HS)", "categories": {}}, {"sheet": "27.10", "stt": "63", "khoa": "NGOẠI TH", "hs": "26208107", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Dung", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "27.10", "stt": "64", "khoa": "KB", "hs": "26227241", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Bs Huệ PK CXK", "sua_khac": "", "categories": {"Sửa cách dùng thuốc, ngày giờ thuốc, vật tư": 1}}, {"sheet": "27.10", "stt": "65", "khoa": "PHỤ SẢN", "hs": "26210980", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "E Ngọc", "sua_khac": "", "categories": {"Lỗi không kết thúc hồ sơ (tồn thuốc, thuốc chưa cấp)": 1}}, {"sheet": "27.10", "stt": "66", "khoa": "CC115", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "Hoa", "sua_khac": "Sửa mã ĐT KCB: cấp cứu ( 02 HS)", "categories": {}}, {"sheet": "27.10", "stt": "67", "khoa": "KB", "hs": "", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "BS Lan Anh PK Mắt", "sua_khac": "Sửa giờ khám bệnh ( 03 HS); sửa mã máy", "categories": {}}, {"sheet": "27.10", "stt": "68", "khoa": "TCKT", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Thuy, Trang", "sua_khac": "", "categories": {"Đăng bài trên website": 3}}, {"sheet": "27.10", "stt": "69", "khoa": "CTXH", "hs": "", "phieu": "", "can_bo": "Đỗ Văn Hiệp", "cb_raw": "Hiệp", "yeu_cau": "Nga, Tùng, Hà", "sua_khac": "", "categories": {"Đăng bài trên website": 11}}, {"sheet": "27.10", "stt": "70", "khoa": "PTGMHS", "hs": "26221979", "phieu": "", "can_bo": "Chu Thị Dương", "cb_raw": "Dương", "yeu_cau": "A Tú", "sua_khac": "", "categories": {"Sửa ngày giờ CLS/PTTT, giờ đọc KQ, người TH, máy CLS": 1}}, {"sheet": "27.10", "stt": "71", "khoa": "NGOẠI TKLN", "hs": "26220268", "phieu": "", "can_bo": "Nguyễn Thị Huyền", "cb_raw": "Huyền", "yeu_cau": "Thắm", "sua_khac": "", "categories": {"Ép thẻ BH": 1}}, {"sheet": "27.10", "stt": "72", "khoa": "NHI", "hs": "26227433", "phieu": "", "can_bo": "Nguyễn Thị Huyền", "cb_raw": "Huyền", "yeu_cau": "Bích", "sua_khac": "", "categories": {"Sửa TT bệnh nhân (NN, Đ/chỉ, tuổi, CCCD...) & công khám, kiểu khám": 1}}, {"sheet": "27.10", "stt": "73", "khoa": "HSTC", "hs": "", "phieu": "", "can_bo": "Nguyễn Thị Huyền", "cb_raw": "Huyền", "yeu_cau": "c Hiền", "sua_khac": "Cài lại vimes", "categories": {}}, {"sheet": "27.10", "stt": "74", "khoa": "NGOẠI TKLN", "hs": "", "phieu": "", "can_bo": "Nguyễn Thị Huyền", "cb_raw": "Huyền", "yeu_cau": "Thắm", "sua_khac": "", "categories": {"Sửa xoá giường, cập nhật giường, thêm giường": 4}}, {"sheet": "27.10", "stt": "75", "khoa": "HSTC", "hs": "", "phieu": "", "can_bo": "Nguyễn Thị Huyền", "cb_raw": "Huyền", "yeu_cau": "c Loan", "sua_khac": "Cài lại vimes", "categories": {}}]}

CNTT_MEMORY_CACHE = dict(INITIAL_CNTT_RECORDS_CACHE)

def fetch_cntt_sheet_live(sheet_name):
    if sheet_name in CNTT_MEMORY_CACHE:
        return CNTT_MEMORY_CACHE[sheet_name]
    try:
        encoded_sheet = urllib.parse.quote(sheet_name)
        url = f"https://docs.google.com/spreadsheets/d/{CNTT_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=12) as resp:
            raw_text = resp.read().decode("utf-8")
        rows = list(csv.reader(io.StringIO(raw_text)))
        recs = []
        def _clean_c(v):
            if not v: return ""
            return str(v).replace("\r", " ").replace("\n", " ").strip()

        for r_idx, row in enumerate(rows[1:], 1):
            stt = _clean_c(row[0]) if len(row) > 0 else ""
            khoa = _clean_c(row[1]) if len(row) > 1 else ""
            hs = _clean_c(row[2]) if len(row) > 2 else ""
            phieu = _clean_c(row[3]) if len(row) > 3 else ""
            other = _clean_c(row[34]) if len(row) > 34 else ""
            yc = _clean_c(row[35]) if len(row) > 35 else ""
            cb_raw = _clean_c(row[36]) if len(row) > 36 else ""
            cb = normalize_cntt_staff(cb_raw)
            cats = {}
            for col_idx, cat_name in CNTT_CAT_COLUMNS:
                if col_idx < len(row):
                    val = row[col_idx].strip()
                    if val:
                        try:
                            cnt = int(float(val))
                        except:
                            cnt = 1
                        cats[cat_name] = cnt
            has_content = bool(cats) or bool(other) or bool(cb_raw)
            if (stt or khoa) and has_content:
                recs.append({
                    "sheet": sheet_name,
                    "stt": stt or str(r_idx),
                    "khoa": khoa,
                    "hs": hs,
                    "phieu": phieu,
                    "can_bo": cb,
                    "cb_raw": cb_raw,
                    "yeu_cau": yc,
                    "sua_khac": other,
                    "categories": cats
                })
        CNTT_MEMORY_CACHE[sheet_name] = recs
        return recs
    except Exception as e:
        print(f"Error fetching live CNTT sheet {sheet_name}: {e}")
        return []

def aggregate_cntt_data(sheets_list):
    all_records = []
    for s in sheets_list:
        recs = fetch_cntt_sheet_live(s)
        all_records.extend(recs)

    total_cases = len(all_records)
    khoa_set = set(r["khoa"] for r in all_records if r["khoa"])
    other_count = sum(1 for r in all_records if r["sua_khac"])

    staff_dict = {}
    dept_counts = {}
    cat_counts = {}

    for r in all_records:
        cb = r["can_bo"]
        if cb not in staff_dict:
            staff_dict[cb] = {
                "name": cb,
                "cases": 0,
                "sheets": set(),
                "departments": set(),
                "other_count": 0,
                "categories": {},
                "other_items": []
            }
        st = staff_dict[cb]
        st["cases"] += 1
        st["sheets"].add(r["sheet"])
        if r["khoa"]:
            st["departments"].add(r["khoa"])
            dept_counts[r["khoa"]] = dept_counts.get(r["khoa"], 0) + 1
        if r["sua_khac"]:
            st["other_count"] += 1
            st["other_items"].append({
                "sheet": r["sheet"],
                "khoa": r["khoa"],
                "content": r["sua_khac"],
                "yeu_cau": r["yeu_cau"],
                "hs": r["hs"]
            })
        for cat_name, cnt in r["categories"].items():
            st["categories"][cat_name] = st["categories"].get(cat_name, 0) + cnt
            cat_counts[cat_name] = cat_counts.get(cat_name, 0) + cnt

    staff_list = []
    for cb, data in staff_dict.items():
        sorted_cats = sorted(data["categories"].items(), key=lambda x: -x[1])
        staff_list.append({
            "name": cb,
            "cases": data["cases"],
            "sheets_count": len(data["sheets"]),
            "sheets_list": sorted(list(data["sheets"]), reverse=True),
            "departments_count": len(data["departments"]),
            "other_count": data["other_count"],
            "categories": [{"name": k, "count": v} for k, v in sorted_cats],
            "other_items": data["other_items"]
        })
    staff_list.sort(key=lambda x: -x["cases"])

    return {
        "success": True,
        "selected_sheets": sheets_list,
        "selected_sheets_count": len(sheets_list),
        "total_cases": total_cases,
        "departments_count": len(khoa_set),
        "staff_count": len(staff_list),
        "other_count": other_count,
        "staff_list": staff_list,
        "records": all_records,
        "dept_rankings": sorted(dept_counts.items(), key=lambda x: -x[1])[:20],
        "cat_rankings": sorted(cat_counts.items(), key=lambda x: -x[1])
    }

def dispatch_request(path=""):
    matched = (
        (path or "") + " " +
        request.path + " " +
        request.headers.get("x-matched-path", "") + " " +
        request.environ.get("HTTP_X_MATCHED_PATH", "") + " " +
        request.args.get("action", "")
    ).lower()

    # Route xử lý Tác vụ 10 (Thống kê lỗi XML) hỗ trợ cả GET lẫn POST
    if "error-tracker/sheets" in matched or request.args.get("action") == "error-tracker-sheets":
        return handle_error_tracker_sheets()
    if "error-tracker/data" in matched or request.args.get("action") == "error-tracker-data":
        return handle_error_tracker_data()
    if "error-tracker/export" in matched or request.args.get("action") == "error-tracker-export":
        return handle_error_tracker_export()
    if "error-tracker/test-connection" in matched or request.args.get("action") == "error-tracker-test-connection":
        return handle_error_tracker_test_connection()
    if "error-tracker/update-status" in matched or request.args.get("action") == "error-tracker-update-status":
        return handle_error_tracker_update_status()
    if "error-tracker/config" in matched or request.args.get("action") == "error-tracker-config":
        return handle_error_tracker_config()

    if "cntt-summary" in matched or request.args.get("action") == "cntt-summary" or request.args.get("action") == "cntt-data":
        sheets_param = request.args.get("sheets", "29.10,28.10,27.10")
        force_sync = request.args.get("sync", "0") == "1"
        if force_sync:
            CNTT_MEMORY_CACHE.clear()
        sheets_list = [s.strip() for s in sheets_param.split(",") if s.strip()]
        if not sheets_list:
            sheets_list = ["29.10", "28.10", "27.10"]
        data = aggregate_cntt_data(sheets_list)
        return jsonify(data)

    # Xác thực mật khẩu nhanh qua endpoint API nếu cần
    if "verify-auth" in matched or request.args.get("action") == "verify-auth":
        data = request.get_json(silent=True) or request.form.to_dict() or {}
        p = (data.get("password") or request.args.get("password") or "").strip()
        if p == "8889":
            resp = jsonify({"success": True, "message": "Xác thực thành công"})
            resp.set_cookie("app_auth_pass", "8889", max_age=864000, path="/", samesite="Lax")
            return resp
        return jsonify({"success": False, "error": "Mật khẩu không chính xác!"}), 401

    if request.method == "POST":
        action = request.args.get("action", "")
        # Bảo vệ toàn bộ các tác vụ ngoài Thống Kê Lỗi XML BHYT bằng mật khẩu 8889
        is_protected = any(act in matched or action == act for act in ["task-1", "task-2", "task-3", "task-4", "task-5", "task-6", "task-7", "task-8", "task-9", "get-sheets", "process-giam-dinh", "process-cntt"])
        if is_protected:
            req_pass = (
                request.headers.get("X-App-Password") or
                request.form.get("auth_pass") or
                request.args.get("auth_pass") or
                request.cookies.get("app_auth_pass") or ""
            ).strip()
            if req_pass != "8889":
                return jsonify({"error": "🔒 Chức năng này được bảo vệ. Vui lòng nhập mật khẩu 8889 trên giao diện để tiếp tục."}), 403
        if "get-sheets" in matched or action == "get-sheets":
            return handle_get_sheets()
        if "task-1" in matched or action == "task-1":
            return handle_task_1()
        if "task-2" in matched or action == "task-2":
            return handle_task_2()
        if "task-3" in matched or action == "task-3":
            return handle_task_3()
        if "task-4" in matched or action == "task-4":
            return handle_task_4()
        if "task-5" in matched or action == "task-5" or "process-giam-dinh" in matched or "date_range" in request.form:
            return handle_task_5()
        if "task-6" in matched or action == "task-6":
            return handle_task_6()
        if "task-7" in matched or action == "task-7":
            return handle_task_7()
        if "task-8" in matched or action == "task-8" or "process-cntt" in matched or "sheet_name" in request.form:
            return handle_task_8()
        if "task-9" in matched or action == "task-9":
            return handle_task_9()
        if "url" in request.form and "sheet_name" not in request.form:
            return handle_get_sheets()

    resp = make_response(render_template_string(HTML_PAGE))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

@app.route("/", methods=["GET", "POST"])
@app.route("/api", methods=["GET", "POST"])
@app.route("/api/index", methods=["GET", "POST"])
@app.route("/api/index.py", methods=["GET", "POST"])
@app.route("/api/get-sheets", methods=["GET", "POST"])
@app.route("/api/task-1", methods=["GET", "POST"])
@app.route("/api/task-2", methods=["GET", "POST"])
@app.route("/api/task-3", methods=["GET", "POST"])
@app.route("/api/task-4", methods=["GET", "POST"])
@app.route("/api/task-5", methods=["GET", "POST"])
@app.route("/api/task-6", methods=["GET", "POST"])
@app.route("/api/task-7", methods=["GET", "POST"])
@app.route("/api/task-8", methods=["GET", "POST"])
@app.route("/api/task-9", methods=["GET", "POST"])
@app.route("/api/error-tracker/sheets", methods=["GET", "POST"])
@app.route("/api/error-tracker/data", methods=["GET", "POST"])
@app.route("/api/error-tracker/export", methods=["GET", "POST"])
@app.route("/api/error-tracker/update-status", methods=["GET", "POST"])
@app.route("/api/error-tracker/config", methods=["GET", "POST"])
@app.route("/api/process-giam-dinh", methods=["GET", "POST"])
@app.route("/api/process-cntt", methods=["GET", "POST"])
@app.route("/<path:path>", methods=["GET", "POST"])
def master_router(path=""):
    return dispatch_request(path)

@app.errorhandler(404)
def error_404(e):
    return dispatch_request(request.path)

if __name__ == "__main__":
    app.run(debug=True, port=5000)
