import io
import os
import re
import csv
import copy
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
    """Trích xuất dữ liệu người dùng từ bảng biểu trong file Word."""
    people = []
    if sso_map is None:
        sso_map = {}
    try:
        doc = Document(doc_stream)
        for t_idx, table in enumerate(doc.tables):
            if not table.rows:
                continue
            header_row_idx = -1
            name_idx, cccd_idx, dob_idx = -1, -1, -1
            issue_date_idx, issue_place_idx, email_idx = -1, -1, -1
            unit_idx, pos_idx, phone_idx = -1, -1, -1

            for r_idx, row in enumerate(table.rows[:min(4, len(table.rows))]):
                header_texts = [clean_text(cell.text) for cell in row.cells]
                t_name, t_cccd, t_dob, t_idate, t_iplace, t_email = -1, -1, -1, -1, -1, -1
                t_unit, t_pos, t_phone = -1, -1, -1

                for idx, col in enumerate(header_texts):
                    c_norm = remove_accents(col)
                    if ("ho va ten" in c_norm or "ho ten" in c_norm) and "nguoi lap" not in c_norm and "truong" not in c_norm and "ky va" not in c_norm:
                        if t_name == -1: t_name = idx
                    elif any(k in c_norm for k in ['cccd', 'cmt', 'cmnd', 'can cuoc']) and "chup anh" not in c_norm and "anh" not in c_norm and "ngay cap" not in c_norm and "noi cap" not in c_norm:
                        if t_cccd == -1: t_cccd = idx
                    elif "ngay sinh" in c_norm or "nam sinh" in c_norm or "dob" in c_norm:
                        if t_dob == -1: t_dob = idx
                    elif "ngay cap" in c_norm and "noi" not in c_norm:
                        if t_idate == -1: t_idate = idx
                    elif "noi cap" in c_norm:
                        if t_iplace == -1: t_iplace = idx
                    elif any(k in c_norm for k in ['email', 'thu dien tu', 'mail', 'hop thu']):
                        if t_email == -1: t_email = idx
                    elif any(k in c_norm for k in ['don vi', 'co quan', 'to chuc', 'khoa', 'phong']) and "tinh" not in c_norm:
                        if t_unit == -1: t_unit = idx
                    elif "chuc vu" in c_norm or "vi tri" in c_norm:
                        if t_pos == -1: t_pos = idx
                    elif any(k in c_norm for k in ['sdt', 'dien thoai', 'di dong', 'phone']):
                        if t_phone == -1: t_phone = idx

                if t_name != -1 and t_cccd != -1:
                    header_row_idx = r_idx
                    name_idx, cccd_idx, dob_idx = t_name, t_cccd, t_dob
                    issue_date_idx, issue_place_idx, email_idx = t_idate, t_iplace, t_email
                    unit_idx, pos_idx, phone_idx = t_unit, t_pos, t_phone
                    break

            if header_row_idx == -1:
                continue

            for r in table.rows[header_row_idx + 1:]:
                vals = [clean_text(cell.text) for cell in r.cells]
                if len(vals) <= max(name_idx, cccd_idx):
                    continue
                name_val = vals[name_idx]
                cccd_val = vals[cccd_idx]
                if not name_val or not cccd_val or "ho va ten" in remove_accents(name_val):
                    continue

                kn = norm_name(name_val)
                kc = norm_cccd(cccd_val)
                sso_info = sso_map.get((kn, kc)) or sso_map.get(kc) or sso_map.get(kn) or {}

                dob = normalize_date(vals[dob_idx]) if dob_idx != -1 and dob_idx < len(vals) else ""
                issue_date = normalize_date(vals[issue_date_idx]) if issue_date_idx != -1 and issue_date_idx < len(vals) else ""
                issue_place = normalize_issue_place(vals[issue_place_idx]) if issue_place_idx != -1 and issue_place_idx < len(vals) else "CCSQLHCVTTXH"
                phone = vals[phone_idx] if phone_idx != -1 and phone_idx < len(vals) else ""
                unit = vals[unit_idx] if unit_idx != -1 and unit_idx < len(vals) else sso_info.get("unit", DEFAULT_ORG3)
                pos = vals[pos_idx] if pos_idx != -1 and pos_idx < len(vals) else sso_info.get("position", "Nhân viên")

                email = sso_info.get("email", "")
                if not email and email_idx != -1 and email_idx < len(vals):
                    email = clean_email(vals[email_idx])
                if not email and kn:
                    # Gợi ý email công vụ nếu chưa có
                    email = f"{kn}@bacninh.gov.vn"

                people.append({
                    "name": name_val, "dob": dob, "cccd": format_cccd_output(cccd_val),
                    "issue_place": issue_place, "issue_date": issue_date, "email": email,
                    "org1": DEFAULT_ORG1, "org2": DEFAULT_ORG2, "org3": unit if unit else DEFAULT_ORG3,
                    "org4": "", "province": DEFAULT_PROVINCE, "position": pos, "phone": phone,
                    "serial_old": "", "network": "", "is_transfer": "", "device_type": ""
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
    """Xuất file Công văn Mẫu 02 đề nghị cấp chứng thư số chuyên dùng Chính phủ."""
    doc = Document()
    now = datetime.now()

    # Quốc hiệu - Tiêu ngữ
    table_top = doc.add_table(rows=1, cols=2)
    table_top.autofit = True
    c_left = table_top.cell(0, 0)
    c_right = table_top.cell(0, 1)

    p_left = c_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_l1 = p_left.add_run("ỦY BAN NHÂN DÂN TỈNH BẮC NINH\n")
    r_l1.bold = True
    r_l1.font.size = Pt(11)
    r_l2 = p_left.add_run("BỆNH VIỆN ĐA KHOA BẮC NINH SỐ 2\n")
    r_l2.bold = True
    r_l2.font.size = Pt(11)
    r_l3 = p_left.add_run("Số:       /CV-BVĐK2")
    r_l3.font.size = Pt(11)

    p_right = c_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_r1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r_r1.bold = True
    r_r1.font.size = Pt(11)
    r_r2 = p_right.add_run("Độc lập - Tự do - Hạnh phúc\n")
    r_r2.bold = True
    r_r2.font.size = Pt(11)
    r_r3 = p_right.add_run(f"Bắc Ninh, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}")
    r_r3.italic = True
    r_r3.font.size = Pt(11)

    doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CÔNG VĂN ĐỀ NGHỊ CẤP CHỨNG THƯ SỐ CHUYÊN DÙNG CHÍNH PHỦ\n")
    r_title.bold = True
    r_title.font.size = Pt(13)
    r_sub = p_title.add_run("(Mẫu 02 quy định ban hành kèm theo Nghị định số 130/2018/NĐ-CP)")
    r_sub.italic = True
    r_sub.font.size = Pt(10)

    p_body = doc.add_paragraph()
    p_body.add_run("Kính gửi: Cục Chứng thực số và Bảo mật thông tin – Ban Cơ yếu Chính phủ.\n\n")
    p_body.add_run(f"Căn cứ nhu cầu sử dụng chứng thư số phục vụ ký số và xác thực điện tử trong công tác khám chữa bệnh và quản lý bệnh viện, Bệnh viện Đa khoa Bắc Ninh số 2 kính đề nghị cấp mới chứng thư số cho các cá nhân (tổng số: {len(people_list):02d} người) có danh sách chi tiết dưới đây:")

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
            p.runs[0].font.size = Pt(9)

    for idx, p_info in enumerate(people_list, 1):
        row = table.add_row()
        cccd_full = f"{p_info['cccd']}; {p_info['issue_date']}; {p_info['issue_place']}".strip("; ")
        vals = [
            str(idx), p_info["name"], p_info["dob"], cccd_full,
            p_info["email"], p_info["org3"], p_info["province"],
            p_info["position"], p_info["phone"]
        ]
        for col_idx, text_val in enumerate(vals):
            c = row.cells[col_idx]
            c.text = text_val
            if c.paragraphs and c.paragraphs[0].runs:
                c.paragraphs[0].runs[0].font.size = Pt(9)

    doc.add_paragraph()
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sign1 = p_sign.add_run("GIÁM ĐỐC / THỦ TRƯỞNG ĐƠN VỊ\n")
    r_sign1.bold = True
    p_sign.add_run("(Ký, ghi rõ họ tên và đóng dấu)")

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
    ws["A1"].font = Font(name="Arial", size=13, bold=True)
    center_cell(ws["A1"])
    ws.row_dimensions[1].height = 28

    headers = ["STT", "Họ và tên", "Ngày sinh", "Số điện thoại", "Số CCCD", "Khoa / Phòng / Đơn vị", "Chức vụ / Vị trí"]
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.border = border
        center_cell(cell, wrap=True)

    for idx, p in enumerate(people_list, 1):
        r = idx + 3
        row_vals = [idx, p["name"], p["dob"], p["phone"], p["cccd"], p["org3"], p["position"]]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            cell.font = Font(name="Arial", size=10)
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

    @media (max-width: 992px) {
      .app-container { flex-direction: column; }
      .sidebar { width: 100%; height: auto; border-right: none; border-bottom: 1px solid #1e293b; }
      .two-col-grid { grid-template-columns: 1fr; }
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

    <!-- SIDEBAR 8 TÁC VỤ ĐƯỢC CHIA THÀNH 2 NHÓM -->
    <div class="sidebar">
      <div class="sidebar-header">
        <div class="sidebar-title">📑 DANH MỤC TÁC VỤ (8 CHỨC NĂNG)</div>
        <div class="sidebar-subtitle">Hệ thống phân thành 2 nhóm chuyên biệt</div>
      </div>

      <div class="sidebar-scrollable">
        <!-- NHÓM 1: CHỨNG THƯ SỐ + EMAIL CÔNG VỤ -->
        <div class="task-group-card card-cks">
          <div class="group-header-box">
            <div class="group-header-title">
              <span>🔐 NHÓM CKS & MAIL CÔNG VỤ</span>
            </div>
            <span class="group-header-badge">4 tác vụ</span>
          </div>
          <ul class="sidebar-menu">
            <li class="active" data-task="1" onclick="switchTask(1)">
              <div class="nav-icon">🎯</div>
              <div class="nav-text">
                <span class="nav-title">1. VGCA (Word + Excel SSO)</span>
                <span class="nav-desc">Đối chiếu cấp CKS Word & Excel SSO</span>
              </div>
            </li>
            <li data-task="2" onclick="switchTask(2)">
              <div class="nav-icon">📝</div>
              <div class="nav-text">
                <span class="nav-title">2. CKS (Công văn Mẫu 02)</span>
                <span class="nav-desc">Xuất Công văn Word Mẫu 02 chuẩn</span>
              </div>
            </li>
            <li data-task="3" onclick="switchTask(3)">
              <div class="nav-icon">📧</div>
              <div class="nav-text">
                <span class="nav-title">3. Cấp Email Công Vụ</span>
                <span class="nav-desc">Tổng hợp danh sách xin cấp Email</span>
              </div>
            </li>
            <li data-task="7" onclick="switchTask(7)">
              <div class="nav-icon">📄</div>
              <div class="nav-text">
                <span class="nav-title">7. VGCA (Từ File Word)</span>
                <span class="nav-desc">Đọc nhiều file Word không cần SSO</span>
              </div>
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
            <li data-task="4" onclick="switchTask(4)">
              <div class="nav-icon">📊</div>
              <div class="nav-text">
                <span class="nav-title">4. Báo Cáo P.CNTT (File)</span>
                <span class="nav-desc">Tổng hợp từ file Excel máy tính</span>
              </div>
            </li>
            <li data-task="5" onclick="switchTask(5)">
              <div class="nav-icon">🏥</div>
              <div class="nav-text">
                <span class="nav-title">5. Giám Định BHYT</span>
                <span class="nav-desc">Lọc người dùng & xuất Excel A4</span>
              </div>
            </li>
            <li data-task="6" onclick="switchTask(6)">
              <div class="nav-icon">🤖</div>
              <div class="nav-text">
                <span class="nav-title">6. P.CNTT Tự Động</span>
                <span class="nav-desc">Quét từ link Google Sheets bất kỳ</span>
              </div>
            </li>
            <li data-task="8" onclick="switchTask(8)">
              <div class="nav-icon">⚡</div>
              <div class="nav-text">
                <span class="nav-title">8. P.CNTT (Sheet Cố Định)</span>
                <span class="nav-desc">Tự động tải Google Sheets của viện</span>
              </div>
            </li>
          </ul>
        </div>
      </div>

      <div class="sidebar-footer">
        <b>Phòng CNTT • BVĐK Bắc Ninh số 2</b><br>
        Phiên bản Web v2.7 • Đã Phân 2 Nhóm Tác Vụ
      </div>
    </div>

    <!-- CONTENT AREA -->
    <div class="content-area">
      <div class="breadcrumb-bar">
        <span class="badge" id="task-badge">TÁC VỤ 1</span>
        <span class="breadcrumb-title" id="task-breadcrumb">🎯 1. VGCA (Word + Excel SSO)  ─  Đối chiếu cấp CKS Word & Excel SSO</span>
      </div>

      <!-- ==================== TÁC VỤ 1 ==================== -->
      <div class="page active" id="page-task-1">
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
            <form onsubmit="handleTask3(event)">
              <div class="form-group">
                <label class="form-label">1️⃣ Chọn các file Word (.docx) xin cấp Email:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t3-word').click()">
                  <div class="drop-icon">📧</div>
                  <div class="drop-text">Nhấp để chọn các file Word (.docx)</div>
                  <input type="file" id="file-t3-word" multiple accept=".docx,.zip" onchange="previewFiles(this, 'chosen-t3-word')">
                </div>
                <div class="file-chosen-badge" id="chosen-t3-word"></div>
              </div>

              <div class="form-group">
                <label class="form-label">2️⃣ Định dạng xuất kết quả:</label>
                <div class="radio-group">
                  <label class="radio-card">
                    <input type="radio" name="email_format" value="excel" checked>
                    <div>
                      <div class="radio-title">Bảng tính Excel (.xlsx) chuẩn Sở TT&TT</div>
                      <div class="radio-desc">Gồm 7 cột: STT, Họ tên, Ngày sinh, SĐT, CCCD, Đơn vị, Chức vụ</div>
                    </div>
                  </label>
                  <label class="radio-card">
                    <input type="radio" name="email_format" value="word">
                    <div>
                      <div class="radio-title">Văn bản Công văn Word (.docx) hoàn chỉnh</div>
                      <div class="radio-desc">Có đầy đủ tiêu ngữ, căn cứ và phần ký tên</div>
                    </div>
                  </label>
                </div>
              </div>

              <button type="submit" class="btn" id="btn-t3">
                <span class="spinner" id="spin-t3"></span>
                <span id="txt-t3">🚀 Bắt Đầu Tổng Hợp Danh Sách Email</span>
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
              <h2>Hướng Dẫn Nghiệp Vụ Tác Vụ 5</h2>
            </div>
            <div class="guide-box">
              <h3>🎯 Mục đích:</h3>
              <p>Đọc file Excel báo cáo cổng giám định BHYT, lọc danh sách người dùng được phép, xuất ra file <b>'DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx'</b> chuẩn khổ A4 dọc, tự động autofit cột và co giãn vừa 1 trang ngang.</p>
              <h3>📝 Các bước thực hiện:</h3>
              <ol>
                <li><b>Bước 1:</b> Tải lên file Excel xuất từ Cổng Giám Định BHYT.</li>
                <li><b>Bước 2:</b> Chọn khoảng ngày cần tổng hợp:
                  <ul>
                    <li><b>Từ ngày 01 đến 14:</b> Lấy nửa đầu tháng.</li>
                    <li><b>Từ ngày 15 đến cuối tháng:</b> Lấy nửa cuối tháng.</li>
                    <li><b>Toàn bộ tháng (01 đến hết tháng):</b> Lấy trọn vẹn dữ liệu.</li>
                  </ul>
                </li>
                <li><b>Bước 3:</b> Nhấn nút <b>"🚀 Xử Lý & Xuất Báo Cáo Giám Định BHYT"</b>.</li>
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
                <label class="form-label">1️⃣ Tải lên file Excel Cổng Giám Định BHYT:</label>
                <div class="file-dropzone" onclick="document.getElementById('file-t5').click()">
                  <div class="drop-icon">🏥</div>
                  <div class="drop-text">Nhấp để chọn file Excel Giám định</div>
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
                <span id="txt-t5">🚀 Xử Lý & Xuất Báo Cáo Giám Định (A4 Dọc)</span>
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

    </div>
  </div>

  <script>
    const taskMeta = {
      1: { badge: 'TÁC VỤ 1', title: '🎯 1. VGCA (Word + Excel SSO)  ─  Đối chiếu cấp CKS Word & Excel SSO' },
      2: { badge: 'TÁC VỤ 2', title: '📝 2. CKS (Công văn Mẫu 02)  ─  Xuất Công văn Word Mẫu 02 chuẩn' },
      3: { badge: 'TÁC VỤ 3', title: '📧 3. Cấp Email Công Vụ  ─  Tổng hợp danh sách xin cấp Email' },
      4: { badge: 'TÁC VỤ 4', title: '📊 4. Báo Cáo P.CNTT (File)  ─  Tổng hợp từ file Excel máy tính' },
      5: { badge: 'TÁC VỤ 5', title: '🏥 5. Giám Định BHYT  ─  Lọc người dùng & xuất Excel A4 dọc' },
      6: { badge: 'TÁC VỤ 6', title: '🤖 6. P.CNTT Tự Động  ─  Quét từ link Google Sheets bất kỳ' },
      7: { badge: 'TÁC VỤ 7', title: '📄 7. VGCA (Từ File Word)  ─  Đọc nhiều file Word không cần SSO' },
      8: { badge: 'TÁC VỤ 8', title: '⚡ 8. P.CNTT (Sheet Cố Định)  ─  Tự động tải Google Sheets của viện' }
    };

    function switchTask(taskId) {
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

      const picker = await promptSaveFile('Cong_van_de_nghi_Mau_02.docx');
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
        const savedName = await writeBlobToHandleOrDownload(blob, picker.handle, 'Cong_van_de_nghi_Mau_02.docx');
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

    // XỬ LÝ TÁC VỤ 3
    async function handleTask3(e) {
      e.preventDefault();
      hideAlert('alert-t3');
      const wInput = document.getElementById('file-t3-word');
      if (!wInput.files || wInput.files.length === 0) {
        showAlert('alert-t3', 'error', '⚠️ Vui lòng chọn các file Word xin cấp Email!');
        return;
      }
      const fmt = document.querySelector('input[name="email_format"]:checked').value;
      const defaultName = fmt === 'word' ? 'De_nghi_cap_email_cong_vu.docx' : 'Danh_sach_cap_email_cong_vu.xlsx';

      const picker = await promptSaveFile(defaultName);
      if (picker.canceled) return;

      const formData = new FormData();
      for (let i = 0; i < wInput.files.length; i++) {
        formData.append('word_files', wInput.files[i]);
      }
      formData.append('format', fmt);

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
          showAlert('alert-t3', 'success', `✅ Đã lưu file "${savedName}" thành công tại vị trí bạn đã chọn!`);
        } else {
          showAlert('alert-t3', 'success', `✅ Đã tổng hợp và tải xuống danh sách cấp Email công vụ thành công!`);
        }
      } catch (err) {
        showAlert('alert-t3', 'error', '❌ Lỗi: ' + err.message);
      } finally {
        setLoading('t3', false, '🚀 Bắt Đầu Tổng Hợp Danh Sách Email');
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
        showAlert('alert-t5', 'error', '⚠️ Vui lòng chọn file Excel Cổng Giám Định BHYT!');
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
        setLoading('t5', false, '🚀 Xử Lý & Xuất Báo Cáo Giám Định (A4 Dọc)');
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
            download_name="Cong_van_de_nghi_Mau_02.docx",
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
            people = parse_people_from_word_stream(fname, dstream)
            all_people.extend(people)

        if not all_people:
            return jsonify({"error": "Không trích xuất được bản ghi nào từ các file Word!"}), 400

        fmt = request.form.get("format", "excel")
        if fmt == "word":
            out_io = build_mau_02_docx(all_people)
            return send_file(
                out_io,
                as_attachment=True,
                download_name="De_nghi_cap_email_cong_vu.docx",
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            out_io = build_email_excel(all_people)
            return send_file(
                out_io,
                as_attachment=True,
                download_name="Danh_sach_cap_email_cong_vu.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
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
            return jsonify({"error": "Vui lòng tải lên file Excel Cổng Giám Định BHYT."}), 400

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
# MASTER DISPATCH ROUTER (TƯƠNG THÍCH MỌI CẤU HÌNH REWRITE VERCEL)
# ============================================================
def dispatch_request(path=""):
    matched = (
        (path or "") + " " +
        request.path + " " +
        request.headers.get("x-matched-path", "") + " " +
        request.environ.get("HTTP_X_MATCHED_PATH", "") + " " +
        request.args.get("action", "")
    ).lower()

    if request.method == "POST":
        action = request.args.get("action", "")
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
