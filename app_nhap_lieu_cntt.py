import os
import re
import shutil
import sqlite3
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# Đặt đường dẫn lưu file CSDL tại thư mục hiện tại
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "quan_ly_cntt.db")
TEMPLATE_NAME = os.path.join(BASE_DIR, "Mau_Import_Khoa_Phong.xlsx")

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

# ============================================================
# 1. KHỞI TẠO CƠ SỞ DỮ LIỆU & TỰ ĐỘNG TẠO BẢNG
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    # Tạo các bảng CSDL
    cursor.executescript("""
        CREATE TABLE IF NOT EXISTS khoa_phong (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_khoa TEXT UNIQUE,
            ten_khoa TEXT NOT NULL UNIQUE,
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS nhan_su (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ho_ten TEXT NOT NULL UNIQUE,
            chuc_vu TEXT DEFAULT 'Kỹ thuật viên',
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS danh_muc_loi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ten_loi TEXT NOT NULL UNIQUE,
            thu_tu INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS nhat_ky_sua_chua (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay_thuc_hien TEXT NOT NULL,
            khoa_phong_id INTEGER NOT NULL,
            nhan_su_id INTEGER NOT NULL,
            loi_khac TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (khoa_phong_id) REFERENCES khoa_phong(id),
            FOREIGN KEY (nhan_su_id) REFERENCES nhan_su(id)
        );

        CREATE TABLE IF NOT EXISTS chi_tiet_loi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nhat_ky_id INTEGER NOT NULL,
            danh_muc_loi_id INTEGER NOT NULL,
            so_luong INTEGER DEFAULT 1,
            FOREIGN KEY (nhat_ky_id) REFERENCES nhat_ky_sua_chua(id) ON DELETE CASCADE,
            FOREIGN KEY (danh_muc_loi_id) REFERENCES danh_muc_loi(id)
        );
    """)

    # Kiểm tra đảm bảo các cột tồn tại trong khoa_phong
    cursor.execute("PRAGMA table_info(khoa_phong);")
    cols = [c[1] for c in cursor.fetchall()]
    if "ma_khoa" not in cols:
        try: cursor.execute("ALTER TABLE khoa_phong ADD COLUMN ma_khoa TEXT;")
        except Exception: pass
    if "active" not in cols:
        try: cursor.execute("ALTER TABLE khoa_phong ADD COLUMN active INTEGER DEFAULT 1;")
        except Exception: pass

    # Kiểm tra đảm bảo các cột tồn tại trong nhan_su
    cursor.execute("PRAGMA table_info(nhan_su);")
    ns_cols = [c[1] for c in cursor.fetchall()]
    if "chuc_vu" not in ns_cols:
        try: cursor.execute("ALTER TABLE nhan_su ADD COLUMN chuc_vu TEXT DEFAULT 'Kỹ thuật viên';")
        except Exception: pass
    if "active" not in ns_cols:
        try: cursor.execute("ALTER TABLE nhan_su ADD COLUMN active INTEGER DEFAULT 1;")
        except Exception: pass


    # Nạp dữ liệu Khoa/Phòng mẫu nếu trống (active = 1)
    cursor.execute("SELECT COUNT(*) FROM khoa_phong")
    if cursor.fetchone()[0] == 0:
        sample_depts = [
            ("K01", "NỘI TH", 1), ("K02", "NGOẠI TH", 1), ("K03", "CẤP CỨU", 1), ("K04", "HSTC", 1), 
            ("K05", "PHỤ SẢN", 1), ("K06", "NHI", 1), ("K07", "KHOA DƯỢC", 1), ("K08", "TCKT", 1), 
            ("K09", "TTBVSK", 1), ("K10", "CĐHA", 1), ("K11", "XNTT", 1), ("K12", "PHCN", 1)
        ]
        cursor.executemany("INSERT INTO khoa_phong (ma_khoa, ten_khoa, active) VALUES (?, ?, ?);", sample_depts)

    # Nạp dữ liệu Nhân sự mẫu nếu trống
    cursor.execute("SELECT COUNT(*) FROM nhan_su")
    if cursor.fetchone()[0] == 0:
        sample_staff = ["Nguyễn Văn A", "Trần Văn B", "Lê Thị C", "Phạm Văn D"]
        cursor.executemany("INSERT INTO nhan_su (ho_ten) VALUES (?);", [(s,) for s in sample_staff])

    # Nạp danh mục Lỗi phần mềm mẫu nếu trống
    cursor.execute("SELECT COUNT(*) FROM danh_muc_loi")
    if cursor.fetchone()[0] == 0:
        sample_errors = [
            "Xoá CLS, PTTT",
            "Hủy CKS",
            "Sửa thông tin BS (NX, Địa chỉ, CCCD, th_sư...)",
            "Sửa sổ khám, sửa phiếu khám",
            "Sửa ngày giờ vào phòng, giờ đọc kết quả, người TH dịch vụ CLS, PTTT",
            "Đăng bài trên website",
            "Sửa cách dùng thuốc, ngày giờ thuốc",
            "Lỗi không kết thúc được hồ sơ",
            "Sửa lỗi xác nhận thuốc",
            "Lỗi in phiếu chỉ định"
        ]
        cursor.executemany("INSERT INTO danh_muc_loi (ten_loi, thu_tu) VALUES (?, ?);", [(e, idx+1) for idx, e in enumerate(sample_errors)])

    conn.commit()
    conn.close()

# ============================================================
# 2. GIAO DIỆN QUẢN LÝ TỔNG HỢP (NHẬP LIỆU & QUẢN LÝ DANH MỤC)
# ============================================================
class AppNhapLieu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HỆ THỐNG NHẬP LIỆU & QUẢN LÝ DANH MỤC P.CNTT")
        self.geometry("1100x760")
        self.minsize(950, 680)
        
        # Khởi tạo CSDL
        init_db()
        
        style = ttk.Style(self)
        style.theme_use('clam')

        # Header Title
        header = tk.Frame(self, bg="#1e3a8a", pady=12)
        header.pack(fill=tk.X)
        tk.Label(
            header, text="HỆ THỐNG NHẬP LIỆU SỬA CHỮA & QUẢN LÝ DANH MỤC P.CNTT", 
            font=("Segoe UI", 14, "bold"), fg="white", bg="#1e3a8a"
        ).pack()

        self.err_vars = {}

        # Notebook chứa các Tab
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: Nhập liệu nhật ký
        self.tab1 = ttk.Frame(self.notebook)
        # Tab 2: Quản lý Khoa/Phòng
        self.tab2 = ttk.Frame(self.notebook)
        # Tab 3: Quản lý Nhân viên
        self.tab3 = ttk.Frame(self.notebook)
        # Tab 4: Quản lý Hạng mục Lỗi
        self.tab4 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab1, text="  📝 Nhập Liệu Nhật Ký Sửa Chữa  ")
        self.notebook.add(self.tab2, text="  🏥 Quản Lý Khoa / Phòng  ")
        self.notebook.add(self.tab3, text="  👨‍💻 Quản Lý Nhân Viên CNTT  ")
        self.notebook.add(self.tab4, text="  ⚙️ Quản Lý Lỗi Phần Mềm  ")

        # Khởi tạo các Tab
        self.setup_tab_nhap_lieu()
        self.setup_tab_khoa_phong()
        self.setup_tab_nhan_vien()
        self.setup_tab_hang_muc_loi()

        # Nạp dữ liệu ban đầu
        self.reload_all_categories()
        self.load_history()

    # ============================================================
    # TAB 1: GIAO DIỆN NHẬP LIỆU NHẬT KÝ
    # ============================================================
    def setup_tab_nhap_lieu(self):
        main_frame = ttk.Frame(self.tab1, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Thông tin chung
        info_box = ttk.LabelFrame(main_frame, text=" 📝 Thông Tin Lượt Sửa Chữa ", padding=10)
        info_box.pack(fill=tk.X, pady=5)

        ttk.Label(info_box, text="Ngày thực hiện:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_ngay = ttk.Entry(info_box, width=15, font=("Segoe UI", 9))
        self.txt_ngay.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.txt_ngay.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(info_box, text="Khoa/Phòng yêu cầu (Chỉ hiện Trạng Thái 1):", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.cb_khoa = ttk.Combobox(info_box, state="readonly", width=22, font=("Segoe UI", 9))
        self.cb_khoa.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        ttk.Label(info_box, text="Nhân sự thực hiện:", font=("Segoe UI", 9, "bold")).grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
        self.cb_nhansu = ttk.Combobox(info_box, state="readonly", width=22, font=("Segoe UI", 9))
        self.cb_nhansu.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)

        # 2. Hạng mục lỗi phần mềm
        self.err_box = ttk.LabelFrame(main_frame, text=" 💻 Hạng Mục Sửa Lỗi Phần Mềm ", padding=10)
        self.err_box.pack(fill=tk.X, pady=5)

        # 3. Sửa lỗi khác
        other_box = ttk.LabelFrame(main_frame, text=" 🛠️ Sửa Lỗi Khác (Diễn giải chi tiết nếu có) ", padding=10)
        other_box.pack(fill=tk.X, pady=5)
        
        self.txt_loi_khac = ttk.Entry(other_box, font=("Segoe UI", 10))
        self.txt_loi_khac.pack(fill=tk.X, expand=True)

        # Button Frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=8)
        
        btn_save = tk.Button(
            btn_frame, text="🚀 LƯU PHIẾU NHẬP LIỆU", font=("Segoe UI", 11, "bold"),
            bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white",
            pady=6, padx=15, command=self.save_entry
        )
        btn_save.pack(side=tk.LEFT, padx=5)

        btn_delete = tk.Button(
            btn_frame, text="❌ XÓA DÒNG ĐÃ CHỌN", font=("Segoe UI", 10, "bold"),
            bg="#dc2626", fg="white", activebackground="#b91c1c", activeforeground="white",
            pady=6, padx=15, command=self.delete_selected_history
        )
        btn_delete.pack(side=tk.RIGHT, padx=5)

        # 4. Nhật ký đã nhập
        tree_box = ttk.LabelFrame(main_frame, text=" 📊 Nhật Ký Sửa Chữa Đã Nhập ", padding=10)
        tree_box.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ("id", "ngay", "khoa", "nhansu", "loi_pm", "loi_khac")
        self.tree = ttk.Treeview(tree_box, columns=columns, show="headings", height=8)
        self.tree.heading("id", text="ID")
        self.tree.heading("ngay", text="Ngày")
        self.tree.heading("khoa", text="Khoa/Phòng Yêu Cầu")
        self.tree.heading("nhansu", text="Nhân Sự Thực Hiện")
        self.tree.heading("loi_pm", text="Sửa Lỗi Phần Mềm")
        self.tree.heading("loi_khac", text="Sửa Lỗi Khác")

        self.tree.column("id", width=45, anchor="center")
        self.tree.column("ngay", width=95, anchor="center")
        self.tree.column("khoa", width=150)
        self.tree.column("nhansu", width=150)
        self.tree.column("loi_pm", width=270)
        self.tree.column("loi_khac", width=250)

        scrollbar = ttk.Scrollbar(tree_box, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

    # ============================================================
    # TAB 2: QUẢN LÝ KHOA / PHÒNG (THÊM / SỬA / XÓA / IMPORT / TRẠNG THÁI 1/0)
    # ============================================================
    def setup_tab_khoa_phong(self):
        main_frame = ttk.Frame(self.tab2, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.selected_khoa_id = None

        # Khung Nhập & Chỉnh Sửa Dữ Liệu Khoa / Phòng
        form_box = ttk.LabelFrame(main_frame, text=" 📝 Thông Tin & Chỉnh Sửa Khoa / Phòng ", padding=10)
        form_box.pack(fill=tk.X, pady=5)

        ttk.Label(form_box, text="Mã Khoa:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_ma_khoa = ttk.Entry(form_box, font=("Segoe UI", 9), width=12)
        self.txt_ma_khoa.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form_box, text="Tên Khoa / Phòng:", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.txt_ten_khoa = ttk.Entry(form_box, font=("Segoe UI", 9), width=32)
        self.txt_ten_khoa.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form_box, text="Trạng Thái (1/0):", font=("Segoe UI", 9, "bold")).grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
        self.cb_trangthai_khoa = ttk.Combobox(form_box, state="readonly", width=25, font=("Segoe UI", 9))
        self.cb_trangthai_khoa["values"] = ["1 - Hoạt động (Hiển thị nhật ký)", "0 - Không hiển thị (Ẩn khỏi nhật ký)"]
        self.cb_trangthai_khoa.current(0)
        self.cb_trangthai_khoa.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)

        # Nút bấm thao tác form
        btn_form_frame = ttk.Frame(form_box)
        btn_form_frame.grid(row=1, column=0, columnspan=6, sticky=tk.W, pady=8, padx=5)

        btn_add = tk.Button(
            btn_form_frame, text="➕ THÊM MỚI KHOA PHÒNG", font=("Segoe UI", 9, "bold"),
            bg="#16a34a", fg="white", activebackground="#15803d", activeforeground="white",
            pady=4, padx=10, command=self.add_khoa_phong
        )
        btn_add.pack(side=tk.LEFT, padx=4)

        btn_edit = tk.Button(
            btn_form_frame, text="✏️ CẬP NHẬT / CHỈNH SỬA", font=("Segoe UI", 9, "bold"),
            bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white",
            pady=4, padx=10, command=self.update_khoa_phong
        )
        btn_edit.pack(side=tk.LEFT, padx=4)

        btn_reset = tk.Button(
            btn_form_frame, text="🧹 XÓA TRẮNG FORM", font=("Segoe UI", 9, "bold"),
            bg="#64748b", fg="white", activebackground="#475569", activeforeground="white",
            pady=4, padx=8, command=self.clear_khoa_form
        )
        btn_reset.pack(side=tk.LEFT, padx=4)

        btn_import = tk.Button(
            btn_form_frame, text="📊 IMPORT TỪ EXCEL", font=("Segoe UI", 9, "bold"),
            bg="#0284c7", fg="white", activebackground="#0369a1", activeforeground="white",
            pady=4, padx=10, command=self.import_khoa_phong_excel
        )
        btn_import.pack(side=tk.LEFT, padx=4)

        btn_template = tk.Button(
            btn_form_frame, text="📄 TẢI FILE MẪU EXCEL", font=("Segoe UI", 9, "bold"),
            bg="#475569", fg="white", activebackground="#334155", activeforeground="white",
            pady=4, padx=10, command=self.download_template_excel
        )
        btn_template.pack(side=tk.LEFT, padx=4)

        # Danh sách bảng
        list_box = ttk.LabelFrame(main_frame, text=" 📋 Danh Sách Khoa / Phòng (Bấm vào dòng để sửa thông tin) ", padding=10)
        list_box.pack(fill=tk.BOTH, expand=True, pady=8)

        self.tree_khoa = ttk.Treeview(list_box, columns=("id", "ma", "ten", "trangthai"), show="headings")
        self.tree_khoa.heading("id", text="ID")
        self.tree_khoa.heading("ma", text="Mã Khoa")
        self.tree_khoa.heading("ten", text="Tên Khoa / Phòng")
        self.tree_khoa.heading("trangthai", text="Trạng Thái (1/0)")
        
        self.tree_khoa.column("id", width=50, anchor="center")
        self.tree_khoa.column("ma", width=100, anchor="center")
        self.tree_khoa.column("ten", width=420)
        self.tree_khoa.column("trangthai", width=160, anchor="center")
        self.tree_khoa.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb_k = ttk.Scrollbar(list_box, orient=tk.VERTICAL, command=self.tree_khoa.yview)
        self.tree_khoa.configure(yscroll=sb_k.set)
        sb_k.pack(side=tk.RIGHT, fill=tk.Y)

        # Sự kiện chọn dòng để điền thông tin vào form chỉnh sửa
        self.tree_khoa.bind("<<TreeviewSelect>>", self.on_select_khoa_phong)

        act_btn_frame = ttk.Frame(main_frame)
        act_btn_frame.pack(fill=tk.X, pady=5)

        btn_toggle = tk.Button(
            act_btn_frame, text="🔄 ĐỔI TRẠNG THÁI LINH HOẠT (1 <-> 0)", font=("Segoe UI", 10, "bold"),
            bg="#d97706", fg="white", activebackground="#b45309", activeforeground="white",
            pady=5, padx=10, command=self.toggle_khoa_status
        )
        btn_toggle.pack(side=tk.LEFT, padx=5)

        btn_del_k = tk.Button(
            act_btn_frame, text="❌ XÓA KHOA/PHÒNG ĐÃ CHỌN", font=("Segoe UI", 10, "bold"),
            bg="#dc2626", fg="white", activebackground="#b91c1c", activeforeground="white",
            pady=5, padx=10, command=self.delete_khoa_phong
        )
        btn_del_k.pack(side=tk.RIGHT, padx=5)

    # ============================================================
    # TAB 3: QUẢN LÝ NHÂN VIÊN CNTT
    # ============================================================
    def setup_tab_nhan_vien(self):
        main_frame = ttk.Frame(self.tab3, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        add_box = ttk.LabelFrame(main_frame, text=" ➕ Thêm Nhân Viên CNTT Mới ", padding=10)
        add_box.pack(fill=tk.X, pady=5)

        ttk.Label(add_box, text="Họ và Tên:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.txt_new_nhansu = ttk.Entry(add_box, font=("Segoe UI", 10), width=30)
        self.txt_new_nhansu.pack(side=tk.LEFT, padx=5)

        ttk.Label(add_box, text="Chức vụ:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.txt_new_chucvu = ttk.Entry(add_box, font=("Segoe UI", 10), width=20)
        self.txt_new_chucvu.insert(0, "Kỹ thuật viên")
        self.txt_new_chucvu.pack(side=tk.LEFT, padx=5)

        btn_add_ns = tk.Button(
            add_box, text="➕ THÊM NHÂN VIÊN", font=("Segoe UI", 10, "bold"),
            bg="#16a34a", fg="white", pady=4, command=self.add_nhan_vien
        )
        btn_add_ns.pack(side=tk.LEFT, padx=10)

        list_box = ttk.LabelFrame(main_frame, text=" 📋 Danh Sách Nhân Viên P.CNTT ", padding=10)
        list_box.pack(fill=tk.BOTH, expand=True, pady=10)

        self.tree_nhansu = ttk.Treeview(list_box, columns=("id", "ten", "chucvu"), show="headings")
        self.tree_nhansu.heading("id", text="ID")
        self.tree_nhansu.heading("ten", text="Họ và Tên Nhân Viên")
        self.tree_nhansu.heading("chucvu", text="Chức Vụ")
        self.tree_nhansu.column("id", width=80, anchor="center")
        self.tree_nhansu.column("ten", width=300)
        self.tree_nhansu.column("chucvu", width=200)
        self.tree_nhansu.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb_ns = ttk.Scrollbar(list_box, orient=tk.VERTICAL, command=self.tree_nhansu.yview)
        self.tree_nhansu.configure(yscroll=sb_ns.set)
        sb_ns.pack(side=tk.RIGHT, fill=tk.Y)

        btn_del_ns = tk.Button(
            main_frame, text="❌ XÓA NHÂN VIÊN ĐÃ CHỌN", font=("Segoe UI", 10, "bold"),
            bg="#dc2626", fg="white", pady=6, command=self.delete_nhan_vien
        )
        btn_del_ns.pack(anchor=tk.W, pady=5)

    # ============================================================
    # TAB 4: QUẢN LÝ HẠNG MỤC LỖI PHẦN MỀM
    # ============================================================
    def setup_tab_hang_muc_loi(self):
        main_frame = ttk.Frame(self.tab4, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        add_box = ttk.LabelFrame(main_frame, text=" ➕ Thêm Hạng Mục Lỗi Phần Mềm Mới ", padding=10)
        add_box.pack(fill=tk.X, pady=5)

        ttk.Label(add_box, text="Tên Hạng Mục Lỗi:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.txt_new_loi = ttk.Entry(add_box, font=("Segoe UI", 10), width=45)
        self.txt_new_loi.pack(side=tk.LEFT, padx=5)

        btn_add_loi = tk.Button(
            add_box, text="➕ THÊM HẠNG MỤC LỖI", font=("Segoe UI", 10, "bold"),
            bg="#16a34a", fg="white", pady=4, command=self.add_hang_muc_loi
        )
        btn_add_loi.pack(side=tk.LEFT, padx=10)

        list_box = ttk.LabelFrame(main_frame, text=" 📋 Danh Sách Hạng Mục Lỗi Phần Mềm ", padding=10)
        list_box.pack(fill=tk.BOTH, expand=True, pady=10)

        self.tree_loi = ttk.Treeview(list_box, columns=("id", "ten"), show="headings")
        self.tree_loi.heading("id", text="ID")
        self.tree_loi.heading("ten", text="Tên Hạng Mục Lỗi Phần Mềm")
        self.tree_loi.column("id", width=80, anchor="center")
        self.tree_loi.column("ten", width=500)
        self.tree_loi.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb_l = ttk.Scrollbar(list_box, orient=tk.VERTICAL, command=self.tree_loi.yview)
        self.tree_loi.configure(yscroll=sb_l.set)
        sb_l.pack(side=tk.RIGHT, fill=tk.Y)

        btn_del_l = tk.Button(
            main_frame, text="❌ XÓA HẠNG MỤC LỖI ĐÃ CHỌN", font=("Segoe UI", 10, "bold"),
            bg="#dc2626", fg="white", pady=6, command=self.delete_hang_muc_loi
        )
        btn_del_l.pack(anchor=tk.W, pady=5)

    # ============================================================
    # LOGIC NẠP DỮ LIỆU & ĐỒNG BỘ GIAO DIỆN
    # ============================================================
    def reload_all_categories(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # 1. Nạp Khoa/Phòng
        # Tab 1: CHỈ LẤY CÁC KHOA CÓ Trạng thái = 1 (active = 1) ĐỂ HIỂN THỊ TRÊN NHẬT KÝ
        cursor.execute("SELECT id, COALESCE(ma_khoa, ''), ten_khoa FROM khoa_phong WHERE active = 1 ORDER BY ten_khoa")
        self.khoa_list = cursor.fetchall()
        self.cb_khoa["values"] = [k[2] for k in self.khoa_list]
        if self.khoa_list: self.cb_khoa.current(0)
        else: self.cb_khoa.set("")

        # Tab 2 Treeview: Lấy TOÀN BỘ Khoa/Phòng (Hiển thị 1 = Hoạt động, 0 = Không hiển thị)
        cursor.execute("SELECT id, COALESCE(ma_khoa, ''), ten_khoa, active FROM khoa_phong ORDER BY active DESC, ten_khoa")
        all_khoa = cursor.fetchall()
        
        for item in self.tree_khoa.get_children(): self.tree_khoa.delete(item)
        for k_id, k_ma, k_ten, k_act in all_khoa:
            status_str = "1 (Hoạt động)" if k_act == 1 else "0 (Không hiển thị)"
            self.tree_khoa.insert("", tk.END, values=(k_id, k_ma, k_ten, status_str))

        # 2. Nạp Nhân sự
        cursor.execute("SELECT id, ho_ten, chuc_vu FROM nhan_su WHERE active = 1 ORDER BY ho_ten")
        self.nhansu_list = cursor.fetchall()
        self.cb_nhansu["values"] = [n[1] for n in self.nhansu_list]
        if self.nhansu_list: self.cb_nhansu.current(0)
        else: self.cb_nhansu.set("")

        for item in self.tree_nhansu.get_children(): self.tree_nhansu.delete(item)
        for n in self.nhansu_list: self.tree_nhansu.insert("", tk.END, values=n)

        # 3. Nạp Danh mục lỗi phần mềm
        for child in self.err_box.winfo_children():
            child.destroy()
        self.err_vars.clear()

        cursor.execute("SELECT id, ten_loi FROM danh_muc_loi ORDER BY thu_tu, id")
        loi_list = cursor.fetchall()
        
        for idx, (loi_id, ten_loi) in enumerate(loi_list):
            var = tk.IntVar(value=0)
            chk = ttk.Checkbutton(self.err_box, text=ten_loi, variable=var)
            col = idx % 2
            row = idx // 2
            chk.grid(row=row, column=col, sticky=tk.W, padx=10, pady=4)
            self.err_vars[loi_id] = var

        for item in self.tree_loi.get_children(): self.tree_loi.delete(item)
        for l in loi_list: self.tree_loi.insert("", tk.END, values=l)

        conn.close()

    # ============================================================
    # THAO TÁC KHOA PHÒNG (THÊM / CHỈNH SỬA / XÓA / ĐỔI TRẠNG THÁI / IMPORT / FILE MẪU)
    # ============================================================
    def on_select_khoa_phong(self, event=None):
        selected = self.tree_khoa.selection()
        if not selected:
            return
        item = self.tree_khoa.item(selected[0])
        k_id, k_ma, k_ten, k_status = item["values"]
        
        self.selected_khoa_id = k_id
        self.txt_ma_khoa.delete(0, tk.END)
        self.txt_ma_khoa.insert(0, str(k_ma or ""))
        
        self.txt_ten_khoa.delete(0, tk.END)
        self.txt_ten_khoa.insert(0, str(k_ten or ""))
        
        if "1" in str(k_status):
            self.cb_trangthai_khoa.current(0)
        else:
            self.cb_trangthai_khoa.current(1)

    def clear_khoa_form(self):
        self.selected_khoa_id = None
        self.txt_ma_khoa.delete(0, tk.END)
        self.txt_ten_khoa.delete(0, tk.END)
        self.cb_trangthai_khoa.current(0)

    def add_khoa_phong(self):
        ma = self.txt_ma_khoa.get().strip()
        ten = self.txt_ten_khoa.get().strip()
        stt_val = self.cb_trangthai_khoa.get()
        active_val = 1 if "1" in stt_val else 0

        if not ten:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên Khoa / Phòng mới!")
            return
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO khoa_phong (ma_khoa, ten_khoa, active) VALUES (?, ?, ?);", (ma or None, ten, active_val))
            conn.commit()
            conn.close()

            self.clear_khoa_form()
            self.reload_all_categories()
            messagebox.showinfo("Thành công", f"Đã thêm mới Khoa/Phòng '{ten}' thành công (Trạng thái = {active_val})!")
        except sqlite3.IntegrityError:
            messagebox.showerror("Lỗi", f"Mã hoặc Tên Khoa/Phòng '{ten}' đã tồn tại trong CSDL!")
        except Exception as ex:
            messagebox.showerror("Lỗi", f"Không thể thêm Khoa/Phòng: {ex}")

    def update_khoa_phong(self):
        if not self.selected_khoa_id:
            selected = self.tree_khoa.selection()
            if selected:
                self.selected_khoa_id = self.tree_khoa.item(selected[0])["values"][0]
            else:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn Khoa / Phòng từ danh sách bên dưới để chỉnh sửa!")
                return

        ma = self.txt_ma_khoa.get().strip()
        ten = self.txt_ten_khoa.get().strip()
        stt_val = self.cb_trangthai_khoa.get()
        active_val = 1 if "1" in stt_val else 0

        if not ten:
            messagebox.showerror("Lỗi", "Tên Khoa / Phòng không được để trống!")
            return

        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE khoa_phong SET ma_khoa = ?, ten_khoa = ?, active = ? WHERE id = ?;",
                (ma or None, ten, active_val, self.selected_khoa_id)
            )
            conn.commit()
            conn.close()

            self.clear_khoa_form()
            self.reload_all_categories()
            messagebox.showinfo("Thành công", f"Đã cập nhật chỉnh sửa Khoa/Phòng '{ten}' thành công!")
        except sqlite3.IntegrityError:
            messagebox.showerror("Lỗi", f"Mã hoặc Tên Khoa/Phòng '{ten}' bị trùng lặp với khoa phòng khác!")
        except Exception as ex:
            messagebox.showerror("Lỗi", f"Không thể cập nhật Khoa/Phòng: {ex}")

    def toggle_khoa_status(self):
        selected = self.tree_khoa.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn Khoa/Phòng cần đổi trạng thái!")
            return
        item = self.tree_khoa.item(selected[0])
        k_id, k_ma, k_ten, k_status = item["values"]
        
        new_active = 0 if "1" in str(k_status) else 1
        new_status_str = "1 (Hoạt động - Hiển thị nhật ký)" if new_active == 1 else "0 (Ngừng hoạt động - Ẩn khỏi nhật ký)"

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("UPDATE khoa_phong SET active = ? WHERE id = ?;", (new_active, k_id))
        conn.commit()
        conn.close()

        self.clear_khoa_form()
        self.reload_all_categories()
        messagebox.showinfo("Thành công", f"Đã đổi trạng thái Khoa/Phòng '{k_ten}' sang:\n{new_status_str}")

    def download_template_excel(self):
        save_path = filedialog.asksaveasfilename(
            title="Lưu file Mẫu Excel Import Khoa/Phòng",
            initialfile="Mau_Import_Khoa_Phong.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if not save_path:
            return

        try:
            if os.path.exists(TEMPLATE_NAME):
                shutil.copyfile(TEMPLATE_NAME, save_path)
            else:
                import openpyxl
                from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "DM_KHOA_PHONG"
                ws.append(["STT", "Mã Khoa", "Tên Khoa / Phòng", "Trạng Thái (1=Hoạt động, 0=Ngừng)"])
                sample_data = [
                    [1, "K01", "KHOA NỘI THỦY", 1],
                    [2, "K02", "KHOA NGOẠI THỦY", 1],
                    [3, "K03", "KHOA CẤP CỨU", 1],
                    [4, "K04", "KHOA HSTC - CHỐNG ĐỘC", 1],
                    [5, "K05", "KHOA PHỤ SẢN", 1],
                    [6, "K06", "KHOA NHI", 1],
                    [7, "K07", "KHOA DƯỢC", 1],
                    [8, "K08", "PHÒNG TÀI CHÍNH KẾ TOÁN", 1],
                    [9, "K09", "TRUNG TÂM BẢO VỆ SỨC KHỎE", 1],
                    [10, "K10", "KHOA CHẨN ĐOÁN HÌNH ẢNH", 0]
                ]
                for row in sample_data: ws.append(row)
                wb.save(save_path)

            messagebox.showinfo("Thành công", f"Đã lưu file mẫu Excel thành công tại:\n{save_path}")
        except Exception as ex:
            messagebox.showerror("Lỗi", f"Không thể lưu file mẫu: {ex}")

    def import_khoa_phong_excel(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel danh sách Khoa / Phòng",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )
        if not file_path:
            return
        
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active

            header_row = 1
            stt_col, ma_col, ten_col, status_col = -1, -1, -1, -1

            # Quét tiêu đề trong 15 dòng đầu
            for r in range(1, min(15, ws.max_row + 1)):
                row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
                for c_idx, val in enumerate(row_vals, 1):
                    val_norm = remove_accents(val).lower()
                    if "stt" in val_norm and stt_col == -1:
                        stt_col = c_idx
                    elif ("ma" in val_norm or "code" in val_norm) and ("khoa" in val_norm or "phong" in val_norm or "don vi" in val_norm) and ma_col == -1:
                        ma_col = c_idx
                    elif ("ten" in val_norm or "khoa" in val_norm or "phong" in val_norm or "don vi" in val_norm) and "ma" not in val_norm and ten_col == -1:
                        ten_col = c_idx
                    elif ("trang thai" in val_norm or "status" in val_norm or "active" in val_norm) and status_col == -1:
                        status_col = c_idx

                if ten_col != -1:
                    header_row = r
                    break

            if ma_col == -1: ma_col = 2
            if ten_col == -1: ten_col = 3
            if status_col == -1: status_col = 4

            count_added = 0
            count_updated = 0

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()

            for r in range(header_row + 1, ws.max_row + 1):
                ma_val = clean_text(ws.cell(r, ma_col).value) if ma_col <= ws.max_column else ""
                ten_val = clean_text(ws.cell(r, ten_col).value) if ten_col <= ws.max_column else ""
                status_val = str(ws.cell(r, status_col).value if status_col <= ws.max_column else "").strip()

                if not ten_val and ma_val:
                    ten_val = ma_val

                if not ten_val:
                    continue

                # Phân tích Trạng thái 1 (Hoạt động) / 0 (Ngừng / Không hiển thị)
                status_norm = remove_accents(status_val).lower()
                if status_val == "0" or any(k in status_norm for k in ("ngung", "inactive", "false", "tam dung", "an", "khoa")):
                    is_active = 0
                else:
                    is_active = 1

                # Kiểm tra khoa đã có chưa
                cursor.execute("SELECT id, ma_khoa, ten_khoa FROM khoa_phong WHERE ten_khoa = ? OR (ma_khoa = ? AND ma_khoa != '')", (ten_val, ma_val))
                row_exist = cursor.fetchone()

                if row_exist:
                    cursor.execute("UPDATE khoa_phong SET ma_khoa = COALESCE(NULLIF(?, ''), ma_khoa), active = ? WHERE id = ?", (ma_val, is_active, row_exist[0]))
                    count_updated += 1
                else:
                    cursor.execute("INSERT INTO khoa_phong (ma_khoa, ten_khoa, active) VALUES (?, ?, ?)", (ma_val or None, ten_val, is_active))
                    count_added += 1

            conn.commit()
            conn.close()

            self.reload_all_categories()
            messagebox.showinfo(
                "Thành Công", 
                f"Đã nạp dữ liệu Khoa/Phòng từ Excel thành công!\n\n"
                f"- Số Khoa/Phòng thêm mới: {count_added}\n"
                f"- Số Khoa/Phòng cập nhật mã/trạng thái: {count_updated}"
            )
        except Exception as ex:
            messagebox.showerror("Lỗi Import Excel", f"Đã xảy ra lỗi khi đọc file Excel: {ex}")

    def delete_khoa_phong(self):
        selected = self.tree_khoa.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn Khoa/Phòng cần xóa!")
            return
        item = self.tree_khoa.item(selected[0])
        k_id, k_ma, k_ten, k_status = item["values"]
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa hẳn Khoa/Phòng '{k_ten}' khỏi hệ thống?\n(Mẹo: Bạn nên đổi Trạng thái sang 0 thay vì xóa hẳn để giữ lại lịch sử sửa chữa)"):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM khoa_phong WHERE id = ?;", (k_id,))
            conn.commit()
            conn.close()
            self.reload_all_categories()
            messagebox.showinfo("Thành công", "Đã xóa Khoa/Phòng thành công!")

    # ============================================================
    # THAO TÁC NHÂN VIÊN & HẠNG MỤC LỖI
    # ============================================================
    def add_nhan_vien(self):
        ten = self.txt_new_nhansu.get().strip()
        chucvu = self.txt_new_chucvu.get().strip() or "Kỹ thuật viên"
        if not ten:
            messagebox.showerror("Lỗi", "Vui lòng nhập họ và tên Nhân viên!")
            return
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO nhan_su (ho_ten, chuc_vu) VALUES (?, ?);", (ten, chucvu))
            conn.commit()
            conn.close()
            self.txt_new_nhansu.delete(0, tk.END)
            self.reload_all_categories()
            messagebox.showinfo("Thành công", f"Đã thêm Nhân viên '{ten}' thành công!")
        except sqlite3.IntegrityError:
            messagebox.showerror("Lỗi", f"Tên Nhân viên '{ten}' đã tồn tại!")

    def delete_nhan_vien(self):
        selected = self.tree_nhansu.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn Nhân viên cần xóa!")
            return
        item = self.tree_nhansu.item(selected[0])
        ns_id, ns_ten, _ = item["values"]
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa Nhân viên '{ns_ten}'?"):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM nhan_su WHERE id = ?;", (ns_id,))
            conn.commit()
            conn.close()
            self.reload_all_categories()
            messagebox.showinfo("Thành công", "Đã xóa Nhân viên thành công!")

    def add_hang_muc_loi(self):
        ten = self.txt_new_loi.get().strip()
        if not ten:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên Hạng mục lỗi phần mềm mới!")
            return
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO danh_muc_loi (ten_loi) VALUES (?);", (ten,))
            conn.commit()
            conn.close()
            self.txt_new_loi.delete(0, tk.END)
            self.reload_all_categories()
            messagebox.showinfo("Thành công", f"Đã thêm Hạng mục lỗi '{ten}' thành công!")
        except sqlite3.IntegrityError:
            messagebox.showerror("Lỗi", f"Hạng mục lỗi '{ten}' đã tồn tại!")

    def delete_hang_muc_loi(self):
        selected = self.tree_loi.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn Hạng mục lỗi cần xóa!")
            return
        item = self.tree_loi.item(selected[0])
        l_id, l_ten = item["values"]
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa Hạng mục lỗi '{l_ten}'?"):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM danh_muc_loi WHERE id = ?;", (l_id,))
            conn.commit()
            conn.close()
            self.reload_all_categories()
            messagebox.showinfo("Thành công", "Đã xóa Hạng mục lỗi thành công!")

    # ============================================================
    # LƯU PHIẾU NHẬP LIỆU & LỊCH SỬ
    # ============================================================
    def save_entry(self):
        ngay = self.txt_ngay.get().strip()
        khoa_idx = self.cb_khoa.current()
        nhansu_idx = self.cb_nhansu.current()
        loi_khac = self.txt_loi_khac.get().strip()

        if khoa_idx == -1 or nhansu_idx == -1:
            messagebox.showerror("Lỗi", "Vui lòng chọn Khoa/Phòng và Nhân sự thực hiện!")
            return

        khoa_id = self.khoa_list[khoa_idx][0]
        nhansu_id = self.nhansu_list[nhansu_idx][0]

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        cursor.execute("""
            INSERT INTO nhat_ky_sua_chua (ngay_thuc_hien, khoa_phong_id, nhan_su_id, loi_khac)
            VALUES (?, ?, ?, ?);
        """, (ngay, khoa_id, nhansu_id, loi_khac))
        nhat_ky_id = cursor.lastrowid

        has_error_selected = False
        for loi_id, var in self.err_vars.items():
            if var.get() == 1:
                cursor.execute("""
                    INSERT INTO chi_tiet_loi (nhat_ky_id, danh_muc_loi_id, so_luong)
                    VALUES (?, ?, 1);
                """, (nhat_ky_id, loi_id))
                has_error_selected = True

        if not has_error_selected and not loi_khac:
            messagebox.showwarning("Cảnh báo", "Bạn chưa chọn lỗi phần mềm hoặc nhập diễn giải lỗi khác!")
            conn.rollback()
            conn.close()
            return

        conn.commit()
        conn.close()

        self.txt_loi_khac.delete(0, tk.END)
        for var in self.err_vars.values(): var.set(0)

        messagebox.showinfo("Thành công", "Đã lưu phiếu sửa chữa thành công!")
        self.load_history()

    def delete_selected_history(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn dòng cần xóa trong bảng nhật ký!")
            return
        
        item = self.tree.item(selected[0])
        nk_id = item["values"][0]

        if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa lượt sửa chữa ID #{nk_id}?"):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM nhat_ky_sua_chua WHERE id = ?;", (nk_id,))
            conn.commit()
            conn.close()
            self.load_history()
            messagebox.showinfo("Thành công", "Đã xóa bản ghi nhật ký thành công!")

    def load_history(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT nk.id, nk.ngay_thuc_hien, kp.ten_khoa, ns.ho_ten, nk.loi_khac
            FROM nhat_ky_sua_chua nk
            JOIN khoa_phong kp ON nk.khoa_phong_id = kp.id
            JOIN nhan_su ns ON nk.nhan_su_id = ns.id
            ORDER BY nk.id DESC
        """)
        rows = cursor.fetchall()
        
        for r in rows:
            nk_id, ngay, khoa, nhansu, loi_khac = r
            cursor.execute("""
                SELECT dml.ten_loi 
                FROM chi_tiet_loi ctl
                JOIN danh_muc_loi dml ON ctl.danh_muc_loi_id = dml.id
                WHERE ctl.nhat_ky_id = ?
            """, (nk_id,))
            loi_pm = ", ".join([l[0] for l in cursor.fetchall()])
            
            self.tree.insert("", tk.END, values=(nk_id, ngay, khoa, nhansu, loi_pm, loi_khac or ""))
            
        conn.close()

if __name__ == "__main__":
    app = AppNhapLieu()
    app.mainloop()
