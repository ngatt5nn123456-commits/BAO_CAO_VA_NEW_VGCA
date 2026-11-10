# -*- coding: utf-8 -*-
"""
DESKTOP STANDALONE RUNNER
HỆ THỐNG TỔNG HỢP THÔNG TIN - BVĐK BẮC NINH SỐ 2
"""
import os
import sys
import time
import socket
import threading
import subprocess
import webbrowser

# Thiết lập đường dẫn môi trường PyInstaller
if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
    app_dir = os.path.dirname(sys.executable)
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))
    app_dir = bundle_dir

if bundle_dir not in sys.path:
    sys.path.insert(0, bundle_dir)
if app_dir not in sys.path:
    sys.path.insert(0, app_dir)

try:
    os.chdir(app_dir)
except Exception:
    pass

import tkinter as tk
from tkinter import messagebox

# Import ứng dụng Flask
from app import app

def find_free_port(start_port=5000):
    for port in range(start_port, start_port + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return start_port

def run_flask_server(port):
    # Tắt log werkzeug để tăng tốc độ và bảo mật
    import logging
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False, threaded=True)

def open_browser_app(port):
    url = f"http://127.0.0.1:{port}"
    time.sleep(1.0)

    # Thử mở bằng Microsoft Edge ở chế độ cửa sổ App độc lập
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe")
    ]
    # Thử mở bằng Google Chrome
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
    ]

    for p in edge_paths + chrome_paths:
        if os.path.exists(p):
            try:
                subprocess.Popen([p, f"--app={url}", "--window-size=1440,900"])
                return
            except Exception:
                pass

    # Nếu không tìm thấy Edge/Chrome, mở trình duyệt mặc định
    webbrowser.open(url)

class DesktopController(tk.Tk):
    def __init__(self, port):
        super().__init__()
        self.port = port
        self.url = f"http://127.0.0.1:{self.port}"

        self.title("Hệ Thống BVĐK Bắc Ninh số 2")
        self.geometry("460x280")
        self.resizable(False, False)
        self.configure(bg="#0f172a")

        # Căn giữa màn hình
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

        # Header
        hdr = tk.Frame(self, bg="#1e293b", padx=16, pady=12)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr, text="🏥 BVĐK BẮC NINH SỐ 2",
            font=("Segoe UI", 11, "bold"), fg="#38bdf8", bg="#1e293b"
        ).pack(anchor="w")
        tk.Label(
            hdr, text="HỆ THỐNG TỔNG HỢP THÔNG TIN (BẢN OFFLINE .EXE)",
            font=("Segoe UI", 9, "bold"), fg="#94a3b8", bg="#1e293b"
        ).pack(anchor="w")

        # Content
        body = tk.Frame(self, bg="#0f172a", padx=20, pady=16)
        body.pack(fill=tk.BOTH, expand=True)

        status_box = tk.Frame(body, bg="#1e293b", padx=12, pady=10, relief=tk.RIDGE, bd=1)
        status_box.pack(fill=tk.X, pady=(0, 14))

        tk.Label(
            status_box, text="🟢 Máy chủ nội bộ đang hoạt động",
            font=("Segoe UI", 10, "bold"), fg="#4ade80", bg="#1e293b"
        ).pack(anchor="w")
        tk.Label(
            status_box, text=f"Địa chỉ: {self.url}",
            font=("Consolas", 10), fg="#93c5fd", bg="#1e293b"
        ).pack(anchor="w", pady=(4, 0))

        # Nút Mở ứng dụng
        btn_open = tk.Button(
            body, text="🌐 MỞ CỬA SỔ ỨNG DỤNG",
            font=("Segoe UI", 10, "bold"), bg="#2563eb", fg="#ffffff",
            activebackground="#1d4ed8", activeforeground="#ffffff",
            cursor="hand2", relief=tk.FLAT, pady=7, command=self.open_app
        )
        btn_open.pack(fill=tk.X, pady=(0, 8))

        # Hàng nút phụ
        row_btn = tk.Frame(body, bg="#0f172a")
        row_btn.pack(fill=tk.X)

        btn_copy = tk.Button(
            row_btn, text="📋 Sao chép Link",
            font=("Segoe UI", 9), bg="#334155", fg="#e2e8f0",
            activebackground="#475569", activeforeground="#ffffff",
            cursor="hand2", relief=tk.FLAT, pady=5, command=self.copy_link
        )
        btn_copy.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        btn_exit = tk.Button(
            row_btn, text="❌ Thoát phần mềm",
            font=("Segoe UI", 9), bg="#7f1d1d", fg="#fca5a5",
            activebackground="#991b1b", activeforeground="#ffffff",
            cursor="hand2", relief=tk.FLAT, pady=5, command=self.on_close
        )
        btn_exit.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(4, 0))

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def open_app(self):
        threading.Thread(target=open_browser_app, args=(self.port,), daemon=True).start()

    def copy_link(self):
        self.clipboard_clear()
        self.clipboard_append(self.url)
        messagebox.showinfo("Đã sao chép", f"Đã sao chép đường link vào bộ nhớ tạm:\n{self.url}")

    def on_close(self):
        self.destroy()
        os._exit(0)

def main():
    port = find_free_port(5000)

    # Khởi động Flask Server ở luồng ngầm
    server_thread = threading.Thread(target=run_flask_server, args=(port,), daemon=True)
    server_thread.start()

    # Mở cửa sổ ứng dụng tự động
    threading.Thread(target=open_browser_app, args=(port,), daemon=True).start()

    # Khởi động Controller GUI
    gui = DesktopController(port)
    gui.mainloop()

if __name__ == '__main__':
    main()

