# HỆ THỐNG TỔNG HỢP CẤP CKS, EMAIL, P.CNTT & GIÁM ĐỊNH BẢO HIỂM

Ứng dụng desktop giao diện đồ họa (GUI Tkinter) được phát triển bằng Python nhằm tự động hóa quy trình tổng hợp, đối chiếu và xuất báo cáo dữ liệu hành chính - CNTT (Chữ ký số VGCA, Email công vụ, Báo cáo công việc P.CNTT, Cổng giám định bảo hiểm y tế) cho **Bệnh viện Đa khoa Bắc Ninh Số 2**.

---

## 🌟 TÍNH NĂNG CHÍNH (7 TÁC VỤ)

Ứng dụng tích hợp 7 phân hệ tác vụ chính trong giao diện điều khiển đa tab (Notebook):

### 1. Đối chiếu VGCA (`Ket_qua.xlsx`)
- **Mục đích**: Tự động đọc dữ liệu từ các đơn đề nghị dạng Word (`.docx`, `.doc`), kết hợp đối chiếu với file danh sách tài khoản SSO (`.xlsx`) để trích xuất email công vụ.
- **Kết quả**: Xuất file `Ket_qua.xlsx` gồm **19 cột thông tin chuẩn định dạng Import VGCA** (Họ tên, Ngày sinh, CCCD, Nơi cấp, Email công vụ, Đơn vị cấp 1-4, Chức vụ, Số điện thoại...).

### 2. Tổng hợp cấp Chữ ký số - Mẫu 02 (`.docx`)
- **Mục đích**: Đọc dữ liệu đơn đề nghị từ file Word và tài khoản SSO từ Excel, sau đó điền tự động vào file Công văn mẫu (`Mẫu 02_CV cap CKS...docx`).
- **Đặc điểm**: Giữ nguyên **100% thể thức văn bản, phông chữ, cỡ chữ, viền bảng và định dạng Word gốc** bằng kĩ thuật nhân bản cấu trúc XML bảng (`_tr`).

### 3. Tổng hợp danh sách xin cấp Email công vụ (`.docx`)
- **Mục đích**: Quét các file Word đề nghị để tổng hợp danh sách cán bộ xin cấp mới địa chỉ email công vụ.
- **Kết quả**: Xuất tự động file Công văn Word `Đề nghị cấp mới địa chỉ email công vụ.docx`.

### 4. Báo cáo tổng hợp công việc P.CNTT (`.xlsx`)
- **Mục đích**: Tổng hợp tình hình xử lý sự cố phần mềm và lỗi kỹ thuật khác theo từng Khoa/Phòng/Trung tâm từ file nhật ký sửa chữa hàng ngày.
- **Đặc điểm nổi bật**:
  - Đọc danh sách Sheet cực nhanh (**0.003 giây**) bằng kỹ thuật giải mã trực tiếp XML từ `zipfile`.
  - **Chuẩn format 100% file mẫu `08.09.2026 TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx`**: Xuất đúng **18 cột tiêu đề cố định**, phông chữ Times New Roman (Tiêu đề 18pt bold, Header 14pt bold, Nội dung 16pt, Tổng cộng 14pt bold), độ rộng cột tiêu chuẩn và độ cao dòng tự điều chỉnh.
  - Xuất file Excel `TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx` chuẩn **khổ giấy A4 ngang (Landscape)**, căn kín chiều dọc khi xem trước khi in (Print Preview), đóng băng dòng tiêu đề (`C4`) và bật Bộ lọc tự động (AutoFilter).
  - Cột nội dung **"Đăng bài trên website"** vẫn hiển thị dữ liệu chi tiết ở các hàng nhưng **không tính số tổng ở hàng TỔNG CỘNG**.
  - Cột **"SỬA LỖI KHÁC"** ở hàng TỔNG CỘNG tính **tổng cộng số lượt sửa chữa của tất cả các hạng mục**.

### 5. Báo cáo Cổng giám định Bảo hiểm y tế (`.xlsx`)
- **Mục đích**: Đọc file theo dõi Cổng giám định, tự động lọc các tài khoản người dùng thuộc các khoa/phòng được phép (`ALLOWED_USERS`).
- **Tùy chọn**: Hỗ trợ lọc theo khoảng ngày linh hoạt (Mặc định: Trước và sau 7 ngày so với ngày hiện tại `Hôm nay ± 7 ngày`, Nửa đầu tháng `01-14`, Nửa cuối tháng `15-end`, hoặc Cả tháng `Full`).
- **Kết quả**: Xuất file Excel `DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx` chuẩn **khổ giấy A4 dọc (Portrait)**, autofit độ rộng cột và tự động định dạng chuẩn font Times New Roman.

### 6. Báo cáo Tổng hợp công việc P.CNTT tự động (`.xlsx`)
- **Mục đích**: Đọc dữ liệu tự động từ **dòng 7** đến khi gặp dòng trống hoàn toàn (kể cả STT), hệ thống tự động cộng dồn lượt sửa phần mềm và lỗi khác theo từng Khoa/Phòng.
- **Kết quả**: Xuất file Excel `TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx` định dạng A4 Ngang chuẩn 18 cột theo file mẫu `08.09.2026 TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx` tương tự Tác vụ 4.

### 7. Xuất kết quả VGCA từ 1 file Word duy nhất (`Ket_qua.xlsx`)
- **Mục đích**: Đọc **duy nhất 1 file Word (.docx, .doc)** đã chứa đầy đủ thông tin (Họ tên, Ngày sinh, CCCD, Ngày cấp, Nơi cấp, Email công vụ, Đơn vị, Chức vụ, SĐT...).
- **Đặc điểm**: Độc lập hoàn toàn, **không cần file Excel SSO**. Tự động tách trường thông tin kể cả khi gộp chung ô (ví dụ: Công văn Mẫu 02) hoặc các cột bảng riêng rẽ.
- **Kết quả**: Xuất file `Ket_qua.xlsx` gồm đúng **19 cột chuẩn định dạng Import VGCA** tương tự Tác vụ 1.


---

## 🛠️ YÊU CẦU HỆ THỐNG & THƯ VIỆN

- **Ngôn ngữ**: Python 3.8+
- **Thư viện phụ thuộc**:
  - `openpyxl`: Xử lý đọc/ghi file Excel `.xlsx`, `.xlsm`, định dạng trang in.
  - `python-docx`: Đọc và chỉnh sửa file Word `.docx`.
  - `pywin32` (`win32com`): Hỗ trợ chuyển đổi file định dạng `.doc` cũ sang `.docx` (nếu có trên Windows).
  - `tkinter` / `ttk`: Thư viện giao diện người dùng mặc định của Python.

---

## 🚀 HƯỚNG DẪN CÀI ĐẶT VÀ CHẠY ỨNG DỤNG

### 1. Cài đặt các thư viện cần thiết
Mở Terminal / Command Prompt / PowerShell và chạy lệnh:

```bash
pip install openpyxl python-docx pywin32
```

### 2. Khởi chạy phần mềm
Chạy trực tiếp file script bằng Python:

```bash
python app_tong_hop.py
```

---

## 📂 THÔNG TIN DỮ LIỆU ĐẦU VÀO & MẪU SỬ DỤNG

| Tác vụ | File đầu vào chính | Mẫu công văn / File hỗ trợ | File kết quả đầu ra |
| :--- | :--- | :--- | :--- |
| **Tab 1: Đối chiếu VGCA** | Các file Word `.docx` đề nghị | File Excel SSO chứa Email | `Ket_qua.xlsx` (19 cột VGCA) |
| **Tab 2: Cấp CKS (Mẫu 02)** | Các file Word `.docx` đề nghị | `Mẫu 02_CV cap CKS...docx` | `Mẫu 02_CV cap CKS...docx` (Đã điền bảng) |
| **Tab 3: Cấp Email công vụ** | Các file Word `.docx` đề nghị | `Đề nghị cấp mới địa chỉ email công vụ.docx` | `Đề nghị cấp mới địa chỉ email công vụ.docx` |
| **Tab 4: Báo cáo P.CNTT** | File Excel Báo cáo sửa chữa | — | `TỔNG HỢP CÔNG VIỆC P.CNTT.xlsx` (A4 ngang) |
| **Tab 5: Cổng Giám Định BHYT** | File Excel Giám định BHYT | — | `DANH SÁCH THEO DÕI GIÁM ĐỊNH BẢO HIỂM.xlsx` (A4 dọc) |
| **Tab 6: P.CNTT Tự động (Dòng 7)** | File Excel Báo cáo sửa chữa | — | `TỔNG HỢP CÔNG VIỆC P.CNTT (TỰ ĐỘNG).xlsx` (A4 ngang) |
| **Tab 7: VGCA (Tổng hợp Word)** | Nhiều file Word / Thư mục chứa file Word (.docx, .doc) đầy đủ thông tin | — | `Ket_qua.xlsx` (19 cột VGCA) |

---

## ⚙️ CÁC HÀM XỬ LÝ CỐT LÕI (CORE LOGIC)

- `clean_text()`, `remove_accents()`, `norm_name()`, `norm_cccd()`: Các hàm chuẩn hóa chuỗi tiếng Việt, bỏ dấu, xóa ký tự đặc biệt, đồng nhất định dạng CCCD/CMND.
- `normalize_date()`: Tự động chuyển đổi định dạng ngày tháng về dạng chuẩn `dd/mm/yyyy`.
- `normalize_issue_place()`: Phân loại chuẩn hóa nơi cấp CCCD (`CCSQLHCVTTXH` hoặc `CCSĐKQLCTVDLQGVDC`).
- `read_sso_excel()`: Quét và nạp nhanh bản đồ ánh xạ (Họ tên, CCCD -> Email, Đơn vị, Chức vụ) từ dữ liệu Excel SSO.
- `get_excel_sheet_names_fast()`: Đọc danh sách sheet Excel bằng `zipfile` XML trong khoảng 3 miligiây.
- `run_task_doi_chieu()`, `run_task_cks()`, `run_task_email()`, `run_task_cntt()`, `run_task_giam_dinh()`, `run_task_cntt_tu_dong()`, `run_task_vgca_word()`: Các hàm thực thi chính cho 7 phân hệ tác vụ.
- `MainApp`: Lớp giao diện GUI chính kế thừa từ `tk.Tk`, quản lý đa luồng (`threading.Thread`) để giao diện luôn phản hồi mượt mà trong quá trình xử lý dữ liệu lớn.

---

## 🗃️ ỨNG DỤNG NHẬP DỮ LIỆU & QUẢN LÝ DANH MỤC (`app_nhap_lieu_cntt.py`)

Song song với bộ công cụ tổng hợp báo cáo, phần mềm còn cung cấp ứng dụng quản trị dữ liệu đầu vào `app_nhap_lieu_cntt.py` kết nối với cơ sở dữ liệu SQLite (`quan_ly_cntt.db`).

### Quản lý Trạng Thái Khoa/Phòng (1 hoặc 0):
- **Trạng thái `1` (Hoạt động)**: Cho phép khoa/phòng hiển thị trong danh sách chọn (Combobox) khi lập phiếu Nhật ký sửa chữa hàng ngày.
- **Trạng thái `0` (Không hiển thị / Ngừng hoạt động)**: Khoa/phòng sẽ bị **ẩn khỏi danh sách chọn nhật ký**, giúp danh mục gọn gàng nhưng vẫn giữ nguyên lịch sử báo cáo sửa chữa đã nhập trước đó.
- **Thao tác linh hoạt & Chỉnh sửa trực tiếp**:
  - Tích hợp các ô nhập liệu: **Mã Khoa**, **Tên Khoa / Phòng** và **Trạng Thái (1/0)**.
  - Khi chọn bất kỳ dòng nào trong bảng danh sách, thông tin sẽ **tự động điền lên các ô nhập liệu**.
  - Người dùng dễ dàng chỉnh sửa thông tin và nhấn nút `✏️ CẬP NHẬT / CHỈNH SỬA` để lưu vào SQLite.
  - Hỗ trợ nút `➕ THÊM MỚI KHOA PHÒNG`, `🧹 XÓA TRẮNG FORM`, `🔄 ĐỔI TRẠNG THÁI (1 <-> 0)` và `📊 IMPORT TỪ EXCEL`.


