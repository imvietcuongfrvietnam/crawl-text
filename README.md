# crawl-text — Tải hồ sơ mời thầu từ muasamcong.mpi.gov.vn

Script tự động:
1. Vào từng link gói thầu trong file Excel
2. Tải **Biểu mẫu mời thầu** (row 1.5) từ tab viewer
3. Tải **tài liệu Yêu cầu kỹ thuật** (row 2.1) từ file đính kèm
4. Bóc tách phần **PHẠM VI CUNG CẤP** (danh mục hàng hoá) từ PDF
5. Ghi nội dung ra file `ket_qua_cuoi_cung.xlsx`

### Cấu trúc thư mục đầu ra

```
data/
  IB2400548541/
    bieu_mau_moi_thau.pdf
    yeu_cau_ky_thuat.docx
  IB2400xxxxxx/
    ...
ket_qua_cuoi_cung.xlsx
```

---

## Cài đặt

### Bước 1 — Cài Conda (nếu chưa có)

Tải Miniconda tại: https://docs.conda.io/en/latest/miniconda.html  
Chọn bản **Windows 64-bit**, cài đặt bình thường (tick "Add to PATH" khi hỏi).

Kiểm tra sau khi cài:
```bash
conda --version
```

---

### Bước 2 — Tạo môi trường ảo

```bash
conda create -n crawl python=3.11 -y
conda activate crawl
```

> Mỗi lần mở terminal mới cần chạy lại `conda activate crawl` trước khi dùng.

---

### Bước 3 — Cài thư viện

```bash
pip install -r requirements.txt
```

---

### Bước 4 — Cài Google Chrome

Tải và cài Chrome tại: https://www.google.com/chrome/  
Script dùng **webdriver-manager** nên sẽ tự tải đúng phiên bản ChromeDriver — không cần cài thủ công.

---

### Bước 5 — Chuẩn bị file Excel đầu vào

File `raw.xlsx` phải có ít nhất 2 cột:

| Mã TBMT | Link chi tiết |
|---------|--------------|
| IB2400548541 | https://muasamcong.mpi.gov.vn/web/guest/contractor-selection/... |

Đặt file `raw.xlsx` cùng thư mục với `crawl_text.py`.

---

### Bước 6 — Chạy script

```bash
python crawl_text.py
```

Trình duyệt Chrome sẽ tự mở và thao tác. Không cần tương tác thêm.

---

## Tuỳ chỉnh

| Nhu cầu | Cách làm |
|---------|----------|
| Chạy ngầm (không hiện Chrome) | Bỏ comment dòng `# opt.add_argument('--headless')` trong `make_chrome_options()` |
| Đổi file đầu vào | Sửa `INPUT_FILE = 'raw.xlsx'` ở đầu file |
| Tăng timeout tải file | Sửa tham số `timeout=90` trong `wait_for_download()` |

---

## Xử lý lỗi thường gặp

| Lỗi | Nguyên nhân | Cách xử lý |
|-----|-------------|------------|
| `ModuleNotFoundError` | Chưa cài thư viện | Chạy lại `pip install -r requirements.txt` trong môi trường `crawl` |
| `invalid session id` | Chrome bị crash | Script tự khởi động lại Chrome, tiếp tục xử lý |
| `Cell contents too long` | PDF quá dài | Script tự cắt tối đa 30.000 ký tự/ô |
| ChromeDriver version mismatch | Chrome mới hơn driver | `pip install --upgrade webdriver-manager` |
