import os
import re
import time
import shutil
import pandas as pd
import fitz  # PyMuPDF
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# CẤU HÌNH CƠ BẢN
# ==========================================
INPUT_FILE  = 'raw.xlsx'
INPUT_SHEET = '1. Raw data'  # Tên sheet chứa dữ liệu đầy đủ (None = sheet đầu tiên)
BASE_DATA  = os.path.join(os.getcwd(), 'data')    # data/{ma_tbmt}/
TEMP_DL    = os.path.join(os.getcwd(), '_temp_dl') # thư mục tải tạm

for _d in (BASE_DATA, TEMP_DL):
    os.makedirs(_d, exist_ok=True)

# ==========================================
# TẠO CHROME OPTIONS
# ==========================================
def make_chrome_options():
    opt = Options()
    # opt.add_argument('--headless')  # bỏ comment để chạy ngầm

    # Tắt popup "Save As" và tự động tải về TEMP_DL
    opt.add_experimental_option("prefs", {
        "download.default_directory": TEMP_DL,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        # Cho phép tải nhiều file tự động (không hỏi)
        "profile.default_content_setting_values.automatic_downloads": 1,
        # Tắt cảnh báo file nguy hiểm
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
    })
    # Tắt download bubble / confirmation bar ở Chrome mới
    opt.add_argument("--disable-features=DownloadBubble,DownloadBubbleV2")
    opt.add_argument("--no-first-run")
    opt.add_argument("--no-default-browser-check")
    return opt

def make_driver():
    svc = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=svc, options=make_chrome_options())
    driver.implicitly_wait(2)
    return driver

# ==========================================
# HELPERS: TẢI FILE
# ==========================================
def clear_temp():
    """Xoá sạch thư mục tải tạm trước mỗi lần tải."""
    for f in os.listdir(TEMP_DL):
        try:
            os.remove(os.path.join(TEMP_DL, f))
        except Exception:
            pass

def wait_for_download(timeout=90):
    """Đợi đến khi có ít nhất 1 file hoàn chỉnh trong TEMP_DL."""
    for _ in range(timeout):
        time.sleep(1)
        files = os.listdir(TEMP_DL)
        done = [f for f in files if not f.endswith('.crdownload') and not f.endswith('.tmp')]
        if done:
            return done[0]
    return None

def move_to_pkg(filename, pkg_folder, dest_name):
    """Di chuyển file từ TEMP_DL sang pkg_folder với tên dest_name."""
    ext = os.path.splitext(filename)[1]
    src = os.path.join(TEMP_DL, filename)
    dst = os.path.join(pkg_folder, f"{dest_name}{ext}")
    if os.path.exists(dst):
        os.remove(dst)
    shutil.move(src, dst)
    return dst

# ==========================================
# HELPERS: SELENIUM
# ==========================================
CLICK_JS = "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();"

def js_click(driver, el):
    driver.execute_script(CLICK_JS, el)

def close_extra_tabs(driver, main_window):
    for h in driver.window_handles:
        if h != main_window:
            driver.switch_to.window(h)
            driver.close()
    driver.switch_to.window(main_window)

def switch_to_new_tab(driver, main_window, timeout=5):
    """Đợi tab mới xuất hiện rồi switch sang. Trả về True nếu thành công."""
    for _ in range(timeout * 2):
        time.sleep(0.5)
        if len(driver.window_handles) > 1:
            new = [h for h in driver.window_handles if h != main_window][0]
            driver.switch_to.window(new)
            return True
    return False

# ==========================================
# BÓC TÁCH PDF — CHỈ LẤY PHẦN PHẠM VI CUNG CẤP
# ==========================================
def extract_danh_muc(pdf_path):
    """Trích nội dung bảng PHẠM VI CUNG CẤP từ file PDF."""
    try:
        full = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                full += page.get_text()

        # Tìm vị trí bắt đầu của phần PHẠM VI CUNG CẤP
        pattern = re.compile(r'PH[ẠA]M VI CUNG C[ẤA]P', re.IGNORECASE)
        m = pattern.search(full)
        if not m:
            return full.strip()  # fallback: trả về toàn bộ text

        section = full[m.start():]

        # Cắt tại section header tiếp theo để không lấy thừa
        end_pattern = re.compile(
            r'\nM[ẫẪ]u s[ốỐ]\s+\d|'
            r'\nCh[ươưƯ][oO]ng\s+[IVXLC\d]|'
            r'\nPh[ầẦ]n\s+\d|'
            r'\nBi[ểỂ]u m[ẫẪ]u|'
            r'\nYêu c[ầẦ]u',
            re.IGNORECASE
        )
        end_m = end_pattern.search(section, 50)  # bỏ qua 50 ký tự header
        if end_m:
            section = section[:end_m.start()]

        return section.strip()
    except Exception as e:
        return f"Lỗi bóc tách: {e}"

# ==========================================
# TẢI FILE TỪ VIEWER (tab mới, có nút Tải về)
# ==========================================
BTN_TAI_VE_XPATH = (
    "//button[normalize-space(.)='Tải về']"
    " | //a[normalize-space(.)='Tải về']"
    " | //button[normalize-space(text())='Tải về']"
    " | //a[normalize-space(text())='Tải về']"
    " | //span[normalize-space(text())='Tải về']/.."
)

def download_from_viewer(driver, wait, main_window, pkg_folder, dest_name):
    """
    Giả sử driver đang ở tab viewer.
    Click Tải về → đợi download → move sang pkg_folder.
    Trả về đường dẫn file đích, hoặc None nếu thất bại.
    """
    try:
        clear_temp()
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, BTN_TAI_VE_XPATH)))
        js_click(driver, btn)
        print("      · Đã click Tải về, đang đợi file...")
        fname = wait_for_download(timeout=90)
        if fname:
            path = move_to_pkg(fname, pkg_folder, dest_name)
            print(f"      · Lưu: {os.path.basename(path)}")
            return path
        print("      · Timeout tải file từ viewer.")
        return None
    except Exception as e:
        print(f"      · Lỗi viewer: {str(e).split(chr(10))[0]}")
        return None
    finally:
        # Luôn đóng tab viewer, quay về main
        if driver.current_window_handle != main_window:
            driver.close()
            driver.switch_to.window(main_window)

# ==========================================
# CHƯƠNG TRÌNH CHÍNH
# ==========================================
def main_process():
    print(f"Đang đọc file: {INPUT_FILE}...")
    try:
        df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)
    except Exception as e:
        print(f"Lỗi đọc file: {e}")
        return

    df = df.drop_duplicates(subset=['Mã TBMT'], keep='first').reset_index(drop=True)
    print(f"Số bản ghi cần xử lý: {len(df)}")

    driver = make_driver()
    wait   = WebDriverWait(driver, 20)

    content_list = []

    for index, row in df.iterrows():
        url     = row['Link chi tiết']
        ma_tbmt = str(row['Mã TBMT']).strip()
        print(f"\n[{index + 1}/{len(df)}] {ma_tbmt}")

        if pd.isna(url) or str(url).strip() == "":
            print("   -> Link trống, bỏ qua.")
            content_list.append("Không có link")
            continue

        # Tạo thư mục riêng cho gói thầu này
        pkg_folder = os.path.join(BASE_DATA, ma_tbmt)
        os.makedirs(pkg_folder, exist_ok=True)

        try:
            # Khởi động lại driver nếu session đã chết
            try:
                driver.current_url
            except (InvalidSessionIdException, WebDriverException):
                print("   !! Session died — khởi động lại trình duyệt...")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = make_driver()
                wait   = WebDriverWait(driver, 20)

            driver.get(url)
            main_window = driver.current_window_handle
            close_extra_tabs(driver, main_window)  # dọn tab thừa từ vòng trước

            # -------------------------------------------------------
            # BƯỚC 1: Click tab "Hồ sơ mời thầu"
            # -------------------------------------------------------
            hsmtt_xpath = (
                "//*[normalize-space(text())='Hồ sơ mời thầu'"
                " and not(descendant::*[normalize-space(text())='Hồ sơ mời thầu'])]"
            )
            hsmtt = wait.until(EC.element_to_be_clickable((By.XPATH, hsmtt_xpath)))
            js_click(driver, hsmtt)
            time.sleep(2)
            print("   -> Tab Hồ sơ mời thầu đã mở.")

            # -------------------------------------------------------
            # BƯỚC 2: Tải Biểu mẫu mời thầu (row 1.5) qua viewer
            # -------------------------------------------------------
            bmmtt_path = None
            try:
                # Ưu tiên tìm trong hàng STT = 1.5
                bmmtt_xpath = (
                    "//tr[td[normalize-space(text())='1.5']]"
                    "//*[normalize-space(text())='Biểu mẫu mời thầu']"
                )
                try:
                    bmmtt = wait.until(EC.element_to_be_clickable((By.XPATH, bmmtt_xpath)))
                except Exception:
                    bmmtt_xpath = (
                        "//*[normalize-space(text())='Biểu mẫu mời thầu'"
                        " and not(descendant::*[normalize-space(text())='Biểu mẫu mời thầu'])]"
                    )
                    bmmtt = wait.until(EC.element_to_be_clickable((By.XPATH, bmmtt_xpath)))

                js_click(driver, bmmtt)
                print("   -> Đã click Biểu mẫu mời thầu, đợi viewer...")

                if switch_to_new_tab(driver, main_window):
                    bmmtt_path = download_from_viewer(
                        driver, wait, main_window, pkg_folder, "bieu_mau_moi_thau"
                    )
                else:
                    print("   -> Không mở được tab viewer cho Biểu mẫu mời thầu.")
            except Exception as e:
                print(f"   -> Lỗi Biểu mẫu mời thầu: {str(e).split(chr(10))[0]}")

            # -------------------------------------------------------
            # BƯỚC 3: Bóc tách danh mục hàng hoá từ PDF
            # -------------------------------------------------------
            extracted = ""
            if bmmtt_path and bmmtt_path.lower().endswith('.pdf'):
                extracted = extract_danh_muc(bmmtt_path)
                print(f"   -> Đã bóc tách danh mục ({len(extracted)} ký tự).")
            elif bmmtt_path:
                extracted = f"File không phải PDF: {os.path.basename(bmmtt_path)}"
            else:
                extracted = "Không tải được Biểu mẫu mời thầu"
            content_list.append(extracted)

            # -------------------------------------------------------
            # BƯỚC 4: Tải tài liệu Yêu cầu kỹ thuật (row 2.1)
            # Row 2.1 có thể có nhiều file với nhiều định dạng khác nhau
            # (pdf, docx, xlsx, ...) → tải tất cả
            # -------------------------------------------------------
            try:
                # Dùng normalize-space(.) thay vì text() để khớp kể cả
                # khi STT nằm trong thẻ con <span> bên trong <td>
                row_xpath = "//tr[td[normalize-space(.)='2.1']]"
                row_el = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, row_xpath))
                )

                # Lấy tất cả phần tử con có thể click được (a, span, button)
                candidates = row_el.find_elements(By.XPATH, ".//a | .//span | .//button")

                FILE_EXTS = (
                    '.pdf', '.doc', '.docx', '.xls', '.xlsx',
                    '.zip', '.rar', '.ppt', '.pptx', '.txt', '.odt',
                )

                # Dùng JS textContent thay vì .text — .text của Selenium
                # trả về rỗng với chip/badge được render bằng CSS
                def get_text(el):
                    return driver.execute_script(
                        "return (arguments[0].textContent || '').trim().toLowerCase();", el
                    )

                all_matched = [
                    el for el in candidates
                    if any(get_text(el).endswith(ext) for ext in FILE_EXTS)
                    or any(ext in get_text(el) for ext in FILE_EXTS)
                ]

                # Dedup theo text: XPath trả về theo document order (cha trước con),
                # nên giữ phần tử đầu tiên của mỗi tên file là tự động giữ phần tử
                # ngoài cùng — bỏ qua <span> con bên trong <a> cùng tên.
                seen_text = set()
                chips = []
                for el in all_matched:
                    txt = get_text(el)
                    if txt not in seen_text:
                        seen_text.add(txt)
                        chips.append(el)

                if not chips:
                    print("   -> Không tìm thấy file đính kèm trong row 2.1.")
                else:
                    names = [driver.execute_script(
                        "return (arguments[0].textContent || '').trim();", c
                    ) for c in chips]
                    print(f"   -> Tìm thấy {len(chips)} file KT: {names}")

                for i, chip in enumerate(chips):
                    suffix = f"_{i + 1}" if len(chips) > 1 else ""
                    dest_name = f"yeu_cau_ky_thuat{suffix}"
                    chip_name = driver.execute_script(
                        "return (arguments[0].textContent || '').trim();", chip
                    )
                    print(f"   -> [{i+1}/{len(chips)}] Tải: {chip_name}")
                    try:
                        clear_temp()            # xoá temp trước khi click
                        js_click(driver, chip)
                        time.sleep(1.5)         # đợi để biết có mở tab mới không

                        if len(driver.window_handles) > 1:
                            switch_to_new_tab(driver, main_window)
                            download_from_viewer(
                                driver, wait, main_window, pkg_folder, dest_name
                            )
                        else:
                            # Download trực tiếp (không qua viewer)
                            fname = wait_for_download(timeout=3)
                            if fname:
                                path = move_to_pkg(fname, pkg_folder, dest_name)
                                print(f"      · Lưu: {os.path.basename(path)}")
                            else:
                                print(f"      · Timeout tải file {i+1}.")
                    except Exception as e_chip:
                        print(f"      · Lỗi file {i+1}: {str(e_chip).split(chr(10))[0]}")

                    close_extra_tabs(driver, main_window)

            except Exception as e:
                print(f"   -> Lỗi Yêu cầu kỹ thuật: {str(e).split(chr(10))[0]}")

            # Dọn sạch tab thừa
            close_extra_tabs(driver, main_window)

        except (InvalidSessionIdException, WebDriverException) as e:
            print(f"   !! Lỗi session: {str(e).split(chr(10))[0]}")
            content_list.append(f"Lỗi session: {str(e).split(chr(10))[0]}")
            try:
                driver.quit()
            except Exception:
                pass
            driver = make_driver()
            wait   = WebDriverWait(driver, 20)

        except Exception as e:
            err = str(e).split('\n')[0]
            print(f"   -> Lỗi: {err}")
            if len(content_list) <= index:
                content_list.append(f"Lỗi: {err}")
            # Dọn tab thừa nếu có
            try:
                close_extra_tabs(driver, driver.current_window_handle)
            except Exception:
                pass

    try:
        driver.quit()
    except Exception:
        pass

    # Ghi kết quả ra Excel
    # Giới hạn 30000 ký tự/ô để tránh cảnh báo Excel
    df['content'] = [str(c)[:30000] for c in content_list]
    output_file = 'ket_qua_cuoi_cung.xlsx'
    df.to_excel(output_file, index=False)
    print(f"\n======================================")
    print(f"Xong! Kết quả: {output_file}")
    print(f"File gốc: data/{{ma_tbmt}}/bieu_mau_moi_thau.*")
    print(f"         data/{{ma_tbmt}}/yeu_cau_ky_thuat.*")
    print(f"======================================")


if __name__ == "__main__":
    main_process()
