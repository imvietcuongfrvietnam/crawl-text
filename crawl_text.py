import os
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    InvalidSessionIdException, WebDriverException, TimeoutException,
)
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# CẤU HÌNH
# ==========================================
INPUT_FILE  = 'raw.xlsx'
INPUT_SHEET = '1. Raw data'
BASE_DATA   = os.path.join(os.getcwd(), 'data')
TEMP_DL     = os.path.join(os.getcwd(), '_temp_dl')
TIMEOUT     = 5  # giây — áp dụng cho mọi thao tác chờ

for _d in (BASE_DATA, TEMP_DL):
    os.makedirs(_d, exist_ok=True)

# ==========================================
# CHROME
# ==========================================
def make_chrome_options():
    opt = Options()
    # opt.add_argument('--headless')
    opt.add_experimental_option("prefs", {
        "download.default_directory": TEMP_DL,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
    })
    opt.add_argument("--disable-features=DownloadBubble,DownloadBubbleV2")
    opt.add_argument("--no-first-run")
    opt.add_argument("--no-default-browser-check")
    return opt

def make_driver():
    svc = Service(ChromeDriverManager().install())
    d = webdriver.Chrome(service=svc, options=make_chrome_options())
    d.implicitly_wait(2)
    return d

# ==========================================
# HELPERS
# ==========================================
def clear_temp():
    for f in os.listdir(TEMP_DL):
        try:
            os.remove(os.path.join(TEMP_DL, f))
        except Exception:
            pass

def wait_for_downloads():
    """Poll mỗi 0.5s, tối đa TIMEOUT giây. Trả về danh sách file đã tải xong."""
    deadline = time.time() + TIMEOUT
    while time.time() < deadline:
        time.sleep(0.5)
        done = [
            f for f in os.listdir(TEMP_DL)
            if not f.endswith('.crdownload') and not f.endswith('.tmp')
        ]
        if done:
            # Đợi thêm 0.5s phòng có thêm file đang tải
            time.sleep(0.5)
            done = [
                f for f in os.listdir(TEMP_DL)
                if not f.endswith('.crdownload') and not f.endswith('.tmp')
            ]
            return done
    return []

def move_all_to_pkg(pkg_folder):
    """Chuyển toàn bộ file trong TEMP_DL sang pkg_folder."""
    moved = []
    for f in os.listdir(TEMP_DL):
        if f.endswith('.crdownload') or f.endswith('.tmp'):
            continue
        src = os.path.join(TEMP_DL, f)
        dst = os.path.join(pkg_folder, f)
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
        moved.append(f)
    return moved

def js_click(driver, el):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", el
    )

def close_extra_tabs(driver, main_window):
    for h in list(driver.window_handles):
        if h != main_window:
            driver.switch_to.window(h)
            driver.close()
    driver.switch_to.window(main_window)

# ==========================================
# MAIN
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
    wait   = WebDriverWait(driver, TIMEOUT)
    status_list = []

    for index, row in df.iterrows():
        url     = row['Link chi tiết']
        ma_tbmt = str(row['Mã TBMT']).strip()
        print(f"\n[{index + 1}/{len(df)}] {ma_tbmt}")

        if pd.isna(url) or str(url).strip() == "":
            print("   -> Link trống, bỏ qua.")
            status_list.append("Không có link")
            continue

        pkg_folder = os.path.join(BASE_DATA, ma_tbmt)
        os.makedirs(pkg_folder, exist_ok=True)

        try:
            # Kiểm tra session còn sống không
            try:
                driver.current_url
            except (InvalidSessionIdException, WebDriverException):
                print("   !! Session died — khởi động lại trình duyệt...")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = make_driver()
                wait   = WebDriverWait(driver, TIMEOUT)

            driver.get(url)
            main_window = driver.current_window_handle
            close_extra_tabs(driver, main_window)

            # --------------------------------------------------
            # BƯỚC 1: Click tab "Hồ sơ mời thầu"
            # --------------------------------------------------
            try:
                hsmt = wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//*[normalize-space(text())='Hồ sơ mời thầu'"
                    " and not(descendant::*[normalize-space(text())='Hồ sơ mời thầu'])]"
                )))
                js_click(driver, hsmt)
                time.sleep(1.5)
                print("   -> Đã click Hồ sơ mời thầu.")
            except TimeoutException:
                print("   -> Timeout: không tìm thấy tab Hồ sơ mời thầu.")
                status_list.append("Timeout: không tìm thấy tab HSMT")
                continue

            # --------------------------------------------------
            # BƯỚC 2: Click "Tải tất cả file đính kèm"
            # --------------------------------------------------
            clear_temp()
            try:
                tai_tat_ca = wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//a[contains(normalize-space(.), 'Tải tất cả')]"
                    " | //button[contains(normalize-space(.), 'Tải tất cả')]"
                    " | //span[contains(normalize-space(.), 'Tải tất cả')]/.."
                )))
                js_click(driver, tai_tat_ca)
                print("   -> Đã click Tải tất cả file đính kèm.")
            except TimeoutException:
                print("   -> Timeout: không tìm thấy nút Tải tất cả.")
                status_list.append("Timeout: không tìm thấy nút tải tất cả")
                continue

            # Đóng tab viewer nếu có mở (một số link mở viewer thay vì download trực tiếp)
            time.sleep(1)
            close_extra_tabs(driver, main_window)

            # --------------------------------------------------
            # BƯỚC 3: Đợi file tải xong rồi chuyển sang pkg_folder
            # --------------------------------------------------
            done = wait_for_downloads()
            if not done:
                print("   -> Timeout: không có file nào tải về.")
                status_list.append("Timeout: không tải được file")
                continue

            moved = move_all_to_pkg(pkg_folder)
            close_extra_tabs(driver, main_window)
            print(f"   -> Lưu {len(moved)} file: {moved}")
            status_list.append(f"OK: {len(moved)} file — {', '.join(moved)}")

        except (InvalidSessionIdException, WebDriverException) as e:
            err = str(e).split('\n')[0]
            print(f"   !! Lỗi session: {err}")
            status_list.append("Lỗi: session died")
            try:
                driver.quit()
            except Exception:
                pass
            driver = make_driver()
            wait   = WebDriverWait(driver, TIMEOUT)

        except Exception as e:
            err = str(e).split('\n')[0]
            print(f"   -> Lỗi: {err}")
            status_list.append(f"Lỗi: {err[:120]}")
            try:
                close_extra_tabs(driver, driver.current_window_handle)
            except Exception:
                pass

    try:
        driver.quit()
    except Exception:
        pass

    df['status'] = status_list
    output_file = 'ket_qua_cuoi_cung.xlsx'
    df.to_excel(output_file, index=False)
    print(f"\n======================================")
    print(f"Xong! Kết quả: {output_file}")
    print(f"File tải về: data/{{ma_tbmt}}/")
    print(f"======================================")


if __name__ == "__main__":
    main_process()
