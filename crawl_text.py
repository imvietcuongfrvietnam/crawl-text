import os
import time
import pandas as pd
import fitz  # PyMuPDF để xử lý PDF
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# CẤU HÌNH CƠ BẢN
# ==========================================
INPUT_FILE = 'raw.xlsx'
DATA_FOLDER = os.path.join(os.getcwd(), 'data')
ALL_FILES_FOLDER = os.path.join(os.getcwd(), 'all_file')

for _folder in (DATA_FOLDER, ALL_FILES_FOLDER):
    if not os.path.exists(_folder):
        os.makedirs(_folder)

# ==========================================
# CẤU HÌNH CHROME
# ==========================================
chrome_options = Options()
# Bỏ comment dòng dưới nếu muốn chạy ngầm:
# chrome_options.add_argument('--headless')

prefs = {
    "download.default_directory": DATA_FOLDER,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

# ==========================================
# HÀM BÓC TÁCH PDF
# ==========================================
def extract_pdf_text(pdf_path):
    try:
        text = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                text += page.get_text()
        return text
    except Exception as e:
        return f"Lỗi bóc tách: {str(e)}"


def wait_for_download(folder, before_files, timeout=60):
    """Đợi 1 file tải xong, trả về tên file mới. None nếu timeout."""
    for _ in range(timeout):
        time.sleep(1)
        after_files = set(os.listdir(folder))
        new_files = list(after_files - before_files)
        completed = [f for f in new_files if not f.endswith('.crdownload')]
        if completed:
            return completed[0]
    return None


def wait_for_all_downloads(folder, before_files, timeout=90, stable_secs=4):
    """Đợi tất cả file tải xong (không còn .crdownload và ổn định).

    Trả về danh sách tên file mới tải về. [] nếu timeout.
    """
    stable_count = 0
    last_completed = set()

    for _ in range(timeout):
        time.sleep(1)
        after_files = set(os.listdir(folder))
        new_files = after_files - before_files
        in_progress = {f for f in new_files if f.endswith('.crdownload')}
        completed = new_files - in_progress

        if completed and not in_progress:
            if completed == last_completed:
                stable_count += 1
            else:
                stable_count = 0
                last_completed = completed
            if stable_count >= stable_secs:
                return list(completed)
        else:
            stable_count = 0
            last_completed = completed

    return list(last_completed) if last_completed else []


# ==========================================
# CHƯƠNG TRÌNH CHÍNH
# ==========================================
def main_process():
    print(f"Đang đọc file: {INPUT_FILE}...")
    try:
        df = pd.read_excel(INPUT_FILE)
    except Exception as e:
        print(f"Lỗi đọc file: {e}")
        return

    df = df.drop_duplicates(subset=['Mã TBMT'], keep='first').reset_index(drop=True)
    print(f"Số lượng bản ghi cần xử lý: {len(df)}")

    print("Đang mở trình duyệt...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)

    content_list = []

    for index, row in df.iterrows():
        url = row['Link chi tiết']
        ma_tbmt = str(row['Mã TBMT']).strip()

        print(f"\n[{index + 1}/{len(df)}] Đang xử lý: {ma_tbmt}")

        if pd.isna(url) or str(url).strip() == "":
            print("   -> Link trống, bỏ qua.")
            content_list.append("Không có link")
            continue

        try:
            driver.get(url)
            main_window = driver.current_window_handle

            # ----------------------------------------------------------
            # BƯỚC 1: Click tab "Hồ sơ mời thầu"
            # Dùng normalize-space(text()) thay vì contains(., ...) để
            # tránh khớp với các element cha chứa text trong subtree.
            # ----------------------------------------------------------
            hsmtt_xpath = (
                "//*["
                "normalize-space(text())='Hồ sơ mời thầu'"
                " and not(descendant::*[normalize-space(text())='Hồ sơ mời thầu'])"
                "]"
            )
            hsmtt = wait.until(EC.element_to_be_clickable((By.XPATH, hsmtt_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", hsmtt)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", hsmtt)
            time.sleep(2)

            # ----------------------------------------------------------
            # BƯỚC 2: Click "Tải tất cả file đính kèm" → lưu vào all_file/{ma_tbmt}/
            # ----------------------------------------------------------
            pkg_folder = os.path.join(ALL_FILES_FOLDER, ma_tbmt)
            os.makedirs(pkg_folder, exist_ok=True)

            try:
                tai_tat_ca_xpath = (
                    "//button[normalize-space(.)='Tải tất cả file đính kèm']"
                    " | //a[normalize-space(.)='Tải tất cả file đính kèm']"
                    " | //span[normalize-space(text())='Tải tất cả file đính kèm']/.."
                    " | //*[normalize-space(text())='Tải tất cả file đính kèm'"
                    " and not(descendant::*[normalize-space(text())='Tải tất cả file đính kèm'])]"
                )
                btn_tai_tat_ca = wait.until(EC.element_to_be_clickable((By.XPATH, tai_tat_ca_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_tai_tat_ca)
                time.sleep(0.5)

                before_all = set(os.listdir(DATA_FOLDER))
                driver.execute_script("arguments[0].click();", btn_tai_tat_ca)
                print("   -> Đang tải tất cả file đính kèm...")

                all_new = wait_for_all_downloads(DATA_FOLDER, before_all, timeout=120)
                if all_new:
                    for fname in all_new:
                        src = os.path.join(DATA_FOLDER, fname)
                        dst = os.path.join(pkg_folder, fname)
                        if os.path.exists(dst):
                            os.remove(dst)
                        os.rename(src, dst)
                    print(f"   -> Đã lưu {len(all_new)} file vào all_file/{ma_tbmt}/")
                else:
                    print("   -> Không tải được file đính kèm (timeout hoặc không có file).")
            except Exception as e_tai:
                print(f"   -> Bỏ qua 'Tải tất cả': {str(e_tai).split(chr(10))[0]}")

            # ----------------------------------------------------------
            # BƯỚC 3: Click "Biểu mẫu mời thầu" (dòng 1.5)
            # ----------------------------------------------------------
            # Thử tìm trong table row chứa "1.5" trước
            bmmtt_xpath = (
                "//tr[td[normalize-space(text())='1.5']]"
                "//*[normalize-space(text())='Biểu mẫu mời thầu']"
                " | //tr[td[normalize-space(.)='1.5']]"
                "//*[normalize-space(.)='Biểu mẫu mời thầu'"
                " and not(descendant::*[normalize-space(.)='Biểu mẫu mời thầu'])]"
            )
            try:
                bmmtt = wait.until(EC.element_to_be_clickable((By.XPATH, bmmtt_xpath)))
            except Exception:
                # Fallback: tìm text node trực tiếp (không phải trong element cha)
                bmmtt_xpath = (
                    "//*["
                    "normalize-space(text())='Biểu mẫu mời thầu'"
                    " and not(descendant::*[normalize-space(text())='Biểu mẫu mời thầu'])"
                    "]"
                )
                bmmtt = wait.until(EC.element_to_be_clickable((By.XPATH, bmmtt_xpath)))

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", bmmtt)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", bmmtt)

            # ----------------------------------------------------------
            # BƯỚC 4: Đợi tab/cửa sổ mới mở ra (viewer)
            # Trang viewer mở ở tab mới → phải switch sang tab đó
            # ----------------------------------------------------------
            time.sleep(3)
            all_handles = driver.window_handles
            if len(all_handles) > 1:
                viewer_tab = [h for h in all_handles if h != main_window][0]
                driver.switch_to.window(viewer_tab)
                print("   -> Đã chuyển sang tab viewer.")
            else:
                print("   -> Viewer mở trong cùng tab.")

            # ----------------------------------------------------------
            # BƯỚC 5: Click nút "Tải về" trong viewer
            # ----------------------------------------------------------
            before_files = set(os.listdir(DATA_FOLDER))

            btn_xpath = (
                "//button[normalize-space(.)='Tải về']"
                " | //a[normalize-space(.)='Tải về']"
                " | //button[normalize-space(text())='Tải về']"
                " | //a[normalize-space(text())='Tải về']"
                " | //*[contains(@class,'download') and normalize-space(.)='Tải về']"
                " | //span[normalize-space(text())='Tải về']/.."
            )
            btn_download = wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_download)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", btn_download)

            # ----------------------------------------------------------
            # BƯỚC 6: Đợi file tải xong, đổi tên, bóc tách PDF
            # ----------------------------------------------------------
            print("   -> Đang đợi tải file...")
            downloaded_file = wait_for_download(DATA_FOLDER, before_files, timeout=60)

            if downloaded_file:
                old_path = os.path.join(DATA_FOLDER, downloaded_file)
                extension = os.path.splitext(downloaded_file)[1]
                new_path = os.path.join(DATA_FOLDER, f"{ma_tbmt}{extension}")

                if os.path.exists(new_path):
                    os.remove(new_path)
                os.rename(old_path, new_path)
                print(f"   -> Tải thành công: {ma_tbmt}{extension}")

                if extension.lower() == '.pdf':
                    pdf_text = extract_pdf_text(new_path)
                    content_list.append(pdf_text)
                    print("   -> Đã bóc tách nội dung PDF.")
                else:
                    content_list.append(f"File không phải PDF (Đuôi: {extension})")
            else:
                print("   -> Lỗi: Tải file quá lâu.")
                content_list.append("Lỗi tải file (Timeout)")

            # Đóng tab viewer, quay về tab chính
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(main_window)

        except Exception as e:
            err_msg = str(e).split('\n')[0]
            print(f"   -> Lỗi thao tác: {err_msg}")
            content_list.append(f"Lỗi: {err_msg}")
            # Đảm bảo luôn quay về tab chính nếu có lỗi
            try:
                if driver.current_window_handle != main_window:
                    driver.close()
                    driver.switch_to.window(main_window)
            except Exception:
                pass

    driver.quit()

    df['content'] = content_list
    output_file = 'ket_qua_cuoi_cung.xlsx'
    df.to_excel(output_file, index=False)
    print(f"\n======================================")
    print(f"Xong! Dữ liệu đã lưu tại: {output_file}")
    print(f"======================================")


if __name__ == "__main__":
    main_process()
