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
INPUT_FILE = 'raw.xlsx'  # Thay bằng tên file Excel của bạn
DATA_FOLDER = os.path.join(os.getcwd(), 'data') # Thư mục lưu file tự động tạo cùng nơi để code

if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)

# ==========================================
# CẤU HÌNH CHROME CHO MÁY CÁ NHÂN
# ==========================================
chrome_options = Options()
# Mình TẮT chế độ headless để bạn có thể nhìn thấy trình duyệt chạy
# Nếu sau này muốn nó chạy ngầm, bạn bỏ dấu # ở dòng dưới:
# chrome_options.add_argument('--headless') 

# Ép tải file về thư mục 'data' mà không hỏi
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

    # Lọc trùng theo Mã TBMT
    df = df.drop_duplicates(subset=['Mã TBMT'], keep='first')
    print(f"Số lượng bản ghi cần xử lý: {len(df)}")

    # Khởi tạo Chrome
    print("Đang mở trình duyệt...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 15)
    
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
            
            # --- ÁP DỤNG CƠ CHẾ ÉP CLICK BẰNG JAVASCRIPT ---
            
            # 1. Click tab Hồ sơ mời thầu
            hsmtt_xpath = "//*[contains(text(), 'Hồ sơ mời thầu') or contains(., 'Hồ sơ mời thầu')]"
            hsmtt = wait.until(EC.presence_of_element_located((By.XPATH, hsmtt_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", hsmtt)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", hsmtt)
            time.sleep(2)
            
            # 2. Click Biểu mẫu mời thầu
            bmmtt_xpath = "//*[contains(text(), 'Biểu mẫu mời thầu') or contains(., 'Biểu mẫu mời thầu')]"
            bmmtt = wait.until(EC.presence_of_element_located((By.XPATH, bmmtt_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", bmmtt)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", bmmtt)
            time.sleep(2.5) # Đợi trang biểu mẫu tải xong
            
            before_files = set(os.listdir(DATA_FOLDER))
            
            # 3. Click nút Tải về
            btn_download_xpath = "//button[contains(., 'Tải về')] | //span[contains(text(), 'Tải về')]"
            btn_download = wait.until(EC.presence_of_element_located((By.XPATH, btn_download_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_download)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", btn_download)
            
            # ------------------------------------------------
            
            print("   -> Đang đợi tải file...")
            timeout = 30
            downloaded = False
            
            for _ in range(timeout):
                time.sleep(1)
                after_files = set(os.listdir(DATA_FOLDER))
                new_files = list(after_files - before_files)
                
                # Bắt file khi tải xong (không còn đuôi .crdownload)
                if new_files and not any(f.endswith('.crdownload') for f in new_files):
                    downloaded_file = new_files[0]
                    old_path = os.path.join(DATA_FOLDER, downloaded_file)
                    extension = os.path.splitext(downloaded_file)[1]
                    new_path = os.path.join(DATA_FOLDER, f"{ma_tbmt}{extension}")
                    
                    # Tránh lỗi ghi đè nếu file đã tồn tại
                    if os.path.exists(new_path):
                        os.remove(new_path)
                        
                    os.rename(old_path, new_path)
                    print(f"   -> Tải thành công: {ma_tbmt}{extension}")
                    
                    # Bóc tách nội dung nếu là file PDF
                    if extension.lower() == '.pdf':
                        pdf_text = extract_pdf_text(new_path)
                        content_list.append(pdf_text)
                        print("   -> Đã bóc tách nội dung PDF.")
                    else:
                        content_list.append(f"File không phải PDF (Đuôi: {extension})")
                        
                    downloaded = True
                    break
            
            if not downloaded:
                print("   -> Lỗi: Tải file quá lâu.")
                content_list.append("Lỗi tải file (Timeout)")
                
        except Exception as e:
            err_msg = str(e).split('\n')[0]
            print(f"   -> Lỗi thao tác: {err_msg}")
            content_list.append(f"Lỗi: {err_msg}")

    driver.quit()

    # Cập nhật DataFrame và xuất Excel
    df['content'] = content_list
    output_file = 'ket_qua_cuoi_cung.xlsx'
    df.to_excel(output_file, index=False)
    print(f"\n======================================")
    print(f"Xong! Dữ liệu đã lưu tại: {output_file}")
    print(f"======================================")

if __name__ == "__main__":
    main_process()