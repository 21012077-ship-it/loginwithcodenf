import time
import pandas as pd
from datetime import datetime # <--- Thêm thư viện xử lý thời gian
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# --- 1. HÀM XỬ LÝ COOKIE (GIỮ NGUYÊN) ---
def parse_netscape_cookies(raw_data):
    cookies = []
    if not isinstance(raw_data, str): return []
    for line in raw_data.strip().split('\n'):
        if line.startswith('#') or not line.strip(): continue
        try:
            parts = line.split('\t')
            domain = parts[0]
            if "netflix.com" not in domain: continue
            cookie = {
                'domain': domain, 'name': parts[5], 'value': parts[6],
                'path': parts[2], 'secure': parts[3] == 'TRUE'
            }
            if len(parts) > 4 and parts[4].isdigit(): cookie['expiry'] = int(parts[4])
            cookies.append(cookie)
        except: pass
    return cookies

# --- 2. HÀM CHECK TỪNG ACCOUNT (GIỮ NGUYÊN) ---
def check_one_account(cookie_raw, index):
    chrome_options = Options()
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--log-level=3")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    status = "UNKNOWN"
    try:
        driver.get("https://www.netflix.com/vn/")
        cookies = parse_netscape_cookies(cookie_raw)
        if not cookies: return "ERROR: Cookie lỗi"
        for cookie in cookies:
            try: driver.add_cookie(cookie)
            except: pass
            
        driver.get("https://www.netflix.com/browse")
        time.sleep(2) 

        if "login" in driver.current_url:
            status = "DIE (Cookie hết hạn)"
        else:
            profiles = driver.find_elements(By.CSS_SELECTOR, "li.profile")
            if not profiles:
                status = "LIVE | Check lại (Không profile)"
            else:
                found_unlocked = False
                unlocked_names = []
                for profile in profiles:
                    try:
                        profile.find_element(By.CSS_SELECTOR, ".svg-icon-profile-lock")
                    except NoSuchElementException:
                        name = profile.find_element(By.CSS_SELECTOR, ".profile-name").text
                        unlocked_names.append(name)
                        found_unlocked = True
                
                if found_unlocked: status = f"LIVE | Mở: {', '.join(unlocked_names)}"
                else: status = "LIVE | Full PIN"
    except Exception as e:
        status = f"ERROR: {str(e)[:30]}"
    finally:
        driver.quit()
        print(f"[{index}] -> {status}")
        return status

# --- 3. CHƯƠNG TRÌNH CHÍNH (CẬP NHẬT TÊN FILE) ---
def main():
    input_file = "Netflix_Accounts_20260124_154427.xlsx"
    
    # --- TẠO TÊN FILE OUTPUT ĐỘNG ---
    # Lấy thời gian hiện tại: NămThángNgày_GiờPhútGiây
    thoi_gian = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"ket_qua_{thoi_gian}.xlsx" 
    
    print(f"--- Đang đọc file {input_file} ---")
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print("❌ Lỗi: Không tìm thấy file accounts.xlsx")
        return

    if "Cookie_Netscape" not in df.columns:
        print("❌ Lỗi: Thiếu cột 'Cookie_Netscape'")
        return

    results = []
    total = len(df)
    print(f"--- Bắt đầu check {total} tài khoản ---")

    for index, row in df.iterrows():
        print(f"Đang xử lý dòng {index + 1}/{total}...", end=" ")
        ket_qua = check_one_account(str(row["Cookie_Netscape"]), index + 1)
        results.append(ket_qua)

    df["Trang_Thai"] = results

    # Lưu ra file mới với tên độc nhất
    df.to_excel(output_file, index=False)
    print(f"\n✅ Đã xong! Kết quả lưu tại file: {output_file}")

if __name__ == "__main__":
    main()