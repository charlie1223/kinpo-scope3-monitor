import requests
import json
import os
from datetime import datetime

# 設定
SHAREPOINT_URL = "https://kinpogroupinc-my.sharepoint.com/personal/tb690310_kinpogroup_com/_layouts/15/onedrive.aspx"
FOLDER_PATH = "/personal/tb690310_kinpogroup_com/Documents/KINPO GROUP Scope 3"
TEAMS_WEBHOOK_URL = os.environ.get("TEAMS_WEBHOOK_URL")
EMAIL = os.environ.get("SHAREPOINT_EMAIL")
PASSWORD = os.environ.get("SHAREPOINT_PASSWORD")

# 要監控的資料夾
FOLDERS_TO_MONITOR = [
    "Castlenet", "CCBP", "CCBR", "CCMX", "CCSD", "CCSZ",
    "CPMA", "CPMY", "CPPE", "CPPH", "CPSG", "CPTH", "Crownpo", "FPIP"
]

def get_folder_info_selenium():
    """使用 Selenium 取得資料夾資訊"""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    import time

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=chrome_options)
    folder_data = {}

    try:
        # 登入 Microsoft
        driver.get("https://login.microsoftonline.com")
        time.sleep(2)

        # 輸入 Email
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "loginfmt"))
        )
        email_input.send_keys(EMAIL)
        driver.find_element(By.ID, "idSIButton9").click()
        time.sleep(2)

        # 輸入密碼
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "passwd"))
        )
        password_input.send_keys(PASSWORD)
        driver.find_element(By.ID, "idSIButton9").click()
        time.sleep(3)

        # 處理「保持登入」提示
        try:
            stay_signed_in = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "idSIButton9"))
            )
            stay_signed_in.click()
            time.sleep(2)
        except:
            pass

        # 前往 SharePoint 資料夾
        sharepoint_url = "https://kinpogroupinc-my.sharepoint.com/personal/tb690310_kinpogroup_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Ftb690310%5Fkinpogroup%5Fcom%2FDocuments%2FKINPO%20GROUP%20Scope%203"
        driver.get(sharepoint_url)
        time.sleep(5)

        # 取得資料夾列表
        rows = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-automationid='DetailsRowFields']"))
        )

        for row in rows:
            try:
                name_element = row.find_element(By.CSS_SELECTOR, "[data-automationid='name']")
                name = name_element.text.strip()

                # 取得修改時間
                date_element = row.find_element(By.CSS_SELECTOR, "[data-automationid='modified']")
                modified_date = date_element.text.strip()

                if name and name in FOLDERS_TO_MONITOR:
                    folder_data[name] = modified_date
            except Exception as e:
                continue

    except Exception as e:
        print(f"Error: {e}")
        send_teams_message(f"❌ 監控程式發生錯誤：{str(e)}")
    finally:
        driver.quit()

    return folder_data

def load_previous_data():
    """載入上次的資料"""
    try:
        with open("folder_data.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_current_data(data):
    """儲存目前的資料"""
    with open("folder_data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def send_teams_message(message):
    """發送 Teams 訊息"""
    if not TEAMS_WEBHOOK_URL:
        print("No Teams webhook URL configured")
        return

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.2",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": message,
                            "wrap": True
                        }
                    ]
                }
            }
        ]
    }

    try:
        response = requests.post(TEAMS_WEBHOOK_URL, json=payload)
        if response.status_code == 200 or response.status_code == 202:
            print("Teams 訊息發送成功")
        else:
            print(f"Teams 訊息發送失敗: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"發送 Teams 訊息時發生錯誤: {e}")

def check_for_updates():
    """檢查是否有更新"""
    print(f"開始檢查... {datetime.now()}")

    current_data = get_folder_info_selenium()
    previous_data = load_previous_data()

    if not current_data:
        print("無法取得資料夾資訊")
        return

    # 比較差異
    updates = []
    for folder_name, modified_date in current_data.items():
        if folder_name in previous_data:
            if previous_data[folder_name] != modified_date:
                updates.append(f"📁 **{folder_name}** - 更新時間：{modified_date}")
        else:
            updates.append(f"📁 **{folder_name}** - 新資料夾，時間：{modified_date}")

    # 發送通知
    if updates:
        now = datetime.now().strftime("%Y/%m/%d %H:%M")
        message = f"🔔 **金寶 Scope3 資料更新通知**\n\n"
        message += f"⏰ 檢查時間：{now}\n\n"
        message += "以下資料夾有更新：\n\n"
        message += "\n".join(updates)
        message += "\n\n[點此查看資料夾](https://kinpogroupinc-my.sharepoint.com/personal/tb690310_kinpogroup_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Ftb690310%5Fkinpogroup%5Fcom%2FDocuments%2FKINPO%20GROUP%20Scope%203)"

        send_teams_message(message)
        print(f"發現 {len(updates)} 個更新")
    else:
        print("沒有發現更新")

    # 儲存目前資料
    save_current_data(current_data)

if __name__ == "__main__":
    check_for_updates()
