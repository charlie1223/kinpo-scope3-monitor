import requests
import json
import os
from datetime import datetime
from pathlib import Path

# 取得腳本所在目錄
SCRIPT_DIR = Path(__file__).parent.absolute()

# 嘗試從 config.py 載入設定，否則用環境變數
try:
    from config import TEAMS_WEBHOOK_URL, SHAREPOINT_EMAIL, SHAREPOINT_PASSWORD
    EMAIL = SHAREPOINT_EMAIL
    PASSWORD = SHAREPOINT_PASSWORD
except ImportError:
    TEAMS_WEBHOOK_URL = os.environ.get("TEAMS_WEBHOOK_URL")
    EMAIL = os.environ.get("SHAREPOINT_EMAIL")
    PASSWORD = os.environ.get("SHAREPOINT_PASSWORD")

# SharePoint 基本網址
BASE_URL = "https://kinpogroupinc-my.sharepoint.com/personal/tb690310_kinpogroup_com/_layouts/15/onedrive.aspx"
FOLDER_ID = "%2Fpersonal%2Ftb690310%5Fkinpogroup%5Fcom%2FDocuments%2FKINPO%20GROUP%20Scope%203"


def create_driver():
    """建立 Selenium driver"""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)


def login_microsoft(driver):
    """登入 Microsoft"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time

    driver.get("https://login.microsoftonline.com")
    time.sleep(3)

    # 輸入 Email
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "loginfmt"))
    )
    email_input.send_keys(EMAIL)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(3)

    # 輸入密碼
    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "passwd"))
    )
    password_input.send_keys(PASSWORD)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(5)

    # 處理「保持登入」提示
    try:
        stay_signed_in = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "idSIButton9"))
        )
        stay_signed_in.click()
        time.sleep(3)
    except:
        pass


def get_items_from_page(driver):
    """從目前頁面取得所有項目（資料夾或檔案）"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time

    items = {}

    # 多等一下讓頁面完全載入
    time.sleep(5)

    # 找所有資料列
    try:
        rows = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[role='row'][data-automationid^='row-']"))
        )
    except:
        rows = []

    for row in rows:
        try:
            # 取得名稱
            name = ""
            try:
                name_element = row.find_element(By.CSS_SELECTOR, "[data-automationid='field-LinkFilename'] span[role='button']")
                name = name_element.text.strip()
            except:
                try:
                    name_element = row.find_element(By.CSS_SELECTOR, "[data-automationid='field-LinkFilename']")
                    name = name_element.text.strip()
                except:
                    pass

            # 取得修改時間
            modified_date = ""
            try:
                date_element = row.find_element(By.CSS_SELECTOR, "[data-automationid='field-Modified']")
                modified_date = date_element.text.strip()
                # 將相對時間轉換成絕對日期
                import re
                from datetime import datetime, timedelta
                if "分鐘前" in modified_date or "剛才" in modified_date:
                    # 幾分鐘前 -> 今天
                    modified_date = datetime.now().strftime("%Y/%m/%d")
                elif "小時前" in modified_date:
                    # 提取小時數，計算實際日期
                    match = re.search(r'(\d+)\s*小時前', modified_date)
                    if match:
                        hours = int(match.group(1))
                        actual_time = datetime.now() - timedelta(hours=hours)
                        modified_date = actual_time.strftime("%Y/%m/%d")
                    else:
                        modified_date = datetime.now().strftime("%Y/%m/%d")
                elif "月" in modified_date and "日" in modified_date:
                    # 12月17日 -> 2025/12/17（加上年份）
                    match = re.search(r'(\d+)月(\d+)日', modified_date)
                    if match:
                        month = int(match.group(1))
                        day = int(match.group(2))
                        year = datetime.now().year
                        modified_date = f"{year}/{month:02d}/{day:02d}"
            except:
                pass

            # 取得修改者
            modified_by = ""
            try:
                modifier_element = row.find_element(By.CSS_SELECTOR, "[data-automationid='field-Editor']")
                modified_by = modifier_element.text.strip()
            except:
                pass

            # 跳過表頭
            if name and name != "名稱":
                items[name] = {"date": modified_date, "by": modified_by}
        except:
            continue

    return items


def get_all_files(driver):
    """取得所有資料夾及其內部檔案"""
    import time
    import urllib.parse

    all_data = {}

    # 先取得主資料夾列表
    main_url = f"{BASE_URL}?id={FOLDER_ID}"
    driver.get(main_url)
    time.sleep(8)

    folders = get_items_from_page(driver)
    print(f"找到 {len(folders)} 個資料夾")

    # 過濾掉檔案（只保留資料夾）
    folder_names = [name for name in folders.keys() if not name.endswith(('.xlsx', '.xls', '.pdf', '.docx', '.doc', '.pptx', '.ppt', '.csv', '.txt'))]

    # 進入每個資料夾取得檔案
    for folder_name in folder_names:
        print(f"  檢查資料夾: {folder_name}")

        folder_url = f"{BASE_URL}?id={FOLDER_ID}%2F{urllib.parse.quote(folder_name)}"
        driver.get(folder_url)
        time.sleep(6)

        files = get_items_from_page(driver)
        all_data[folder_name] = files
        print(f"    找到 {len(files)} 個檔案")

    return all_data


def load_previous_data():
    """載入上次的資料"""
    try:
        data_file = SCRIPT_DIR / "folder_data.json"
        with open(data_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}


def save_current_data(data):
    """儲存目前的資料"""
    data_file = SCRIPT_DIR / "folder_data.json"
    with open(data_file, "w", encoding="utf-8") as f:
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


def send_teams_message_with_mention(message, exclude_names=None):
    """發送 Teams 訊息並 @ Joy 和 Noah（排除修改者）"""
    if not TEAMS_WEBHOOK_URL:
        print("No Teams webhook URL configured")
        return

    if exclude_names is None:
        exclude_names = []

    # 要 @ 的人員列表
    all_people = [
        {"email": "joy.lu@cfgreen-energy.com", "name": "Joy", "sharepoint_name": "Joy Lu"},
        {"email": "noah.lin@cfgreen-energy.com", "name": "Noah", "sharepoint_name": "Noah Lin"},
    ]

    # 過濾掉修改者
    people = [p for p in all_people if p["sharepoint_name"] not in exclude_names]

    # 建立 entities 列表
    entities = []
    for person in people:
        entities.append({
            "type": "mention",
            "text": f"<at>{person['name']}</at>",
            "mentioned": {
                "id": person["email"],
                "name": person["name"]
            }
        })

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
                    ],
                    "msteams": {
                        "entities": entities
                    }
                }
            }
        ]
    }

    # 產生 @ 的文字
    mention_text = " ".join([f"<at>{p['name']}</at>" for p in people])

    try:
        response = requests.post(TEAMS_WEBHOOK_URL, json=payload)
        if response.status_code == 200 or response.status_code == 202:
            mentioned_names = ", ".join([p['name'] for p in people]) if people else "無"
            print(f"Teams 訊息發送成功（已 @ {mentioned_names}）")
        else:
            print(f"Teams 訊息發送失敗: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"發送 Teams 訊息時發生錯誤: {e}")

    return mention_text


def check_for_updates():
    """檢查是否有更新"""
    print(f"開始檢查... {datetime.now()}")

    driver = create_driver()

    try:
        print("登入中...")
        login_microsoft(driver)

        print("掃描所有資料夾...")
        current_data = get_all_files(driver)

    except Exception as e:
        print(f"Error: {e}")
        driver.quit()
        return
    finally:
        driver.quit()

    previous_data = load_previous_data()

    if not current_data:
        print("無法取得資料夾資訊")
        return

    # 比較差異 - 按修改者分組
    updates_by_modifier = {}  # {修改者: [{folder, name, type}]}

    for folder_name, files in current_data.items():
        previous_files = previous_data.get(folder_name, {})

        for file_name, file_info in files.items():
            if file_name in previous_files:
                prev_info = previous_files[file_name]
                # 相容舊格式（純字串）和新格式（dict）
                prev_date = prev_info.get("date", prev_info) if isinstance(prev_info, dict) else prev_info
                curr_date = file_info.get("date", file_info) if isinstance(file_info, dict) else file_info
                if prev_date != curr_date:
                    modifier = file_info.get("by", "") if isinstance(file_info, dict) else ""
                    modifier = modifier or "未知"
                    if modifier not in updates_by_modifier:
                        updates_by_modifier[modifier] = []
                    updates_by_modifier[modifier].append({
                        "folder": folder_name,
                        "name": file_name,
                        "type": "修改"
                    })
            else:
                # 只有當之前有記錄時，才通知新檔案
                if previous_data:
                    modifier = file_info.get("by", "") if isinstance(file_info, dict) else ""
                    modifier = modifier or "未知"
                    if modifier not in updates_by_modifier:
                        updates_by_modifier[modifier] = []
                    updates_by_modifier[modifier].append({
                        "folder": folder_name,
                        "name": file_name,
                        "type": "新增"
                    })

    # 檢查新增的資料夾
    new_folders = []
    if previous_data:
        for folder_name in current_data.keys():
            if folder_name not in previous_data:
                new_folders.append(folder_name)

    # 發送通知 - 每個修改者分開發送
    # 自己人名單（這些人修改時會排除自己）
    all_people = [
        {"name": "Joy", "sharepoint_name": "Joy Lu"},
        {"name": "Noah", "sharepoint_name": "Noah Lin"},
        # 未來可加入：{"name": "Stacy", "sharepoint_name": "Stacy Hung 洪翠禪"},
    ]

    # 自己人的 SharePoint 名稱列表
    our_team_names = [p["sharepoint_name"] for p in all_people]

    notification_count = 0
    skip_notification = False  # 正常發送通知

    for modifier, changes in updates_by_modifier.items():
        now = datetime.now().strftime("%Y/%m/%d %H:%M")

        # 只有當修改者是「自己人」時，才排除那個人
        # 如果是外部人員（如 Stacy），則通知所有人
        if modifier in our_team_names:
            people_to_mention = [p for p in all_people if p["sharepoint_name"] != modifier]
        else:
            people_to_mention = all_people

        mention_text = " ".join([f"<at>{p['name']}</at>" for p in people_to_mention])

        # 按資料夾整理
        folders_dict = {}
        for change in changes:
            folder = change["folder"]
            if folder not in folders_dict:
                folders_dict[folder] = []
            emoji = "🆕" if change["type"] == "新增" else "✏️"
            folders_dict[folder].append(f"  {emoji} {change['type']}: {change['name']}")

        # 組合訊息
        updates_text = []
        for folder, items in folders_dict.items():
            updates_text.append(f"📁 **{folder}**\n" + "\n".join(items))

        message = f"🔔 **金寶 Scope3 資料更新通知**\n\n"
        if mention_text:
            message += f"{mention_text} 請查看 **{modifier}** 的更新：\n\n"
        else:
            message += f"請查看 **{modifier}** 的更新：\n\n"
        message += f"⏰ 檢查時間：{now}\n\n"
        message += "\n\n".join(updates_text)
        message += f"\n\n[點此查看資料夾]({BASE_URL}?id={FOLDER_ID})"

        if not skip_notification:
            send_teams_message_with_mention(message, exclude_names=[modifier])
            notification_count += 1

    # 新增資料夾通知（這個不排除任何人）
    if new_folders and not skip_notification:
        now = datetime.now().strftime("%Y/%m/%d %H:%M")
        mention_text = " ".join([f"<at>{p['name']}</at>" for p in all_people])

        folders_text = "\n".join([f"📁 **{f}** - 🆕 新增資料夾" for f in new_folders])

        message = f"🔔 **金寶 Scope3 資料更新通知**\n\n"
        message += f"{mention_text} 請查看以下新增資料夾：\n\n"
        message += f"⏰ 檢查時間：{now}\n\n"
        message += folders_text
        message += f"\n\n[點此查看資料夾]({BASE_URL}?id={FOLDER_ID})"

        send_teams_message_with_mention(message)
        notification_count += 1

    if notification_count > 0:
        print(f"發現更新，已發送 {notification_count} 則通知")
    elif not previous_data:
        print("首次執行，建立基準資料（不發送通知）")
    else:
        print("沒有發現更新")

    # 儲存目前資料
    save_current_data(current_data)
    print("完成檢查")


if __name__ == "__main__":
    check_for_updates()
