import requests
import json
import os
import re
from datetime import datetime, timedelta
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


def parse_relative_time(time_str):
    """將相對時間轉換成 datetime 物件"""
    now = datetime.now()

    if "剛才" in time_str:
        return now

    # N 分鐘前
    match = re.search(r'(\d+)\s*分鐘前', time_str)
    if match:
        minutes = int(match.group(1))
        return now - timedelta(minutes=minutes)

    # N 小時前
    match = re.search(r'(\d+)\s*小時前', time_str)
    if match:
        hours = int(match.group(1))
        return now - timedelta(hours=hours)

    # N 天前
    match = re.search(r'(\d+)\s*天前', time_str)
    if match:
        days = int(match.group(1))
        return now - timedelta(days=days)

    # 昨天 HH:MM AM/PM
    if "昨天" in time_str:
        yesterday = now - timedelta(days=1)
        # 嘗試解析時間部分
        time_match = re.search(r'(\d+):(\d+)\s*(AM|PM)?', time_str)
        if time_match:
            hour = int(time_match.group(1))
            minute = int(time_match.group(2))
            ampm = time_match.group(3)
            if ampm == "PM" and hour != 12:
                hour += 12
            elif ampm == "AM" and hour == 12:
                hour = 0
            return yesterday.replace(hour=hour, minute=minute, second=0, microsecond=0)
        return yesterday

    # 星期X 格式
    weekday_map = {
        "星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3,
        "星期五": 4, "星期六": 5, "星期日": 6
    }
    for weekday_name, weekday_num in weekday_map.items():
        if weekday_name in time_str:
            today_weekday = now.weekday()
            days_diff = (today_weekday - weekday_num) % 7
            if days_diff == 0:
                days_diff = 7  # 上週同一天
            return now - timedelta(days=days_diff)

    # MM月DD日 或 YY年MM月DD日 格式
    match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日', time_str)
    if match:
        year = int(match.group(1)) + 2000 if match.group(1) else now.year
        month = int(match.group(2))
        day = int(match.group(3))
        return datetime(year, month, day)

    # 無法解析，返回 None
    return None


def get_activity_from_panel(driver):
    """從詳細資料面板取得活動紀錄"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time

    activities = []

    # 開啟主資料夾
    main_url = f"{BASE_URL}?id={FOLDER_ID}"
    driver.get(main_url)
    time.sleep(8)

    # 點擊「詳細資料」按鈕 - 嘗試多種選擇器
    detail_clicked = False

    # 方法1: data-automationid
    try:
        detail_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-automationid='detailsPane']"))
        )
        detail_button.click()
        detail_clicked = True
        print("使用 data-automationid 點擊成功")
    except:
        pass

    # 方法2: 包含「詳細資料」文字的按鈕
    if not detail_clicked:
        try:
            detail_button = driver.find_element(By.XPATH, "//button[contains(., '詳細資料')]")
            detail_button.click()
            detail_clicked = True
            print("使用 XPATH 文字匹配點擊成功")
        except:
            pass

    # 方法3: aria-label
    if not detail_clicked:
        try:
            detail_button = driver.find_element(By.XPATH, "//button[contains(@aria-label, '詳細資料') or contains(@aria-label, 'Details')]")
            detail_button.click()
            detail_clicked = True
            print("使用 aria-label 點擊成功")
        except:
            pass

    # 方法4: i 標籤圖示按鈕
    if not detail_clicked:
        try:
            detail_button = driver.find_element(By.CSS_SELECTOR, "button[name='Details pane'], button[name='詳細資料窗格']")
            detail_button.click()
            detail_clicked = True
            print("使用 name 屬性點擊成功")
        except:
            pass

    # 方法5: 使用鍵盤快捷鍵 Alt+D
    if not detail_clicked:
        try:
            from selenium.webdriver.common.keys import Keys
            from selenium.webdriver.common.action_chains import ActionChains
            actions = ActionChains(driver)
            actions.key_down(Keys.ALT).send_keys('d').key_up(Keys.ALT).perform()
            detail_clicked = True
            print("使用鍵盤快捷鍵")
        except:
            pass

    if not detail_clicked:
        print("無法開啟詳細資料面板，嘗試直接從頁面讀取活動")

    # 等待面板載入完成
    time.sleep(5)

    # 嘗試滾動面板以載入更多內容
    try:
        driver.execute_script("""
            var panels = document.querySelectorAll('[class*="DetailPane"], [class*="detailPane"], aside, [role="complementary"]');
            panels.forEach(function(panel) {
                panel.scrollTop = panel.scrollHeight;
            });
        """)
        time.sleep(2)
    except:
        pass

    # 截圖以便除錯
    screenshot_path = SCRIPT_DIR / "debug_screenshot.png"
    driver.save_screenshot(str(screenshot_path))
    print(f"已儲存截圖到 {screenshot_path}")

    # 列印頁面 HTML 的一部分以便除錯
    try:
        page_source = driver.page_source
        if "活動" in page_source:
            print("頁面中有「活動」文字")
        if "Activity" in page_source:
            print("頁面中有「Activity」文字")
        if "已編輯" in page_source:
            print("頁面原始碼中有「已編輯」文字")
        else:
            print("頁面原始碼中沒有「已編輯」文字")

        # 檢查 iframe 並切換到包含活動的 iframe
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"找到 {len(iframes)} 個 iframe")
        for i, iframe in enumerate(iframes):
            try:
                driver.switch_to.frame(iframe)
                iframe_source = driver.page_source
                if "已編輯" in iframe_source:
                    print(f"iframe {i} 中有「已編輯」文字，保持在此 iframe")
                    # 不切回 default_content，讓後續程式碼在這個 iframe 中執行
                    break
                driver.switch_to.default_content()
            except:
                driver.switch_to.default_content()

    except Exception as e:
        print(f"除錯資訊取得失敗: {e}")

    # 等待活動面板載入
    time.sleep(3)

    # 嘗試找到活動區域
    try:
        # 從截圖看，活動項目在右側面板中，每個活動有圖示、文字和時間
        # 嘗試找到包含活動的容器

        # 方法1: 找包含「已編輯」或「中建立」的元素
        activity_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '已編輯') or contains(text(), '中建立')]")
        print(f"找到 {len(activity_elements)} 個包含活動文字的元素")

        # 取得整個面板的文字內容
        panel_text = ""

        # 嘗試多種選擇器找面板
        panel_selectors = [
            "[role='complementary']",
            "[class*='DetailPane']",
            "[class*='detailsPane']",
            "[class*='od-ItemContent']",
            "[data-automationid='DetailPane']",
            "aside",
        ]

        for selector in panel_selectors:
            try:
                panel = driver.find_element(By.CSS_SELECTOR, selector)
                panel_text = panel.text
                if panel_text and len(panel_text) > 50:
                    print(f"使用選擇器 '{selector}' 找到面板，文字長度: {len(panel_text)}")
                    break
            except:
                continue

        # 備用：用 JavaScript 取得右側面板
        if not panel_text or "已編輯" not in panel_text:
            try:
                # 用 JS 找到右側面板 - 遍歷所有元素包括 Shadow DOM
                js_result = driver.execute_script("""
                    function getAllText(element) {
                        var text = '';
                        if (element.shadowRoot) {
                            text += getAllText(element.shadowRoot);
                        }
                        if (element.innerText) {
                            text += element.innerText;
                        }
                        return text;
                    }

                    // 嘗試找右側面板的各種可能
                    var selectors = [
                        '.od-DetailPane',
                        '[class*="DetailPane"]',
                        '[class*="detailPane"]',
                        '[class*="ItemActivity"]',
                        '[class*="activityFeed"]',
                        'aside',
                        '[role="complementary"]',
                    ];

                    for (var i = 0; i < selectors.length; i++) {
                        var elements = document.querySelectorAll(selectors[i]);
                        for (var j = 0; j < elements.length; j++) {
                            var text = getAllText(elements[j]);
                            if (text && (text.includes('已編輯') || text.includes('中建立'))) {
                                return text;
                            }
                        }
                    }

                    // 最後嘗試：找所有包含活動文字的元素
                    var all = document.querySelectorAll('*');
                    for (var i = 0; i < all.length; i++) {
                        var text = all[i].innerText || '';
                        if (text.includes('已編輯') && text.includes('小時前') && text.length < 3000) {
                            return text;
                        }
                    }

                    return '';
                """)
                if js_result:
                    panel_text = js_result
                    print(f"使用 JS 結果，長度: {len(panel_text)}")
                else:
                    print("JS 沒有找到包含活動的元素")
            except Exception as e:
                print(f"JS 執行失敗: {e}")

        # 印出面板文字前 500 字以便除錯（可移除）
        # if panel_text:
        #     print(f"面板文字前 500 字:\n{panel_text[:500]}")

        if panel_text:
            # 解析面板文字
            # 格式範例：
            # 「Noah Lin」已編輯SBTi Net Zreo Member List_20251217 .xlsx
            # 16 小時前
            # 「Noah Lin」已在 [CPSG] 中建立「To-do list (CPSG)_20260105.xlsx」
            # 17 小時前

            lines = panel_text.split('\n')

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                # 跳過時間分類標題（昨天、本週、上週等）
                if line in ["昨天", "今天", "本週", "上週", "活動"]:
                    i += 1
                    continue

                # 找到包含「已編輯」或「中建立」或「刪除」的行
                if "已編輯" in line or "中建立" in line or "刪除" in line:
                    activity = {
                        "raw_text": line,
                        "modifier": "",
                        "action": "",
                        "file_name": "",
                        "folder": "",
                        "time_str": "",
                        "time": None
                    }

                    # 解析修改者和動作
                    if "已編輯" in line:
                        # 格式：「Noah Lin」已編輯SBTi Net Zreo Member List_20251217 .xlsx
                        # 或：您已編輯xxx.xlsx
                        match = re.search(r'「(.+?)」已編輯(.+)', line)
                        if match:
                            activity["modifier"] = match.group(1)
                            activity["file_name"] = match.group(2).strip()
                        elif line.startswith("您已編輯"):
                            activity["modifier"] = "您"
                            activity["file_name"] = line.replace("您已編輯", "").strip()
                        activity["action"] = "編輯"

                    elif "中建立" in line:
                        # 格式：「Noah Lin」已在 [CPSG] 中建立「To-do list (CPSG)_20260105.xlsx」
                        match = re.search(r'「(.+?)」已在\s*\[(.+?)\]\s*中建立「(.+?)」', line)
                        if match:
                            activity["modifier"] = match.group(1)
                            activity["folder"] = match.group(2)
                            activity["file_name"] = match.group(3)
                        elif "您已在" in line:
                            match = re.search(r'您已在\s*\[(.+?)\]\s*中建立「(.+?)」', line)
                            if match:
                                activity["modifier"] = "您"
                                activity["folder"] = match.group(1)
                                activity["file_name"] = match.group(2)
                        activity["action"] = "新增"

                    elif "刪除" in line:
                        # 格式：「Noah Lin」已從 [FPIP] 刪除「xxx.xlsx」
                        match = re.search(r'「(.+?)」已從\s*\[(.+?)\]\s*刪除「(.+?)」', line)
                        if match:
                            activity["modifier"] = match.group(1)
                            activity["folder"] = match.group(2)
                            activity["file_name"] = match.group(3)
                        activity["action"] = "刪除"

                    # 往下找時間（通常是下一行）
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        # 時間格式：16 小時前、昨天 1:48 AM、3 天前等
                        if "前" in next_line or "昨天" in next_line or "AM" in next_line or "PM" in next_line:
                            activity["time_str"] = next_line
                            activity["time"] = parse_relative_time(next_line)
                            i += 1  # 跳過時間行

                    # 只保留有效的活動（有修改者和檔案名稱）
                    if activity["modifier"] and activity["file_name"]:
                        activities.append(activity)
                        print(f"  解析活動: {activity['modifier']} {activity['action']} {activity['file_name']} ({activity['time_str']})")

                i += 1

        print(f"找到 {len(activities)} 個活動項目")

    except Exception as e:
        print(f"讀取活動紀錄時發生錯誤: {e}")
        import traceback
        traceback.print_exc()

    return activities


def load_previous_data():
    """載入上次的資料"""
    try:
        data_file = SCRIPT_DIR / "activity_data.json"
        with open(data_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {"last_check": None, "activities": []}


def save_current_data(data):
    """儲存目前的資料"""
    data_file = SCRIPT_DIR / "activity_data.json"
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

    try:
        response = requests.post(TEAMS_WEBHOOK_URL, json=payload)
        if response.status_code == 200 or response.status_code == 202:
            mentioned_names = ", ".join([p['name'] for p in people]) if people else "無"
            print(f"Teams 訊息發送成功（已 @ {mentioned_names}）")
        else:
            print(f"Teams 訊息發送失敗: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"發送 Teams 訊息時發生錯誤: {e}")


def check_for_updates():
    """檢查是否有更新"""
    print(f"開始檢查... {datetime.now()}")

    driver = create_driver()

    try:
        print("登入中...")
        login_microsoft(driver)

        print("讀取活動紀錄...")
        activities = get_activity_from_panel(driver)
        print(f"取得 {len(activities)} 個活動")

        # 印出活動內容以便除錯
        for act in activities:
            print(f"  - {act.get('modifier', '?')} {act.get('action', '?')}: {act.get('file_name', '?')} ({act.get('time_str', '?')})")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        driver.quit()
        return
    finally:
        driver.quit()

    # 載入上次資料
    previous_data = load_previous_data()
    last_check_str = previous_data.get("last_check")

    if last_check_str:
        last_check = datetime.fromisoformat(last_check_str)
    else:
        last_check = None

    # 過濾出新的活動（上次檢查之後的）
    new_activities = []
    for act in activities:
        act_time = act.get("time")
        if act_time and last_check:
            if act_time > last_check:
                new_activities.append(act)
        elif not last_check:
            # 首次執行，不通知
            pass

    # 按修改者分組
    updates_by_modifier = {}
    for act in new_activities:
        modifier = act.get("modifier", "未知")
        if modifier not in updates_by_modifier:
            updates_by_modifier[modifier] = []
        updates_by_modifier[modifier].append(act)

    # 發送通知
    all_people = [
        {"name": "Joy", "sharepoint_name": "Joy Lu"},
        {"name": "Noah", "sharepoint_name": "Noah Lin"},
    ]
    our_team_names = [p["sharepoint_name"] for p in all_people]

    notification_count = 0

    for modifier, changes in updates_by_modifier.items():
        now = datetime.now().strftime("%Y/%m/%d %H:%M")

        # 排除修改者本人
        if modifier in our_team_names:
            people_to_mention = [p for p in all_people if p["sharepoint_name"] != modifier]
        else:
            people_to_mention = all_people

        mention_text = " ".join([f"<at>{p['name']}</at>" for p in people_to_mention])

        # 組合訊息
        changes_text = []
        for change in changes:
            action = change.get("action", "更新")
            file_name = change.get("file_name", "未知檔案")
            folder = change.get("folder", "")
            emoji = "🆕" if action == "新增" else "✏️"
            if folder:
                changes_text.append(f"  {emoji} {action}: [{folder}] {file_name}")
            else:
                changes_text.append(f"  {emoji} {action}: {file_name}")

        message = f"🔔 **金寶 Scope3 資料更新通知**\n\n"
        if mention_text:
            message += f"{mention_text} 請查看 **{modifier}** 的更新：\n\n"
        else:
            message += f"請查看 **{modifier}** 的更新：\n\n"
        message += f"⏰ 檢查時間：{now}\n\n"
        message += "\n".join(changes_text)
        message += f"\n\n[點此查看資料夾]({BASE_URL}?id={FOLDER_ID})"

        send_teams_message_with_mention(message, exclude_names=[modifier])
        notification_count += 1

    if notification_count > 0:
        print(f"發現更新，已發送 {notification_count} 則通知")
    elif not last_check:
        print("首次執行，建立基準資料（不發送通知）")
    else:
        print("沒有發現更新")

    # 儲存目前資料
    current_data = {
        "last_check": datetime.now().isoformat(),
        "activities": [
            {
                "modifier": act.get("modifier", ""),
                "action": act.get("action", ""),
                "file_name": act.get("file_name", ""),
                "folder": act.get("folder", ""),
                "time_str": act.get("time_str", "")
            }
            for act in activities
        ]
    }
    save_current_data(current_data)
    print("完成檢查")


if __name__ == "__main__":
    check_for_updates()
