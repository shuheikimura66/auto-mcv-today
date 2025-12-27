import os
import json
import time
import glob
import csv
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 環境変数 ---
USER_ID = os.environ["USER_ID"]
PASSWORD = os.environ["USER_PASS"]
json_creds = json.loads(os.environ["GCP_JSON"])

# --- 設定 ---
TARGET_URL = "https://asp1.six-pack.xyz/admin/log/click/list"
DRIVE_FOLDER_ID = "1R49uIPjJ0amEr2He4qQTZ1Z1rQy7nTHf"
SPREADSHEET_ID = "1xzIbMw-YqGn7_KG-xJ_vpWK3MYQY9Zbnm6THTkcZtdM" # 転記先スプレッドシートID
SHEET_NAME = "raw_cv_当日" # 転記先シート名

def get_google_service(service_name, version):
    """Google APIサービスを取得するヘルパー関数"""
    # DriveとSheets両方の権限を設定
    scopes = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets'
    ]
    creds = Credentials.from_service_account_info(json_creds, scopes=scopes)
    return build(service_name, version, credentials=creds)

def upload_to_drive(file_path):
    """Google DriveにCSVファイルをアップロードする関数"""
    print(f"ドライブへのアップロードを開始: {file_path}")
    service = get_google_service('drive', 'v3')

    file_name = os.path.basename(file_path)
    
    file_metadata = {
        'name': file_name,
        'parents': [DRIVE_FOLDER_ID]
    }
    media = MediaFileUpload(file_path, mimetype='text/csv')

    # 共有ドライブ対応 (supportsAllDrives=True)
    file = service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id', 
        supportsAllDrives=True
    ).execute()
    
    print(f"アップロード完了 File ID: {file.get('id')}")

def update_google_sheet(csv_path):
    """CSVの中身を読み込んでスプレッドシートに張り付ける関数"""
    print(f"スプレッドシートへの転記を開始: {SHEET_NAME}")
    service = get_google_service('sheets', 'v4')

    # 1. CSVデータの読み込み (文字コード判定付き)
    csv_data = []
    try:
        # まずUTF-8で試行
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            csv_data = list(reader)
    except UnicodeDecodeError:
        print("UTF-8での読み込みに失敗しました。Shift_JIS(CP932)で再試行します。")
        try:
            # 失敗したらShift_JIS(CP932)で試行 (日本のASPによくある形式)
            with open(csv_path, 'r', encoding='cp932') as f:
                reader = csv.reader(f)
                csv_data = list(reader)
        except Exception as e:
            print(f"CSV読み込みエラー: {e}")
            return

    if not csv_data:
        print("CSVデータが空のため転記をスキップします。")
        return

    # 2. シートのクリア (古いデータを消す)
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME
        ).execute()
        print("既存データをクリアしました。")
    except Exception as e:
        print(f"シートクリアエラー(シートが存在しない可能性があります): {e}")
        # シートがない場合は新規作成などの処理が必要ですが、今回はエラーログのみ

    # 3. データの書き込み
    body = {
        'values': csv_data
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1",
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()

    print(f"スプレッドシート更新完了: {result.get('updatedCells')} セル更新")

def get_today_jst():
    """日本時間の【当日】を計算して文字列(YYYY年MM月DD日)で返す"""
    JST = timezone(timedelta(hours=+9), 'JST')
    now = datetime.now(JST)
    # yesterday = now - timedelta(days=1) # 前日取得ロジックを削除
    return now.strftime("%Y年%m月%d日")

def main():
    print("=== MCV取得処理開始(当日分) ===")
    
    download_dir = os.path.join(os.getcwd(), "downloads_mcv")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- 1. ログイン ---
        safe_user = quote(USER_ID, safe='')
        safe_pass = quote(PASSWORD, safe='')
        url_body = TARGET_URL.replace("https://", "").replace("http://", "")
        auth_url = f"https://{safe_user}:{safe_pass}@{url_body}"
        
        print(f"アクセス中: {TARGET_URL}")
        driver.get(auth_url)
        time.sleep(3)

        # 画面リフレッシュ
        print("画面を再読み込みします...")
        driver.get(auth_url)
        time.sleep(5) 

        # --- 2. 検索メニューを開く ---
        print("検索メニューを開きます...")
        try:
            filter_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '絞り込み検索')]")))
            filter_btn.click()
            time.sleep(3)
        except:
            pass

        # --- 3. 日付入力（当日） ---
        try:
            # 日本時間の「当日」を取得
            today_str = get_today_jst()
            date_range_str = f"{today_str} - {today_str}"
            print(f"日付範囲を指定します: {date_range_str}")
            
            # 日付入力欄を探す
            date_label = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'クリック日時')]")))
            date_input = date_label.find_element(By.XPATH, "./following::input[1]")
            
            # JSでクリア＆入力
            driver.execute_script("arguments[0].click();", date_input)
            driver.execute_script("arguments[0].value = '';", date_input)
            time.sleep(0.5)
            
            date_input.send_keys(date_range_str)
            date_input.send_keys(Keys.ENTER)
            print("日付を入力しました")
            time.sleep(2)
            
        except Exception as e:
            print(f"日付入力エラー: {e}")
            import traceback
            traceback.print_exc()

        # --- 4. パートナー入力 ---
        print("パートナーを入力します...")
        try:
            partner_label = driver.find_element(By.XPATH, "//div[contains(text(), 'パートナー')] | //label[contains(text(), 'パートナー')]")
            partner_target = partner_label.find_element(By.XPATH, "./following::input[contains(@placeholder, '選択')][1]")
            partner_target.click()
            time.sleep(1)
            
            active_elem = driver.switch_to.active_element
            active_elem.send_keys("株式会社フルアウト")
            time.sleep(3)
            active_elem.send_keys(Keys.ENTER)
            print("パートナーを選択しました")
            time.sleep(2)

        except Exception as e:
            print(f"パートナー入力エラー: {e}")

        # --- 5. 検索ボタン実行 ---
        print("検索ボタンを探して押します...")
        try:
            search_btns = driver.find_elements(By.XPATH, "//input[@value='検索'] | //button[contains(text(), '検索')]")
            target_search_btn = None
            for btn in search_btns:
                if btn.is_displayed():
                    target_search_btn = btn
            
            if target_search_btn:
                driver.execute_script("arguments[0].click();", target_search_btn)
                print("検索ボタンをクリックしました")
            else:
                webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()

        except Exception as e:
            print(f"検索ボタン操作エラー: {e}")
            webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
        
        # --- 6. 検索結果の反映待ち ---
        print("検索結果を待機中...")
        time.sleep(15)

        # --- 7. CSVダウンロード ---
        print("CSV作成ボタンを押します...")
        try:
            csv_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'CSV') or @value='CSV作成' or @value='CSV生成']")))
            driver.execute_script("arguments[0].click();", csv_btn)
            print("CSVボタンをクリックしました")
        except Exception as e:
            print(f"CSVボタンエラー: {e}")
            return
        
        # ダウンロード待ち
        time.sleep(5)
        for i in range(20):
            files = glob.glob(os.path.join(download_dir, "*.csv"))
            if files:
                break
            time.sleep(3)
            
        files = glob.glob(os.path.join(download_dir, "*.csv"))
        if not files:
            print("【エラー】CSVファイルが見つかりません。")
            return
        
        csv_file_path = files[0]
        print(f"ダウンロード成功: {csv_file_path}")

        # --- 8. データの保存処理 ---
        
        # 1. Google Driveへアップロード (元の機能も維持)
        upload_to_drive(csv_file_path)

        # 2. Google SpreadSheetへ転記 (追加機能)
        update_google_sheet(csv_file_path)

    except Exception as e:
        print(f"【エラー発生】: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
