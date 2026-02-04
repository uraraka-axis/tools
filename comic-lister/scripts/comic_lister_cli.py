"""
コミックリスター自動化ツール（CLI版）
GitHub Actionsで毎週自動実行され、コミックリスターでCSVを生成します。
"""

import time
import os
import sys
import datetime
import re
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd

# 環境変数から設定を取得
HARU_USERNAME = os.environ.get("HARU_USERNAME", "haruuser")
HARU_PASSWORD = os.environ.get("HARU_PASSWORD", "Haru@9999")
GOOGLE_SHEETS_URL = os.environ.get("GOOGLE_SHEETS_URL",
    "https://docs.google.com/spreadsheets/d/1ivKBwOnyHi88F_-OjDPTu0-P2OW81qpS/edit?gid=1315015327#gid=1315015327")
ASSIGNEE_NAME = os.environ.get("ASSIGNEE_NAME", "笹山")
ISBN_SETTING = os.environ.get("ISBN_SETTING", "1st")  # lst, dat, max, 1st

# 出力ディレクトリ
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "/tmp/comic-lister-output"))


def log(message):
    """ログメッセージを出力"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def get_comic_numbers_from_sheets():
    """Google SheetsからコミックNo.を取得"""
    try:
        log("Google Sheetsからデータ取得中...")

        # 公開CSVアクセス用URLを生成
        if "edit" in GOOGLE_SHEETS_URL:
            sheet_id = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', GOOGLE_SHEETS_URL).group(1)
            gid = re.search(r'gid=([0-9]+)', GOOGLE_SHEETS_URL)
            gid = gid.group(1) if gid else "0"
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        else:
            csv_url = GOOGLE_SHEETS_URL

        df = pd.read_csv(csv_url, header=None)

        # C列（2列目）の3行目以降からコミックNo.を取得
        comic_numbers = []
        if len(df.columns) >= 3:
            for i in range(2, len(df)):
                value = df.iloc[i, 2]
                if pd.notna(value) and str(value).strip():
                    comic_numbers.append(str(value).strip())

        log(f"取得したコミックNo.: {len(comic_numbers)}件")
        return comic_numbers

    except Exception as e:
        log(f"Google Sheetsからのデータ取得エラー: {e}")
        return []


def create_list_csv(comic_numbers):
    """コミックNo.リストからlist.csvを作成"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    list_csv_path = OUTPUT_DIR / "list.csv"

    log(f"list.csv作成中... ({len(comic_numbers)}件)")

    # DataFrameを作成（J列にコミックNo.、K列に1）
    csv_data = []
    for comic_no in comic_numbers:
        row = [''] * 9 + [comic_no, 1]
        csv_data.append(row)

    df = pd.DataFrame(csv_data)
    df.to_csv(list_csv_path, index=False, header=False, encoding='utf-8')

    log(f"list.csv保存完了: {list_csv_path}")
    return str(list_csv_path)


class ComicListerAutomator:
    def __init__(self, config):
        self.base_url = "https://haru-u.biz/comic/index.html"
        self.username = HARU_USERNAME
        self.password = HARU_PASSWORD
        self.driver = None
        self.config = config

    def setup_browser(self):
        """ブラウザセットアップ（ヘッドレスモード）"""
        log("ブラウザを起動中...")

        chrome_options = Options()
        chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

        # ダウンロードディレクトリの設定
        prefs = {
            "download.default_directory": str(OUTPUT_DIR),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(10)

        log("ブラウザの起動が完了しました")

    def navigate_to_site(self):
        """サイトにアクセス（ベーシック認証付き）"""
        log("サイトにアクセス中...")

        auth_url = f"https://{self.username}:{self.password}@haru-u.biz/comic/index.html"
        self.driver.get(auth_url)

        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "engn17"))
        )
        log("サイトアクセス完了")

    def click_comic_lister_button(self):
        """コミックリスターボタンをクリック"""
        log("コミックリスターボタンをクリック...")

        comic_lister_btn = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "engn17"))
        )
        comic_lister_btn.click()

        time.sleep(3)
        log("コミックリスターが起動しました")

    def switch_to_iframe(self):
        """コミックリスターのiframeに切り替え"""
        try:
            iframe = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "infrm17"))
            )
            self.driver.switch_to.frame(iframe)
            return True
        except Exception as e:
            log(f"iframe切り替えエラー: {e}")
            return False

    def switch_to_default(self):
        """メインコンテンツに戻る"""
        try:
            alert = self.driver.switch_to.alert
            alert.accept()
        except:
            pass
        self.driver.switch_to.default_content()

    def upload_csv_file(self):
        """CSVファイルをアップロード"""
        log("CSVファイルをアップロード中...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            file_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]'))
            )
            file_input.send_keys(self.config['csv_path'])

            upload_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='アップロード']"))
            )
            upload_btn.click()

            # リスト名入力アラートを処理
            try:
                alert = WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                alert_text = alert.text

                if "名称を入力" in alert_text:
                    alert.send_keys(self.config['list_name'])
                    alert.accept()
                    log(f"リスト名「{self.config['list_name']}」を入力しました")
                else:
                    alert.accept()
                    try:
                        alert2 = WebDriverWait(self.driver, 3).until(EC.alert_is_present())
                        alert2.send_keys(self.config['list_name'])
                        alert2.accept()
                    except TimeoutException:
                        pass

            except TimeoutException:
                pass

            time.sleep(3)
            log("CSVファイルのアップロードが完了しました")

        finally:
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
            except:
                pass
            self.switch_to_default()

    def click_work_button_and_handle_assignee(self):
        """作業ボタンをクリックして担当者入力"""
        log("作業ボタンをクリックして担当者入力処理...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            work_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='作業']"))
            )
            work_btn.click()
            log("作業ボタンをクリックしました")

            time.sleep(2)

            for attempt in range(3):
                try:
                    alert = WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                    alert_text = alert.text

                    if "担当する人の名前" in alert_text or "リストを担当する" in alert_text:
                        alert.send_keys(self.config['assignee_name'])
                        alert.accept()
                        log("担当者入力が完了しました")
                        time.sleep(3)
                        return True
                    else:
                        alert.accept()
                        time.sleep(1)

                except TimeoutException:
                    time.sleep(2)
                    continue

            return False

        finally:
            try:
                alert = WebDriverWait(self.driver, 2).until(EC.alert_is_present())
                alert.accept()
            except:
                pass

            try:
                self.driver.switch_to.default_content()
            except:
                pass

    def click_initial_survey_1(self):
        """初期調査1ボタンをクリック"""
        log("初期調査1を実行...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            survey_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='初期調査1']"))
            )
            survey_btn.click()

            try:
                alert = WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                alert.accept()
            except TimeoutException:
                pass

            log("初期調査1の処理を待機中...")
            time.sleep(30)

            try:
                alert = WebDriverWait(self.driver, 10).until(EC.alert_is_present())
                alert.accept()
            except TimeoutException:
                try:
                    ok_btn = self.driver.find_element(By.XPATH, "//input[@value='OK']")
                    ok_btn.click()
                except NoSuchElementException:
                    pass

            log("初期調査1が完了しました")

        finally:
            self.switch_to_default()

    def click_list_creation_complete(self):
        """リスト作成完了ボタンをクリック"""
        log("リスト作成完了ボタンをクリック...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            complete_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='リスト作成完了']"))
            )
            complete_btn.click()

            try:
                alert = WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                alert_text = alert.text

                if "提出日" in alert_text or "日付" in alert_text:
                    alert.send_keys(self.config['submission_date'])
                    alert.accept()
                    log(f"提出日「{self.config['submission_date']}」を入力しました")
                else:
                    alert.accept()

            except TimeoutException:
                pass

            time.sleep(3)
            log("リスト作成が完了しました")

        finally:
            self.switch_to_default()

    def select_assignee_from_list(self):
        """一覧から担当者の行を選択して詳細画面に遷移"""
        log(f"一覧から「{self.config['assignee_name']}」のボタンを選択...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            assignee_buttons = self.driver.find_elements(
                By.XPATH, f"//input[@type='button' and @value='{self.config['assignee_name']}']"
            )
            log(f"「{self.config['assignee_name']}」のボタン数: {len(assignee_buttons)}")

            if assignee_buttons:
                assignee_buttons[0].click()
                log(f"一番上の{self.config['assignee_name']}ボタンをクリックしました")
            else:
                raise Exception(f"{self.config['assignee_name']}ボタンが見つかりませんでした")

            time.sleep(3)
            log("詳細画面に遷移しました")

        finally:
            self.switch_to_default()

    def go_to_download_options(self):
        """詳細画面から出力オプション画面に遷移"""
        log("詳細画面から出力オプション画面に遷移中...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            download_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='ダウンロード']"))
            )
            download_btn.click()
            log("詳細画面のダウンロードボタンをクリックしました")

            time.sleep(3)
            log("出力オプション画面に遷移しました")

        finally:
            self.switch_to_default()

    def download_list_from_options(self):
        """出力オプション画面で設定してダウンロード"""
        log("出力オプション画面で設定とダウンロード中...")

        if not self.switch_to_iframe():
            raise Exception("iframeに切り替えできませんでした")

        try:
            isbn_setting = self.config['isbn_setting']
            isbn_names = {
                'lst': 'リスト巻数',
                'dat': '提出日時点',
                'max': '最終巻',
                '1st': '第1巻'
            }

            log(f"ISBN「{isbn_names.get(isbn_setting, isbn_setting)}」を選択中...")

            try:
                isbn_radio = self.driver.find_element(
                    By.XPATH, f"//input[@type='radio' and @name='isb' and @value='{isbn_setting}']"
                )
                isbn_radio.click()
                log(f"「{isbn_names.get(isbn_setting, isbn_setting)}」のラジオボタン選択完了")
            except NoSuchElementException:
                log("ラジオボタンが見つかりませんでした。デフォルトのまま進行します。")

            time.sleep(1)

            download_btn = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='ダウンロード']"))
            )
            download_btn.click()
            log("ダウンロード処理が開始されました")

            time.sleep(5)
            log("ダウンロード処理完了")

        finally:
            self.switch_to_default()

    def run_automation(self):
        """メイン自動化処理"""
        try:
            log("=== コミックリスター自動化開始 ===")

            self.setup_browser()
            self.navigate_to_site()
            self.click_comic_lister_button()
            self.upload_csv_file()

            if not self.click_work_button_and_handle_assignee():
                raise Exception("作業開始または担当者入力に失敗しました")

            self.click_initial_survey_1()
            self.click_list_creation_complete()
            self.select_assignee_from_list()
            self.go_to_download_options()
            self.download_list_from_options()

            log("=== 自動化処理完了 ===")

        except Exception as e:
            log(f"エラー発生: {e}")
            raise
        finally:
            self.cleanup()

    def cleanup(self):
        """リソースクリーンアップ"""
        try:
            if self.driver:
                self.driver.quit()
            log("リソースのクリーンアップが完了しました")
        except Exception as e:
            log(f"クリーンアップエラー: {e}")


def main():
    log("=" * 50)
    log("Comic Lister CLI Started")
    log("=" * 50)

    # 環境変数チェック
    log(f"HARU_USERNAME set: {bool(HARU_USERNAME)}")
    log(f"HARU_PASSWORD set: {bool(HARU_PASSWORD)}")
    log(f"GOOGLE_SHEETS_URL set: {bool(GOOGLE_SHEETS_URL)}")
    log(f"ASSIGNEE_NAME: {ASSIGNEE_NAME}")
    log(f"ISBN_SETTING: {ISBN_SETTING}")
    log(f"OUTPUT_DIR: {OUTPUT_DIR}")

    # 出力ディレクトリ作成
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Google SheetsからコミックNo.を取得
    comic_numbers = get_comic_numbers_from_sheets()

    if not comic_numbers:
        log("コミックNo.が取得できませんでした。終了します。")
        sys.exit(1)

    # 2. list.csvを作成
    csv_path = create_list_csv(comic_numbers)

    # 3. 設定を準備
    today = datetime.datetime.now().strftime("%Y/%m/%d")
    config = {
        'csv_path': csv_path,
        'list_name': f"不足画像リスト_{datetime.datetime.now().strftime('%Y%m%d')}",
        'assignee_name': ASSIGNEE_NAME,
        'submission_date': today,
        'isbn_setting': ISBN_SETTING
    }

    log(f"設定: {config}")

    # 4. 自動化処理を実行
    automator = ComicListerAutomator(config)
    automator.run_automation()

    # 5. 出力ファイルを確認
    log("出力ファイル一覧:")
    for file in OUTPUT_DIR.iterdir():
        log(f"  - {file.name}")

    log("=" * 50)
    log("Comic Lister CLI Completed")
    log("=" * 50)


if __name__ == "__main__":
    main()
