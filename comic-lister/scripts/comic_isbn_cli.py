"""
コミックISBN検索自動化ツール（CLI版）
GitHub Actionsで実行され、missing_comics.csvからis_list.csvを生成してGitHubにアップロードします。
"""

import time
import os
import sys
import datetime
import base64
from pathlib import Path
import urllib.parse

import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select

# 環境変数から設定を取得
HARU_USERNAME = os.environ.get("HARU_USERNAME", "haruuser")
HARU_PASSWORD = os.environ.get("HARU_PASSWORD", "Haru@9999")

# GitHub設定
GITHUB_REPO = "uraraka-axis/tools"
GITHUB_INPUT_PATH = "comic-lister/data/missing_comics.csv"
GITHUB_OUTPUT_PATH = "comic-lister/data/is_list.csv"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")

# 出力ディレクトリ
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "/tmp/comic-lister-output"))


def log(message):
    """ログメッセージを出力"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def upload_to_github(file_path: str, github_path: str, message: str) -> bool:
    """ファイルをGitHubにアップロード（上書き更新）"""
    if not GITHUB_TOKEN:
        log("GITHUB_TOKEN未設定のためアップロードをスキップ")
        return False

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    # ファイル内容を読み込み
    with open(file_path, 'rb') as f:
        content = f.read()

    # 既存ファイルのSHAを取得（更新時に必要）
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{github_path}"
    sha = None

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            sha = response.json().get("sha")
    except:
        pass

    # ファイルをアップロード
    data = {
        "message": message,
        "content": base64.b64encode(content).decode('utf-8'),
        "branch": "master"
    }
    if sha:
        data["sha"] = sha

    try:
        response = requests.put(url, headers=headers, json=data)
        if response.status_code in [200, 201]:
            log(f"GitHubアップロード成功: {github_path}")
            return True
        else:
            log(f"GitHubアップロード失敗: HTTP {response.status_code}")
            return False
    except Exception as e:
        log(f"GitHubアップロードエラー: {e}")
        return False


def get_comic_numbers_from_github():
    """GitHubからmissing_comics.csvを取得してコミックNo.を抽出"""
    try:
        log("GitHubからデータ取得中...")

        raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/master/{GITHUB_INPUT_PATH}"

        headers = {}
        if GITHUB_TOKEN:
            headers["Authorization"] = f"token {GITHUB_TOKEN}"

        response = requests.get(raw_url, headers=headers)

        if response.status_code != 200:
            log(f"GitHub取得エラー: HTTP {response.status_code}")
            return []

        # CSVをパース（J列=9列目にコミックNo.が入っている）
        df = pd.read_csv(pd.io.common.StringIO(response.text), header=None)

        comic_numbers = []
        for i in range(len(df)):
            if len(df.columns) > 9:
                value = df.iloc[i, 9]  # J列（0始まりで9）
                if pd.notna(value) and str(value).strip():
                    comic_numbers.append(str(value).strip())

        log(f"取得したコミックNo.: {len(comic_numbers)}件")
        return comic_numbers

    except Exception as e:
        log(f"GitHubからのデータ取得エラー: {e}")
        return []


class ComicISBNSearchAutomation:
    def __init__(self, download_folder=None):
        """コミックISBN検索自動化クラス（ヘッドレス版）"""
        self.username = HARU_USERNAME
        self.password = HARU_PASSWORD

        if download_folder is None:
            self.download_folder = OUTPUT_DIR
        else:
            self.download_folder = Path(download_folder)

        self.driver = None

    def setup_driver(self):
        """Seleniumドライバーの設定（ヘッドレス版）"""
        chrome_options = Options()

        # ダウンロード設定
        prefs = {
            "download.default_directory": str(self.download_folder),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        # ヘッドレスモード設定
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(10)
        log("Chromeドライバー起動成功（ヘッドレスモード）")
        return self.driver

    def wait_for_download_complete(self, timeout=120):
        """ダウンロード完了を待機"""
        log(f"CSVダウンロード完了を待機中（最大{timeout}秒）...")
        start_time = time.time()
        initial_files = set(self.download_folder.glob("*"))

        while time.time() - start_time < timeout:
            current_files = set(self.download_folder.glob("*"))
            new_files = current_files - initial_files

            for file_path in new_files:
                if "is_list" in file_path.name.lower() and file_path.suffix.lower() == ".csv":
                    if file_path.stat().st_size > 0:
                        log(f"CSVダウンロード完了: {file_path}")
                        return file_path

            time.sleep(2)

        log("CSVダウンロードがタイムアウトしました")
        return None

    def search_and_download_csv(self, comic_numbers):
        """コミックサーチで検索してCSVをダウンロード"""
        if not comic_numbers:
            log("検索するコミックNo.がありません")
            return None

        try:
            if self.driver is None:
                self.setup_driver()

            # ISBN検索機能に直接アクセス
            isbn_url = f"https://{self.username}:{urllib.parse.quote(self.password)}@haru-u.biz/comic/final/cn_search_isbn.asp"
            log("ISBN検索ページにアクセス中...")
            self.driver.get(isbn_url)
            time.sleep(3)

            log(f"現在のURL: {self.driver.current_url}")

            # textareaにコミックNo.を入力
            log("コミックNo.を入力中...")
            textarea = self.driver.find_element(By.CSS_SELECTOR, "textarea[name='isbn']")
            textarea.clear()
            textarea.send_keys('\n'.join(comic_numbers))
            log(f"コミックNo. {len(comic_numbers)}件を入力しました")

            # 検索オプション設定
            log("検索オプションを設定中...")

            # 快活並びにチェック
            kaikatsu_checkbox = self.driver.find_element(By.CSS_SELECTOR, "input[name='c_va'][type='checkbox']")
            if not kaikatsu_checkbox.is_selected():
                kaikatsu_checkbox.click()
            log("快活並び (c_va) にチェックしました")

            # シリーズにチェック
            series_checkbox = self.driver.find_element(By.CSS_SELECTOR, "input[name='c_se'][type='checkbox']")
            if not series_checkbox.is_selected():
                series_checkbox.click()
            log("シリーズ (c_se) にチェックしました")

            # 展開リスト選択
            disp_select = self.driver.find_element(By.CSS_SELECTOR, "select[name='disp']")
            select_obj = Select(disp_select)
            select_obj.select_by_value("tenk")
            log("展開リスト (tenk) を選択しました")

            # 検索実行
            log("検索実行中...")
            search_btn = self.driver.find_element(By.CSS_SELECTOR, "input[type='button'][value='検索']")
            search_btn.click()
            log("検索ボタンをクリックしました")

            # 検索結果ページの読み込み待機
            time.sleep(10)
            log(f"検索後のURL: {self.driver.current_url}")

            # 結果ファイルのダウンロードリンクを探してクリック
            log("結果ファイルのダウンロードリンクを探索中...")
            try:
                download_link = self.driver.find_element(By.XPATH, "//a[contains(@href, 'cn_search_dlf.asp') and contains(text(), '結果ファイル')]")
                download_url = download_link.get_attribute("href")
                log(f"結果ファイルリンクを発見: {download_url}")

                download_link.click()
                log("結果ファイルリンクをクリックしました")

            except Exception as e:
                log(f"結果ファイルリンクが見つかりません: {e}")

                # フォールバック: cn_search_dlf.aspを含むリンクを探す
                try:
                    all_links = self.driver.find_elements(By.TAG_NAME, "a")
                    for link in all_links:
                        href = link.get_attribute("href") or ""
                        if "cn_search_dlf.asp" in href and "is_list.csv" in href:
                            log(f"フォールバックでダウンロードリンクを発見: {href}")
                            link.click()
                            log("ダウンロードリンクをクリックしました")
                            break
                    else:
                        log("ダウンロードリンクが見つかりませんでした")
                        return None

                except Exception as e2:
                    log(f"フォールバック処理でもエラー: {e2}")
                    return None

            # CSVダウンロード完了を待機
            downloaded_file = self.wait_for_download_complete(timeout=120)

            if downloaded_file:
                log(f"ダウンロード成功: {downloaded_file}")
                return downloaded_file
            else:
                log("CSVダウンロードに失敗しました")
                return None

        except Exception as e:
            log(f"検索・ダウンロードエラー: {e}")
            import traceback
            log(f"詳細エラー: {traceback.format_exc()}")
            return None

    def cleanup(self):
        """リソースのクリーンアップ"""
        if self.driver:
            self.driver.quit()
            self.driver = None
            log("ドライバーをクリーンアップしました")


def main():
    log("=" * 50)
    log("Comic ISBN Search CLI Started")
    log("=" * 50)

    # 環境変数チェック
    log(f"HARU_USERNAME set: {bool(HARU_USERNAME)}")
    log(f"HARU_PASSWORD set: {bool(HARU_PASSWORD)}")
    log(f"GITHUB_TOKEN set: {bool(GITHUB_TOKEN)}")
    log(f"GITHUB_INPUT_PATH: {GITHUB_INPUT_PATH}")
    log(f"GITHUB_OUTPUT_PATH: {GITHUB_OUTPUT_PATH}")
    log(f"OUTPUT_DIR: {OUTPUT_DIR}")

    # 出力ディレクトリ作成
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 1. GitHubからコミックNo.を取得
    comic_numbers = get_comic_numbers_from_github()

    if not comic_numbers:
        log("コミックNo.が取得できませんでした。終了します。")
        sys.exit(1)

    # 2. ISBN検索を実行してis_list.csvをダウンロード
    automation = ComicISBNSearchAutomation(download_folder=OUTPUT_DIR)

    try:
        downloaded_file = automation.search_and_download_csv(comic_numbers)

        if downloaded_file:
            # 3. is_list.csvをGitHubにアップロード
            log(f"GitHubにアップロード: {downloaded_file.name}")
            today = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            upload_to_github(
                str(downloaded_file),
                GITHUB_OUTPUT_PATH,
                f"Update is_list.csv ({len(comic_numbers)}件) - {today}"
            )
        else:
            log("is_list.csvの生成に失敗しました")
            sys.exit(1)

    finally:
        automation.cleanup()

    # 出力ファイル一覧
    log("出力ファイル一覧:")
    for file in OUTPUT_DIR.iterdir():
        log(f"  - {file.name}")

    log("=" * 50)
    log("Comic ISBN Search CLI Completed")
    log("=" * 50)


if __name__ == "__main__":
    main()
