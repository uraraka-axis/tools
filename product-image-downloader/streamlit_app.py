#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import streamlit as st
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import time
from datetime import datetime
import random
import re
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO

# ===== ページ設定 =====
st.set_page_config(page_title="商品画像ダウンローダー", page_icon="📦", layout="wide")


# ===== パスワード認証 =====
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    password_input = st.text_input("パスワードを入力してください", type="password")
    if password_input:
        if password_input == st.secrets.get("password", ""):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません")
    return False


if not check_password():
    st.stop()


# ===== 定数 =====
SURUGAYA_SEARCH_URL = "https://www.suruga-ya.jp/kaitori/search_buy"
BOOKOFF_BASE_URL = "https://shopping.bookoff.co.jp/search/keyword/"
NO_IMAGE_PATTERNS = ['item_ll.gif', 'no_image', 'noimage', 'no-image', 'now_printing']
MIN_FILE_SIZE = 2 * 1024  # 2KB
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}


# ===== Selenium セットアップ =====
def setup_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--log-level=3')
    options.add_argument('--disable-blink-features=AutomationControlled')

    # Streamlit Cloud (Debian) のシステム Chromium を検索
    for chrome_path in ['/usr/bin/chromium-browser', '/usr/bin/chromium', '/usr/bin/google-chrome']:
        if os.path.exists(chrome_path):
            options.binary_location = chrome_path
            break

    driver = None
    for driver_path in ['/usr/bin/chromedriver', '/usr/lib/chromium-browser/chromedriver',
                        '/usr/lib/chromium/chromedriver']:
        if os.path.exists(driver_path):
            service = Service(driver_path)
            driver = webdriver.Chrome(service=service, options=options)
            break

    if driver is None:
        # ローカル環境: webdriver-manager でフォールバック
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

    # 駿河屋セーフサーチ設定
    driver.get("https://www.suruga-ya.jp/")
    driver.add_cookie({'name': 'safe_search_option', 'value': '3', 'domain': '.suruga-ya.jp'})
    return driver


# ===== ヘルパー関数 =====
def is_no_image(url: str) -> bool:
    url_lower = url.lower()
    return any(p in url_lower for p in NO_IMAGE_PATTERNS)


def sanitize(name) -> str | None:
    if not name or str(name).strip() == "":
        return None
    return re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip('. ')


def build_genre_path(h_dir: Path, genres: list) -> Path:
    valid_genres = [g for g in genres if g is not None]
    if not valid_genres:
        return h_dir / "未分類"
    target_path = h_dir
    for genre in valid_genres:
        target_path = target_path / genre
    return target_path


# ===== 画像取得関数 =====
def get_amazon_images(driver, asin: str, main_only: bool) -> list:
    url = f"https://www.amazon.co.jp/dp/{asin}"
    try:
        driver.get(url)
        time.sleep(random.uniform(2, 3))
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        images = []

        main_img = soup.find('img', {'id': 'landingImage'})
        if main_img:
            src = main_img.get('data-old-hires') or main_img.get('src')
            if src:
                src = re.sub(r'_AC_[A-Z]{2}\d+_', '_AC_SL1500_', src)
                images.append(src)

        if not main_only:
            alt_div = soup.find('div', {'id': 'altImages'})
            if alt_div:
                for thumb in alt_div.find_all('img'):
                    t_src = thumb.get('src')
                    if t_src and 'video' not in t_src.lower():
                        h_res = re.sub(r'_AC_[A-Z]{2}\d+,?\d*_', '_AC_SL1500_', t_src)
                        h_res = re.sub(r'\._[A-Z]{2}\d+,?\d*_\.', '._SL1500_.', h_res)
                        if h_res not in images and not is_no_image(h_res):
                            images.append(h_res)
        return images
    except Exception:
        return []


def get_surugaya_images(driver, jan: str) -> list:
    url = f"{SURUGAYA_SEARCH_URL}?search_word={jan}&key_flag=1"
    try:
        driver.get(url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        title_a = soup.select_one('div.title a')
        if not title_a:
            return []

        detail_url = title_a['href']
        if detail_url.startswith('/'):
            detail_url = "https://www.suruga-ya.jp" + detail_url

        driver.get(detail_url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        img_up = soup.find('div', {'id': 'imgUp'})
        if img_up and img_up.find('a'):
            img_url = img_up.find('a')['href']
            if img_url.startswith('/'):
                img_url = "https://www.suruga-ya.jp" + img_url
            return [img_url]
    except Exception:
        pass
    return []


def get_bookoff_images(driver, jan: str) -> list:
    url = f"{BOOKOFF_BASE_URL}{jan}"
    try:
        driver.get(url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        img_tag = soup.select_one('img.js-gridImg, .productItem__image img')
        if img_tag and img_tag.get('src'):
            img_url = img_tag['src'].replace('/SS/', '/LL/').replace('SS.jpg', 'LL.jpg')
            return [img_url]
    except Exception:
        pass
    return []


# ===== ダウンロード＆フィルタ =====
def download_and_filter_images(session, images, base_fname, h_dir, f_dir, genres, main_only):
    valid_images = []
    for idx, url in enumerate(images):
        if idx > 0 and main_only:
            break
        try:
            resp = session.get(url, timeout=30, headers=HEADERS)
            if resp.status_code == 200:
                content = resp.content
                if len(content) > MIN_FILE_SIZE:
                    valid_images.append(content)
        except Exception:
            pass

    if not valid_images:
        return 0

    target_path = build_genre_path(h_dir, genres)
    target_path.mkdir(parents=True, exist_ok=True)

    saved_count = 0
    for idx, content in enumerate(valid_images):
        suffix = "" if idx == 0 else f"_{idx}"
        fname = f"{base_fname}{suffix}.jpg"
        save_h = target_path / fname
        save_f = f_dir / fname
        try:
            with open(save_h, 'wb') as f:
                f.write(content)
            shutil.copy2(save_h, save_f)
            saved_count += 1
        except Exception:
            pass
    return saved_count


# ===== ZIP 作成 =====
def create_zip(base_dir: Path) -> BytesIO:
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(base_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(base_dir)
                zf.write(file_path, arcname)
    zip_buffer.seek(0)
    return zip_buffer


# ===== メイン処理 =====
def process(uploaded_file, main_only):
    session = requests.Session()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir = Path(tmp_dir)
        excel_path = tmp_dir / uploaded_file.name
        with open(excel_path, 'wb') as f:
            f.write(uploaded_file.getvalue())

        wb = load_workbook(excel_path)
        ws = wb.active

        h_dir = tmp_dir / "1_階層管理"
        f_dir = tmp_dir / "2_一括管理"
        h_dir.mkdir(exist_ok=True)
        f_dir.mkdir(exist_ok=True)

        rows = [r for r in range(2, ws.max_row + 1) if ws[f'C{r}'].value or ws[f'D{r}'].value]
        total = len(rows)

        if total == 0:
            st.warning("処理対象の行が見つかりませんでした（C列またはD列にデータが必要です）")
            return None

        # ブラウザ起動
        with st.status("処理中...", expanded=True) as status:
            status.update(label="ブラウザを起動中...")
            try:
                driver = setup_driver()
            except Exception as e:
                st.error(f"ブラウザ起動に失敗しました: {e}")
                return None

            progress_bar = st.progress(0)
            status_text = st.empty()
            log_area = st.empty()

            stats = {'total': total, 'success': 0, 'not_found': 0}
            logs = []

            try:
                for i, r in enumerate(rows, 1):
                    progress_bar.progress(i / total)
                    status.update(label=f"処理中: {i} / {total} ({i / total * 100:.1f}%)")
                    status_text.text(f"処理中: {i} / {total}")

                    seq = str(ws[f'B{r}'].value or "0").zfill(6)
                    jan = str(ws[f'C{r}'].value or "").strip()
                    asin = str(ws[f'D{r}'].value or "").strip()
                    genres = [sanitize(ws[f'{col}{r}'].value) for col in ['F', 'G', 'H', 'I']]
                    shelf = sanitize(ws[f'K{r}'].value) or "00"
                    base_code = sanitize(ws[f'M{r}'].value) or "XX"
                    base_fname = f"{shelf}-{base_code}-{seq}"

                    product_name = str(ws[f'E{r}'].value or "").strip()
                    if len(product_name) > 30:
                        product_name = product_name[:30] + "..."

                    images = []
                    source_site = ""

                    # Amazon → 駿河屋 → ブックオフ
                    if asin and asin not in ["-", ""]:
                        images = get_amazon_images(driver, asin, main_only)
                        if images:
                            source_site = "Amazon"

                    if not images and jan:
                        images = get_surugaya_images(driver, jan)
                        if images:
                            source_site = "駿河屋"

                    if not images and jan:
                        images = get_bookoff_images(driver, jan)
                        if images:
                            source_site = "ブックオフ"

                    timestamp = datetime.now().strftime("%H:%M:%S")

                    if images:
                        downloaded_count = download_and_filter_images(
                            session, images, base_fname, h_dir, f_dir, genres, main_only
                        )
                        if downloaded_count > 0:
                            ws[f'J{r}'].value = downloaded_count
                            stats['success'] += 1
                            log_msg = f"[{timestamp}] [{i}/{total}] ✅ {source_site} / {product_name} / 画像{downloaded_count}枚"
                        else:
                            stats['not_found'] += 1
                            log_msg = f"[{timestamp}] [{i}/{total}] ⚠️ {source_site} / {product_name} / 有効な画像なし"
                    else:
                        stats['not_found'] += 1
                        log_msg = f"[{timestamp}] [{i}/{total}] ❌ 取得失敗 / {product_name}"

                    logs.append(log_msg)
                    log_area.code("\n".join(logs[-50:]))  # 直近50件表示

                    time.sleep(random.uniform(0.5, 1.0))

            finally:
                driver.quit()

            status.update(label="✅ 処理完了", state="complete", expanded=False)

        # Excel 保存
        wb.save(excel_path)

        # ZIP 作成（メモリ上に保持）
        zip_buffer = create_zip(tmp_dir)

        return {
            'zip': zip_buffer,
            'stats': stats,
            'logs': logs,
            'filename': Path(uploaded_file.name).stem
        }


# ===== UI =====
st.title("📦 商品画像一括ダウンロード")
st.caption("Excelリストに基づき、ネット上の商品画像を自動収集・整理します")

with st.container(border=True):
    uploaded_file = st.file_uploader("Excelファイル", type=["xlsx"])
    mode = st.radio("取得モード", ["全画像を取得", "メインのみ"], horizontal=True)
    st.info(
        "**保存仕様**: 1\_階層管理（ジャンル別）／ 2\_一括管理（全集約）  \n"
        "**命名規則**: 棚番-拠点コード-連番.jpg  \n"
        "**フィルタ**: 2KB以下の画像は自動スキップ"
    )

if uploaded_file:
    if st.button("▶ 実行開始", type="primary", use_container_width=True):
        result = process(uploaded_file, mode == "メインのみ")

        if result:
            st.divider()
            col1, col2, col3 = st.columns(3)
            col1.metric("合計", result['stats']['total'])
            col2.metric("成功", result['stats']['success'])
            col3.metric("未取得", result['stats']['not_found'])

            st.download_button(
                label="📥 結果をダウンロード（ZIP）",
                data=result['zip'],
                file_name=f"{result['filename']}_images.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
