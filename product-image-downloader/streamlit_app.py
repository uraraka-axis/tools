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
MIN_FILE_SIZE = 2 * 1024  # 2KB（これ以下はスキップ）
MAX_FOLDER_SIZE = 23 * 1024 * 1024  # 23MB（フォルダ上限）
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


# ===== フォルダ振り分け管理 =====
class FolderManager:
    """23MB以下でフォルダを自動振り分け"""

    def __init__(self, output_base: Path):
        self.output_base = output_base
        self.folder_index = 1
        self.folder_size = 0
        self.folder_names = []  # 各フォルダのパスを記録

    def get_folder(self, files_size: int, base_fname: str) -> Path:
        """現在のフォルダを返す。容量超過なら次のフォルダへ。
        画像セットは同一フォルダに格納する。
        フォルダ名は最初に格納される商品のbase_fnameを使用。"""
        if self.folder_size > 0 and self.folder_size + files_size > MAX_FOLDER_SIZE:
            self.folder_index += 1
            self.folder_size = 0

        if self.folder_index > len(self.folder_names):
            # 新しいフォルダ: 最初の商品名で命名
            folder_name = base_fname.lower()
            folder_path = self.output_base / folder_name
            folder_path.mkdir(parents=True, exist_ok=True)
            self.folder_names.append(folder_path)
        else:
            folder_path = self.folder_names[self.folder_index - 1]

        return folder_path

    def add_size(self, size: int):
        self.folder_size += size


# ===== ダウンロード＆フィルタ =====
def download_and_filter_images(session, images, base_fname, folder_mgr, main_only):
    """画像をダウンロードし、2KB以下をスキップ、フォルダに振り分け保存"""
    valid_images = []  # (fname, content) のリスト
    for idx, url in enumerate(images):
        if idx > 0 and main_only:
            break
        try:
            resp = session.get(url, timeout=30, headers=HEADERS)
            if resp.status_code == 200:
                content = resp.content
                if len(content) > MIN_FILE_SIZE:
                    suffix = "" if len(valid_images) == 0 else f"_{len(valid_images)}"
                    fname = f"{base_fname}{suffix}.jpg".lower()
                    valid_images.append((fname, content))
        except Exception:
            pass

    if not valid_images:
        return 0

    # セット全体のサイズを計算
    total_size = sum(len(c) for _, c in valid_images)

    # フォルダを取得（23MB制限で自動振り分け）
    folder = folder_mgr.get_folder(total_size, base_fname)

    saved_count = 0
    saved_size = 0
    for fname, content in valid_images:
        try:
            with open(folder / fname, 'wb') as f:
                f.write(content)
            saved_count += 1
            saved_size += len(content)
        except Exception:
            pass

    folder_mgr.add_size(saved_size)
    return saved_count


# ===== ZIP 作成 =====
def create_zip_files(output_base: Path, folder_mgr) -> Path:
    """各フォルダの中身をフラットなZIPに圧縮。
    ZIP名 = フォルダ内の最初のファイル名（拡張子なし）.zip"""
    zip_dir = output_base / "圧縮ファイル"
    zip_dir.mkdir(parents=True, exist_ok=True)

    for folder_path in folder_mgr.folder_names:
        if not folder_path.exists():
            continue

        files = sorted([f for f in folder_path.iterdir() if f.is_file()], key=lambda x: x.name)
        if not files:
            continue

        zip_filename = f"{files[0].stem}.zip"
        zip_path = zip_dir / zip_filename

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_path in files:
                zf.write(file_path, file_path.name)  # フラット格納

    return zip_dir


def create_final_zip(zip_dir: Path, excel_path: Path) -> BytesIO:
    """個別ZIPファイル + 更新済みExcelを1つのZIPにまとめる"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        # 個別ZIPを格納
        for f in sorted(zip_dir.iterdir()):
            if f.is_file() and f.suffix == '.zip':
                zf.write(f, f.name)
        # 更新済みExcelを格納
        zf.write(excel_path, excel_path.name)
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

        # 画像保存用ディレクトリ
        img_dir = tmp_dir / "images"
        img_dir.mkdir(exist_ok=True)

        rows = [r for r in range(2, ws.max_row + 1) if ws[f'C{r}'].value or ws[f'D{r}'].value]
        total = len(rows)

        if total == 0:
            st.warning("処理対象の行が見つかりませんでした（C列またはD列にデータが必要です）")
            return None

        # フォルダ振り分け管理
        folder_mgr = FolderManager(img_dir)

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
                            session, images, base_fname, folder_mgr, main_only
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

            # ZIP圧縮
            status.update(label="ZIP圧縮中...")
            zip_dir = create_zip_files(img_dir, folder_mgr)

            status.update(label="✅ 処理完了", state="complete", expanded=False)

        # Excel 保存
        wb.save(excel_path)

        # 最終ZIP作成（個別ZIP + Excel）
        zip_buffer = create_final_zip(zip_dir, excel_path)

        stats['folder_count'] = len(folder_mgr.folder_names)

        return {
            'zip': zip_buffer,
            'stats': stats,
            'logs': logs,
            'filename': Path(uploaded_file.name).stem
        }


# ===== ひな形Excel生成 =====
@st.cache_data
def create_template_excel() -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "商品リスト"

    FONT = "メイリオ"
    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0'),
    )

    # 列定義: (ヘッダー名, 必須レベル, 幅)
    # 必須レベル: "required"=必須, "optional"=任意, "auto"=自動, "unused"=未使用
    columns = {
        'A': ('', 'unused', 4),
        'B': ('連番', 'required', 8),
        'C': ('JANコード ★', 'required', 18),
        'D': ('ASIN ★', 'required', 16),
        'E': ('品名', 'optional', 30),
        'F': ('ディレクトリ1', 'optional', 14),
        'G': ('ディレクトリ2', 'optional', 14),
        'H': ('ディレクトリ3', 'optional', 14),
        'I': ('ディレクトリ4', 'optional', 14),
        'J': ('DL枚数', 'auto', 10),
        'K': ('棚番', 'required', 8),
        'L': ('', 'unused', 4),
        'M': ('拠点コード', 'required', 12),
        'N': ('ファイル名（自動）', 'auto', 28),
    }

    # ヘッダー色の定義
    fills = {
        'required': PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),
        'optional': PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid"),
        'auto':     PatternFill(start_color="A9D18E", end_color="A9D18E", fill_type="solid"),
        'unused':   PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
    }
    font_colors = {
        'required': Font(name=FONT, bold=True, size=10, color="FFFFFF"),
        'optional': Font(name=FONT, bold=True, size=10, color="1F4E79"),
        'auto':     Font(name=FONT, bold=True, size=10, color="375623"),
        'unused':   Font(name=FONT, bold=True, size=10, color="808080"),
    }

    # 1行目: ヘッダー
    for col, (title, level, width) in columns.items():
        cell = ws[f'{col}1']
        cell.value = title
        cell.font = font_colors[level]
        cell.fill = fills[level]
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        ws.column_dimensions[col].width = width

    # 2行目: 補足説明
    notes = {
        'A': '▼ 3行目から入力',
        'B': '※必須', 'C': '※C or D 必須', 'D': '※C or D 必須',
        'E': '任意', 'F': '任意', 'G': '任意', 'H': '任意', 'I': '任意',
        'J': '※自動記入', 'K': '※必須', 'M': '※必須', 'N': '※自動表示',
    }
    note_font = Font(name=FONT, size=8, color="808080")
    note_highlight = Font(name=FONT, size=8, bold=True, color="FF0000")
    for col, note in notes.items():
        cell = ws[f'{col}2']
        cell.value = note
        cell.font = note_highlight if col == 'A' else note_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # 3行目: サンプルデータ
    sample = {
        'B': 1, 'C': '4902370549560', 'D': 'B0CHY3GYW4',
        'E': 'サンプル商品名', 'F': 'ゲーム', 'G': 'Switch',
        'K': '999', 'M': 'TK',
    }
    data_font = Font(name=FONT, size=10)
    sample_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    for col_letter in columns:
        cell = ws[f'{col_letter}3']
        cell.value = sample.get(col_letter)
        cell.font = data_font
        cell.fill = sample_fill
        cell.border = thin_border

    # N列: ファイル名プレビュー関数（3行目〜102行目）
    formula_font = Font(name=FONT, size=10, color="808080")
    for row in range(3, 103):
        cell = ws[f'N{row}']
        cell.value = f'=IF(B{row}="","",LOWER(K{row}&"-"&M{row}&"-"&TEXT(B{row},"000000")&".jpg"))'
        cell.font = formula_font
        if row == 3:
            cell.fill = sample_fill
        cell.border = thin_border

    # 行の高さ
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 20

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ===== UI =====
st.title("📦 商品画像一括ダウンロード")
st.caption("Excelリストに基づき、ネット上の商品画像を自動収集・整理します")

with st.container(border=True):
    col_upload, col_template = st.columns([3, 1])
    with col_upload:
        uploaded_file = st.file_uploader("Excelファイル", type=["xlsx"])
    with col_template:
        st.write("")  # スペーサー
        st.write("")
        st.download_button(
            label="📄 ひな形ダウンロード",
            data=create_template_excel(),
            file_name="商品画像DL_テンプレート.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    mode = st.radio("取得モード", ["全画像を取得", "メインのみ"], horizontal=True)
    st.info(
        "**命名規則**: 棚番-拠点コード-連番.jpg（小文字）  \n"
        "**フォルダ振り分け**: 各フォルダ23MB以下で自動分割  \n"
        "**ZIP形式**: フォルダごとにフラットなZIP（ZIP名＝先頭ファイル名）  \n"
        "**フィルタ**: 2KB以下の画像は自動スキップ"
    )

if uploaded_file:
    if st.button("▶ 実行開始", type="primary", use_container_width=True):
        result = process(uploaded_file, mode == "メインのみ")

        if result:
            st.divider()
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("合計", result['stats']['total'])
            col2.metric("成功", result['stats']['success'])
            col3.metric("未取得", result['stats']['not_found'])
            col4.metric("ZIPファイル数", result['stats']['folder_count'])

            st.download_button(
                label="📥 結果をダウンロード（ZIP）",
                data=result['zip'],
                file_name=f"{result['filename']}_images.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
