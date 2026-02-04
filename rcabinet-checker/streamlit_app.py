"""
R-Cabinet 管理ツール
- フォルダ画像一覧：R-Cabinetのフォルダ毎に画像を一覧表示
- 画像存在チェック：コミックNoを入力して存在確認
"""

# バージョン（デプロイ確認用）
APP_VERSION = "2.1.0"

import streamlit as st
import requests
import base64
import xml.etree.ElementTree as ET
import pandas as pd
import time
from io import BytesIO
from datetime import datetime

# 重いライブラリは遅延読み込み（起動高速化）
_bs4_module = None
_openpyxl_styles = None
_openpyxl_utils = None
_supabase_module = None
_zipfile_module = None
_random_module = None

# Gemini AI（遅延読み込み - 起動高速化のため）
GEMINI_AVAILABLE = None
_genai_module = None


def get_bs4():
    """BeautifulSoupを遅延読み込み"""
    global _bs4_module
    if _bs4_module is None:
        from bs4 import BeautifulSoup
        _bs4_module = BeautifulSoup
    return _bs4_module


def get_openpyxl_styles():
    """openpyxlスタイルを遅延読み込み"""
    global _openpyxl_styles, _openpyxl_utils
    if _openpyxl_styles is None:
        from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        _openpyxl_styles = {'Font': Font, 'Border': Border, 'Side': Side, 'PatternFill': PatternFill, 'Alignment': Alignment}
        _openpyxl_utils = {'get_column_letter': get_column_letter}
    return _openpyxl_styles, _openpyxl_utils


def get_supabase_module():
    """Supabaseを遅延読み込み"""
    global _supabase_module
    if _supabase_module is None:
        from supabase import create_client
        _supabase_module = create_client
    return _supabase_module


def get_zipfile():
    """zipfileを遅延読み込み"""
    global _zipfile_module
    if _zipfile_module is None:
        import zipfile
        _zipfile_module = zipfile
    return _zipfile_module


def get_random():
    """randomを遅延読み込み"""
    global _random_module
    if _random_module is None:
        import random
        _random_module = random
    return _random_module

# ページ設定
st.set_page_config(
    page_title="R-Cabinet 管理ツール",
    page_icon="🖼️",
    layout="wide"
)

# 認証情報（Streamlit Secretsから取得）
APP_PASSWORD = st.secrets.get("password", "")
SERVICE_SECRET = st.secrets.get("RMS_SERVICE_SECRET", "")
LICENSE_KEY = st.secrets.get("RMS_LICENSE_KEY", "")
BASE_URL = "https://api.rms.rakuten.co.jp/es/1.0"

# Supabase接続情報
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")

# GitHub接続情報
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO = "uraraka-axis/tools"
GITHUB_MISSING_CSV_PATH = "comic-lister/data/missing_comics.csv"
GITHUB_IS_LIST_PATH = "comic-lister/data/is_list.csv"
GITHUB_COMIC_LIST_PATH = "comic-lister/data/comic_list.csv"
GITHUB_FOLDER_HIERARCHY_PATH = "comic-lister/data/folder_hierarchy.xlsx"

# Gemini API設定（セルフヒーリング用）
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")


def get_gemini_model():
    """Gemini AIモデルを遅延読み込みで取得"""
    global GEMINI_AVAILABLE, _genai_module

    if GEMINI_AVAILABLE is None:
        try:
            import google.generativeai as genai
            _genai_module = genai
            GEMINI_AVAILABLE = True
        except ImportError:
            GEMINI_AVAILABLE = False
            return None

    if not GEMINI_AVAILABLE or not GEMINI_API_KEY:
        return None

    if _genai_module:
        _genai_module.configure(api_key=GEMINI_API_KEY)
        return _genai_module.GenerativeModel('gemini-2.0-flash')

    return None


def upload_to_github(content: str, path: str, message: str) -> dict:
    """GitHubにファイルをアップロード（上書き更新）"""
    if not GITHUB_TOKEN:
        return {"success": False, "error": "GITHUB_TOKEN未設定"}

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 既存ファイルのSHAを取得（更新時に必要）
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
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
        "content": base64.b64encode(content.encode('utf-8')).decode('utf-8'),
        "branch": "master"
    }
    if sha:
        data["sha"] = sha

    try:
        response = requests.put(url, headers=headers, json=data)
        if response.status_code in [200, 201]:
            return {"success": True, "url": response.json().get("content", {}).get("html_url", "")}
        else:
            return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:200]}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def upload_binary_to_github(content: bytes, path: str, message: str) -> dict:
    """バイナリファイルをGitHubにアップロード（上書き更新）"""
    if not GITHUB_TOKEN:
        return {"success": False, "error": "GITHUB_TOKEN未設定"}

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 既存ファイルのSHAを取得（更新時に必要）
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
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
            return {"success": True, "url": response.json().get("content", {}).get("html_url", "")}
        else:
            return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:200]}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def download_from_github(path: str) -> dict:
    """GitHubからファイルをダウンロード"""
    if not GITHUB_TOKEN:
        return {"success": False, "error": "GITHUB_TOKEN未設定"}

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3.raw"
    }

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return {"success": True, "content": response.content, "path": path}
        elif response.status_code == 404:
            return {"success": False, "error": f"ファイルが見つかりません: {path}"}
        else:
            return {"success": False, "error": f"HTTP {response.status_code}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_github_file_info(path: str) -> dict:
    """GitHubファイルの情報（更新日時など）を取得"""
    if not GITHUB_TOKEN:
        return {}

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    url = f"https://api.github.com/repos/{GITHUB_REPO}/commits?path={path}&per_page=1"

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200 and response.json():
            commit = response.json()[0]
            date_str = commit.get("commit", {}).get("committer", {}).get("date", "")
            if date_str:
                # ISO形式をパースして日本時間に変換（+9時間）
                from datetime import datetime, timedelta, timezone
                dt_utc = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
                dt_jst = dt_utc + timedelta(hours=9)
                return {"last_updated": dt_jst.strftime("%Y-%m-%d %H:%M"), "exists": True}
        return {"exists": False}
    except:
        return {"exists": False}


def trigger_github_actions(workflow_file: str) -> dict:
    """GitHub Actionsワークフローを手動実行"""
    if not GITHUB_TOKEN:
        return {"success": False, "error": "GITHUB_TOKEN未設定"}

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    url = f"https://api.github.com/repos/{GITHUB_REPO}/actions/workflows/{workflow_file}/dispatches"

    try:
        response = requests.post(url, headers=headers, json={"ref": "master"})
        if response.status_code == 204:
            return {"success": True, "message": "ワークフローを開始しました"}
        elif response.status_code == 404:
            return {"success": False, "error": "ワークフローが見つかりません"}
        else:
            return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:200]}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_workflow_runs(workflow_file: str, limit: int = 3) -> list:
    """GitHub Actionsワークフローの実行履歴を取得"""
    if not GITHUB_TOKEN:
        return []

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    url = f"https://api.github.com/repos/{GITHUB_REPO}/actions/workflows/{workflow_file}/runs?per_page={limit}"

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            runs = response.json().get("workflow_runs", [])
            result = []
            for run in runs:
                created = run.get("created_at", "")
                if created:
                    dt = datetime.fromisoformat(created.replace("Z", "+00:00"))
                    created = dt.strftime("%Y-%m-%d %H:%M")
                result.append({
                    "status": run.get("status"),
                    "conclusion": run.get("conclusion"),
                    "created_at": created,
                    "html_url": run.get("html_url")
                })
            return result
        return []
    except:
        return []


@st.cache_resource
def get_supabase_client():
    """Supabaseクライアントを取得（遅延読み込み）"""
    if SUPABASE_URL and SUPABASE_KEY:
        create_client = get_supabase_module()
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    return None


def fetch_all_from_supabase(supabase: Client, table: str, columns: str = "*", filter_col: str = None, filter_val: str = None) -> list:
    """Supabaseから全件取得（ページネーション対応）"""
    all_data = []
    page_size = 1000
    offset = 0

    while True:
        query = supabase.table(table).select(columns).range(offset, offset + page_size - 1)
        if filter_col and filter_val:
            query = query.ilike(filter_col, f"%{filter_val}%")
        response = query.execute()

        if not response.data:
            break

        all_data.extend(response.data)

        if len(response.data) < page_size:
            break

        offset += page_size

    return all_data


def sync_images_to_db(images: list) -> dict:
    """画像一覧をDBに同期（upsert）"""
    supabase = get_supabase_client()
    if not supabase:
        return {"success": False, "error": "Supabase未設定"}

    try:
        # file_nameごとにグループ化（重複検出）
        file_dict = {}
        for img in images:
            file_name = img.get("FileName", "")
            folder_name = img.get("FolderName", "")
            if file_name in file_dict:
                # 重複: folder_namesに追加
                existing_folders = file_dict[file_name]["folder_names"].split(", ")
                if folder_name not in existing_folders:
                    file_dict[file_name]["folder_names"] += f", {folder_name}"
            else:
                file_dict[file_name] = {
                    "file_name": file_name,
                    "folder_names": folder_name,
                    "file_url": img.get("FileUrl", ""),
                    "file_size": img.get("FileSize", 0),
                    "file_timestamp": img.get("TimeStamp", "")
                }

        # 既存データを取得（ページネーション対応）
        existing_data = fetch_all_from_supabase(supabase, "rcabinet_images", "file_name, file_timestamp")
        existing_dict = {row["file_name"]: row["file_timestamp"] for row in existing_data}

        # 差分計算
        new_count = 0
        updated_count = 0
        duplicate_count = 0
        unchanged_count = 0

        records_to_upsert = []
        for file_name, record in file_dict.items():
            # 重複チェック（複数フォルダにある）
            if ", " in record["folder_names"]:
                duplicate_count += 1

            if file_name not in existing_dict:
                new_count += 1
                records_to_upsert.append(record)
            elif existing_dict[file_name] != record["file_timestamp"]:
                updated_count += 1
                records_to_upsert.append(record)
            else:
                unchanged_count += 1

        # 削除済み検出（DBにあるがAPIにない）
        deleted_files = set(existing_dict.keys()) - set(file_dict.keys())
        deleted_count = len(deleted_files)

        # upsert実行（100件ずつ）
        for i in range(0, len(records_to_upsert), 100):
            batch = records_to_upsert[i:i+100]
            supabase.table("rcabinet_images").upsert(
                batch, on_conflict="file_name"
            ).execute()

        # 削除済みファイルをDBから削除
        if deleted_files:
            for file_name in deleted_files:
                supabase.table("rcabinet_images").delete().eq("file_name", file_name).execute()

        return {
            "success": True,
            "new": new_count,
            "updated": updated_count,
            "duplicate": duplicate_count,
            "unchanged": unchanged_count,
            "deleted": deleted_count,
            "total": len(file_dict)
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def load_images_from_db() -> tuple[list, str]:
    """DBから画像一覧を読み込み（ページネーション対応）"""
    supabase = get_supabase_client()
    if not supabase:
        return [], "Supabase未設定"

    try:
        all_data = fetch_all_from_supabase(supabase, "rcabinet_images", "*")
        images = []
        for row in all_data:
            images.append({
                "FolderName": row.get("folder_names", ""),
                "FileName": row.get("file_name", ""),
                "FileUrl": row.get("file_url", ""),
                "FileSize": row.get("file_size", 0),
                "TimeStamp": row.get("file_timestamp", "")
            })
        return images, f"{len(images)}件を読み込みました"
    except Exception as e:
        return [], str(e)


def get_db_stats() -> dict:
    """DBの統計情報を取得（ページネーション対応）"""
    supabase = get_supabase_client()
    if not supabase:
        return {}

    try:
        all_data = fetch_all_from_supabase(supabase, "rcabinet_images", "folder_names, created_at")
        total = len(all_data)
        duplicates = sum(1 for row in all_data if ", " in row.get("folder_names", ""))

        # 最終更新日時を取得
        last_updated = None
        if all_data:
            dates = [row.get("created_at") for row in all_data if row.get("created_at")]
            if dates:
                last_updated = max(dates)[:16].replace("T", " ")  # "2025-02-04 10:30"形式

        return {"total": total, "duplicates": duplicates, "last_updated": last_updated}
    except Exception:
        return {}


def load_images_from_db_by_folder(folder_name: str) -> list:
    """DBから特定フォルダの画像を読み込み（ページネーション対応）"""
    supabase = get_supabase_client()
    if not supabase:
        return []

    try:
        all_data = fetch_all_from_supabase(supabase, "rcabinet_images", "*", "folder_names", folder_name)
        images = []
        for row in all_data:
            images.append({
                "FolderName": row.get("folder_names", ""),
                "FileName": row.get("file_name", ""),
                "FileUrl": row.get("file_url", ""),
                "FileSize": row.get("file_size", 0),
                "TimeStamp": row.get("file_timestamp", "")
            })
        return images
    except Exception:
        return []


def check_password():
    """パスワード認証"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    password_input = st.text_input("パスワードを入力してください", type="password")

    if password_input:
        if password_input == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません")

    return False


# パスワード認証
if not check_password():
    st.stop()


def get_auth_header():
    """ESA認証ヘッダーを生成"""
    credentials = f"{SERVICE_SECRET}:{LICENSE_KEY}"
    encoded = base64.b64encode(credentials.encode()).decode()
    return {"Authorization": f"ESA {encoded}"}


def safe_int(value, default=0):
    """安全にintに変換"""
    try:
        return int(value) if value else default
    except (ValueError, TypeError):
        return default


def style_excel(ws, num_columns=4, url_column=None):
    """Excelワークシートにスタイルを適用"""
    styles, utils = get_openpyxl_styles()
    Font = styles['Font']
    Border = styles['Border']
    Side = styles['Side']
    PatternFill = styles['PatternFill']
    Alignment = styles['Alignment']
    get_column_letter = utils['get_column_letter']

    # フォント設定
    meiryo_font = Font(name='Meiryo UI')
    header_font = Font(name='Meiryo UI', bold=True, color='FFFFFF')
    # 罫線設定
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    # ヘッダー背景色（濃い青）
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')

    # 全セルにフォントと罫線を適用
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=num_columns):
        for cell in row:
            cell.font = meiryo_font
            cell.border = thin_border

    # ヘッダー行のスタイル（1行目）
    for cell in ws[1]:
        if cell.column <= num_columns:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 列幅を自動調整
    for col_idx in range(1, num_columns + 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
        # URL列は固定幅、それ以外は自動調整
        if url_column and col_idx == url_column:
            ws.column_dimensions[column_letter].width = 70
        else:
            ws.column_dimensions[column_letter].width = min(max_length * 1.5 + 2, 40)


def merge_csv_data(is_df, cl_df):
    """IS検索とCL検索の結果をマージ"""
    # comic_list.csvから辞書を作成（N列=CNO, S列=出版社, Y列=シリーズ）
    cl_dict = {}
    for i in range(1, len(cl_df)):
        try:
            cno = str(cl_df.iloc[i, 13]).strip() if len(cl_df.columns) > 13 else ''  # N列
            publisher = str(cl_df.iloc[i, 18]).strip() if len(cl_df.columns) > 18 else ''  # S列
            series = str(cl_df.iloc[i, 24]).strip() if len(cl_df.columns) > 24 else ''  # Y列

            if cno and cno != 'nan':
                cl_dict[cno] = {
                    'publisher': publisher if publisher != 'nan' else '',
                    'series': series if series != 'nan' else ''
                }
        except Exception:
            continue

    # is_list.csvの出版社とシリーズを置換
    for i in range(1, len(is_df)):
        try:
            cno = str(is_df.iloc[i, 6]).strip() if len(is_df.columns) > 6 else ''  # G列（コミックNo）
            if cno in cl_dict:
                if cl_dict[cno]['publisher'] and len(is_df.columns) > 11:
                    is_df.iloc[i, 11] = cl_dict[cno]['publisher']  # L列
                if cl_dict[cno]['series'] and len(is_df.columns) > 13:
                    is_df.iloc[i, 13] = cl_dict[cno]['series']  # N列
        except Exception:
            continue

    return is_df


def normalize_jan_code(value):
    """JANコードを正規化（数値の.0除去、nan除去）"""
    if pd.isna(value):
        return ''
    jan_str = str(value).strip()
    # '.0' を除去（pandasで数値として読み込まれた場合）
    if jan_str.endswith('.0'):
        jan_str = jan_str[:-2]
    # 'nan' は空文字に
    if jan_str.lower() == 'nan':
        return ''
    return jan_str


def extract_first_volumes(merged_df):
    """1巻のみを抽出して整形"""
    first_vol_dict = {}
    latest_vol_dict = {}
    comic_info_dict = {}  # comic_noごとの情報を保持

    # パス1: 全行を処理して first_vol_dict と latest_vol_dict を構築
    for i in range(1, len(merged_df)):
        try:
            comic_no = normalize_jan_code(merged_df.iloc[i, 6]) if len(merged_df.columns) > 6 else ''  # G列
            if not comic_no:
                continue

            # JAN情報（正規化）
            jan_code = normalize_jan_code(merged_df.iloc[i, 5]) if len(merged_df.columns) > 5 else ''  # F列
            if jan_code:
                latest_vol_dict[comic_no] = jan_code

            # 1巻チェック（J列）
            volume = str(merged_df.iloc[i, 9]).strip() if len(merged_df.columns) > 9 else ''
            if volume == '1' or volume == '1.0':
                if comic_no not in first_vol_dict and jan_code:
                    first_vol_dict[comic_no] = jan_code

            # comic_noの最初の出現行の情報を保持
            if comic_no not in comic_info_dict:
                comic_info_dict[comic_no] = {
                    'kaikatsu_narabi': str(merged_df.iloc[i, 3]).strip() if len(merged_df.columns) > 3 else '',
                    'first_isbn': str(merged_df.iloc[i, 4]).strip() if len(merged_df.columns) > 4 else '',
                    'comic_no': comic_no,
                    'genre': str(merged_df.iloc[i, 7]).strip() if len(merged_df.columns) > 7 else '',
                    'title': str(merged_df.iloc[i, 8]).strip() if len(merged_df.columns) > 8 else '',
                    'publisher': str(merged_df.iloc[i, 11]).strip() if len(merged_df.columns) > 11 else '',
                    'author': str(merged_df.iloc[i, 12]).strip() if len(merged_df.columns) > 12 else '',
                    'series': str(merged_df.iloc[i, 13]).strip() if len(merged_df.columns) > 13 else '',
                }
        except Exception:
            continue

    # パス2: result_dataを構築（全行処理後にfirst_janを設定）
    result_data = []
    for comic_no, info in comic_info_dict.items():
        # 1巻のJAN > 最新巻のJAN > 空 の優先順位
        first_jan = first_vol_dict.get(comic_no, latest_vol_dict.get(comic_no, ''))
        info['first_jan'] = first_jan
        result_data.append(info)

    # 快活並びでソート
    result_data.sort(key=lambda x: int(float(x['kaikatsu_narabi'])) if x['kaikatsu_narabi'] and x['kaikatsu_narabi'] != 'nan' else 999999)
    return result_data


def add_folder_hierarchy_info(result_data, hierarchy_df):
    """フォルダ階層情報を付与"""
    hierarchy_list = []
    for i in range(1, len(hierarchy_df)):
        try:
            row = hierarchy_df.iloc[i]
            hierarchy_list.append({
                'genre': str(row[0]).strip() if pd.notna(row[0]) else '',
                'publisher': str(row[1]).strip() if pd.notna(row[1]) else '',
                'series': str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else '',
                'main_folder': str(row[3]).strip() if len(row) > 3 and pd.notna(row[3]) else '',
                'sub_folder': str(row[4]).strip() if len(row) > 4 and pd.notna(row[4]) else ''
            })
        except Exception:
            continue

    for data in result_data:
        matched = False
        for h in hierarchy_list:
            if data['genre'] == h['genre'] and data['publisher'] == h['publisher']:
                if data['series'] and h['series']:
                    if data['series'] == h['series']:
                        data['main_folder'] = h['main_folder']
                        data['sub_folder'] = h['sub_folder']
                        matched = True
                        break
                elif not h['series']:
                    data['main_folder'] = h['main_folder']
                    data['sub_folder'] = h['sub_folder']
                    matched = True
                    break
        if not matched:
            data['main_folder'] = ''
            data['sub_folder'] = ''

    return result_data


def get_bookoff_image(jan_code, session):
    """ブックオフから画像URL取得"""
    url = f"https://shopping.bookoff.co.jp/search/keyword/{jan_code}"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}

    NO_IMAGE_PATTERNS = ['item_ll.gif', 'no_image', 'noimage', 'no-image', 'dummy', 'blank', 'spacer']
    BeautifulSoup = get_bs4()

    try:
        response = session.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')
        img_tag = soup.select_one('.productItem__image img, .js-gridImg')

        if img_tag and img_tag.get('src'):
            image_url = img_tag['src']
            if any(no_img in image_url.lower() for no_img in NO_IMAGE_PATTERNS):
                return None
            return image_url
        return None
    except Exception:
        return None


def get_amazon_image(jan_code, session):
    """Amazonから画像URL取得（複数セレクタ対応）"""
    search_url = f"https://www.amazon.co.jp/s?k={jan_code}&i=stripbooks"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
    }

    # 複数のセレクタを試す（サイト構造変更に対応）
    SELECTORS = [
        '.s-image',
        'img[data-image-latency]',
        '.s-product-image img',
        '[data-component-type="s-product-image"] img',
        '.s-result-item img[src*="images-na"]',
        '.s-result-item img[src*="m.media-amazon"]',
    ]
    BeautifulSoup = get_bs4()

    try:
        response = session.get(search_url, headers=headers, timeout=15)
        if response.status_code == 503:
            return None
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # 複数のセレクタを順番に試す
        for selector in SELECTORS:
            img_tags = soup.select(selector)
            for img_tag in img_tags:
                src = img_tag.get('src') or img_tag.get('data-src')
                if src and ('images-na' in src or 'm.media-amazon' in src or 'images-amazon' in src):
                    # NO IMAGE系を除外
                    if 'no-img' not in src.lower() and 'no_image' not in src.lower():
                        # 高解像度版に変換
                        if '_AC_' in src:
                            src = src.split('._AC_')[0] + '._SY466_.jpg'
                        elif '_SX' in src or '_SY' in src:
                            # サイズ指定を大きくする
                            import re
                            src = re.sub(r'\._S[XY]\d+_', '._SY466_', src)
                        return src

        # フォールバック: 正規表現でAmazon画像URLを探す
        import re
        amazon_img_pattern = r'(https?://[^"\']+(?:images-na\.ssl-images-amazon|m\.media-amazon|images-amazon)[^"\'\s]+\.(?:jpg|jpeg|png))'
        matches = re.findall(amazon_img_pattern, response.text)
        for match in matches:
            if 'no-img' not in match.lower() and 'no_image' not in match.lower() and 'sprite' not in match.lower():
                if '_AC_' in match:
                    match = match.split('._AC_')[0] + '._SY466_.jpg'
                return match

        return None
    except Exception:
        return None


def get_rakuten_image(jan_code, session):
    """楽天ブックスから画像URL取得（Amazonのフォールバック）"""
    search_url = f"https://books.rakuten.co.jp/search?g=001&isbn={jan_code}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }
    BeautifulSoup = get_bs4()

    try:
        response = session.get(search_url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # 楽天ブックスの画像セレクタ
        selectors = [
            '.rbcomp__item-list__item__image img',
            '.item-image img',
            'img[src*="thumbnail.image.rakuten"]',
        ]

        for selector in selectors:
            img_tag = soup.select_one(selector)
            if img_tag:
                src = img_tag.get('src') or img_tag.get('data-src')
                if src and 'noimage' not in src.lower():
                    # 大きいサイズに変換
                    src = src.replace('_ex=64x64', '_ex=200x200').replace('_ex=100x100', '_ex=200x200')
                    return src

        return None
    except Exception:
        return None


def get_image_with_gemini_ai(jan_code, session, source_name="amazon"):
    """Gemini AIを使って画像URLを抽出（セルフヒーリング機能）"""
    # Geminiモデルを遅延読み込み
    model = get_gemini_model()
    if not model:
        return None

    # ソース別のURL設定
    if source_name == "amazon":
        search_url = f"https://www.amazon.co.jp/s?k={jan_code}&i=stripbooks"
    elif source_name == "rakuten":
        search_url = f"https://books.rakuten.co.jp/search?g=001&isbn={jan_code}"
    elif source_name == "bookoff":
        search_url = f"https://shopping.bookoff.co.jp/search/keyword/{jan_code}"
    else:
        return None

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
    BeautifulSoup = get_bs4()

    try:
        response = session.get(search_url, headers=headers, timeout=15)
        if response.status_code != 200:
            return None

        # HTMLの重要部分だけを抽出（トークン節約）
        soup = BeautifulSoup(response.content, 'html.parser')

        # スクリプトとスタイルを削除
        for tag in soup(['script', 'style', 'noscript', 'header', 'footer', 'nav']):
            tag.decompose()

        # 商品画像が含まれそうな部分を抽出
        main_content = soup.find('main') or soup.find('div', {'id': 'search'}) or soup.find('body')
        if main_content:
            html_snippet = str(main_content)[:8000]  # 最大8000文字に制限
        else:
            html_snippet = str(soup)[:8000]

        prompt = f"""以下のHTMLから、JANコード「{jan_code}」の本の表紙画像URLを1つだけ抽出してください。

条件:
- 画像URLのみを返してください（説明不要）
- NO IMAGE、noimage、placeholder等のダミー画像は除外
- https://で始まる完全なURLで返してください
- 見つからない場合は「NOT_FOUND」とだけ返してください

HTML:
{html_snippet}"""

        response = model.generate_content(prompt)
        result = response.text.strip()

        # 結果を検証
        if result and result != "NOT_FOUND" and result.startswith("http"):
            # NO IMAGE系を最終チェック
            no_image_patterns = ['no_image', 'noimage', 'no-image', 'dummy', 'blank', 'spacer', 'placeholder']
            if not any(p in result.lower() for p in no_image_patterns):
                return result

        return None

    except Exception as e:
        # エラーログ（デバッグ用）
        print(f"Gemini AI error: {e}")
        return None


def download_image(image_url, session):
    """画像をダウンロードしてバイトデータを返す（NO IMAGE検出付き）"""
    try:
        response = session.get(image_url, timeout=10)
        response.raise_for_status()
        content = response.content

        # 画像サイズが小さすぎる場合はNO IMAGEの可能性が高い（5KB未満）
        if len(content) < 5000:
            return None

        # 特定のパターンをURLで再チェック
        no_image_patterns = ['no_image', 'noimage', 'no-image', 'dummy', 'blank', 'spacer', 'placeholder']
        if any(pattern in image_url.lower() for pattern in no_image_patterns):
            return None

        return content
    except Exception:
        return None


@st.cache_data(ttl=600, show_spinner=False)
def get_all_folders():
    """R-Cabinetの全フォルダ一覧を取得"""
    url = f"{BASE_URL}/cabinet/folders/get"
    headers = get_auth_header()

    all_folders = []
    offset = 1  # 1始まり（ページ番号）
    limit = 100  # APIの上限は100件

    while True:
        params = {"offset": offset, "limit": limit}

        try:
            response = requests.get(url, headers=headers, params=params, timeout=30)
        except requests.exceptions.RequestException as e:
            return None, f"接続エラー: {str(e)}"

        if response.status_code != 200:
            return None, f"エラー: {response.status_code} - {response.text[:200]}"

        try:
            root = ET.fromstring(response.text)
        except ET.ParseError as e:
            return None, f"XMLパースエラー: {str(e)}"

        # エラーチェック
        system_status = root.findtext('.//systemStatus', '')
        if system_status != 'OK':
            message = root.findtext('.//message', 'Unknown error')
            return None, f"APIエラー: {message}"

        folders = root.findall('.//folder')

        for folder in folders:
            all_folders.append({
                'FolderId': folder.findtext('FolderId', ''),
                'FolderName': folder.findtext('FolderName', ''),
                'FolderPath': folder.findtext('FolderPath', ''),
                'FileCount': safe_int(folder.findtext('FileCount', '0')),
            })

        # 取得件数がlimit未満なら終了（最終ページ）
        if len(folders) < limit:
            break
        offset += 1  # 次のページへ
        time.sleep(0.3)

    return all_folders, None


@st.cache_data(ttl=300, show_spinner=False)
def get_folder_files(folder_id: int, max_retries: int = 3):
    """指定フォルダ内の画像一覧を取得（リトライ機能付き）"""
    url = f"{BASE_URL}/cabinet/folder/files/get"
    headers = get_auth_header()

    all_files = []
    offset = 1  # 1始まり（ページ番号）
    limit = 100  # APIの上限は100件

    while True:
        params = {"folderId": folder_id, "offset": offset, "limit": limit}

        # リトライ処理
        for retry in range(max_retries):
            try:
                response = requests.get(url, headers=headers, params=params, timeout=30)
            except requests.exceptions.RequestException as e:
                if retry < max_retries - 1:
                    time.sleep(2)  # 2秒待ってリトライ
                    continue
                return None, f"接続エラー: {str(e)}"

            if response.status_code == 200:
                break  # 成功
            elif response.status_code == 403 and retry < max_retries - 1:
                time.sleep(3)  # 403の場合は3秒待ってリトライ
                continue
            else:
                if retry == max_retries - 1:
                    return None, f"エラー: {response.status_code}"

        try:
            root = ET.fromstring(response.text)
        except ET.ParseError as e:
            return None, f"XMLパースエラー: {str(e)}"

        system_status = root.findtext('.//systemStatus', '')
        if system_status != 'OK':
            message = root.findtext('.//message', 'Unknown error')
            return None, f"APIエラー: {message}"

        files = root.findall('.//file')

        for f in files:
            all_files.append({
                'FileId': f.findtext('FileId', ''),
                'FileName': f.findtext('FileName', ''),
                'FileUrl': f.findtext('FileUrl', ''),
                'FilePath': f.findtext('FilePath', ''),
                'FileSize': f.findtext('FileSize', ''),
                'TimeStamp': f.findtext('TimeStamp', ''),
            })

        # 取得件数がlimit未満なら終了（最終ページ）
        if len(files) < limit:
            break
        offset += 1  # 次のページへ
        time.sleep(0.3)

    return all_files, None


def search_image_by_name(file_name: str):
    """画像名で検索"""
    url = f"{BASE_URL}/cabinet/files/search"
    headers = get_auth_header()
    params = {"fileName": file_name}

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        root = ET.fromstring(response.text)
        files = root.findall('.//file')

        results = []
        for f in files:
            results.append({
                'FileId': f.findtext('FileId', ''),
                'FileName': f.findtext('FileName', ''),
                'FileUrl': f.findtext('FileUrl', ''),
                'FolderName': f.findtext('FolderName', ''),
                'FolderPath': f.findtext('FolderPath', ''),
            })
        return results
    return []


def is_exact_match(file_name: str, comic_no: str) -> bool:
    """ファイル名がコミックNoと完全一致するかチェック（拡張子除く）"""
    # 拡張子を除去
    name_without_ext = file_name.rsplit('.', 1)[0] if '.' in file_name else file_name
    # 完全一致のみ
    return name_without_ext == comic_no


def check_comic_images(comic_numbers: list, progress_bar=None, status_text=None):
    """コミックNoリストの画像存在チェック"""
    results = []
    total = len(comic_numbers)

    for i, comic_no in enumerate(comic_numbers):
        if progress_bar:
            progress_bar.progress((i + 1) / total)
        if status_text:
            status_text.text(f"チェック中: {comic_no} ({i + 1}/{total})")

        found_files = search_image_by_name(str(comic_no))

        # 完全一致でフィルタリング
        matched_files = [f for f in found_files if is_exact_match(f['FileName'], str(comic_no))]

        if matched_files:
            for f in matched_files:
                results.append({
                    'コミックNo': comic_no,
                    '存在': '✅ あり',
                    'ファイル名': f['FileName'],
                    'フォルダ': f['FolderName'],
                    'URL': f['FileUrl'],
                })
        else:
            results.append({
                'コミックNo': comic_no,
                '存在': '❌ なし',
                'ファイル名': '-',
                'フォルダ': '-',
                'URL': '-',
            })

        time.sleep(0.4)

    return results


# 認証情報チェック
if not SERVICE_SECRET or not LICENSE_KEY:
    st.error("⚠️ RMS API認証情報が設定されていません。Streamlit Secretsに設定してください。")
    st.stop()


# サイドバー：モード切替
with st.sidebar:
    st.title("🖼️ R-Cabinet")
    st.caption(f"v{APP_VERSION}")

    st.markdown("<br>", unsafe_allow_html=True)

    mode = st.radio(
        "機能を選択",
        ["📂 画像一覧取得", "🔍 画像存在チェック", "📥 不足画像取得"],
        label_visibility="collapsed"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()
    st.markdown("<br>", unsafe_allow_html=True)


# メインコンテンツ
if mode == "📂 画像一覧取得":
    st.title("📂 画像一覧取得")
    st.markdown("R-Cabinetのフォルダを選択して、画像を一覧表示します。")

    # セッション状態の初期化
    if "folders_loaded" not in st.session_state:
        st.session_state.folders_loaded = False
        st.session_state.folders_data = None
        st.session_state.folders_error = None
    if "images_loaded" not in st.session_state:
        st.session_state.images_loaded = False
        st.session_state.images_data = None

    # ステップ1: フォルダ一覧を取得（まだの場合）
    if not st.session_state.folders_loaded:
        st.markdown("### ステップ1: フォルダ一覧を取得")
        if st.button("📂 フォルダ一覧を取得", type="primary"):
            with st.spinner("フォルダ一覧を取得中..."):
                folders, error = get_all_folders()
            st.session_state.folders_data = folders
            st.session_state.folders_error = error
            st.session_state.folders_loaded = True
            st.rerun()
        st.stop()

    folders = st.session_state.folders_data
    error = st.session_state.folders_error

    if error:
        st.error(error)
        if st.button("🔄 再試行"):
            st.session_state.folders_loaded = False
            st.cache_data.clear()
            st.rerun()
        st.stop()

    if not folders:
        st.warning("フォルダがありません。")
        st.stop()

    # 総ファイル数を計算
    total_files = sum(f['FileCount'] for f in folders)

    # サイドバーにフォルダ情報
    with st.sidebar:
        st.success(f"📁 {len(folders)} フォルダ")
        st.info(f"📷 {total_files} 画像（全体）")
        if st.button("🔄 フォルダ再取得"):
            st.session_state.folders_loaded = False
            st.session_state.images_loaded = False
            st.cache_data.clear()
            st.rerun()

    # ステップ2: フォルダ選択
    st.markdown("### フォルダを選択")

    folder_options = {"📁 すべて（全フォルダ）": None}
    folder_options.update({f"{f['FolderName']} ({f['FileCount']}件)": f for f in folders})

    selected_folder_name = st.selectbox(
        "取得するフォルダ",
        list(folder_options.keys()),
        label_visibility="collapsed"
    )

    # DB統計情報を表示
    db_stats = get_db_stats()
    if db_stats.get("total", 0) > 0:
        stat_cols = st.columns(4)
        with stat_cols[0]:
            st.metric("DB登録数", db_stats.get("total", 0))
        with stat_cols[1]:
            st.metric("重複ファイル", db_stats.get("duplicates", 0))
        with stat_cols[2]:
            st.metric("API総数", total_files)
        with stat_cols[3]:
            last_updated = db_stats.get("last_updated", "-")
            st.metric("最終更新", last_updated if last_updated else "-")

    # ステップ3: 操作ボタン（2つ）
    btn_col1, btn_col2, _ = st.columns([1.2, 1.2, 2])
    with btn_col1:
        show_db_btn = st.button(
            "📂 保存済み一覧を表示",
            disabled=(db_stats.get("total", 0) == 0),
            help="DBに保存された一覧を表示（高速）"
        )
    with btn_col2:
        fetch_api_btn = st.button(
            "🔄 最新一覧を取得",
            type="primary",
            help="APIから最新データを取得してDBに同期"
        )

    st.divider()

    # ボタン押下時の処理
    if show_db_btn:
        # DBから読み込み
        st.session_state.data_source = "db"
        if selected_folder_name == "📁 すべて（全フォルダ）":
            loaded_images, msg = load_images_from_db()
        else:
            folder_name = folder_options[selected_folder_name]['FolderName']
            loaded_images = load_images_from_db_by_folder(folder_name)
            msg = f"{len(loaded_images)}件を読み込みました"

        if loaded_images:
            st.session_state.images_data = loaded_images
            st.session_state.images_loaded = True
            st.session_state.error_folders = []
            st.success(f"📂 DBから{msg}")
        else:
            st.warning("DBにデータがありません")

    if fetch_api_btn:
        # APIから取得してDB同期
        st.session_state.data_source = "api"
        st.session_state.images_loaded = False
        st.session_state.images_data = None

        if selected_folder_name == "📁 すべて（全フォルダ）":
            # 全フォルダの画像を取得
            all_files = []
            error_folders = []
            expected_total = sum(f['FileCount'] for f in folders)
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, folder in enumerate(folders):
                status_text.text(f"取得中: {folder['FolderName']} ({i + 1}/{len(folders)}) - {folder['FileCount']}件")
                progress_bar.progress((i + 1) / len(folders))

                files, err = get_folder_files(int(folder['FolderId']))
                time.sleep(0.5)

                if err:
                    error_folders.append({
                        'FolderName': folder['FolderName'],
                        'FolderId': folder['FolderId'],
                        'FileCount': folder['FileCount'],
                        'Error': err
                    })
                if files:
                    for f in files:
                        f['FolderName'] = folder['FolderName']
                    all_files.extend(files)

            progress_bar.empty()
            status_text.empty()

            # DB同期
            with st.spinner("DBに同期中..."):
                sync_result = sync_images_to_db(all_files)

            if sync_result.get("success"):
                st.success(f"🔄 API取得完了・DB同期済み（新規: {sync_result['new']} / 更新: {sync_result['updated']} / 重複: {sync_result['duplicate']}）")
                if sync_result['duplicate'] > 0:
                    st.warning(f"⚠️ {sync_result['duplicate']}件のファイルが複数フォルダに存在")
            else:
                st.error(f"DB同期エラー: {sync_result.get('error')}")

            st.session_state.images_data = all_files
            st.session_state.error_folders = error_folders
            st.session_state.expected_total = expected_total
            st.session_state.images_loaded = True
        else:
            # 個別フォルダの場合
            selected_folder = folder_options[selected_folder_name]
            folder_id = int(selected_folder['FolderId'])

            with st.spinner(f"「{selected_folder['FolderName']}」の画像を取得中..."):
                files, error = get_folder_files(folder_id)

            if error:
                st.error(error)
            elif files:
                for f in files:
                    f['FolderName'] = selected_folder['FolderName']

                # DB同期
                with st.spinner("DBに同期中..."):
                    sync_result = sync_images_to_db(files)

                if sync_result.get("success"):
                    st.success(f"🔄 取得完了（{len(files)}件）・DB同期済み")

                st.session_state.images_data = files
                st.session_state.error_folders = []
                st.session_state.images_loaded = True

    # 画像一覧表示
    if st.session_state.images_loaded and st.session_state.images_data:
        all_files = st.session_state.images_data
        error_folders = st.session_state.get('error_folders', [])

        if all_files:
            # サマリー表示
            st.success(f"📷 {len(all_files)} 件の画像")

            # エラーフォルダがあれば表示
            if error_folders:
                with st.expander(f"⚠️ エラーが発生したフォルダ ({len(error_folders)}件)", expanded=False):
                    for ef in error_folders:
                        st.write(f"- **{ef['FolderName']}** ({ef['FileCount']}件): {ef['Error']}")

            # 検索フィルター
            search_term = st.text_input("🔍 ファイル名で絞り込み", placeholder="検索キーワード")

            display_files = all_files
            if search_term:
                display_files = [f for f in all_files if search_term.lower() in f['FileName'].lower()]
                st.info(f"絞り込み結果: {len(display_files)} 件")

            # データフレーム表示
            df = pd.DataFrame(display_files)
            df = df[['FolderName', 'FileName', 'FileUrl', 'FileSize', 'TimeStamp']]
            df.columns = ['フォルダ', 'ファイル名', 'URL', 'サイズ(KB)', '更新日時']

            st.dataframe(df, use_container_width=True, height=500)

            # Excelダウンロード
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                style_excel(writer.sheets['Sheet1'], num_columns=5, url_column=3)
            excel_buffer.seek(0)
            st.download_button(
                label="📥 Excelでダウンロード",
                data=excel_buffer,
                file_name="rcabinet_images.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("画像がありません。")


elif mode == "🔍 画像存在チェック":
    st.title("🔍 画像存在チェック")
    st.markdown("コミックNoを入力して、R-Cabinetに画像が存在するか確認します。")

    # セッション状態の初期化
    if "check_results" not in st.session_state:
        st.session_state.check_results = None

    st.divider()

    # 入力方法の選択
    input_method = st.radio(
        "入力方法を選択",
        ["テキスト入力", "CSVアップロード"],
        horizontal=True
    )

    comic_numbers = []

    if input_method == "テキスト入力":
        st.markdown("### コミックNo入力")
        st.markdown("1行に1つのコミックNoを入力してください。")

        text_input = st.text_area(
            "コミックNo（改行区切り）",
            height=200,
            placeholder="123456\n234567\n345678"
        )

        if text_input:
            comic_numbers = [line.strip() for line in text_input.split('\n') if line.strip()]
            st.info(f"入力されたコミックNo: {len(comic_numbers)}件")

    else:
        st.markdown("### CSVファイルアップロード")
        st.markdown("コミックNo列を含むCSVファイルをアップロードしてください。")

        uploaded_file = st.file_uploader("CSVファイルを選択", type=['csv'])

        if uploaded_file:
            try:
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            except:
                df = pd.read_csv(uploaded_file, encoding='cp932')

            st.markdown("#### プレビュー")
            st.dataframe(df.head(10), use_container_width=True)

            columns = df.columns.tolist()
            selected_column = st.selectbox("コミックNo列を選択", columns, index=0)

            if selected_column:
                comic_numbers = df[selected_column].dropna().astype(str).tolist()
                st.info(f"読み込んだコミックNo: {len(comic_numbers)}件")

    st.divider()

    # チェック実行ボタン（常に表示）
    check_button = st.button("🔍 チェック実行", type="primary")

    if check_button:
        if not comic_numbers:
            st.warning("コミックNoを入力またはCSVをアップロードしてください。")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            results = check_comic_images(comic_numbers, progress_bar, status_text)

            progress_bar.empty()
            status_text.empty()

            # 結果をsession_stateに保存
            st.session_state.check_results = results

    # 結果表示（session_stateから）
    if st.session_state.check_results:
        results = st.session_state.check_results
        df_results = pd.DataFrame(results)

        st.markdown("### チェック結果")

        exists_count = len([r for r in results if r['存在'] == '✅ あり'])
        not_exists_count = len([r for r in results if r['存在'] == '❌ なし'])

        col1, col2, col3 = st.columns(3)
        col1.metric("総数", len(results))
        col2.metric("存在あり", exists_count)
        col3.metric("存在なし", not_exists_count)

        st.divider()

        filter_option = st.radio(
            "表示フィルター",
            ["すべて", "存在あり", "存在なし"],
            horizontal=True
        )

        if filter_option == "存在あり":
            df_display = df_results[df_results['存在'] == '✅ あり']
        elif filter_option == "存在なし":
            df_display = df_results[df_results['存在'] == '❌ なし']
        else:
            df_display = df_results

        st.dataframe(df_display, use_container_width=True, height=400)

        # ダウンロードボタン（1行目：左寄せ）
        dl_col1, dl_col2, _ = st.columns([1, 1.5, 2])

        with dl_col1:
            # Comic Search検索用CSVダウンロード（存在なしのコミックNoのみ）
            not_exists_comics = [r['コミックNo'] for r in results if r['存在'] == '❌ なし']
            if not_exists_comics:
                # list_コミックナンバー.csv形式で作成
                is_csv_data = []
                for comic_no in not_exists_comics:
                    is_csv_data.append({
                        'ジャンル': '',
                        'タイトル': '',
                        '出版社': '',
                        '著者': '',
                        '完結': '',
                        '巻数': '',
                        'ＩＳＢＮ': '',
                        '棚番': '',
                        'コメント': '',
                        'コミ№': comic_no,
                        '冊数': '1'
                    })
                df_is_csv = pd.DataFrame(is_csv_data)
                csv_buffer = BytesIO()
                df_is_csv.to_csv(csv_buffer, index=False, encoding='cp932')
                csv_buffer.seek(0)
                st.download_button(
                    label="📥 Comic Search検索用CSV",
                    data=csv_buffer,
                    file_name="list_コミックナンバー.csv",
                    mime="text/csv"
                )
            else:
                st.button("📥 Comic Search検索用CSV", disabled=True, help="存在なしのコミックNoがありません")

        with dl_col2:
            # Excelダウンロード（スタイル付き）
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_results.to_excel(writer, index=False, sheet_name='Sheet1')
                style_excel(writer.sheets['Sheet1'], num_columns=5, url_column=5)
            excel_buffer.seek(0)
            st.download_button(
                label="📥 結果ファイルをExcelでダウンロード",
                data=excel_buffer,
                file_name="rcabinet_check_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 2行目：GitHubアップロード、結果クリア
        btn_col3, btn_col4, _ = st.columns([1.5, 1, 2])

        with btn_col3:
            # GitHubにアップロードボタン
            if not_exists_comics:
                if st.button("📤 GitHubにアップロード", help="コミックリスター用にGitHubへアップロード"):
                    # コミックリスター用CSV形式（J列にコミックNo.、K列に1）
                    csv_lines = []
                    for comic_no in not_exists_comics:
                        row = [''] * 9 + [str(comic_no), '1']
                        csv_lines.append(','.join(row))
                    csv_content = '\n'.join(csv_lines)

                    with st.spinner("GitHubにアップロード中..."):
                        today = datetime.now().strftime("%Y-%m-%d %H:%M")
                        result = upload_to_github(
                            csv_content,
                            GITHUB_MISSING_CSV_PATH,
                            f"Update missing_comics.csv ({len(not_exists_comics)}件) - {today}"
                        )

                    if result.get("success"):
                        st.success(f"GitHubにアップロード完了（{len(not_exists_comics)}件）")
                    else:
                        st.error(f"アップロード失敗: {result.get('error')}")
            else:
                st.button("📤 GitHubにアップロード", disabled=True, help="存在なしのコミックNoがありません")

        with btn_col4:
            # 結果クリアボタン
            if st.button("🗑️ 結果をクリア"):
                st.session_state.check_results = None
                st.rerun()


elif mode == "📥 不足画像取得":
    st.title("📥 不足画像取得")
    st.markdown("IS検索結果からJANコードで画像を取得し、ZIPでダウンロードします。")

    st.divider()

    # セッション状態の初期化
    if "github_is_list" not in st.session_state:
        st.session_state.github_is_list = None
    if "github_comic_list" not in st.session_state:
        st.session_state.github_comic_list = None
    if "github_folder_hierarchy" not in st.session_state:
        st.session_state.github_folder_hierarchy = None
    if "image_download_result" not in st.session_state:
        st.session_state.image_download_result = None

    st.markdown("### ステップ0: GitHubからファイル取得")
    st.markdown("GitHub Actionsで生成されたファイルを取得します。")

    # 自動ダウンロードフラグ（無限ループ防止）
    if "auto_download_tried" not in st.session_state:
        st.session_state.auto_download_tried = False

    # まだセッションに読み込まれていない場合は自動ダウンロード（1回だけ試行）
    not_loaded_yet = not st.session_state.github_is_list or not st.session_state.github_comic_list or not st.session_state.github_folder_hierarchy

    if not_loaded_yet and not st.session_state.auto_download_tried:
        st.session_state.auto_download_tried = True
        with st.spinner("GitHubからファイルを自動取得中..."):
            auto_errors = []
            if not st.session_state.github_is_list:
                result = download_from_github(GITHUB_IS_LIST_PATH)
                if result.get("success"):
                    st.session_state.github_is_list = result["content"]
                else:
                    auto_errors.append(f"is_list.csv: {result.get('error', '不明')}")
            if not st.session_state.github_comic_list:
                result = download_from_github(GITHUB_COMIC_LIST_PATH)
                if result.get("success"):
                    st.session_state.github_comic_list = result["content"]
                else:
                    auto_errors.append(f"comic_list.csv: {result.get('error', '不明')}")
            if not st.session_state.github_folder_hierarchy:
                result = download_from_github(GITHUB_FOLDER_HIERARCHY_PATH)
                if result.get("success"):
                    st.session_state.github_folder_hierarchy = result["content"]
                else:
                    auto_errors.append(f"フォルダ階層リスト: {result.get('error', '不明')}")
            if auto_errors:
                st.warning(f"自動取得エラー: {', '.join(auto_errors)}")
        st.rerun()

    # GitHubファイル情報を取得（表示用）
    is_info = get_github_file_info(GITHUB_IS_LIST_PATH)
    cl_info = get_github_file_info(GITHUB_COMIC_LIST_PATH)
    fh_info = get_github_file_info(GITHUB_FOLDER_HIERARCHY_PATH)

    # GitHubファイル情報を表示
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        if is_info.get("exists"):
            st.success(f"is_list.csv\n更新: {is_info.get('last_updated', '不明')}")
        else:
            st.warning("is_list.csv\n未生成")
    with col_info2:
        if cl_info.get("exists"):
            st.success(f"comic_list.csv\n更新: {cl_info.get('last_updated', '不明')}")
        else:
            st.warning("comic_list.csv\n未生成")
    with col_info3:
        if fh_info.get("exists"):
            st.success(f"フォルダ階層リスト\n更新: {fh_info.get('last_updated', '不明')}")
        else:
            st.warning("フォルダ階層リスト\n未配置")

    # フォルダ階層リストのアップロード機能
    hierarchy_upload = st.file_uploader(
        "フォルダ階層リストをアップロード（更新）",
        type=['xlsx'],
        key="hierarchy_quick_upload",
        help="フォルダ階層リスト.xlsxをドラッグ&ドロップしてGitHubにアップロード"
    )
    if hierarchy_upload:
        if st.button("📤 フォルダ階層リストを更新", type="secondary"):
            hierarchy_upload.seek(0)
            content = hierarchy_upload.read()
            result = upload_binary_to_github(
                content,
                GITHUB_FOLDER_HIERARCHY_PATH,
                f"Update folder_hierarchy.xlsx - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            )
            if result.get("success"):
                st.success("フォルダ階層リストを更新しました")
                st.session_state.github_folder_hierarchy = content
                st.rerun()
            else:
                st.error(f"アップロード失敗: {result.get('error')}")

    # CSV生成・取得セクション
    st.markdown("#### CSVファイル操作")

    # 最新の実行履歴を表示（日本時間に変換）
    runs = get_workflow_runs("weekly-comic-lister.yml", limit=1)
    if runs:
        latest = runs[0]
        status_icon = "🟢" if latest["conclusion"] == "success" else "🔴" if latest["conclusion"] == "failure" else "🟡"
        # 日本時間に変換（+9時間）
        from datetime import timedelta
        try:
            dt_utc = datetime.strptime(latest['created_at'], "%Y-%m-%d %H:%M")
            dt_jst = dt_utc + timedelta(hours=9)
            jst_str = dt_jst.strftime("%Y-%m-%d %H:%M")
        except:
            jst_str = latest['created_at']
        status_text = "完了" if latest["conclusion"] == "success" else "失敗" if latest["conclusion"] == "failure" else "処理中..."
        st.caption(f"前回生成: {jst_str} {status_icon} {status_text}")

    # ボタンを横並びに配置（左を目立つ色に）
    btn_col1, btn_col2, _ = st.columns([3, 2, 3])

    with btn_col1:
        run_actions = st.button("📊 is_list / comic_list 生成", type="primary", help="不足コミックのCSVファイルを自動生成します", use_container_width=True)

    with btn_col2:
        fetch_files = st.button("📥 ダウンロード", type="secondary", help="生成済みのファイルをダウンロードします", use_container_width=True)

    # GitHub Actions 実行処理
    if run_actions:
        with st.spinner("CSVファイル生成を開始中..."):
            result = trigger_github_actions("weekly-comic-lister.yml")
        if result.get("success"):
            st.success("CSVファイルの生成を開始しました（完了まで2〜3分お待ちください）")
        else:
            st.error(f"生成開始に失敗しました: {result.get('error')}")

    # GitHubから一括取得処理
    if fetch_files:
        with st.spinner("GitHubからファイルを取得中..."):
            errors = []

            # is_list.csv
            result = download_from_github(GITHUB_IS_LIST_PATH)
            if result.get("success"):
                st.session_state.github_is_list = result["content"]
            else:
                errors.append(f"is_list.csv: {result.get('error')}")

            # comic_list.csv
            result = download_from_github(GITHUB_COMIC_LIST_PATH)
            if result.get("success"):
                st.session_state.github_comic_list = result["content"]
            else:
                errors.append(f"comic_list.csv: {result.get('error')}")

            # folder_hierarchy.xlsx
            result = download_from_github(GITHUB_FOLDER_HIERARCHY_PATH)
            if result.get("success"):
                st.session_state.github_folder_hierarchy = result["content"]
            else:
                errors.append(f"フォルダ階層リスト: {result.get('error')}")

        if errors:
            for err in errors:
                st.warning(err)
        else:
            st.success("全ファイルの取得が完了しました")
        st.rerun()

    # 取得済みファイルの表示
    status_cols = st.columns(3)
    with status_cols[0]:
        if st.session_state.github_is_list:
            st.info("✅ is_list.csv 取得済み")
    with status_cols[1]:
        if st.session_state.github_comic_list:
            st.info("✅ comic_list.csv 取得済み")
    with status_cols[2]:
        if st.session_state.github_folder_hierarchy:
            st.info("✅ フォルダ階層リスト 取得済み")

    st.divider()

    # 使用するファイルを決定（GitHubから取得したもの）
    use_is_list = BytesIO(st.session_state.github_is_list) if st.session_state.github_is_list else None
    use_comic_list = BytesIO(st.session_state.github_comic_list) if st.session_state.github_comic_list else None
    use_hierarchy = BytesIO(st.session_state.github_folder_hierarchy) if st.session_state.github_folder_hierarchy else None

    # ファイルのプレビュー
    if use_is_list:
        st.markdown("### is_list.csv プレビュー")
        try:
            use_is_list.seek(0)
            # UTF-8を先に試し、失敗したらcp932
            try:
                df_is_preview = pd.read_csv(use_is_list, encoding='utf-8', header=None)
            except:
                use_is_list.seek(0)
                df_is_preview = pd.read_csv(use_is_list, encoding='cp932', header=None)
            st.dataframe(df_is_preview.head(10), use_container_width=True, height=200)
            st.info(f"読み込み件数: {len(df_is_preview)}行")
        except Exception as e:
            st.error(f"CSVの読み込みエラー: {e}")

    st.divider()

    st.markdown("### 画像取得")

    # 全ファイルが利用可能かチェック
    all_files_ready = use_is_list and use_comic_list and use_hierarchy

    if not all_files_ready:
        missing = []
        if not use_is_list:
            missing.append("is_list.csv")
        if not use_comic_list:
            missing.append("comic_list.csv")
        if not use_hierarchy:
            missing.append("フォルダ階層リスト.xlsx")
        st.info(f"以下のファイルが必要です: {', '.join(missing)}\n\n「GitHubから一括取得」ボタンを押すか、手動でアップロードしてください。")
    else:
        # 画像取得ボタン
        if st.button("🖼️ 画像取得開始", type="primary"):
            try:
                # ファイル読み込み（UTF-8を先に試し、失敗したらcp932）
                use_is_list.seek(0)
                use_comic_list.seek(0)
                use_hierarchy.seek(0)

                with st.spinner("ファイルを読み込み中..."):
                    # is_list.csv
                    try:
                        df_is = pd.read_csv(use_is_list, encoding='utf-8', header=None)
                    except:
                        use_is_list.seek(0)
                        df_is = pd.read_csv(use_is_list, encoding='cp932', header=None)

                    # comic_list.csv
                    try:
                        df_cl = pd.read_csv(use_comic_list, encoding='utf-8', header=None)
                    except:
                        use_comic_list.seek(0)
                        df_cl = pd.read_csv(use_comic_list, encoding='cp932', header=None)

                    df_hierarchy = pd.read_excel(use_hierarchy, sheet_name="フォルダ階層リスト", header=None)

                st.success(f"ファイル読み込み完了: IS={len(df_is)}行, CL={len(df_cl)}行, 階層={len(df_hierarchy)}行")

                # データ統合
                with st.spinner("データを統合中..."):
                    merged_df = merge_csv_data(df_is.copy(), df_cl)
                    result_data = extract_first_volumes(merged_df)
                    result_data = add_folder_hierarchy_info(result_data, df_hierarchy)

                # JANコードの状態を確認
                jan_count = sum(1 for d in result_data if d.get('first_jan') and normalize_jan_code(d.get('first_jan', '')))
                no_jan_count = len(result_data) - jan_count
                st.success(f"データ統合完了: {len(result_data)}件（JANあり: {jan_count}件, JANなし: {no_jan_count}件）")

                # JANコードがない場合は詳細を表示
                if no_jan_count > 0:
                    no_jan_items = [d for d in result_data if not normalize_jan_code(d.get('first_jan', ''))]
                    with st.expander(f"⚠️ JANコードなし: {no_jan_count}件（詳細）"):
                        for item in no_jan_items[:10]:  # 最大10件表示
                            st.write(f"- {item.get('comic_no', '?')}: {item.get('title', '?')} (first_jan='{item.get('first_jan', '')}')")

                # 画像ダウンロード
                st.markdown("### 画像ダウンロード中...")

                # Gemini AI状態を表示
                if GEMINI_API_KEY:
                    st.info("🤖 Gemini AI セルフヒーリング: 有効（APIキー設定済み）")
                else:
                    st.warning("🤖 Gemini AI セルフヒーリング: 無効（GEMINI_API_KEY未設定）")

                progress_bar = st.progress(0)
                status_text = st.empty()

                session = requests.Session()
                downloaded_images = []
                stats = {'total': len(result_data), 'success': 0, 'bookoff': 0, 'amazon': 0, 'rakuten': 0, 'gemini_ai': 0, 'failed': 0}

                random = get_random()
                for i, data in enumerate(result_data):
                    jan_code = normalize_jan_code(data['first_jan'])
                    comic_no = data['comic_no']

                    progress_bar.progress((i + 1) / len(result_data))
                    status_text.text(f"処理中: {comic_no} ({i + 1}/{len(result_data)}) JAN: {jan_code or '(なし)'}")

                    if not jan_code:
                        stats['failed'] += 1
                        stats['failed_no_jan'] = stats.get('failed_no_jan', 0) + 1
                        continue

                    # 1. ブックオフで検索
                    image_url = get_bookoff_image(jan_code, session)
                    source = 'bookoff'

                    # 2. Amazonで検索
                    if not image_url:
                        time.sleep(random.uniform(0.5, 1.0))
                        image_url = get_amazon_image(jan_code, session)
                        source = 'amazon'

                    # 3. 楽天ブックスで検索（フォールバック）
                    if not image_url:
                        time.sleep(random.uniform(0.3, 0.6))
                        image_url = get_rakuten_image(jan_code, session)
                        source = 'rakuten'

                    # 4. Gemini AIでセルフヒーリング（全て失敗した場合）
                    # デバッグ: AI修復条件を記録
                    ai_condition = f"image_url={bool(image_url)}, GEMINI_API_KEY={bool(GEMINI_API_KEY)}"
                    if not image_url and GEMINI_API_KEY:
                        time.sleep(random.uniform(0.5, 1.0))
                        status_text.text(f"処理中: {comic_no} ({i + 1}/{len(result_data)}) - AI解析中...")
                        stats['gemini_tried'] = stats.get('gemini_tried', 0) + 1
                        # Amazonを再試行（AIでHTML解析）
                        ai_result = get_image_with_gemini_ai(jan_code, session, "amazon")
                        if ai_result:
                            image_url = ai_result
                            source = 'gemini_ai'
                    elif not image_url and not GEMINI_API_KEY:
                        # GEMINI_API_KEYがないためスキップ
                        stats['ai_skipped_no_key'] = stats.get('ai_skipped_no_key', 0) + 1

                    if image_url:
                        image_data = download_image(image_url, session)
                        if image_data:
                            downloaded_images.append({
                                'filename': f"{comic_no}.jpg",
                                'data': image_data,
                                'comic_no': comic_no,
                                'jan': jan_code,
                                'title': data['title']
                            })
                            stats['success'] += 1
                            stats[source] += 1
                        else:
                            stats['failed'] += 1
                            stats['failed_download'] = stats.get('failed_download', 0) + 1
                            # デバッグ: ダウンロード失敗のURLを記録
                            stats['debug_failed_urls'] = stats.get('debug_failed_urls', [])
                            stats['debug_failed_urls'].append({'comic_no': comic_no, 'url': image_url[:100]})
                    else:
                        stats['failed'] += 1
                        stats['failed_not_found'] = stats.get('failed_not_found', 0) + 1

                    time.sleep(0.3)

                progress_bar.empty()
                status_text.empty()

                # 結果をsession_stateに保存
                st.session_state.image_download_result = {
                    'stats': stats,
                    'downloaded_images': downloaded_images,
                    'result_data': result_data
                }
                st.rerun()

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
                import traceback
                st.code(traceback.format_exc())

    # 結果表示（session_stateから）
    if st.session_state.image_download_result:
        result = st.session_state.image_download_result
        stats = result['stats']
        downloaded_images = result['downloaded_images']
        result_data = result['result_data']

        # 結果サマリー
        st.markdown("### 結果")
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.metric("総数", stats['total'])
        col2.metric("成功", stats['success'])
        col3.metric("ブックオフ", stats['bookoff'])
        col4.metric("Amazon", stats['amazon'])
        col5.metric("楽天", stats.get('rakuten', 0))
        col6.metric("AI修復", stats.get('gemini_ai', 0))

        # Gemini AI試行回数を表示
        gemini_tried = stats.get('gemini_tried', 0)
        failed_no_jan = stats.get('failed_no_jan', 0)
        failed_not_found = stats.get('failed_not_found', 0)
        failed_download = stats.get('failed_download', 0)

        if stats['failed'] > 0:
            # 失敗の詳細
            failed_details = []
            if failed_no_jan > 0:
                failed_details.append(f"JANコードなし: {failed_no_jan}件")
            if failed_not_found > 0:
                failed_details.append(f"画像見つからず: {failed_not_found}件")
            if failed_download > 0:
                failed_details.append(f"ダウンロード失敗: {failed_download}件")

            # 詳細がない場合は古い結果の可能性
            if not failed_details:
                failed_details.append("詳細不明（古い結果？→クリアして再実行してください）")

            st.warning(f"取得できなかった画像: {stats['failed']}件 ({', '.join(failed_details)})")

            # AI修復の状態
            ai_skipped_no_key = stats.get('ai_skipped_no_key', 0)
            if GEMINI_API_KEY:
                if gemini_tried > 0:
                    st.info(f"🤖 Gemini AI試行: {gemini_tried}回 → 成功: {stats.get('gemini_ai', 0)}回")
                elif failed_no_jan == stats['failed']:
                    st.info("🤖 AI修復: JANコードがないためスキップ（AI修復にもJANコードが必要です）")
                elif ai_skipped_no_key > 0:
                    st.warning(f"🤖 AI修復: APIキーが実行時に空だった（{ai_skipped_no_key}件スキップ）")
                elif failed_not_found > 0:
                    st.warning("🤖 AI修復が試行されませんでした（要調査：画像が見つからないのにAIが発動していない）")
            else:
                st.warning("🤖 Gemini APIキーが未設定のため、AI修復はスキップされました")

            # デバッグ情報
            with st.expander("🔧 デバッグ情報（詳細）"):
                st.write(f"**stats全体:** {stats}")
                st.write(f"**GEMINI_API_KEY設定:** {'あり' if GEMINI_API_KEY else 'なし'}")
                if stats.get('debug_failed_urls'):
                    st.write("**ダウンロード失敗URL:**")
                    for item in stats['debug_failed_urls'][:5]:
                        st.write(f"  - {item['comic_no']}: {item['url']}")

        # ZIPダウンロード
        if downloaded_images:
            st.divider()
            st.markdown("### ダウンロード")

            # ZIP作成
            zipfile = get_zipfile()
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for img in downloaded_images:
                    zf.writestr(img['filename'], img['data'])
            zip_buffer.seek(0)

            # 振り分けマップExcel作成
            excel_data = []
            for i, data in enumerate(result_data, 1):
                excel_data.append({
                    '連番': i,
                    'コミックNo': data['comic_no'],
                    '1巻JAN': data['first_jan'],
                    'タイトル': data['title'],
                    'ジャンル': data['genre'],
                    '出版社': data['publisher'],
                    '著者': data['author'],
                    'シリーズ': data['series'],
                    'メインフォルダ': data.get('main_folder', ''),
                    'サブフォルダ': data.get('sub_folder', '')
                })

            df_excel = pd.DataFrame(excel_data)
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='振り分けマップ')
                style_excel(writer.sheets['振り分けマップ'], num_columns=10)
            excel_buffer.seek(0)

            # ダウンロードボタンを横並びに
            dl_col1, dl_col2, dl_col3 = st.columns([2, 2, 1])
            with dl_col1:
                st.download_button(
                    label=f"📥 画像ZIP ({len(downloaded_images)}件)",
                    data=zip_buffer,
                    file_name="comic_images.zip",
                    mime="application/zip",
                    key="zip_download"
                )
            with dl_col2:
                st.download_button(
                    label="📥 振り分けマップExcel",
                    data=excel_buffer,
                    file_name="振り分けマップ.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="excel_download"
                )
            with dl_col3:
                if st.button("🗑️ クリア"):
                    st.session_state.image_download_result = None
                    st.rerun()
