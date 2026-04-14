"""
R-Cabinet 管理ツール
- フォルダ画像一覧：R-Cabinetのフォルダ毎に画像を一覧表示
- 画像存在チェック：コミックNoを入力して存在確認
"""

# バージョン（デプロイ確認用）
APP_VERSION = "3.0.0"

import streamlit as st
import requests
import base64
import xml.etree.ElementTree as ET
import pandas as pd
import time
import json
from io import BytesIO
from datetime import datetime, timezone, timedelta

JST = timezone(timedelta(hours=9))

# 重いライブラリは遅延読み込み（起動高速化）
_bs4_module = None
_openpyxl_styles = None
_openpyxl_utils = None
_supabase_module = None
_zipfile_module = None
_random_module = None
_pil_module = None

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


def get_pil():
    """PILを遅延読み込み"""
    global _pil_module
    if _pil_module is None:
        from PIL import Image
        _pil_module = Image
    return _pil_module

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
GITHUB_MISSING_TANPIN_PATH = "comic-lister/data/missing_tanpin.csv"  # 単品用
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
                # ISO形式をパースして日本時間に変換
                dt_utc = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
                dt_jst = dt_utc.astimezone(JST)
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
                    created = dt.astimezone(JST).strftime("%Y-%m-%d %H:%M")
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


def fetch_all_from_supabase(supabase, table: str, columns: str = "*", filter_col: str = None, filter_val: str = None) -> list:
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
        # file_urlはJSONで {フォルダ名: URL} 形式で保存（複数フォルダの場合でも正確なURLを保持）
        file_dict = {}
        for img in images:
            file_name = img.get("FileName", "")
            folder_name = img.get("FolderName", "")
            file_url = img.get("FileUrl", "")
            if file_name in file_dict:
                # 重複: folder_namesに追加、URLもフォルダ別に記録
                existing_folders = file_dict[file_name]["folder_names"].split(", ")
                if folder_name not in existing_folders:
                    file_dict[file_name]["folder_names"] += f", {folder_name}"
                    url_dict = json.loads(file_dict[file_name]["file_url"])
                    url_dict[folder_name] = file_url
                    file_dict[file_name]["file_url"] = json.dumps(url_dict, ensure_ascii=False)
            else:
                file_dict[file_name] = {
                    "file_name": file_name,
                    "folder_names": folder_name,
                    "file_url": json.dumps({folder_name: file_url}, ensure_ascii=False),
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
            folder_names_str = row.get("folder_names", "")
            file_url_raw = row.get("file_url", "")
            file_name = row.get("file_name", "")
            file_size = row.get("file_size", 0)
            file_timestamp = row.get("file_timestamp", "")
            # file_urlをJSONとして解析（新形式: {フォルダ名: URL} の辞書）
            try:
                url_data = json.loads(file_url_raw)
                if isinstance(url_data, dict) and url_data:
                    # 新形式: フォルダ別に行を展開（各フォルダが正確なURLを持つ）
                    for folder in folder_names_str.split(", "):
                        folder = folder.strip()
                        if folder:
                            url = url_data.get(folder, next(iter(url_data.values()), ""))
                            images.append({
                                "FolderName": folder,
                                "FileName": file_name,
                                "FileUrl": url,
                                "FileSize": file_size,
                                "TimeStamp": file_timestamp
                            })
                else:
                    images.append({
                        "FolderName": folder_names_str,
                        "FileName": file_name,
                        "FileUrl": file_url_raw,
                        "FileSize": file_size,
                        "TimeStamp": file_timestamp
                    })
            except (json.JSONDecodeError, TypeError):
                # 旧形式: URLが文字列のまま（DB移行前のデータ）
                images.append({
                    "FolderName": folder_names_str,
                    "FileName": file_name,
                    "FileUrl": file_url_raw,
                    "FileSize": file_size,
                    "TimeStamp": file_timestamp
                })
        return images, f"{len(images)}件を読み込みました"
    except Exception as e:
        return [], str(e)


def get_db_stats() -> dict:
    """DBの統計情報を取得（カウントクエリで高速化）"""
    supabase = get_supabase_client()
    if not supabase:
        return {}

    try:
        # 全件数をカウント（全件フェッチせず高速）
        total_resp = supabase.table("rcabinet_images").select("*", count="exact").limit(1).execute()
        total = total_resp.count or 0

        # 重複ファイル数（folder_namesにカンマを含む件数）
        dup_resp = supabase.table("rcabinet_images").select("*", count="exact").ilike("folder_names", "%, %").limit(1).execute()
        duplicates = dup_resp.count or 0

        return {"total": total, "duplicates": duplicates, "last_updated": None}
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
            folder_names_str = row.get("folder_names", "")
            file_url_raw = row.get("file_url", "")
            file_name = row.get("file_name", "")
            file_size = row.get("file_size", 0)
            file_timestamp = row.get("file_timestamp", "")
            try:
                url_data = json.loads(file_url_raw)
                if isinstance(url_data, dict) and url_data:
                    # 新形式: フォルダ別に行を展開
                    for folder in folder_names_str.split(", "):
                        folder = folder.strip()
                        if folder:
                            url = url_data.get(folder, next(iter(url_data.values()), ""))
                            images.append({
                                "FolderName": folder,
                                "FileName": file_name,
                                "FileUrl": url,
                                "FileSize": file_size,
                                "TimeStamp": file_timestamp
                            })
                else:
                    images.append({
                        "FolderName": folder_names_str,
                        "FileName": file_name,
                        "FileUrl": file_url_raw,
                        "FileSize": file_size,
                        "TimeStamp": file_timestamp
                    })
            except (json.JSONDecodeError, TypeError):
                images.append({
                    "FolderName": folder_names_str,
                    "FileName": file_name,
                    "FileUrl": file_url_raw,
                    "FileSize": file_size,
                    "TimeStamp": file_timestamp
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

    NO_IMAGE_PATTERNS = ['item_ll', 'no_image', 'noimage', 'no-image', 'dummy', 'blank', 'spacer', 'placeholder']
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


def resize_to_square(image_data: bytes, size: int = 600):
    """画像を正方形にリサイズ（アスペクト比維持、白背景パディング、表紙を大きく表示）"""
    Image = get_pil()
    img = Image.open(BytesIO(image_data)).convert("RGB")

    # バッジ領域を考慮した表示領域（右側にバッジがあるため左寄せ）
    display_w = int(size * 0.75)  # 横は75%まで使用
    display_h = int(size * 0.90)  # 縦は90%まで使用

    # アスペクト比を維持して表示領域にフィット（拡大も許可）
    ratio = min(display_w / img.width, display_h / img.height)
    new_w = int(img.width * ratio)
    new_h = int(img.height * ratio)
    img = img.resize((new_w, new_h), Image.LANCZOS)

    # 白背景の正方形キャンバスに左寄せ・上下中央配置
    canvas = Image.new("RGB", (size, size), (255, 255, 255))
    x = (display_w - new_w) // 2 + int(size * 0.02)  # 左寄せ（少し余白）
    y = (size - new_h) // 2
    canvas.paste(img, (x, y))

    return canvas


def add_shipping_badge(base_image, badge_path: str):
    """送料無料バッジを合成（白背景を透明化してオーバーレイ）"""
    Image = get_pil()
    import numpy as np

    base = base_image.convert("RGBA")
    badge = Image.open(badge_path).convert("RGBA")

    # バッジを元画像サイズにリサイズ
    badge = badge.resize(base.size, Image.LANCZOS)

    # 白背景を透明化（RGB各チャンネルが240以上を透明に）
    badge_data = np.array(badge)
    white_mask = (badge_data[:, :, 0] > 240) & (badge_data[:, :, 1] > 240) & (badge_data[:, :, 2] > 240)
    badge_data[white_mask, 3] = 0
    badge = Image.fromarray(badge_data, "RGBA")

    # 合成
    result = Image.alpha_composite(base, badge)
    return result.convert("RGB")


def image_to_bytes(image, quality: int = 95) -> bytes:
    """PIL ImageをJPEGバイトデータに変換"""
    buf = BytesIO()
    image.save(buf, format="JPEG", quality=quality)
    return buf.getvalue()


def process_workflow_images(missing_comics: list, is_list_content: str, comic_list_content: str, badge_path: str, progress_bar=None, status_text=None, log_container=None):
    """ワークフロー用：不足画像を取得してバッジ合成まで行う"""
    import os

    # CSVをDataFrameに変換
    try:
        is_df = pd.read_csv(BytesIO(is_list_content.encode('utf-8')), header=None)
    except:
        is_df = pd.read_csv(BytesIO(is_list_content.encode('cp932')), header=None)

    try:
        cl_df = pd.read_csv(BytesIO(comic_list_content.encode('utf-8')), header=None)
    except:
        cl_df = pd.read_csv(BytesIO(comic_list_content.encode('cp932')), header=None)

    # CSVマージ＋JAN抽出
    merged_df = merge_csv_data(is_df, cl_df)
    result_data = extract_first_volumes(merged_df)

    # missing_comicsでフィルタリング（.0除去で正規化して比較）
    missing_set = set(normalize_jan_code(c) for c in missing_comics if normalize_jan_code(c))
    # セット品のみ（_なし）でresult_dataをフィルタリング
    missing_set_only = set(c for c in missing_set if '_' not in c)
    target_data = [d for d in result_data if str(d.get('comic_no', '')).strip() in missing_set_only]

    # 単品（_あり）はis_listから直接JAN検索
    # is_dfから (comic_no, volume) → JAN の辞書を構築
    is_jan_lookup = {}  # (comic_no, volume) → JAN
    for i in range(1, len(is_df)):
        try:
            cno = str(is_df.iloc[i, 6]).strip() if pd.notna(is_df.iloc[i, 6]) else ''
            cno = cno.replace('.0', '')
            vol = str(is_df.iloc[i, 9]).strip() if pd.notna(is_df.iloc[i, 9]) else ''
            vol = vol.replace('.0', '')
            jan = normalize_jan_code(is_df.iloc[i, 5])
            if cno and jan:
                is_jan_lookup[(cno, vol)] = jan
        except:
            continue

    result_data_dict = {str(d.get('comic_no', '')).strip(): d for d in result_data}

    tanpin_comics = [c for c in missing_comics if '_' in str(c)]
    for tc in tanpin_comics:
        parts = str(tc).split('_')
        base_no = parts[0]
        vol_num = int(parts[1]) if len(parts) > 1 else 1

        # (comic_no, volume) 完全一致でJAN検索
        jan_code = is_jan_lookup.get((base_no, str(vol_num)), '')

        base_info = result_data_dict.get(base_no, {})
        if jan_code:
            target_data.append({
                'comic_no': tc,
                'first_jan': jan_code,
                'is_tanpin': True,
                'genre': base_info.get('genre', ''),
                'publisher': base_info.get('publisher', ''),
                'series': base_info.get('series', ''),
                'title': base_info.get('title', ''),
            })

    # セット品フラグ追加
    for d in target_data:
        if 'is_tanpin' not in d:
            d['is_tanpin'] = '_' in str(d.get('comic_no', ''))

    if not target_data:
        return {'success': False, 'error': '処理対象がありません', 'images': [], 'stats': {}}

    # 画像取得
    session = requests.Session()
    random = get_random()
    downloaded_images = []
    stats = {'total': len(target_data), 'success': 0, 'failed': 0, 'bookoff': 0, 'amazon': 0, 'rakuten': 0, 'gemini_ai': 0}
    logs = []

    for i, data in enumerate(target_data):
        comic_no = str(data.get('comic_no', '')).strip()
        jan_code = normalize_jan_code(data.get('first_jan', ''))

        if progress_bar:
            progress_bar.progress((i + 1) / len(target_data))
        if status_text:
            status_text.text(f"処理中: {comic_no} ({i + 1}/{len(target_data)})")

        if not jan_code:
            logs.append(f"⚠️ {comic_no}: JANコードなし - スキップ")
            stats['failed'] += 1
            continue

        # 画像取得（優先順）
        image_url = None
        source = ''

        # 1. ブックオフ
        image_url = get_bookoff_image(jan_code, session)
        source = 'bookoff'

        # 2. Amazon
        if not image_url:
            time.sleep(random.uniform(0.5, 1.0))
            image_url = get_amazon_image(jan_code, session)
            source = 'amazon'

        # 3. 楽天
        if not image_url:
            time.sleep(random.uniform(0.3, 0.6))
            image_url = get_rakuten_image(jan_code, session)
            source = 'rakuten'

        # 4. Gemini AI
        if not image_url and GEMINI_API_KEY:
            time.sleep(random.uniform(0.5, 1.0))
            ai_result = get_image_with_gemini_ai(jan_code, session, "amazon")
            if ai_result:
                image_url = ai_result
                source = 'gemini_ai'

        if not image_url:
            logs.append(f"❌ {comic_no} (JAN: {jan_code}): 画像が見つかりません")
            stats['failed'] += 1
            continue

        # ダウンロード
        image_data = download_image(image_url, session)
        if not image_data:
            logs.append(f"❌ {comic_no}: ダウンロード失敗 ({source})")
            stats['failed'] += 1
            continue

        is_tanpin = data.get('is_tanpin', False)

        if is_tanpin:
            # 単品：取得した画像をそのまま使用（加工なし）
            final_bytes = image_data
            badge_status = "加工なし"
        else:
            # セット品：600x600リサイズ＋バッジ合成
            resized = resize_to_square(image_data, 600)
            if os.path.exists(badge_path):
                final_image = add_shipping_badge(resized, badge_path)
                badge_status = "バッジ付き"
            else:
                final_image = resized
                badge_status = "バッジ画像なし"
            final_bytes = image_to_bytes(final_image)

        downloaded_images.append({
            'comic_no': comic_no,
            'jan_code': jan_code,
            'image_data': final_bytes,
            'source': source,
            'is_tanpin': is_tanpin,
            'badge': not is_tanpin,
            'genre': data.get('genre', ''),
            'publisher': data.get('publisher', ''),
            'series': data.get('series', ''),
            'title': data.get('title', '')
        })

        stats['success'] += 1
        stats[source] += 1
        logs.append(f"✅ {comic_no} (JAN: {jan_code}): {source} - {badge_status}")

        # レート制限
        time.sleep(random.uniform(0.3, 0.8))

    return {'success': True, 'images': downloaded_images, 'stats': stats, 'logs': logs}


def prepare_yahoo_zips(images: list, excel_set_df, excel_tanpin_df, additional_dir: str) -> dict:
    """ヤフー用にリネーム＋追加画像＋ZIP分割生成"""
    import os
    zipfile = get_zipfile()
    Image = get_pil()

    MAX_ZIP_SIZE = 25 * 1024 * 1024  # 25MB

    # コミックNo → 商品コード のマッピング構築
    comic_to_product = {}

    # セット品シート: A列=商品コード, D列=コミックNo
    if excel_set_df is not None:
        for i in range(len(excel_set_df)):
            try:
                product_code = str(excel_set_df.iloc[i, 0]).strip()
                comic_no = str(excel_set_df.iloc[i, 3]).strip().replace('.0', '')
                if product_code and comic_no and product_code != 'nan' and comic_no != 'nan':
                    comic_to_product[comic_no] = {'code': product_code, 'type': 'set'}
            except:
                continue

    # 単品シート: A列=商品コード, E列=コミックNo
    if excel_tanpin_df is not None:
        for i in range(len(excel_tanpin_df)):
            try:
                product_code = str(excel_tanpin_df.iloc[i, 0]).strip()
                comic_no = str(excel_tanpin_df.iloc[i, 4]).strip().replace('.0', '')
                if product_code and comic_no and product_code != 'nan' and comic_no != 'nan':
                    comic_to_product[comic_no] = {'code': product_code, 'type': 'tanpin'}
            except:
                continue

    # 追加画像読み込み
    additional_1_path = os.path.join(additional_dir, "additional_1.jpg")
    additional_2_path = os.path.join(additional_dir, "additional_2.jpg")
    additional_1_data = None
    additional_2_data = None
    if os.path.exists(additional_1_path):
        with open(additional_1_path, 'rb') as f:
            additional_1_data = f.read()
    if os.path.exists(additional_2_path):
        with open(additional_2_path, 'rb') as f:
            additional_2_data = f.read()

    # ファイルリスト構築
    file_entries = []  # [(filename, bytes)]
    mapped_count = 0
    unmapped = []
    logs = []

    for img in images:
        comic_no = str(img['comic_no']).strip()
        mapping = comic_to_product.get(comic_no)

        if not mapping:
            unmapped.append(comic_no)
            logs.append(f"⚠️ {comic_no}: 商品コードが見つかりません")
            continue

        product_code = mapping['code']
        is_set = mapping['type'] == 'set'

        # メイン画像
        file_entries.append((f"{product_code}.jpg", img['image_data']))

        # セット品のみ追加画像
        if is_set:
            if additional_1_data:
                file_entries.append((f"{product_code}_1.jpg", additional_1_data))
            if additional_2_data:
                file_entries.append((f"{product_code}_2.jpg", additional_2_data))

        mapped_count += 1
        logs.append(f"✅ {comic_no} → {product_code} {'(セット品+追加画像)' if is_set else '(単品)'}")

    # ZIP分割生成
    zip_buffers = []
    current_files = []
    current_size = 0

    for filename, data in file_entries:
        file_size = len(data)
        if current_size + file_size > MAX_ZIP_SIZE and current_files:
            # 現在のZIPを保存
            buf = BytesIO()
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fn, fd in current_files:
                    zf.writestr(fn, fd)
            zip_buffers.append(buf.getvalue())
            current_files = []
            current_size = 0

        current_files.append((filename, data))
        current_size += file_size

    # 残りを保存
    if current_files:
        buf = BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fn, fd in current_files:
                zf.writestr(fn, fd)
        zip_buffers.append(buf.getvalue())

    return {
        'zips': zip_buffers,
        'mapped': mapped_count,
        'unmapped': unmapped,
        'total_files': len(file_entries),
        'logs': logs
    }


def prepare_rakuten_queue(images: list, hierarchy_df, folders: list) -> dict:
    """楽天用にフォルダマッピング＋アップロードキュー生成"""
    # フォルダパス → FolderId の辞書
    folder_map = {}
    for f in folders:
        folder_map[f['FolderName']] = f['FolderId']

    # 階層情報パース
    hierarchy_list = []
    if hierarchy_df is not None:
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
            except:
                continue

    # 画像ごとにフォルダを決定
    queue = []
    unmapped = []
    logs = []

    # is_listのresult_dataが必要だが、workflow_dataから取得する形
    # ここではcomic_noベースでシンプルにマッチング
    for img in images:
        comic_no = str(img['comic_no']).strip()
        file_name = f"{comic_no}.jpg"

        # sub_folderをフォルダ名として使用
        matched_folder = None
        matched_main_folder = None
        matched_folder_id = None

        # result_dataにgenre/publisher情報がある場合はマッチング
        genre = img.get('genre', '')
        publisher = img.get('publisher', '')
        series = img.get('series', '')

        # 画像データに既にmain_folder/sub_folderが付与済みの場合はそちらを優先
        if img.get('main_folder') and img.get('sub_folder'):
            matched_main_folder = img['main_folder']
            matched_folder = img['sub_folder'] or img['main_folder']
        else:
            for h in hierarchy_list:
                if genre == h['genre'] and publisher == h['publisher']:
                    if series and h['series'] and series == h['series']:
                        matched_folder = h['sub_folder'] or h['main_folder']
                        matched_main_folder = h['main_folder']
                        break
                    elif not h['series']:
                        matched_folder = h['sub_folder'] or h['main_folder']
                        matched_main_folder = h['main_folder']
                        break

        if matched_folder and matched_folder in folder_map:
            matched_folder_id = folder_map[matched_folder]
            queue.append({
                'file_bytes': img['image_data'],
                'folder_id': matched_folder_id,
                'folder_name': matched_folder,
                'main_folder': matched_main_folder,
                'file_name': file_name,
                'comic_no': comic_no
            })
            logs.append(f"✅ {comic_no} → {matched_main_folder}/{matched_folder} (ID: {matched_folder_id})")
        else:
            unmapped.append(comic_no)
            if matched_folder:
                logs.append(f"⚠️ {comic_no}: フォルダ '{matched_folder}' がR-Cabinetに見つかりません")
            else:
                logs.append(f"⚠️ {comic_no}: フォルダ階層にマッチしません")

    return {
        'queue': queue,
        'mapped': len(queue),
        'unmapped': unmapped,
        'logs': logs
    }


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


def create_folder(folder_name, directory_name=None, upper_folder_id=None):
    """R-Cabinetにフォルダを1件作成（cabinet.folder.insert）"""
    url = f"{BASE_URL}/cabinet/folder/insert"
    headers = get_auth_header()
    headers["Content-Type"] = "text/xml;charset=UTF-8"

    # XMLリクエストボディを構築
    folder_elements = f"<folderName>{folder_name}</folderName>"
    if directory_name:
        folder_elements += f"<directoryName>{directory_name}</directoryName>"
    if upper_folder_id:
        folder_elements += f"<upperFolderId>{upper_folder_id}</upperFolderId>"

    xml_body = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        "<request><folderInsertRequest><folder>"
        f"{folder_elements}"
        "</folder></folderInsertRequest></request>"
    )

    try:
        response = requests.post(url, headers=headers, data=xml_body.encode('utf-8'), timeout=30)
    except requests.exceptions.RequestException as e:
        return {"success": False, "error": f"接続エラー: {str(e)}"}

    if response.status_code != 200:
        return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:500]}"}

    try:
        root = ET.fromstring(response.text)
    except ET.ParseError as e:
        return {"success": False, "error": f"XMLパースエラー: {str(e)}"}

    system_status = root.findtext('.//systemStatus', '')
    if system_status != 'OK':
        message = root.findtext('.//message', 'Unknown error')
        return {"success": False, "error": f"APIエラー: {message}"}

    folder_id = root.findtext('.//FolderId', '')
    return {"success": True, "folder_id": folder_id}


def upload_image(file_data, file_name, folder_id, file_path_name=None, overwrite=False):
    """R-Cabinetに画像を1枚アップロード（cabinet.file.insert）"""
    url = f"{BASE_URL}/cabinet/file/insert"
    headers = get_auth_header()
    # Content-Typeはrequestsが自動設定（multipart/form-data + boundary）

    # XMLパラメータ部分を構築
    xml_elements = f"<fileName>{file_name}</fileName>"
    xml_elements += f"<folderId>{folder_id}</folderId>"
    if file_path_name:
        xml_elements += f"<filePath>{file_path_name}</filePath>"
    if overwrite:
        xml_elements += "<overWrite>true</overWrite>"

    xml_body = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        "<request><cabinetFileInsertRequest>"
        f"{xml_elements}"
        "</cabinetFileInsertRequest></request>"
    )

    try:
        response = requests.post(
            url,
            headers=headers,
            data={"xml": xml_body},
            files={"file": (file_name, file_data, "application/octet-stream")},
            timeout=60,
        )
    except requests.exceptions.RequestException as e:
        return {"success": False, "error": f"接続エラー: {str(e)}"}

    if response.status_code != 200:
        return {"success": False, "error": f"HTTP {response.status_code}: {response.text[:500]}"}

    try:
        root = ET.fromstring(response.text)
    except ET.ParseError as e:
        return {"success": False, "error": f"XMLパースエラー: {str(e)}"}

    system_status = root.findtext('.//systemStatus', '')
    if system_status != 'OK':
        message = root.findtext('.//message', 'Unknown error')
        return {"success": False, "error": f"APIエラー: {message}"}

    file_url = root.findtext('.//FileUrl', '')
    file_id = root.findtext('.//FileId', '')
    return {"success": True, "file_url": file_url, "file_id": file_id}


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
    """コミックNoリストの画像存在チェック（DB参照版 - 高速）"""
    results = []
    total = len(comic_numbers)

    # DBから全画像データを取得（1回だけ）
    if status_text:
        status_text.text("DBからデータを読み込み中...")
    if progress_bar:
        progress_bar.progress(0.1)

    all_images, _ = load_images_from_db()

    if not all_images:
        # DBにデータがない場合はエラー
        return None

    if progress_bar:
        progress_bar.progress(0.3)

    # ファイル名（拡張子除く）→ 画像情報の辞書を作成
    if status_text:
        status_text.text("検索インデックスを作成中...")

    image_dict = {}
    for img in all_images:
        file_name = img.get('FileName', '')
        name_without_ext = file_name.rsplit('.', 1)[0] if '.' in file_name else file_name
        if name_without_ext not in image_dict:
            image_dict[name_without_ext] = []
        image_dict[name_without_ext].append(img)

    if progress_bar:
        progress_bar.progress(0.5)

    # 各コミックNoをチェック（メモリ内検索なので高速）
    if status_text:
        status_text.text("チェック中...")

    for i, comic_no in enumerate(comic_numbers):
        comic_no_str = str(comic_no).strip()

        if comic_no_str in image_dict:
            # 存在する場合
            for img in image_dict[comic_no_str]:
                results.append({
                    'コミックNo': comic_no,
                    '存在': '✅ あり',
                    'ファイル名': img.get('FileName', ''),
                    'フォルダ': img.get('FolderName', ''),
                    'URL': img.get('FileUrl', ''),
                })
        else:
            # 存在しない場合
            results.append({
                'コミックNo': comic_no,
                '存在': '❌ なし',
                'ファイル名': '-',
                'フォルダ': '-',
                'URL': '-',
            })

        # 進捗更新（100件ごと）
        if progress_bar and (i + 1) % 100 == 0:
            progress_bar.progress(0.5 + 0.5 * (i + 1) / total)

    if progress_bar:
        progress_bar.progress(1.0)

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
        ["🔄 画像ワークフロー", "📂 画像一覧取得", "🔍 画像存在チェック", "🖼️ 新規画像取得", "📁 フォルダ一括作成", "📤 画像アップロード"],
        label_visibility="collapsed"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()
    st.markdown("<br>", unsafe_allow_html=True)


# ============================================================
# 画像ワークフロー（統合モード）のカスタムCSS
# ============================================================
WORKFLOW_CSS = """
<style>
/* ステップナビゲーション */
.workflow-nav {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 1.5rem 2rem;
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border-radius: 16px;
    margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
    position: relative;
    overflow: hidden;
}

.workflow-nav::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 3px;
    background: linear-gradient(90deg, #667eea, #764ba2, #f093fb);
}

.step-item {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 0.5rem;
    flex: 1;
    position: relative;
    z-index: 1;
}

.step-circle {
    width: 48px;
    height: 48px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 1.1rem;
    transition: all 0.3s ease;
}

.step-circle.completed {
    background: linear-gradient(135deg, #00d9a5 0%, #00b894 100%);
    color: white;
    box-shadow: 0 4px 15px rgba(0, 217, 165, 0.4);
}

.step-circle.active {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    box-shadow: 0 4px 20px rgba(102, 126, 234, 0.5);
    transform: scale(1.1);
    animation: pulse 2s infinite;
}

.step-circle.pending {
    background: rgba(255, 255, 255, 0.1);
    color: rgba(255, 255, 255, 0.5);
    border: 2px dashed rgba(255, 255, 255, 0.2);
}

@keyframes pulse {
    0%, 100% { box-shadow: 0 4px 20px rgba(102, 126, 234, 0.5); }
    50% { box-shadow: 0 4px 30px rgba(102, 126, 234, 0.8); }
}

.step-label {
    font-size: 0.75rem;
    color: rgba(255, 255, 255, 0.7);
    text-align: center;
    max-width: 80px;
}

.step-label.active {
    color: white;
    font-weight: 600;
}

.step-connector {
    flex: 0.5;
    height: 3px;
    background: rgba(255, 255, 255, 0.1);
    position: relative;
    margin-top: -24px;
}

.step-connector.completed {
    background: linear-gradient(90deg, #00d9a5, #00b894);
}

/* ステップコンテンツカード */
.step-card {
    background: linear-gradient(145deg, #ffffff 0%, #f8f9fa 100%);
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
    border: 1px solid rgba(0, 0, 0, 0.05);
}

.step-card-header {
    display: flex;
    align-items: center;
    gap: 1rem;
    margin-bottom: 1.5rem;
    padding-bottom: 1rem;
    border-bottom: 2px solid #f0f0f0;
}

.step-card-icon {
    width: 56px;
    height: 56px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.5rem;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
}

.step-card-title {
    font-size: 1.4rem;
    font-weight: 700;
    color: #1a1a2e;
    margin: 0;
}

.step-card-desc {
    font-size: 0.9rem;
    color: #666;
    margin: 0;
}

/* アクションボタン */
.action-btn-primary {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    border: none;
    padding: 0.8rem 2rem;
    border-radius: 12px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
}

.action-btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
}

/* 結果サマリーカード */
.result-card {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    border-radius: 12px;
    padding: 1.2rem;
    text-align: center;
    border: 1px solid rgba(0, 0, 0, 0.05);
}

.result-card.success {
    background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
    border-color: rgba(40, 167, 69, 0.2);
}

.result-card.warning {
    background: linear-gradient(135deg, #fff3cd 0%, #ffeeba 100%);
    border-color: rgba(255, 193, 7, 0.2);
}

.result-card.error {
    background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
    border-color: rgba(220, 53, 69, 0.2);
}

.result-value {
    font-size: 2rem;
    font-weight: 700;
    color: #1a1a2e;
}

.result-label {
    font-size: 0.8rem;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* プログレスバー */
.progress-container {
    background: rgba(0, 0, 0, 0.05);
    border-radius: 10px;
    height: 12px;
    overflow: hidden;
    margin: 1rem 0;
}

.progress-bar {
    height: 100%;
    border-radius: 10px;
    background: linear-gradient(90deg, #667eea, #764ba2);
    transition: width 0.3s ease;
}

/* ログエリア */
.log-area {
    background: #1a1a2e;
    border-radius: 12px;
    padding: 1rem;
    max-height: 300px;
    overflow-y: auto;
    font-family: 'JetBrains Mono', 'Fira Code', monospace;
    font-size: 0.85rem;
}

.log-entry {
    padding: 0.3rem 0;
    border-bottom: 1px solid rgba(255, 255, 255, 0.05);
}

.log-entry.success { color: #00d9a5; }
.log-entry.error { color: #ff6b6b; }
.log-entry.info { color: #74b9ff; }

</style>
"""


def render_workflow_step_nav(current_step: int, completed_steps: list):
    """ワークフローのステップナビゲーションを描画"""
    steps = [
        ("①", "不足特定"),
        ("②", "JAN取得"),
        ("③", "画像取得"),
        ("④", "準備"),
        ("⑤", "アップロード")
    ]

    html = '<div class="workflow-nav">'

    for i, (num, label) in enumerate(steps):
        step_num = i + 1
        if step_num in completed_steps:
            status = "completed"
            icon = "✓"
        elif step_num == current_step:
            status = "active"
            icon = num
        else:
            status = "pending"
            icon = num

        label_class = "active" if step_num == current_step else ""

        html += f'''
        <div class="step-item">
            <div class="step-circle {status}">{icon}</div>
            <div class="step-label {label_class}">{label}</div>
        </div>
        '''

        # コネクター（最後以外）
        if i < len(steps) - 1:
            conn_status = "completed" if step_num in completed_steps else ""
            html += f'<div class="step-connector {conn_status}"></div>'

    html += '</div>'
    return html


# メインコンテンツ
if mode == "🔄 画像ワークフロー":
    st.markdown(WORKFLOW_CSS, unsafe_allow_html=True)

    st.title("🔄 画像ワークフロー")
    st.markdown("不足画像の特定から楽天・ヤフーへのアップロードまで、一気通貫で処理します。")

    # セッション状態の初期化
    if "workflow_step" not in st.session_state:
        st.session_state.workflow_step = 1
    if "workflow_completed" not in st.session_state:
        st.session_state.workflow_completed = []
    if "workflow_data" not in st.session_state:
        st.session_state.workflow_data = {}

    current_step = st.session_state.workflow_step
    completed_steps = st.session_state.workflow_completed

    # ステップナビゲーション
    st.markdown(render_workflow_step_nav(current_step, completed_steps), unsafe_allow_html=True)

    # ステップ選択（手動でジャンプ可能）
    with st.expander("ステップを直接選択", expanded=False):
        step_options = {
            "① 不足特定": 1,
            "② JAN取得": 2,
            "③ 画像取得": 3,
            "④ アップロード準備": 4,
            "⑤ API連携": 5
        }
        selected = st.selectbox(
            "移動先",
            list(step_options.keys()),
            index=current_step - 1,
            label_visibility="collapsed"
        )
        if st.button("移動"):
            st.session_state.workflow_step = step_options[selected]
            st.rerun()

    st.divider()

    # ============================================================
    # Step 1: 不足特定
    # ============================================================
    if current_step == 1:
        st.markdown("""
        <div class="step-card">
            <div class="step-card-header">
                <div class="step-card-icon">🔍</div>
                <div>
                    <p class="step-card-title">Step ① 不足特定</p>
                    <p class="step-card-desc">R-Cabinetに存在しない画像を特定します</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 入力方法の選択
        input_method = st.radio(
            "入力方法",
            ["出品シートExcel", "テキスト入力"],
            horizontal=True
        )

        comic_numbers = []

        if input_method == "出品シートExcel":
            st.info(
                "**出品シートExcelをアップロード（複数可）**\n\n"
                "- セット品 → A列: 商品コード ／ D列: コミックNo\n"
                "- 単品 → A列: 商品コード ／ E列: コミックNo\n\n"
                "各ファイルの **Sheet1** が読み込まれます。Step④のヤフーマッピングにも自動連携されます。"
            )
            excel_files = st.file_uploader("出品シートExcel", type=['xlsx', 'xls'], key="step1_excel", accept_multiple_files=True)

            if excel_files:
                set_comics = []
                tanpin_comics = []
                yahoo_files_data = []  # ヤフー連携用

                for idx, excel_file in enumerate(excel_files):
                    file_type = st.radio(
                        f"📄 {excel_file.name}",
                        ["セット品", "単品"],
                        horizontal=True,
                        key=f"step1_ftype_{idx}"
                    )

                    try:
                        df = pd.read_excel(excel_file, sheet_name=0, header=None)
                        excel_file.seek(0)
                        file_bytes = excel_file.read()

                        if file_type == "セット品":
                            col_idx = 3  # D列
                            for i in range(len(df)):
                                try:
                                    cno = str(df.iloc[i, col_idx]).strip().replace('.0', '')
                                    if cno and cno != 'nan' and cno.replace('_', '').isdigit():
                                        set_comics.append(cno)
                                except:
                                    continue
                            yahoo_files_data.append({'bytes': file_bytes, 'type': 'set'})
                        else:
                            col_idx = 4  # E列
                            for i in range(len(df)):
                                try:
                                    cno = str(df.iloc[i, col_idx]).strip().replace('.0', '')
                                    if cno and cno != 'nan' and cno.replace('_', '').isdigit():
                                        tanpin_comics.append(cno)
                                except:
                                    continue
                            yahoo_files_data.append({'bytes': file_bytes, 'type': 'tanpin'})

                        st.caption(f"→ {len(df)}行 / コミックNo抽出済み")

                    except Exception as e:
                        st.error(f"{excel_file.name} 読み込みエラー: {e}")

                comic_numbers = list(set(set_comics + tanpin_comics))
                if comic_numbers:
                    st.info(f"合計: {len(comic_numbers)}件（セット: {len(set_comics)}件, 単品: {len(tanpin_comics)}件）")

                # ヤフーマッピング用にセッションに保存
                st.session_state.workflow_data['yahoo_excel_files'] = yahoo_files_data
        else:
            # テキスト入力
            text_input = st.text_area(
                "コミックNo（改行区切り）",
                height=150,
                placeholder="123456\n234567\n19763_003"
            )
            if text_input:
                comic_numbers = [line.strip() for line in text_input.split('\n') if line.strip()]
                st.info(f"入力: {len(comic_numbers)}件")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("🔍 チェック実行", type="primary", disabled=not comic_numbers):
                progress = st.progress(0)
                status = st.empty()

                results = check_comic_images(comic_numbers, progress, status)

                progress.empty()
                status.empty()

                if results:
                    st.session_state.workflow_data['check_results'] = results
                    exists_count = len([r for r in results if r['存在'] == '✅ あり'])
                    missing_count = len([r for r in results if r['存在'] == '❌ なし'])

                    cols = st.columns(3)
                    cols[0].metric("総数", len(results))
                    cols[1].metric("存在あり", exists_count)
                    cols[2].metric("存在なし", missing_count)

                    if missing_count > 0:
                        st.success(f"不足画像 {missing_count}件 を特定しました")
                else:
                    st.error("DBにデータがありません")

        # 結果がある場合
        if 'check_results' in st.session_state.workflow_data:
            results = st.session_state.workflow_data['check_results']
            exists_items = [r for r in results if r['存在'] == '✅ あり']
            # RECフォルダを除外した画像
            exists_items_no_rec = [r for r in exists_items if 'REC' not in (r.get('フォルダ', '') or '').upper()]
            missing = [r for r in results if r['存在'] == '❌ なし']

            # 存在あり画像のダウンロード
            if exists_items:
                rec_count = len(exists_items) - len(exists_items_no_rec)
                if exists_items_no_rec:
                    expander_label = f"📦 存在あり画像をダウンロード（{len(exists_items_no_rec)}件、REC {rec_count}件除外）"
                else:
                    expander_label = f"📦 存在あり画像（{len(exists_items)}件すべてRECフォルダ）"
                with st.expander(expander_label):
                    if 'wf_rcab_dl_result' not in st.session_state:
                        st.session_state.wf_rcab_dl_result = None

                    if not exists_items_no_rec:
                        st.info("REC除外後の対象が0件のためダウンロードできません")

                    if st.button("🖼️ R-Cabinetから画像を取得", type="primary", key="wf_rcab_dl_btn", disabled=not exists_items_no_rec):
                        _zipfile = get_zipfile()
                        progress = st.progress(0)
                        status = st.empty()
                        session = requests.Session()

                        downloaded = []
                        failed = []
                        for i, item in enumerate(exists_items_no_rec):
                            comic_no = str(item['コミックNo'])
                            url = item.get('URL', '')
                            folder = item.get('フォルダ', '')
                            file_name = item.get('ファイル名', '') or f"{comic_no}.jpg"
                            if '.' not in file_name:
                                file_name = f"{file_name}.jpg"

                            status.text(f"ダウンロード中: {comic_no} ({i+1}/{len(exists_items_no_rec)})")
                            progress.progress((i + 1) / len(exists_items_no_rec))

                            if not url or url == '-':
                                failed.append(comic_no)
                                continue
                            try:
                                resp = session.get(url, timeout=15)
                                if resp.status_code == 200 and len(resp.content) > 100:
                                    downloaded.append({
                                        'comic_no': comic_no,
                                        'file_name': file_name,
                                        'folder': folder,
                                        'data': resp.content,
                                    })
                                else:
                                    failed.append(comic_no)
                            except Exception:
                                failed.append(comic_no)

                        progress.empty()
                        status.empty()
                        st.session_state.wf_rcab_dl_result = {'downloaded': downloaded, 'failed': failed}
                        st.rerun()

                    if st.session_state.wf_rcab_dl_result:
                        dl_result = st.session_state.wf_rcab_dl_result
                        downloaded = dl_result['downloaded']
                        failed = dl_result['failed']

                        st.success(f"取得完了: {len(downloaded)}件成功" + (f", {len(failed)}件失敗" if failed else ""))

                        if downloaded:
                            _zipfile = get_zipfile()
                            dl_cols = st.columns(2)

                            with dl_cols[0]:
                                buf_flat = BytesIO()
                                with _zipfile.ZipFile(buf_flat, 'w', _zipfile.ZIP_DEFLATED) as zf:
                                    for img in downloaded:
                                        zf.writestr(img['file_name'], img['data'])
                                buf_flat.seek(0)
                                st.download_button(
                                    label=f"📦 フラットZIP（{len(downloaded)}件）",
                                    data=buf_flat,
                                    file_name="rcabinet_images_flat.zip",
                                    mime="application/zip",
                                    key="wf_rcab_dl_flat"
                                )
                                st.caption("全画像を直下に配置")

                            with dl_cols[1]:
                                buf_folder = BytesIO()
                                with _zipfile.ZipFile(buf_folder, 'w', _zipfile.ZIP_DEFLATED) as zf:
                                    for img in downloaded:
                                        folder = img['folder'] if img['folder'] and img['folder'] != '-' else 'その他'
                                        zf.writestr(f"{folder}/{img['file_name']}", img['data'])
                                buf_folder.seek(0)
                                st.download_button(
                                    label=f"📂 フォルダ付きZIP（{len(downloaded)}件）",
                                    data=buf_folder,
                                    file_name="rcabinet_images_by_folder.zip",
                                    mime="application/zip",
                                    key="wf_rcab_dl_folder"
                                )
                                st.caption("R-Cabinetのフォルダ構成を保持")

            if missing:
                st.divider()
                st.markdown("### 不足画像一覧")

                df_missing = pd.DataFrame(missing)
                st.dataframe(df_missing, use_container_width=True, height=200)

                col1, col2 = st.columns([1, 1])
                with col1:
                    if st.button("📤 GitHubにアップロード", type="secondary"):
                        # セット品と単品を分離
                        set_comics = [r['コミックNo'] for r in missing if '_' not in str(r['コミックNo'])]
                        tanpin_comics = [r['コミックNo'] for r in missing if '_' in str(r['コミックNo'])]

                        today = datetime.now(JST).strftime('%Y-%m-%d %H:%M')
                        upload_results = []

                        if set_comics:
                            content = '\n'.join([str(c) for c in set_comics])
                            result = upload_to_github(content, GITHUB_MISSING_CSV_PATH, f"Update missing_comics.csv ({len(set_comics)}件) - {today}")
                            if result.get("success"):
                                upload_results.append(f"セット品: {len(set_comics)}件 ✅")

                        if tanpin_comics:
                            content = '\n'.join([str(c) for c in tanpin_comics])
                            result = upload_to_github(content, GITHUB_MISSING_TANPIN_PATH, f"Update missing_tanpin.csv ({len(tanpin_comics)}件) - {today}")
                            if result.get("success"):
                                upload_results.append(f"単品: {len(tanpin_comics)}件 ✅")

                        if upload_results:
                            st.success(", ".join(upload_results))
                            st.session_state.workflow_data['missing_uploaded'] = True

                with col2:
                    if st.button("次へ進む →", type="primary"):
                        if 1 not in st.session_state.workflow_completed:
                            st.session_state.workflow_completed.append(1)
                        st.session_state.workflow_step = 2
                        st.rerun()

    # ============================================================
    # Step 2: JAN取得
    # ============================================================
    elif current_step == 2:
        st.markdown("""
        <div class="step-card">
            <div class="step-card-header">
                <div class="step-card-icon">📊</div>
                <div>
                    <p class="step-card-title">Step ② JAN取得</p>
                    <p class="step-card-desc">コミックリスターでJAN情報を取得します（GitHub Actions）</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 最新の実行履歴を表示
        runs = get_workflow_runs("weekly-comic-lister.yml", limit=3)
        if runs:
            st.markdown("### 実行履歴")
            for run in runs:
                status_icon = "🟢" if run["conclusion"] == "success" else "🔴" if run["conclusion"] == "failure" else "🟡"
                st.write(f"{status_icon} {run['created_at']} - {run['conclusion'] or '実行中'}")

        if st.button("📊 CSV生成・取得", type="primary"):
            status_area = st.empty()
            progress_area = st.empty()

            # 1. 不足リストが未アップロードなら自動アップロード
            if not st.session_state.workflow_data.get('missing_uploaded'):
                check_results = st.session_state.workflow_data.get('check_results', [])
                missing = [r for r in check_results if r.get('存在') == '❌ なし']
                if missing:
                    set_comics = [r['コミックNo'] for r in missing if '_' not in str(r['コミックNo'])]
                    tanpin_comics = [r['コミックNo'] for r in missing if '_' in str(r['コミックNo'])]
                    today = datetime.now(JST).strftime('%Y-%m-%d %H:%M')
                    status_area.info("📤 不足リストをGitHubにアップロード中...")
                    if set_comics:
                        content = '\n'.join([str(c) for c in set_comics])
                        upload_to_github(content, GITHUB_MISSING_CSV_PATH, f"Update missing_comics.csv ({len(set_comics)}件) - {today}")
                    if tanpin_comics:
                        content = '\n'.join([str(c) for c in tanpin_comics])
                        upload_to_github(content, GITHUB_MISSING_TANPIN_PATH, f"Update missing_tanpin.csv ({len(tanpin_comics)}件) - {today}")
                    st.session_state.workflow_data['missing_uploaded'] = True

            # 2. GitHub Actionsを起動
            status_area.info("🚀 GitHub Actionsを起動中...")
            result = trigger_github_actions("weekly-comic-lister.yml")
            if not result.get("success"):
                status_area.error(f"エラー: {result.get('error')}")
            else:
                # 3. ワークフロー完了をポーリング（最大5分）
                import time as _time
                max_wait = 300
                poll_interval = 15
                elapsed = 0
                completed = False

                status_area.info("⏳ ワークフロー実行中... 完了まで2〜3分かかります")

                # トリガー直後は少し待つ（新しいrunが作られるまで）
                _time.sleep(5)
                elapsed += 5

                while elapsed < max_wait:
                    _time.sleep(poll_interval)
                    elapsed += poll_interval
                    minutes = elapsed // 60
                    seconds = elapsed % 60
                    time_str = f"{minutes}分{seconds}秒" if minutes > 0 else f"{seconds}秒"
                    progress_area.progress(min(elapsed / max_wait, 0.95), text=f"⏳ ワークフロー実行中... ({time_str}経過)")

                    runs = get_workflow_runs("weekly-comic-lister.yml", limit=1)
                    if not runs:
                        continue
                    run = runs[0]
                    # まだ実行中なら待機を続ける
                    if run.get("status") in ("queued", "in_progress"):
                        continue
                    # 完了した場合
                    if run.get("conclusion") == "success":
                        completed = True
                    break

                progress_area.empty()

                if completed:
                    # 4. 完了後にCSVを自動ダウンロード
                    status_area.info("📥 ワークフロー完了。CSVをダウンロード中...")
                    is_result = download_from_github(GITHUB_IS_LIST_PATH)
                    cl_result = download_from_github(GITHUB_COMIC_LIST_PATH)

                    if is_result.get("success") and cl_result.get("success"):
                        is_content = is_result["content"]
                        if isinstance(is_content, bytes):
                            is_content = is_content.decode('utf-8', errors='replace')
                        cl_content = cl_result["content"]
                        if isinstance(cl_content, bytes):
                            cl_content = cl_content.decode('utf-8', errors='replace')
                        st.session_state.workflow_data['is_list'] = is_content
                        st.session_state.workflow_data['comic_list'] = cl_content
                        status_area.success("✅ CSV生成・取得が完了しました。Step ③に進めます")
                        st.balloons()
                    else:
                        status_area.error("CSVのダウンロードに失敗しました")
                elif elapsed >= max_wait:
                    status_area.warning("⏱️ タイムアウト（5分）しました。完了後に再度ボタンを押してください")
                else:
                    status_area.error("❌ ワークフローが失敗しました")

        # ファイル状態表示
        st.divider()
        st.markdown("### ファイル状態")
        col1, col2 = st.columns(2)
        with col1:
            if st.session_state.workflow_data.get('is_list'):
                st.success("✅ is_list.csv 取得済み")
            else:
                st.warning("⬜ is_list.csv 未取得")
        with col2:
            if st.session_state.workflow_data.get('comic_list'):
                st.success("✅ comic_list.csv 取得済み")
            else:
                st.warning("⬜ comic_list.csv 未取得")

        # 次へ進むボタン
        st.divider()
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("← 戻る"):
                st.session_state.workflow_step = 1
                st.rerun()
        with col2:
            files_ready = st.session_state.workflow_data.get('is_list') and st.session_state.workflow_data.get('comic_list')
            if st.button("次へ進む →", type="primary", disabled=not files_ready):
                if 2 not in st.session_state.workflow_completed:
                    st.session_state.workflow_completed.append(2)
                st.session_state.workflow_step = 3
                st.rerun()

    # ============================================================
    # Step 3: 画像取得
    # ============================================================
    elif current_step == 3:
        import os
        st.markdown("""
        <div class="step-card">
            <div class="step-card-header">
                <div class="step-card-icon">🖼️</div>
                <div>
                    <p class="step-card-title">Step ③ 画像取得＋加工</p>
                    <p class="step-card-desc">JANコードで画像を取得し、送料無料バッジを合成します（600×600px）</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # --- 入力データの確認 ---
        st.markdown("### 入力データ")

        # 不足リスト（Step ①の結果 or GitHubから取得済み）
        missing_comics = []
        if 'check_results' in st.session_state.workflow_data:
            missing = [r for r in st.session_state.workflow_data['check_results'] if r['存在'] == '❌ なし']
            missing_comics = [normalize_jan_code(r['コミックNo']) for r in missing]
        elif st.session_state.workflow_data.get('missing_from_github'):
            missing_comics = st.session_state.workflow_data['missing_from_github']

        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            if missing_comics:
                st.success(f"✅ 不足リスト: {len(missing_comics)}件")
            else:
                st.warning("⬜ 不足リスト: なし")
                st.caption("Step ①を実行するか、GitHubから取得してください")
        with col_s2:
            if st.session_state.workflow_data.get('is_list'):
                st.success("✅ is_list.csv 取得済み")
            else:
                st.warning("⬜ is_list.csv 未取得")
        with col_s3:
            if st.session_state.workflow_data.get('comic_list'):
                st.success("✅ comic_list.csv 取得済み")
            else:
                st.warning("⬜ comic_list.csv 未取得")

        # GitHubから取得ボタン群
        need_fetch = not missing_comics or not st.session_state.workflow_data.get('is_list') or not st.session_state.workflow_data.get('comic_list')
        if need_fetch:
            st.divider()
            fetch_cols = st.columns(3)
            with fetch_cols[0]:
                if not missing_comics and st.button("📥 不足リスト取得"):
                    with st.spinner("GitHubから取得中..."):
                        parsed_comics = []

                        # セット品（missing_comics.csv）
                        result_set = download_from_github(GITHUB_MISSING_CSV_PATH)
                        if result_set.get("success"):
                            content = result_set.get("content", b"")
                            if isinstance(content, bytes):
                                content = content.decode('utf-8', errors='replace')
                            for l in content.strip().split('\n'):
                                l = l.strip()
                                if not l:
                                    continue
                                if ',' in l:
                                    fields = [f.strip() for f in l.split(',') if f.strip()]
                                    for f in fields:
                                        if f.replace('_', '').replace('.0', '').isdigit() and len(f.replace('.0', '')) > 1:
                                            parsed_comics.append(f.replace('.0', ''))
                                            break
                                else:
                                    parsed_comics.append(l)

                        # 単品（missing_tanpin.csv）
                        result_tanpin = download_from_github(GITHUB_MISSING_TANPIN_PATH)
                        if result_tanpin.get("success"):
                            content = result_tanpin.get("content", b"")
                            if isinstance(content, bytes):
                                content = content.decode('utf-8', errors='replace')
                            for l in content.strip().split('\n'):
                                l = l.strip()
                                if l:
                                    parsed_comics.append(l)

                    if parsed_comics:
                        missing_comics = parsed_comics
                        st.session_state.workflow_data['missing_from_github'] = missing_comics
                        set_c = len([c for c in parsed_comics if '_' not in str(c)])
                        tanpin_c = len([c for c in parsed_comics if '_' in str(c)])
                        st.success(f"{len(parsed_comics)}件取得（セット品: {set_c}件 / 単品: {tanpin_c}件）")
                        st.rerun()
                    else:
                        st.error("取得失敗またはデータなし")
            with fetch_cols[1]:
                if not st.session_state.workflow_data.get('is_list') and st.button("📥 is_list.csv取得"):
                    with st.spinner("取得中..."):
                        result = download_from_github(GITHUB_IS_LIST_PATH)
                    if result.get("success"):
                        content = result["content"]
                        if isinstance(content, bytes):
                            content = content.decode('utf-8', errors='replace')
                        st.session_state.workflow_data['is_list'] = content
                        st.success("取得完了")
                        st.rerun()
                    else:
                        st.error("取得失敗")
            with fetch_cols[2]:
                if not st.session_state.workflow_data.get('comic_list') and st.button("📥 comic_list.csv取得"):
                    with st.spinner("取得中..."):
                        result = download_from_github(GITHUB_COMIC_LIST_PATH)
                    if result.get("success"):
                        content = result["content"]
                        if isinstance(content, bytes):
                            content = content.decode('utf-8', errors='replace')
                        st.session_state.workflow_data['comic_list'] = content
                        st.success("取得完了")
                        st.rerun()
                    else:
                        st.error("取得失敗")

        # GitHubから取得した不足リストも反映
        if not missing_comics and st.session_state.workflow_data.get('missing_from_github'):
            missing_comics = st.session_state.workflow_data['missing_from_github']

        # --- CSVダウンロード ---
        with st.expander("📄 CSVファイルをダウンロード"):
            dl_cols = st.columns(4)
            csv_files = [
                ("missing_tanpin.csv", GITHUB_MISSING_TANPIN_PATH),
                ("missing_comics.csv", GITHUB_MISSING_CSV_PATH),
                ("is_list.csv", GITHUB_IS_LIST_PATH),
                ("comic_list.csv", GITHUB_COMIC_LIST_PATH),
            ]
            # キャッシュキー
            if 'csv_downloads' not in st.session_state:
                st.session_state.csv_downloads = {}
            for idx, (fname, gpath) in enumerate(csv_files):
                with dl_cols[idx]:
                    if fname in st.session_state.csv_downloads:
                        data = st.session_state.csv_downloads[fname]
                        st.download_button(
                            label=f"💾 {fname}",
                            data=data,
                            file_name=fname,
                            mime="text/csv",
                            key=f"dl_{fname}"
                        )
                    else:
                        if st.button(f"📥 {fname}", key=f"fetch_{fname}"):
                            with st.spinner("取得中..."):
                                result = download_from_github(gpath)
                            if result.get("success"):
                                content = result["content"]
                                if isinstance(content, bytes):
                                    st.session_state.csv_downloads[fname] = content
                                else:
                                    st.session_state.csv_downloads[fname] = content.encode('utf-8')
                                st.rerun()
                            else:
                                st.error(f"取得失敗: {result.get('error')}")

        # --- 実行セクション ---
        st.divider()
        all_ready = missing_comics and st.session_state.workflow_data.get('is_list') and st.session_state.workflow_data.get('comic_list')

        # セット品と単品の内訳
        if missing_comics:
            set_count = len([c for c in missing_comics if '_' not in str(c)])
            tanpin_count = len([c for c in missing_comics if '_' in str(c)])
            st.markdown(f"**対象: {len(missing_comics)}件**（セット品: {set_count}件 / 単品: {tanpin_count}件）")

        if st.button("🖼️ 画像取得開始", type="primary", disabled=not all_ready):
            badge_path = os.path.join(os.path.dirname(__file__), "images", "badge_free_shipping.jpg")

            progress = st.progress(0)
            status = st.empty()

            result = process_workflow_images(
                missing_comics=missing_comics,
                is_list_content=st.session_state.workflow_data['is_list'],
                comic_list_content=st.session_state.workflow_data['comic_list'],
                badge_path=badge_path,
                progress_bar=progress,
                status_text=status
            )

            progress.empty()
            status.empty()

            if result.get('success'):
                st.session_state.workflow_data['downloaded_images'] = result['images']
                st.session_state.workflow_data['image_stats'] = result['stats']
                st.session_state.workflow_data['image_logs'] = result['logs']
                st.rerun()
            else:
                st.error(result.get('error', '画像取得に失敗しました'))

        # --- 結果表示セクション ---
        if 'downloaded_images' in st.session_state.workflow_data:
            images = st.session_state.workflow_data['downloaded_images']
            stats = st.session_state.workflow_data.get('image_stats', {})
            logs = st.session_state.workflow_data.get('image_logs', [])

            st.divider()
            st.markdown("### 取得結果")

            # 統計
            stat_cols = st.columns(6)
            stat_cols[0].metric("対象", stats.get('total', 0))
            stat_cols[1].metric("成功", stats.get('success', 0))
            stat_cols[2].metric("失敗", stats.get('failed', 0))
            stat_cols[3].metric("ブックオフ", stats.get('bookoff', 0))
            stat_cols[4].metric("Amazon", stats.get('amazon', 0))
            stat_cols[5].metric("楽天/AI", f"{stats.get('rakuten', 0)}/{stats.get('gemini_ai', 0)}")

            # 画像プレビュー（3列グリッド、最大6件）
            if images:
                preview_count = min(len(images), 6)
                st.markdown(f"### プレビュー（{preview_count}/{len(images)}件）")
                for row_start in range(0, preview_count, 3):
                    cols = st.columns(3)
                    for j, col in enumerate(cols):
                        idx = row_start + j
                        if idx < len(images):
                            img = images[idx]
                            with col:
                                st.image(img['image_data'], width=200)
                                badge_label = "🏷️ バッジ付き" if img['badge'] else "📷 バッジなし"
                                st.caption(f"{img['comic_no']} ({img['source']}) {badge_label}")

                # ZIPダウンロード
                st.divider()
                zipfile = get_zipfile()
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for img in images:
                        filename = f"{img['comic_no']}.jpg"
                        zf.writestr(filename, img['image_data'])

                st.download_button(
                    label=f"📦 ZIPダウンロード（{len(images)}件）",
                    data=zip_buffer.getvalue(),
                    file_name="workflow_images.zip",
                    mime="application/zip"
                )

            # ログ
            with st.expander("ログ", expanded=False):
                for log in logs:
                    st.text(log)

        # ナビゲーション
        st.divider()
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("← 戻る"):
                st.session_state.workflow_step = 2
                st.rerun()
        with col2:
            has_images = 'downloaded_images' in st.session_state.workflow_data and st.session_state.workflow_data['downloaded_images']
            if st.button("次へ進む →", type="primary", disabled=not has_images):
                if 3 not in st.session_state.workflow_completed:
                    st.session_state.workflow_completed.append(3)
                st.session_state.workflow_step = 4
                st.rerun()

    # ============================================================
    # Step 4: アップロード準備
    # ============================================================
    elif current_step == 4:
        import os as _os
        st.markdown("""
        <div class="step-card">
            <div class="step-card-header">
                <div class="step-card-icon">📦</div>
                <div>
                    <p class="step-card-title">Step ④ アップロード準備</p>
                    <p class="step-card-desc">楽天・ヤフー向けにファイルを整形します</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Step ③の画像確認
        images = st.session_state.workflow_data.get('downloaded_images', [])
        if not images:
            st.warning("Step ③で画像を取得してください。")
            if st.button("← Step ③に戻る"):
                st.session_state.workflow_step = 3
                st.rerun()
        else:
            st.info(f"対象画像: {len(images)}件（セット品: {len([i for i in images if not i.get('is_tanpin')])}件 / 単品: {len([i for i in images if i.get('is_tanpin')])}件）")

            tab_yahoo, tab_rakuten = st.tabs(["🛒 ヤフー準備", "🏪 楽天準備"])

            # ============================================
            # ヤフータブ
            # ============================================
            with tab_yahoo:
                st.markdown("### データフォーマットExcel")

                excel_set_df = None
                excel_tanpin_df = None

                # Step ①で出品シートExcelが読み込み済みか確認
                yahoo_from_step1 = st.session_state.workflow_data.get('yahoo_excel_files')

                if yahoo_from_step1:
                    st.success(f"Step ①で出品シートExcel {len(yahoo_from_step1)}件 を読み込み済み")
                    try:
                        set_dfs = []
                        tanpin_dfs = []
                        for fd in yahoo_from_step1:
                            df = pd.read_excel(BytesIO(fd['bytes']), sheet_name=0, header=None)
                            if fd['type'] == 'set':
                                set_dfs.append(df)
                            else:
                                tanpin_dfs.append(df)
                        if set_dfs:
                            excel_set_df = pd.concat(set_dfs, ignore_index=True)
                        if tanpin_dfs:
                            excel_tanpin_df = pd.concat(tanpin_dfs, ignore_index=True)
                        st.caption(f"セット品: {len(set_dfs)}ファイル / 単品: {len(tanpin_dfs)}ファイル")
                    except Exception as e:
                        st.error(f"Excel読み込みエラー: {e}")
                        yahoo_from_step1 = None

                if not yahoo_from_step1:
                    st.info(
                        "**出品シートExcelをアップロード（複数可）**\n\n"
                        "- セット品 → A列: 商品コード ／ D列: コミックNo\n"
                        "- 単品 → A列: 商品コード ／ E列: コミックNo\n\n"
                        "各ファイルの **Sheet1** が読み込まれます。"
                    )
                    yahoo_excels = st.file_uploader("Excelファイル", type=['xlsx', 'xls'], key="yahoo_excel", accept_multiple_files=True)

                    if yahoo_excels:
                        set_dfs = []
                        tanpin_dfs = []
                        for idx, yf in enumerate(yahoo_excels):
                            file_type = st.radio(
                                f"📄 {yf.name}",
                                ["セット品", "単品"],
                                horizontal=True,
                                key=f"yahoo_ftype_{idx}"
                            )
                            try:
                                df = pd.read_excel(yf, sheet_name=0, header=None)
                                if file_type == "セット品":
                                    set_dfs.append(df)
                                else:
                                    tanpin_dfs.append(df)
                            except Exception as e:
                                st.error(f"{yf.name} 読み込みエラー: {e}")

                        if set_dfs:
                            excel_set_df = pd.concat(set_dfs, ignore_index=True)
                        if tanpin_dfs:
                            excel_tanpin_df = pd.concat(tanpin_dfs, ignore_index=True)

                # プレビュー（どちらの読み込み方法でも共通）
                if excel_set_df is not None:
                    with st.expander("マッピングプレビュー", expanded=True):
                        set_mappings = []
                        for i in range(len(excel_set_df)):
                            try:
                                code = str(excel_set_df.iloc[i, 0]).strip()
                                cno = str(excel_set_df.iloc[i, 3]).strip().replace('.0', '')
                                if code != 'nan' and cno != 'nan' and code and cno and cno.replace('_', '').isdigit():
                                    matched = any(str(img['comic_no']) == cno for img in images)
                                    set_mappings.append({'商品コード': code, 'コミックNo': cno, '対象画像': '✅' if matched else '❌'})
                            except:
                                continue

                        tanpin_mappings = []
                        if excel_tanpin_df is not None:
                            for i in range(len(excel_tanpin_df)):
                                try:
                                    code = str(excel_tanpin_df.iloc[i, 0]).strip()
                                    cno = str(excel_tanpin_df.iloc[i, 4]).strip().replace('.0', '')
                                    if code != 'nan' and cno != 'nan' and code and cno and cno.replace('_', '').isdigit():
                                        matched = any(str(img['comic_no']) == cno for img in images)
                                        tanpin_mappings.append({'商品コード': code, 'コミックNo': cno, '対象画像': '✅' if matched else '❌'})
                                except:
                                    continue

                        if set_mappings:
                            st.markdown(f"**セット品: {len(set_mappings)}件**")
                            st.dataframe(pd.DataFrame(set_mappings), use_container_width=True, height=150)
                        if tanpin_mappings:
                            st.markdown(f"**単品: {len(tanpin_mappings)}件**")
                            st.dataframe(pd.DataFrame(tanpin_mappings), use_container_width=True, height=150)

                # ZIP生成ボタン
                st.divider()
                can_generate = excel_set_df is not None
                if st.button("📦 ZIP生成", type="primary", disabled=not can_generate, key="yahoo_zip_btn"):
                    additional_dir = _os.path.join(_os.path.dirname(__file__), "images")
                    result = prepare_yahoo_zips(images, excel_set_df, excel_tanpin_df, additional_dir)

                    st.session_state.workflow_data['yahoo_zips'] = result
                    st.rerun()

                # ZIP結果表示
                if 'yahoo_zips' in st.session_state.workflow_data:
                    result = st.session_state.workflow_data['yahoo_zips']
                    zips = result['zips']

                    col_m1, col_m2, col_m3 = st.columns(3)
                    col_m1.metric("マッピング成功", result['mapped'])
                    col_m2.metric("未マッチ", len(result['unmapped']))
                    col_m3.metric("ZIP数", len(zips))

                    for i, zip_data in enumerate(zips):
                        size_mb = len(zip_data) / (1024 * 1024)
                        st.download_button(
                            label=f"📥 yahoo_upload_{i+1:03d}.zip ({size_mb:.1f}MB)",
                            data=zip_data,
                            file_name=f"yahoo_upload_{i+1:03d}.zip",
                            mime="application/zip",
                            key=f"yahoo_zip_dl_{i}"
                        )

                    with st.expander("ログ", expanded=False):
                        for log in result['logs']:
                            st.text(log)

            # ============================================
            # 楽天タブ
            # ============================================
            with tab_rakuten:
                st.markdown("### フォルダマッピング")

                # フォルダ階層リスト取得
                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    if st.button("📥 フォルダ階層リスト取得", key="rakuten_hierarchy_btn"):
                        with st.spinner("GitHubから取得中..."):
                            result = download_from_github(GITHUB_FOLDER_HIERARCHY_PATH)
                        if result.get("success"):
                            content = result["content"]
                            if isinstance(content, bytes):
                                st.session_state.workflow_data['hierarchy_xlsx'] = content
                            st.success("フォルダ階層リスト取得完了")
                            st.rerun()
                        else:
                            st.error(f"取得失敗: {result.get('error')}")

                with col_r2:
                    if st.button("📂 楽天フォルダ一覧取得", key="rakuten_folders_btn"):
                        with st.spinner("R-Cabinet APIから取得中..."):
                            folders, error = get_all_folders()
                        if error:
                            st.error(error)
                        elif folders:
                            st.session_state.workflow_data['rakuten_folders'] = folders
                            st.success(f"{len(folders)}フォルダ取得")
                            st.rerun()

                # 状態表示
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    if st.session_state.workflow_data.get('hierarchy_xlsx'):
                        st.success("✅ フォルダ階層リスト取得済み")
                    else:
                        st.warning("⬜ フォルダ階層リスト未取得")
                with col_s2:
                    folders = st.session_state.workflow_data.get('rakuten_folders')
                    if folders:
                        st.success(f"✅ 楽天フォルダ: {len(folders)}件")
                    else:
                        st.warning("⬜ 楽天フォルダ未取得")

                # データ確認エリア
                with st.expander("データ確認（中身を見る）", expanded=False):
                    # 1. 画像に紐づく情報（merge_csv_data結果）
                    st.markdown("**画像の出版社・シリーズ情報**")
                    img_info_df = pd.DataFrame([
                        {'コミックNo': img['comic_no'], 'ジャンル': img.get('genre', ''), '出版社': img.get('publisher', ''), 'シリーズ': img.get('series', ''), 'タイトル': img.get('title', '')}
                        for img in images
                    ])
                    st.dataframe(img_info_df, use_container_width=True, height=150)

                    # 2. フォルダ階層リスト
                    if st.session_state.workflow_data.get('hierarchy_xlsx'):
                        st.markdown("**フォルダ階層リスト**")
                        try:
                            h_df = pd.read_excel(BytesIO(st.session_state.workflow_data['hierarchy_xlsx']), sheet_name="フォルダ階層リスト", header=None)
                            col_names = ['ジャンル', '出版社', 'シリーズ', 'メインフォルダ', 'サブフォルダ'] + [f'列{i}' for i in range(6, 20)]
                            h_df.columns = col_names[:len(h_df.columns)]
                            st.dataframe(h_df, use_container_width=True, height=200)
                        except Exception as e:
                            st.error(f"表示エラー: {e}")

                    # 3. 楽天フォルダ一覧
                    if st.session_state.workflow_data.get('rakuten_folders'):
                        st.markdown("**楽天フォルダ一覧（R-Cabinet API）**")
                        f_df = pd.DataFrame(st.session_state.workflow_data['rakuten_folders'])
                        st.dataframe(f_df, use_container_width=True, height=200)

                # キュー生成
                st.divider()
                hierarchy_ready = st.session_state.workflow_data.get('hierarchy_xlsx')
                folders_ready = st.session_state.workflow_data.get('rakuten_folders')

                if st.button("🏪 アップロードキュー生成", type="primary", disabled=not (hierarchy_ready and folders_ready), key="rakuten_queue_btn"):
                    hierarchy_bytes = st.session_state.workflow_data['hierarchy_xlsx']
                    hierarchy_df = pd.read_excel(BytesIO(hierarchy_bytes), sheet_name="フォルダ階層リスト", header=None)
                    folders = st.session_state.workflow_data['rakuten_folders']

                    result = prepare_rakuten_queue(images, hierarchy_df, folders)
                    st.session_state.workflow_data['rakuten_queue'] = result
                    st.rerun()

                # キュー結果表示
                if 'rakuten_queue' in st.session_state.workflow_data:
                    result = st.session_state.workflow_data['rakuten_queue']

                    col_m1, col_m2 = st.columns(2)
                    col_m1.metric("マッピング成功", result['mapped'])
                    col_m2.metric("未マッチ", len(result['unmapped']))

                    if result['queue']:
                        queue_df = pd.DataFrame([
                            {'コミックNo': q['comic_no'], 'メインフォルダ': q.get('main_folder', ''), 'サブフォルダ': q['folder_name'], 'FolderId': q['folder_id']}
                            for q in result['queue']
                        ])
                        st.dataframe(queue_df, use_container_width=True, height=200)

                    with st.expander("ログ", expanded=False):
                        for log in result['logs']:
                            st.text(log)

            # ナビゲーション
            st.divider()
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("← 戻る"):
                    st.session_state.workflow_step = 3
                    st.rerun()
            with col2:
                yahoo_ready = 'yahoo_zips' in st.session_state.workflow_data
                rakuten_ready = 'rakuten_queue' in st.session_state.workflow_data
                if st.button("次へ進む →", type="primary", disabled=not (yahoo_ready or rakuten_ready)):
                    if 4 not in st.session_state.workflow_completed:
                        st.session_state.workflow_completed.append(4)
                    st.session_state.workflow_step = 5
                    st.rerun()

    # ============================================================
    # Step 5: API連携
    # ============================================================
    elif current_step == 5:
        st.markdown("""
        <div class="step-card">
            <div class="step-card-header">
                <div class="step-card-icon">🚀</div>
                <div>
                    <p class="step-card-title">Step ⑤ API連携</p>
                    <p class="step-card-desc">楽天・ヤフーにAPIで画像をアップロードします</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        tab_manual, tab_api = st.tabs(["📁 手動アップロード", "🤖 APIアップロード"])

        with tab_manual:
            st.markdown("### 手動アップロード用ZIP")
            st.caption("フォルダ構成どおりに画像を振り分けたZIPをダウンロードし、R-Cabinetに手動でアップロードします。")

            rakuten_queue = st.session_state.workflow_data.get('rakuten_queue')
            yahoo_zips = st.session_state.workflow_data.get('yahoo_zips')

            # 楽天ZIP生成
            st.markdown("#### 楽天（R-Cabinet）")
            if rakuten_queue and rakuten_queue.get('queue'):
                queue = rakuten_queue['queue']
                st.info(f"対象: {len(queue)}件（{len(rakuten_queue.get('unmapped', []))}件未マッチ）")

                if st.button("📦 楽天用ZIP生成", type="primary", key="rakuten_manual_zip"):
                    import zipfile
                    import openpyxl
                    zip_buffer = BytesIO()
                    mapping_rows = []
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for q in queue:
                            main_folder = q.get('main_folder', '')
                            sub_folder = q.get('folder_name', '')
                            file_name = q['file_name']
                            # sub_folderとmain_folderが同じ場合はサブフォルダなし
                            if sub_folder and sub_folder != main_folder:
                                zip_path = f"{main_folder}/{sub_folder}/{file_name}"
                            elif main_folder:
                                zip_path = f"{main_folder}/{file_name}"
                            else:
                                zip_path = f"{sub_folder}/{file_name}"
                            zf.writestr(zip_path, q['file_bytes'])
                            mapping_rows.append({
                                'コミックNo': q['comic_no'],
                                'メインフォルダ': main_folder,
                                'サブフォルダ': sub_folder if sub_folder != main_folder else '',
                                'ファイル名': file_name,
                                'ZIPパス': zip_path
                            })

                        # マッピング一覧Excelを生成してZIP直下に配置
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "振り分け一覧"
                        font = openpyxl.styles.Font(name='Meiryo UI')
                        headers = ['No', 'コミックNo', 'メインフォルダ', 'サブフォルダ', 'ファイル名', 'ZIPパス']
                        ws.append(headers)
                        for idx, row in enumerate(mapping_rows, 1):
                            ws.append([idx, row['コミックNo'], row['メインフォルダ'], row['サブフォルダ'], row['ファイル名'], row['ZIPパス']])
                        # フォント適用・列幅調整（全角文字は幅2として計算）
                        def calc_display_width(text):
                            return sum(2 if ord(c) > 0x7F else 1 for c in str(text))

                        for col_idx, header in enumerate(headers, 1):
                            max_w = calc_display_width(header)
                            for r in range(1, ws.max_row + 1):
                                cell = ws.cell(row=r, column=col_idx)
                                cell.font = font
                                max_w = max(max_w, calc_display_width(cell.value or ''))
                            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max_w + 2, 80)
                        xlsx_buffer = BytesIO()
                        wb.save(xlsx_buffer)
                        zf.writestr("振り分け一覧.xlsx", xlsx_buffer.getvalue())

                    st.session_state.workflow_data['rakuten_manual_zip'] = zip_buffer.getvalue()
                    st.rerun()

                if st.session_state.workflow_data.get('rakuten_manual_zip'):
                    zip_data = st.session_state.workflow_data['rakuten_manual_zip']
                    size_mb = len(zip_data) / (1024 * 1024)
                    st.download_button(
                        label=f"📥 rakuten_upload.zip ({size_mb:.1f}MB)",
                        data=zip_data,
                        file_name="rakuten_upload.zip",
                        mime="application/zip",
                        key="rakuten_manual_zip_dl"
                    )
            else:
                st.warning("Step ④でアップロードキューを生成してください。")

            # ヤフーZIP
            st.divider()
            st.markdown("#### ヤフー")
            if yahoo_zips and yahoo_zips.get('zips'):
                zips = yahoo_zips['zips']
                st.info(f"マッピング成功: {yahoo_zips['mapped']}件 / ZIP数: {len(zips)}")
                for i, zip_data in enumerate(zips):
                    size_mb = len(zip_data) / (1024 * 1024)
                    st.download_button(
                        label=f"📥 yahoo_upload_{i+1:03d}.zip ({size_mb:.1f}MB)",
                        data=zip_data,
                        file_name=f"yahoo_upload_{i+1:03d}.zip",
                        mime="application/zip",
                        key=f"yahoo_manual_zip_dl_{i}"
                    )
            else:
                st.warning("Step ④でヤフー用ZIPを生成してください。")

        with tab_api:
            st.markdown("### APIアップロード")
            st.info("🚧 この機能は現在実装中です。")

            tab_api_yahoo, tab_api_rakuten = st.tabs(["🛒 ヤフー", "🏪 楽天"])

            with tab_api_yahoo:
                st.warning("Yahoo認証トークンが未設定です。設定後にアップロードが可能になります。")

            with tab_api_rakuten:
                st.success("楽天API認証は設定済みです。実装予定。")

        st.divider()
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("← 戻る"):
                st.session_state.workflow_step = 4
                st.rerun()
        with col2:
            if st.button("ワークフローをリセット", type="secondary"):
                st.session_state.workflow_step = 1
                st.session_state.workflow_completed = []
                st.session_state.workflow_data = {}
                st.rerun()


elif mode == "📂 画像一覧取得":
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

    # フォルダ一覧を自動取得（初回のみ）
    if not st.session_state.folders_loaded:
        with st.spinner("フォルダ一覧を取得中..."):
            folders, error = get_all_folders()
        st.session_state.folders_data = folders
        st.session_state.folders_error = error
        st.session_state.folders_loaded = True
        st.rerun()

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

    # サイドバーにフォルダ情報（再取得ボタンは最新一覧を取得に統合したため不要）
    with st.sidebar:
        st.success(f"📁 {len(folders)} フォルダ")
        st.info(f"📷 {total_files} 画像（全体）")

    # DB統計情報を表示（常に表示）
    db_stats = get_db_stats()
    stat_cols = st.columns(4)
    with stat_cols[0]:
        st.metric("DB登録数", db_stats.get("total", 0))
    with stat_cols[1]:
        st.metric("重複ファイル", db_stats.get("duplicates", 0))
    with stat_cols[2]:
        st.metric("API総数", total_files)
    with stat_cols[3]:
        last_updated = st.session_state.get("last_sync_time") or "-"
        st.metric("最終更新", last_updated)

    # 操作ボタン（2つ）
    btn_col1, btn_col2, _ = st.columns([1.5, 1.5, 1])
    with btn_col1:
        show_db_btn = st.button(
            "📋 前回取得データを表示",
            help="前回APIから取得してDBに保存したデータを表示（高速）"
        )
    with btn_col2:
        fetch_api_btn = st.button(
            "🔄 最新一覧を取得",
            type="primary",
            help="APIから最新データを取得してDBに同期（フォルダ一覧も更新）"
        )

    st.divider()

    # フォルダ選択（絞り込み）
    folder_options = {"📁 すべて（全フォルダ）": None}
    folder_options.update({f"{f['FolderName']} ({f['FileCount']}件)": f for f in folders})

    selected_folder_name = st.selectbox(
        "フォルダで絞り込み",
        list(folder_options.keys()),
    )

    # ボタン押下時の処理
    if show_db_btn:
        st.session_state.images_loaded = False
        st.session_state.images_data = None

        if selected_folder_name == "📁 すべて（全フォルダ）":
            # 全フォルダ: DBから読み込み（高速）
            st.session_state.data_source = "db"
            loaded_images, msg = load_images_from_db()
            if loaded_images:
                st.session_state.images_data = loaded_images
                st.session_state.images_loaded = True
                st.session_state.error_folders = []
                st.success(f"📂 DBから{msg}")
            else:
                st.warning("DBにデータがありません")
        else:
            # 個別フォルダ: APIから直接取得（高速）
            st.session_state.data_source = "api"
            selected_folder = folder_options[selected_folder_name]
            folder_id = int(selected_folder['FolderId'])
            folder_name = selected_folder['FolderName']

            with st.spinner(f"「{folder_name}」を取得中..."):
                files, error = get_folder_files(folder_id)

            if error:
                st.error(error)
            elif files:
                for f in files:
                    f['FolderName'] = folder_name
                st.session_state.images_data = files
                st.session_state.images_loaded = True
                st.session_state.error_folders = []
                st.success(f"📂 APIから{len(files)}件を取得しました")
            else:
                st.warning("画像がありません")

    if fetch_api_btn:
        # APIから取得してDB同期
        st.session_state.data_source = "api"
        st.session_state.images_loaded = False
        st.session_state.images_data = None

        # フォルダ一覧も最新化
        with st.spinner("フォルダ一覧を更新中..."):
            new_folders, folder_error = get_all_folders()
        if not folder_error and new_folders:
            st.session_state.folders_data = new_folders
            folders = new_folders
            total_files = sum(f['FileCount'] for f in folders)

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
                st.session_state.last_sync_time = datetime.now(JST).strftime("%Y-%m-%d %H:%M")
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
                    st.session_state.last_sync_time = datetime.now(JST).strftime("%Y-%m-%d %H:%M")

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

            if results is None:
                st.error("DBにデータがありません。先に「📂 画像一覧取得」で「最新一覧を取得」を実行してください。")
            else:
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

        # --- 存在あり画像のダウンロード ---
        if exists_count > 0:
            exists_items = [r for r in results if r['存在'] == '✅ あり']
            exists_items_no_rec = [r for r in exists_items if 'REC' not in (r.get('フォルダ', '') or '').upper()]
            with st.expander(f"📦 存在あり画像をダウンロード（{len(exists_items_no_rec)}件、REC除外）"):
                if 'rcab_dl_result' not in st.session_state:
                    st.session_state.rcab_dl_result = None

                if st.button("🖼️ R-Cabinetから画像を取得", type="primary", key="rcab_dl_btn"):
                    _zipfile = get_zipfile()
                    progress = st.progress(0)
                    status = st.empty()
                    session = requests.Session()

                    downloaded = []
                    failed = []
                    for i, item in enumerate(exists_items_no_rec):
                        comic_no = str(item['コミックNo'])
                        url = item.get('URL', '')
                        folder = item.get('フォルダ', '')
                        file_name = item.get('ファイル名', '') or f"{comic_no}.jpg"
                        if '.' not in file_name:
                            file_name = f"{file_name}.jpg"

                        status.text(f"ダウンロード中: {comic_no} ({i+1}/{len(exists_items_no_rec)})")
                        progress.progress((i + 1) / len(exists_items_no_rec))

                        if not url or url == '-':
                            failed.append(comic_no)
                            continue
                        try:
                            resp = session.get(url, timeout=15)
                            if resp.status_code == 200 and len(resp.content) > 100:
                                downloaded.append({
                                    'comic_no': comic_no,
                                    'file_name': file_name,
                                    'folder': folder,
                                    'data': resp.content,
                                })
                            else:
                                failed.append(comic_no)
                        except Exception:
                            failed.append(comic_no)

                    progress.empty()
                    status.empty()
                    st.session_state.rcab_dl_result = {'downloaded': downloaded, 'failed': failed}
                    st.rerun()

                if st.session_state.rcab_dl_result:
                    dl_result = st.session_state.rcab_dl_result
                    downloaded = dl_result['downloaded']
                    failed = dl_result['failed']

                    st.success(f"取得完了: {len(downloaded)}件成功" + (f", {len(failed)}件失敗" if failed else ""))

                    if downloaded:
                        _zipfile = get_zipfile()
                        dl_cols = st.columns(2)

                        # フラット（直下）
                        with dl_cols[0]:
                            buf_flat = BytesIO()
                            with _zipfile.ZipFile(buf_flat, 'w', _zipfile.ZIP_DEFLATED) as zf:
                                for img in downloaded:
                                    zf.writestr(img['file_name'], img['data'])
                            buf_flat.seek(0)
                            st.download_button(
                                label=f"📦 フラットZIP（{len(downloaded)}件）",
                                data=buf_flat,
                                file_name="rcabinet_images_flat.zip",
                                mime="application/zip",
                                key="rcab_dl_flat"
                            )
                            st.caption("全画像を直下に配置")

                        # フォルダ構成保持
                        with dl_cols[1]:
                            buf_folder = BytesIO()
                            with _zipfile.ZipFile(buf_folder, 'w', _zipfile.ZIP_DEFLATED) as zf:
                                for img in downloaded:
                                    folder = img['folder'] if img['folder'] and img['folder'] != '-' else 'その他'
                                    zf.writestr(f"{folder}/{img['file_name']}", img['data'])
                            buf_folder.seek(0)
                            st.download_button(
                                label=f"📂 フォルダ付きZIP（{len(downloaded)}件）",
                                data=buf_folder,
                                file_name="rcabinet_images_by_folder.zip",
                                mime="application/zip",
                                key="rcab_dl_folder"
                            )
                            st.caption("R-Cabinetのフォルダ構成を保持")

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
        # セット品と単品を分離
        set_comics = [c for c in not_exists_comics if '_' not in str(c)]  # セット品（_なし）
        tanpin_comics = [c for c in not_exists_comics if '_' in str(c)]   # 単品（_あり）

        # 単品からベースのコミックNoを抽出（重複排除）
        tanpin_base_comics = list(set([str(c).split('_')[0] for c in tanpin_comics]))
        # セット品と単品ベースをマージ（重複排除）
        all_base_comics = list(set(set_comics + tanpin_base_comics))

        btn_col3, btn_col4, _ = st.columns([1.5, 1, 2])

        with btn_col3:
            # GitHubにアップロードボタン（セット品・単品を分けてアップロード）
            if set_comics or tanpin_comics:
                upload_label = "📤 GitHubにアップロード"
                upload_help = f"セット品: {len(set_comics)}件, 単品: {len(tanpin_comics)}件"
                if tanpin_base_comics:
                    # 単品のベースも追加される旨を表示
                    new_bases = [b for b in tanpin_base_comics if b not in set_comics]
                    if new_bases:
                        upload_help += f" (単品ベース追加: {len(new_bases)}件)"
                if st.button(upload_label, help=upload_help):
                    today = datetime.now(JST).strftime("%Y-%m-%d %H:%M")
                    upload_results = []

                    # セット品＋単品ベースをアップロード（missing_comics.csv）
                    if all_base_comics:
                        csv_lines = []
                        for comic_no in all_base_comics:
                            row = [''] * 9 + [str(comic_no), '1']
                            csv_lines.append(','.join(row))
                        csv_content = '\n'.join(csv_lines)

                        with st.spinner(f"コミックリスター用をアップロード中... ({len(all_base_comics)}件)"):
                            result = upload_to_github(
                                csv_content,
                                GITHUB_MISSING_CSV_PATH,
                                f"Update missing_comics.csv ({len(all_base_comics)}件) - {today}"
                            )
                        if result.get("success"):
                            upload_results.append(f"コミックリスター用: {len(all_base_comics)}件 ✅")
                        else:
                            upload_results.append(f"コミックリスター用: 失敗 ❌ {result.get('error')}")

                    # 単品をアップロード（missing_tanpin.csv）
                    if tanpin_comics:
                        # 単品CSV形式: コミックNo_巻数
                        tanpin_content = '\n'.join([str(c) for c in tanpin_comics])

                        with st.spinner(f"単品をアップロード中... ({len(tanpin_comics)}件)"):
                            result = upload_to_github(
                                tanpin_content,
                                GITHUB_MISSING_TANPIN_PATH,
                                f"Update missing_tanpin.csv ({len(tanpin_comics)}件) - {today}"
                            )
                        if result.get("success"):
                            upload_results.append(f"単品: {len(tanpin_comics)}件 ✅")
                        else:
                            upload_results.append(f"単品: 失敗 ❌ {result.get('error')}")

                    # 結果表示
                    if upload_results:
                        st.success("アップロード完了: " + ", ".join(upload_results))
            else:
                st.button("📤 GitHubにアップロード", disabled=True, help="存在なしのコミックNoがありません")

        with btn_col4:
            # 結果クリアボタン
            if st.button("🗑️ 結果をクリア"):
                st.session_state.check_results = None
                st.rerun()


elif mode == "🖼️ 新規画像取得":
    st.title("🖼️ 新規画像取得")
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

    st.markdown("### ステップ1: 必要なファイルをそろえよう")
    st.markdown("GitHub Actionsで生成されたファイルを取得します。")

    # まだセッションに読み込まれていないファイルがあれば自動ダウンロード
    need_is_list = not st.session_state.github_is_list
    need_comic_list = not st.session_state.github_comic_list
    need_hierarchy = not st.session_state.github_folder_hierarchy

    if need_is_list or need_comic_list or need_hierarchy:
        with st.spinner("GitHubからファイルを自動取得中..."):
            downloaded_any = False
            auto_errors = []

            if need_is_list:
                result = download_from_github(GITHUB_IS_LIST_PATH)
                if result.get("success"):
                    st.session_state.github_is_list = result["content"]
                    downloaded_any = True
                else:
                    auto_errors.append(f"is_list.csv: {result.get('error', '不明')}")

            if need_comic_list:
                result = download_from_github(GITHUB_COMIC_LIST_PATH)
                if result.get("success"):
                    st.session_state.github_comic_list = result["content"]
                    downloaded_any = True
                else:
                    auto_errors.append(f"comic_list.csv: {result.get('error', '不明')}")

            if need_hierarchy:
                result = download_from_github(GITHUB_FOLDER_HIERARCHY_PATH)
                if result.get("success"):
                    st.session_state.github_folder_hierarchy = result["content"]
                    downloaded_any = True
                else:
                    auto_errors.append(f"フォルダ階層リスト: {result.get('error', '不明')}")

            if auto_errors:
                st.warning(f"自動取得エラー: {', '.join(auto_errors)}")

        if downloaded_any:
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
                f"Update folder_hierarchy.xlsx - {datetime.now(JST).strftime('%Y-%m-%d %H:%M')}"
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

# ============================================================
# フォルダ一括作成
# ============================================================
elif mode == "📁 フォルダ一括作成":
    st.header("📁 フォルダ一括作成")
    st.markdown("R-Cabinetに複数のフォルダをまとめて作成します。パス形式で階層構造を指定でき、上位フォルダのIDは自動で解決されます。")

    st.divider()

    # CSV入力説明
    st.subheader("CSV入力")
    st.caption("カンマ区切りで「フォルダパス, ディレクトリ名」を入力")
    st.markdown("""
| 列 | 項目 | 必須 | 説明 |
|---|---|---|---|
| 1列目 | **フォルダパス** | ○ | `/`区切りで階層を表現。最下層がフォルダ名（最大50バイト） |
| 2列目 | ディレクトリ名 | | a-z, 0-9, -, _ のみ。最大20文字。省略時は自動採番 |

- 上位フォルダは**パスから自動判定**されます（IDの手動指定は不要）
- 上位フォルダが同じCSV内にあれば、作成後のIDを自動で引き継ぎます
- 上位フォルダがR-Cabinetに既存の場合も自動で検出します
""")

    # テンプレートCSVダウンロード
    template_csv = "フォルダパス,ディレクトリ名\nコミック,comic\nコミック/セット,set\nコミック/セット/セット1,set1\nコミック/セット/セット2,set2\nコミック/セット/セット3,set3\nコミック/単品,tanpin\nコミック/単品/単品1,tanpin1\nコミック/単品/単品2,tanpin2\nコミック/単品/単品3,tanpin3\nコミック/予約,yoyaku\nコミック/予約/予約1,yoyaku1\nコミック/予約/予約2,yoyaku2\nコミック/予約/予約3,yoyaku3"
    st.download_button(
        label="📥 テンプレートCSVをダウンロード",
        data=template_csv.encode('utf-8-sig'),
        file_name="folder_template.csv",
        mime="text/csv"
    )

    st.markdown("")

    csv_input = st.text_area(
        "CSV入力",
        height=250,
        placeholder="例:\nコミック,comic\nコミック/セット,set\nコミック/セット/セット1,set1\nコミック/セット/セット2,set2\nコミック/単品,tanpin\nコミック/単品/単品1,tanpin1",
        label_visibility="collapsed"
    )

    # CSVファイルアップロード
    uploaded_csv = st.file_uploader("またはCSVファイルをアップロード", type=['csv'])

    # 入力ソースの決定
    raw_lines = []
    if uploaded_csv:
        try:
            content = uploaded_csv.read().decode('utf-8')
        except UnicodeDecodeError:
            uploaded_csv.seek(0)
            content = uploaded_csv.read().decode('cp932')
        raw_lines = content.strip().splitlines()
    elif csv_input:
        raw_lines = csv_input.strip().splitlines()

    # ヘッダー行をスキップ
    if raw_lines and ("フォルダパス" in raw_lines[0] or "フォルダ名" in raw_lines[0]):
        raw_lines = raw_lines[1:]

    # パース
    import re as _re
    folder_entries = []
    for line in raw_lines:
        if not line.strip():
            continue
        parts = [p.strip() for p in line.split(",")]
        path = parts[0] if len(parts) > 0 else ""
        directory = parts[1] if len(parts) > 1 and parts[1] else None
        if path:
            # パスから階層を分解
            path_parts = [p.strip() for p in path.split("/") if p.strip()]
            folder_name = path_parts[-1]  # 最下層がフォルダ名
            parent_path = "/".join(path_parts[:-1]) if len(path_parts) > 1 else None
            folder_entries.append({
                "path": "/".join(path_parts),
                "name": folder_name,
                "directory": directory,
                "parent_path": parent_path
            })

    # プレビュー表示
    if folder_entries:
        st.divider()
        st.subheader(f"作成予定: {len(folder_entries)} フォルダ")

        preview_data = []
        for i, entry in enumerate(folder_entries, 1):
            # 階層の深さに応じてインデント
            depth = entry["path"].count("/")
            indent = "　" * depth
            preview_data.append({
                "No": i,
                "フォルダ構造": f"{indent}{entry['name']}",
                "パス": entry["path"],
                "ディレクトリ名": entry["directory"] or "（自動採番）",
                "上位フォルダ": entry["parent_path"] or "（ルート）"
            })

        st.dataframe(
            pd.DataFrame(preview_data),
            use_container_width=True,
            hide_index=True
        )

        # バリデーション
        errors = []
        path_set = {e["path"] for e in folder_entries}
        for i, entry in enumerate(folder_entries, 1):
            if len(entry["name"].encode('utf-8')) > 50:
                errors.append(f"No.{i}「{entry['name']}」: フォルダ名が50バイトを超えています")
            if entry["directory"]:
                if len(entry["directory"]) > 20:
                    errors.append(f"No.{i}「{entry['directory']}」: ディレクトリ名が20文字を超えています")
                elif not _re.fullmatch(r'[a-z0-9_-]+', entry["directory"]):
                    errors.append(f"No.{i}「{entry['directory']}」: ディレクトリ名に使用不可の文字（a-z, 0-9, -, _ のみ）")
            if entry["parent_path"] and entry["parent_path"] not in path_set:
                errors.append(f"No.{i}「{entry['path']}」: 上位フォルダ「{entry['parent_path']}」がCSV内に見つかりません")

        if errors:
            for err in errors:
                st.error(err)
        else:
            st.divider()
            if st.button("🚀 一括作成を実行", type="primary"):
                # 既存フォルダ一覧を取得（既存フォルダのID解決用）
                existing_folders, folder_list_error = get_all_folders()
                existing_name_to_id = {}
                if existing_folders:
                    for f in existing_folders:
                        existing_name_to_id[f['FolderName']] = f['FolderId']

                # パス→フォルダIDのマッピング（作成済みフォルダのID記録用）
                path_to_id = {}

                results = []
                progress = st.progress(0, text="フォルダ作成中...")
                total = len(folder_entries)

                for i, entry in enumerate(folder_entries):
                    progress.progress((i + 1) / total, text=f"作成中... ({i + 1}/{total}) {entry['name']}")

                    # 上位フォルダIDを解決
                    upper_folder_id = None
                    if entry["parent_path"]:
                        if entry["parent_path"] in path_to_id:
                            # 同じCSV内で先に作成済み
                            upper_folder_id = path_to_id[entry["parent_path"]]
                        else:
                            # 既存フォルダから検索（フォルダ名で一致）
                            parent_name = entry["parent_path"].split("/")[-1]
                            if parent_name in existing_name_to_id:
                                upper_folder_id = existing_name_to_id[parent_name]

                        if not upper_folder_id:
                            results.append({
                                "フォルダ構造": entry["path"],
                                "フォルダ名": entry["name"],
                                "ディレクトリ名": entry["directory"] or "（自動）",
                                "結果": "❌ スキップ",
                                "フォルダID": "",
                                "エラー": f"上位フォルダ「{entry['parent_path']}」のIDが見つかりません"
                            })
                            continue

                    result = create_folder(
                        folder_name=entry["name"],
                        directory_name=entry["directory"],
                        upper_folder_id=upper_folder_id
                    )

                    if result["success"]:
                        path_to_id[entry["path"]] = result["folder_id"]

                    results.append({
                        "フォルダ構造": entry["path"],
                        "フォルダ名": entry["name"],
                        "ディレクトリ名": entry["directory"] or "（自動）",
                        "結果": "✅ 成功" if result["success"] else "❌ 失敗",
                        "フォルダID": result.get("folder_id", ""),
                        "エラー": result.get("error", "")
                    })
                    if i < total - 1:
                        time.sleep(0.5)  # API負荷軽減

                progress.empty()

                # 結果表示
                success_count = sum(1 for r in results if r["結果"] == "✅ 成功")
                fail_count = total - success_count

                if fail_count == 0:
                    st.success(f"全 {total} フォルダの作成が完了しました！")
                elif success_count == 0:
                    st.error(f"全 {total} フォルダの作成に失敗しました")
                else:
                    st.warning(f"成功: {success_count} / 失敗: {fail_count}")

                st.dataframe(
                    pd.DataFrame(results),
                    use_container_width=True,
                    hide_index=True
                )

# ============================
# 📤 画像アップロード
# ============================
elif mode == "📤 画像アップロード":
    st.markdown("## 📤 画像アップロード")
    st.markdown("R-Cabinetのフォルダに画像をアップロードします。")

    # フォルダ一覧を取得
    with st.spinner("フォルダ一覧を取得中..."):
        folders, folder_error = get_all_folders()

    if folder_error:
        st.error(f"フォルダ一覧の取得に失敗しました: {folder_error}")
    elif not folders:
        st.warning("フォルダが見つかりません。先にフォルダを作成してください。")
    else:
        # フォルダ選択
        folder_options = {f"{f['FolderName']}（ID: {f['FolderId']}）": f for f in folders}
        selected_label = st.selectbox(
            "アップロード先フォルダ",
            options=list(folder_options.keys()),
            help="画像をアップロードするフォルダを選択してください"
        )
        selected_folder = folder_options[selected_label]

        # 画像ファイル選択
        uploaded_files = st.file_uploader(
            "画像ファイルを選択",
            type=["jpg", "jpeg", "png", "gif", "tiff", "bmp"],
            accept_multiple_files=True,
            help="対応形式: JPEG, PNG, GIF, TIFF, BMP（1ファイル2MBまで、最大3840x3840px）"
        )

        if uploaded_files:
            # バリデーション
            valid_files = []
            for f in uploaded_files:
                if f.size > 2 * 1024 * 1024:
                    st.warning(f"⚠️ {f.name}: ファイルサイズが2MBを超えています（{f.size / 1024 / 1024:.1f}MB）")
                else:
                    valid_files.append(f)

            if valid_files:
                st.info(f"📎 {len(valid_files)} ファイル選択済み")

                # プレビュー
                cols = st.columns(min(len(valid_files), 4))
                for i, f in enumerate(valid_files[:4]):
                    with cols[i]:
                        st.image(f, caption=f.name, width=150)
                        st.caption(f"{f.size / 1024:.0f} KB")
                if len(valid_files) > 4:
                    st.caption(f"他 {len(valid_files) - 4} ファイル...")

                # 上書きオプション
                overwrite = st.checkbox("同名ファイルが存在する場合は上書きする", value=False)

                # アップロード実行
                if st.button("📤 アップロード実行", type="primary"):
                    progress = st.progress(0, text="アップロード中...")
                    results = []
                    total = len(valid_files)

                    for i, f in enumerate(valid_files):
                        progress.progress((i + 1) / total, text=f"アップロード中... ({i + 1}/{total}) {f.name}")

                        # ファイル名をAPIの制限（50バイト）に合わせる
                        api_file_name = f.name
                        if len(api_file_name.encode('utf-8')) > 50:
                            # 拡張子を保持しつつ切り詰め
                            name_part, ext = api_file_name.rsplit('.', 1) if '.' in api_file_name else (api_file_name, '')
                            while len(f"{name_part}.{ext}".encode('utf-8')) > 50 and name_part:
                                name_part = name_part[:-1]
                            api_file_name = f"{name_part}.{ext}" if ext else name_part

                        f.seek(0)
                        result = upload_image(
                            file_data=f.read(),
                            file_name=api_file_name,
                            folder_id=selected_folder["FolderId"],
                            overwrite=overwrite,
                        )

                        results.append({
                            "ファイル名": f.name,
                            "結果": "✅ 成功" if result["success"] else "❌ 失敗",
                            "URL": result.get("file_url", ""),
                            "エラー": result.get("error", ""),
                        })

                        # API負荷軽減（3 req/sec制限）
                        if i < total - 1:
                            time.sleep(0.35)

                    progress.empty()

                    # 結果表示
                    success_count = sum(1 for r in results if r["結果"] == "✅ 成功")
                    fail_count = total - success_count

                    if fail_count == 0:
                        st.success(f"全 {total} ファイルのアップロードが完了しました！")
                    elif success_count == 0:
                        st.error(f"全 {total} ファイルのアップロードに失敗しました")
                    else:
                        st.warning(f"成功: {success_count} / 失敗: {fail_count}")

                    st.dataframe(
                        pd.DataFrame(results),
                        use_container_width=True,
                        hide_index=True
                    )
