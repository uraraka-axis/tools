#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
画像振り分け・依頼リスト更新ツール（Streamlit Cloud版）
Google Drive上のファイルを操作し、依頼リストを更新する
"""

import streamlit as st
import pandas as pd
import re
import io
import time
from datetime import datetime
from typing import Optional, List, Dict, Tuple
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# ページ設定
st.set_page_config(
    page_title="画像振り分け・依頼リスト更新ツール",
    page_icon="🖼️",
    layout="wide"
)


def check_password():
    """パスワード認証"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.title("🔐 ログイン")
    password = st.text_input("パスワードを入力してください", type="password")

    if st.button("ログイン", type="primary"):
        if password == st.secrets.get("password", ""):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません")

    return False


# 認証チェック
if not check_password():
    st.stop()

# 定数
RAKUTEN_RMS_SHEET = "Rakuten RMS"
IRAI_BUN_SHEET = "依頼分"
COMIC_DB_SHEET = "コミック画像DB一覧"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


def extract_file_id(url_or_id: str) -> str:
    """GoogleドライブのURLまたはファイルIDからファイルIDを抽出"""
    url_or_id = url_or_id.strip()

    # 純粋なIDの場合
    if re.match(r'^[a-zA-Z0-9_-]+$', url_or_id):
        return url_or_id

    # URL形式の場合
    patterns = [
        r'/spreadsheets/d/([a-zA-Z0-9_-]+)',
        r'/file/d/([a-zA-Z0-9_-]+)',
        r'/folders/([a-zA-Z0-9_-]+)',
        r'id=([a-zA-Z0-9_-]+)',
    ]

    for pattern in patterns:
        match = re.search(pattern, url_or_id)
        if match:
            return match.group(1)

    raise ValueError(f"有効なGoogle DriveのURLまたはIDではありません: {url_or_id}")


@st.cache_resource
def get_google_services():
    """Google APIサービスを取得（キャッシュ）"""
    try:
        credentials_info = st.secrets["gcp_service_account"]
        credentials = service_account.Credentials.from_service_account_info(
            credentials_info,
            scopes=SCOPES
        )
        sheets_service = build('sheets', 'v4', credentials=credentials)
        drive_service = build('drive', 'v3', credentials=credentials)
        return sheets_service, drive_service
    except Exception as e:
        st.error(f"Google API接続エラー: {e}")
        return None, None


def log_message(message: str, log_container):
    """ログメッセージを追加"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    if "logs" not in st.session_state:
        st.session_state.logs = []
    st.session_state.logs.append(f"[{timestamp}] {message}")
    log_container.code("\n".join(st.session_state.logs), language="")


def get_sheet_data(sheets_service, file_id: str, sheet_name: str, range_notation: str) -> List[List]:
    """シートからデータを取得"""
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=file_id,
            range=f"'{sheet_name}'!{range_notation}"
        ).execute()
        return result.get('values', [])
    except Exception as e:
        return []


def get_sheet_id(sheets_service, file_id: str, sheet_name: str) -> Optional[int]:
    """シート名からシートIDを取得"""
    try:
        spreadsheet = sheets_service.spreadsheets().get(
            spreadsheetId=file_id
        ).execute()

        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        return None
    except Exception:
        return None


def download_file_from_drive(drive_service, file_id: str) -> bytes:
    """Google Driveからファイルをダウンロード"""
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    buffer.seek(0)
    return buffer.read()


def get_input_data(drive_service, file_id: str, log_container) -> pd.DataFrame:
    """Google DriveのExcel/CSVファイルからデータ取得"""
    log_message("📊 入力ファイル読み込み中...", log_container)

    try:
        # ファイル情報取得
        file_info = drive_service.files().get(fileId=file_id, fields='name,mimeType', supportsAllDrives=True).execute()
        file_name = file_info.get('name', '')
        mime_type = file_info.get('mimeType', '')

        log_message(f"   ファイル名: {file_name}", log_container)

        # ファイルダウンロード
        file_content = download_file_from_drive(drive_service, file_id)

        # ファイル形式に応じて読み込み
        if 'spreadsheet' in mime_type or file_name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(io.BytesIO(file_content), header=None, engine='openpyxl')
        elif 'csv' in mime_type or file_name.endswith('.csv'):
            try:
                df = pd.read_csv(io.BytesIO(file_content), header=None, encoding='utf-8')
            except:
                df = pd.read_csv(io.BytesIO(file_content), header=None, encoding='cp932')
        else:
            # Google Sheets形式の場合はエクスポート
            request = drive_service.files().export_media(
                fileId=file_id,
                mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            buffer = io.BytesIO()
            downloader = MediaIoBaseDownload(buffer, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            buffer.seek(0)
            df = pd.read_excel(buffer, header=None, engine='openpyxl')

        log_message(f"✅ 読み込み完了: {len(df)}行 × {len(df.columns)}列", log_container)
        return df

    except Exception as e:
        raise Exception(f"入力ファイル読み込みエラー: {e}")


def parse_input_file(df: pd.DataFrame, log_container) -> List[Dict]:
    """入力ファイルを解析"""
    data_list = []

    for i in range(1, len(df)):
        try:
            comic_no = str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else ''

            if not comic_no or comic_no == 'nan' or comic_no == '':
                break

            main_folder = str(df.iloc[i, 10]).strip() if pd.notna(df.iloc[i, 10]) else ''
            sub_folder = str(df.iloc[i, 11]).strip() if pd.notna(df.iloc[i, 11]) else ''

            if main_folder == 'nan':
                main_folder = ''
            if sub_folder == 'nan':
                sub_folder = ''

            data_list.append({
                'comic_no': comic_no,
                'main_folder': main_folder,
                'sub_folder': sub_folder
            })

        except Exception:
            continue

    log_message(f"📖 データ解析完了: {len(data_list)}件", log_container)
    return data_list


def list_files_in_folder(drive_service, folder_id: str) -> Dict[str, Dict]:
    """フォルダ内のファイル一覧を取得"""
    files = {}
    page_token = None

    while True:
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            spaces='drive',
            fields='nextPageToken, files(id, name, mimeType)',
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        for file in response.get('files', []):
            files[file['name']] = {
                'id': file['id'],
                'mimeType': file['mimeType']
            }

        page_token = response.get('nextPageToken')
        if not page_token:
            break

    return files


def get_parent_folder_id(drive_service, folder_id: str) -> Optional[str]:
    """フォルダの親フォルダIDを取得"""
    try:
        file_info = drive_service.files().get(
            fileId=folder_id,
            fields='parents',
            supportsAllDrives=True
        ).execute()
        parents = file_info.get('parents', [])
        return parents[0] if parents else None
    except Exception:
        return None


def find_or_create_folder(drive_service, parent_id: str, folder_name: str) -> str:
    """フォルダを検索、なければ作成"""
    # 既存フォルダを検索
    response = drive_service.files().list(
        q=f"'{parent_id}' in parents and name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        spaces='drive',
        fields='files(id, name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()

    files = response.get('files', [])
    if files:
        return files[0]['id']

    # フォルダを作成
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = drive_service.files().create(
        body=file_metadata,
        fields='id',
        supportsAllDrives=True
    ).execute()

    return folder.get('id')


def copy_images(drive_service, data_list: List[Dict], input_folder_id: str,
                output_folder_id: str, log_container, progress_bar) -> Tuple[Dict, List[str]]:
    """画像ファイルをGoogle Drive内で振り分けコピー"""
    log_message("", log_container)
    log_message("=" * 50, log_container)
    log_message("🖼️ 画像ファイル振り分け", log_container)
    log_message("=" * 50, log_container)

    stats = {
        'total': len(data_list),
        'success': 0,
        'not_found': 0,
        'failed': 0
    }

    success_comic_nos = []

    log_message(f"📋 処理対象: {len(data_list)}件", log_container)

    # 入力フォルダのファイル一覧を取得
    log_message("📁 入力フォルダのファイル一覧を取得中...", log_container)
    input_files = list_files_in_folder(drive_service, input_folder_id)
    log_message(f"   {len(input_files)}ファイル発見", log_container)

    # フォルダキャッシュ
    folder_cache = {}

    for i, data in enumerate(data_list, 1):
        comic_no = data['comic_no']
        main_folder = data['main_folder']
        sub_folder = data['sub_folder']

        # 画像ファイルを探す
        source_file = None
        source_name = None
        for ext in ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']:
            candidate_name = f"{comic_no}{ext}"
            if candidate_name in input_files:
                source_file = input_files[candidate_name]
                source_name = candidate_name
                break

        if not source_file:
            stats['not_found'] += 1
            progress_bar.progress(i / stats['total'])
            continue

        # 出力先フォルダを取得/作成
        try:
            if main_folder:
                cache_key = f"{main_folder}/{sub_folder}" if sub_folder else main_folder

                if cache_key not in folder_cache:
                    main_folder_id = find_or_create_folder(drive_service, output_folder_id, main_folder)

                    if sub_folder:
                        dest_folder_id = find_or_create_folder(drive_service, main_folder_id, sub_folder)
                    else:
                        dest_folder_id = main_folder_id

                    folder_cache[cache_key] = dest_folder_id

                dest_folder_id = folder_cache[cache_key]
            else:
                dest_folder_id = output_folder_id

            # ファイルをコピー
            file_metadata = {
                'name': source_name,
                'parents': [dest_folder_id]
            }

            drive_service.files().copy(
                fileId=source_file['id'],
                body=file_metadata,
                supportsAllDrives=True
            ).execute()

            stats['success'] += 1
            success_comic_nos.append(comic_no)

            # 10件ごとに進捗表示
            if stats['success'] % 10 == 0:
                log_message(f"   📦 {stats['success']}件コピー完了...", log_container)

        except Exception as e:
            log_message(f"   ❌ {comic_no}: コピー失敗 - {e}", log_container)
            stats['failed'] += 1

        progress_bar.progress(i / stats['total'])

    log_message(f"✅ 完了: 成功={stats['success']}, 未発見={stats['not_found']}, 失敗={stats['failed']}", log_container)
    return stats, success_comic_nos


def find_insert_position(sheet_data: List[List], main_folder: str, sub_folder: str) -> int:
    """同じフォルダ構成の最後の行を特定"""
    last_match_row = -1

    for i in range(3, len(sheet_data)):
        row = sheet_data[i]
        c_val = str(row[2]).strip() if len(row) > 2 and row[2] else ''
        d_val = str(row[3]).strip() if len(row) > 3 and row[3] else ''

        if c_val == main_folder and d_val == sub_folder:
            last_match_row = i

    return last_match_row


def update_rakuten_rms(sheets_service, file_id: str, data_list: List[Dict],
                       success_comic_nos: List[str], log_container, progress_bar) -> Dict:
    """Rakuten RMSシートに行を挿入"""
    log_message("", log_container)
    log_message("=" * 50, log_container)
    log_message("📝 Rakuten RMSシート更新開始", log_container)
    log_message("=" * 50, log_container)

    filtered_data_list = [data for data in data_list if data['comic_no'] in success_comic_nos]

    stats = {
        'total': len(filtered_data_list),
        'success': 0,
        'skipped': 0,
        'failed': 0,
        'duplicate': 0
    }

    try:
        sheet_data = get_sheet_data(sheets_service, file_id, RAKUTEN_RMS_SHEET, "A:F")

        if not sheet_data:
            log_message(f"⚠️ シート '{RAKUTEN_RMS_SHEET}' が見つかりません", log_container)
            stats['failed'] = len(filtered_data_list)
            return stats

        log_message(f"📊 対象シート: {RAKUTEN_RMS_SHEET}（{len(sheet_data)}行）", log_container)
        log_message(f"📋 処理対象: {len(filtered_data_list)}件", log_container)

        # 既存コミックNoをチェック
        log_message("🔍 既存コミックNoをチェック中...", log_container)
        existing_comic_nos = set()
        for row in sheet_data:
            if len(row) > 4 and row[4]:
                existing_comic_nos.add(str(row[4]).strip())

        log_message(f"   既存コミックNo: {len(existing_comic_nos)}件", log_container)

        sheet_id = get_sheet_id(sheets_service, file_id, RAKUTEN_RMS_SHEET)
        if sheet_id is None:
            raise Exception(f"シート '{RAKUTEN_RMS_SHEET}' のIDが取得できません")

        requests = []
        values_to_update = []

        prev_main_folder = None
        prev_sub_folder = None
        prev_insert_row = -1

        for i, data in enumerate(filtered_data_list, 1):
            comic_no = data['comic_no']
            main_folder = data['main_folder']
            sub_folder = data['sub_folder']

            if comic_no in existing_comic_nos:
                stats['duplicate'] += 1
                progress_bar.progress(i / stats['total']) if stats['total'] > 0 else None
                continue

            if main_folder == prev_main_folder and sub_folder == prev_sub_folder and prev_insert_row != -1:
                actual_insert_row = prev_insert_row + 1
            else:
                insert_row = find_insert_position(sheet_data, main_folder, sub_folder)

                if insert_row == -1:
                    log_message(f"   ⚠️ [{i}] {comic_no}: フォルダ構成が見つかりません", log_container)
                    stats['skipped'] += 1
                    progress_bar.progress(i / stats['total']) if stats['total'] > 0 else None
                    continue

                actual_insert_row = insert_row + 1

            try:
                requests.append({
                    'insertDimension': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'ROWS',
                            'startIndex': actual_insert_row,
                            'endIndex': actual_insert_row + 1
                        },
                        'inheritFromBefore': True
                    }
                })

                row_num = actual_insert_row + 1
                values_to_update.append({
                    'range': f"'{RAKUTEN_RMS_SHEET}'!B{row_num}:F{row_num}",
                    'values': [[
                        f'=ROW()-3',
                        main_folder,
                        sub_folder,
                        comic_no,
                        f"=VLOOKUP(E{row_num},'{COMIC_DB_SHEET}'!$C:$C,1,FALSE)"
                    ]]
                })

                existing_comic_nos.add(comic_no)
                prev_main_folder = main_folder
                prev_sub_folder = sub_folder
                prev_insert_row = actual_insert_row
                stats['success'] += 1

                # 10件ごとにバッチ実行
                if len(requests) >= 10:
                    execute_batch_update(sheets_service, file_id, requests, values_to_update)
                    log_message(f"   📝 {stats['success']}件挿入完了...", log_container)
                    requests = []
                    values_to_update = []

            except Exception as e:
                log_message(f"   ❌ [{i}] {comic_no}: 挿入失敗 - {e}", log_container)
                stats['failed'] += 1
                prev_main_folder = None
                prev_sub_folder = None
                prev_insert_row = -1

            progress_bar.progress(i / stats['total']) if stats['total'] > 0 else None

        # 残りのリクエストを実行
        if requests:
            execute_batch_update(sheets_service, file_id, requests, values_to_update)

        log_message(f"✅ 完了: 成功={stats['success']}, 重複={stats['duplicate']}, スキップ={stats['skipped']}, 失敗={stats['failed']}", log_container)
        return stats

    except Exception as e:
        log_message(f"❌ エラー: {e}", log_container)
        stats['failed'] = len(filtered_data_list) - stats['success']
        return stats


def execute_batch_update(sheets_service, file_id: str, requests: List[Dict], values_to_update: List[Dict]):
    """バッチ更新を実行"""
    if requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=file_id,
            body={'requests': requests}
        ).execute()
        time.sleep(0.5)

    if values_to_update:
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=file_id,
            body={
                'valueInputOption': 'USER_ENTERED',
                'data': values_to_update
            }
        ).execute()
        time.sleep(0.5)


def delete_processed_rows(sheets_service, file_id: str, success_comic_nos: List[str],
                          log_container) -> bool:
    """依頼分シートから処理済み行を削除"""
    log_message("", log_container)
    log_message("=" * 50, log_container)
    log_message("🗑️ 依頼分シート 処理済み行削除", log_container)
    log_message("=" * 50, log_container)

    try:
        sheet_data = get_sheet_data(sheets_service, file_id, IRAI_BUN_SHEET, "A:E")

        if not sheet_data:
            log_message(f"⚠️ シート '{IRAI_BUN_SHEET}' が見つかりません", log_container)
            return False

        success_set = set(success_comic_nos)

        # 削除対象行を特定
        end_row = -1
        for i in range(2, len(sheet_data)):
            row = sheet_data[i]
            if len(row) > 4 and row[4]:
                end_row = i
                log_message(f"   📍 削除範囲: 3～{end_row + 1}行目", log_container)
                break

        if end_row == -1:
            log_message("⚠️ E列にデータが見つかりませんでした", log_container)
            return False

        rows_to_delete = []

        for i in range(2, end_row + 1):
            row = sheet_data[i]

            if len(row) > 2 and row[2]:
                c_raw = row[2]

                if isinstance(c_raw, (int, float)):
                    c_val = str(int(c_raw))
                else:
                    c_val = str(c_raw).strip()
                    if '.' in c_val and c_val.replace('.', '').replace('-', '').isdigit():
                        try:
                            c_val = str(int(float(c_val)))
                        except:
                            pass

                if c_val and c_val in success_set:
                    rows_to_delete.append(i)

        if not rows_to_delete:
            log_message("⚠️ 削除対象の行がありません", log_container)
            return False

        log_message(f"🗑️ {len(rows_to_delete)}行を削除します", log_container)

        sheet_id = get_sheet_id(sheets_service, file_id, IRAI_BUN_SHEET)
        if sheet_id is None:
            raise Exception(f"シート '{IRAI_BUN_SHEET}' のIDが取得できません")

        requests = []
        for row_idx in reversed(rows_to_delete):
            requests.append({
                'deleteDimension': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': row_idx,
                        'endIndex': row_idx + 1
                    }
                }
            })

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=file_id,
            body={'requests': requests}
        ).execute()

        log_message(f"✅ {len(rows_to_delete)}行を削除しました", log_container)
        return True

    except Exception as e:
        log_message(f"❌ 行削除エラー: {e}", log_container)
        return False


def update_comic_db(sheets_service, file_id: str, data_list: List[Dict],
                    success_comic_nos: List[str], log_container, progress_bar) -> Dict:
    """コミック画像DB一覧にコミックNoを追加"""
    log_message("", log_container)
    log_message("=" * 50, log_container)
    log_message("📚 コミック画像DB一覧更新", log_container)
    log_message("=" * 50, log_container)

    filtered_data_list = [data for data in data_list if data['comic_no'] in success_comic_nos]

    stats = {
        'total': len(filtered_data_list),
        'success': 0,
        'failed': 0
    }

    try:
        sheet_data = get_sheet_data(sheets_service, file_id, COMIC_DB_SHEET, "A:C")

        if not sheet_data:
            log_message(f"⚠️ シート '{COMIC_DB_SHEET}' が見つかりません", log_container)
            stats['failed'] = len(filtered_data_list)
            return stats

        # 最終行を特定
        last_row = 2
        for i in range(len(sheet_data) - 1, -1, -1):
            row = sheet_data[i]
            if len(row) > 2 and row[2]:
                last_row = i
                break

        start_row = last_row + 2
        log_message(f"📍 追加開始位置: {start_row}行目（{len(filtered_data_list)}件追加）", log_container)

        values_to_add = []
        for i, data in enumerate(filtered_data_list):
            comic_no = data['comic_no']
            row_num = start_row + i

            values_to_add.append([
                f'=ROW()-2',
                comic_no
            ])

            stats['success'] += 1
            progress_bar.progress((i + 1) / stats['total']) if stats['total'] > 0 else None

        if values_to_add:
            range_notation = f"'{COMIC_DB_SHEET}'!B{start_row}:C{start_row + len(values_to_add) - 1}"

            sheets_service.spreadsheets().values().update(
                spreadsheetId=file_id,
                range=range_notation,
                valueInputOption='USER_ENTERED',
                body={'values': values_to_add}
            ).execute()

        log_message(f"✅ {stats['success']}件追加完了", log_container)

    except Exception as e:
        log_message(f"❌ 追加失敗: {e}", log_container)
        stats['failed'] = len(filtered_data_list) - stats['success']

    return stats


def init_session_state():
    """セッション状態の初期化"""
    if "saved_spreadsheet_url" not in st.session_state:
        st.session_state.saved_spreadsheet_url = ""
    if "saved_input_file_url" not in st.session_state:
        st.session_state.saved_input_file_url = ""
    if "saved_input_folder_url" not in st.session_state:
        st.session_state.saved_input_folder_url = ""


def main():
    st.title("🖼️ 画像振り分け・依頼リスト更新ツール")
    st.caption("Google Drive上のファイルを操作し、依頼リストを更新します")

    # セッション状態初期化
    init_session_state()

    # サービス取得
    sheets_service, drive_service = get_google_services()

    if sheets_service is None or drive_service is None:
        st.error("Google APIに接続できません。Secretsの設定を確認してください。")
        st.stop()

    st.success("✅ Google API接続完了")

    # 設定入力
    st.subheader("設定")

    spreadsheet_url = st.text_input(
        "依頼リストURL",
        value=st.session_state.saved_spreadsheet_url,
        placeholder="GoogleスプレッドシートのURL または ファイルID",
        help="Rakuten RMS、依頼分、コミック画像DB一覧シートを含むスプレッドシート"
    )

    input_file_url = st.text_input(
        "振り分けマップ（Excel/CSV）",
        value=st.session_state.saved_input_file_url,
        placeholder="Google DriveのファイルURL または ファイルID",
        help="E列=コミックNo, K列=メインフォルダ, L列=サブフォルダ"
    )

    input_folder_url = st.text_input(
        "入力画像フォルダ",
        value=st.session_state.saved_input_folder_url,
        placeholder="Google DriveのフォルダURL または フォルダID",
        help="コピー元の画像が格納されているフォルダ（同階層に出力フォルダを自動作成）"
    )

    st.info("📁 出力フォルダは入力フォルダと同じ階層に「出力_yyyymmdd」として自動作成されます")

    st.divider()

    # 処理実行
    if st.button("▶ 処理開始", type="primary", use_container_width=True):
        # バリデーション
        if not spreadsheet_url:
            st.error("依頼リストURLを入力してください")
            st.stop()

        if not input_file_url:
            st.error("振り分けマップを入力してください")
            st.stop()

        if not input_folder_url:
            st.error("入力画像フォルダを入力してください")
            st.stop()

        # 入力値を保存
        st.session_state.saved_spreadsheet_url = spreadsheet_url
        st.session_state.saved_input_file_url = input_file_url
        st.session_state.saved_input_folder_url = input_folder_url

        # ID抽出
        try:
            spreadsheet_id = extract_file_id(spreadsheet_url)
            input_file_id = extract_file_id(input_file_url)
            input_folder_id = extract_file_id(input_folder_url)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        # ログ初期化
        st.session_state.logs = []

        # 処理開始
        st.subheader("処理ログ")
        log_container = st.empty()
        progress_bar = st.progress(0)

        start_time = time.time()

        try:
            log_message("=" * 50, log_container)
            log_message("📥 処理開始", log_container)
            log_message("=" * 50, log_container)
            log_message(f"📋 スプレッドシートID: {spreadsheet_id}", log_container)

            # 入力ファイル読み込み
            input_df = get_input_data(drive_service, input_file_id, log_container)
            data_list = parse_input_file(input_df, log_container)

            if not data_list:
                st.error("入力データがありません")
                st.stop()

            # 出力フォルダを自動作成
            log_message("📁 出力フォルダを作成中...", log_container)
            parent_folder_id = get_parent_folder_id(drive_service, input_folder_id)
            if not parent_folder_id:
                st.error("入力フォルダの親フォルダが取得できません")
                st.stop()

            output_folder_name = f"出力_{datetime.now().strftime('%Y%m%d')}"
            output_folder_id = find_or_create_folder(drive_service, parent_folder_id, output_folder_name)
            log_message(f"   出力フォルダ: {output_folder_name}", log_container)

            # 画像コピー
            image_stats, success_comic_nos = copy_images(
                drive_service, data_list, input_folder_id, output_folder_id,
                log_container, progress_bar
            )

            if not success_comic_nos:
                log_message("⚠️ 画像コピー成功分がないため、依頼リスト編集をスキップします", log_container)
            else:
                # Rakuten RMS更新
                rms_stats = update_rakuten_rms(
                    sheets_service, spreadsheet_id, data_list, success_comic_nos,
                    log_container, progress_bar
                )

                # 依頼分シート行削除
                delete_result = delete_processed_rows(
                    sheets_service, spreadsheet_id, success_comic_nos, log_container
                )

                # コミック画像DB更新
                db_stats = update_comic_db(
                    sheets_service, spreadsheet_id, data_list, success_comic_nos,
                    log_container, progress_bar
                )

            processing_time = time.time() - start_time

            log_message("", log_container)
            log_message("=" * 50, log_container)
            log_message(f"✅ 全処理完了（{processing_time:.1f}秒）", log_container)
            log_message("=" * 50, log_container)

            progress_bar.progress(1.0)

            # 結果表示
            st.success("処理が完了しました！")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("画像コピー", f"{image_stats['success']}件")
            with col2:
                if success_comic_nos:
                    st.metric("RMS更新", f"{rms_stats['success']}件")
                else:
                    st.metric("RMS更新", "0件")
            with col3:
                if success_comic_nos:
                    st.metric("DB更新", f"{db_stats['success']}件")
                else:
                    st.metric("DB更新", "0件")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
            log_message(f"❌ エラー: {e}", log_container)


if __name__ == "__main__":
    main()
