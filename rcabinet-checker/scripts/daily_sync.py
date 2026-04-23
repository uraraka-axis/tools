"""
R-Cabinet 日次同期スクリプト
GitHub Actionsで毎日自動実行され、R-Cabinet APIからデータを取得してSupabaseに保存します。
"""

import os
import sys
import base64
import json
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
import requests
from supabase import create_client, Client

# 環境変数から認証情報を取得
SERVICE_SECRET = os.environ.get("RMS_SERVICE_SECRET", "")
LICENSE_KEY = os.environ.get("RMS_LICENSE_KEY", "")
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")

BASE_URL = "https://api.rms.rakuten.co.jp/es/1.0"

# 同期対象のルートパス prefix（この配下のフォルダ・サブフォルダのみ同期対象）
# 3階層目の連番フォルダ（例: comic-set-set10, dvdblu-dvdblu2 等）は自動で含まれる
ALLOWED_ROOT_PREFIXES = [
    "/comic",
    "/dvdblu",
    "/toy",
    "/calenda",
    "/st",
    "/bk",
    "/kagu",
]


def is_target_folder(folder_path: str) -> bool:
    """folder_path が同期対象（ALLOWED_ROOT_PREFIXES 配下）かどうか判定"""
    if not folder_path:
        return False
    for prefix in ALLOWED_ROOT_PREFIXES:
        if folder_path == prefix or folder_path.startswith(prefix + '/'):
            return True
    return False


def get_auth_header():
    """ESA認証ヘッダーを生成"""
    auth_string = f"ESA {base64.b64encode(f'{SERVICE_SECRET}:{LICENSE_KEY}'.encode()).decode()}"
    return {"Authorization": auth_string}


def get_all_folders():
    """全フォルダ一覧を取得"""
    all_folders = []
    page = 1

    while True:
        url = f"{BASE_URL}/cabinet/folders/get"
        params = {"limit": 100, "offset": page}

        print(f"  API Request: {url} (page={page})")
        response = requests.get(url, headers=get_auth_header(), params=params)
        print(f"  Response status: {response.status_code}")

        if response.status_code != 200:
            print(f"Error: {response.status_code}")
            print(f"Response: {response.text[:500]}")
            return all_folders

        root = ET.fromstring(response.content)
        status = root.find(".//resultCode")
        print(f"  Result code: {status.text if status is not None else 'None'}")
        # R-Cabinet APIでは resultCode=0 が成功
        if status is None or status.text not in ["0", "N000"]:
            print(f"  API Error: {status.text if status else 'None'}")
            break

        folders = root.findall(".//folder")
        if not folders:
            break

        for folder in folders:
            folder_data = {
                'FolderId': folder.findtext('FolderId', ''),
                'FolderName': folder.findtext('FolderName', ''),
                'FolderPath': folder.findtext('FolderPath', ''),
                'FileCount': int(folder.findtext('FileCount', '0'))
            }
            if not is_target_folder(folder_data['FolderPath']):
                continue
            all_folders.append(folder_data)

        if len(folders) < 100:
            break

        page += 1
        time.sleep(0.3)

    return all_folders


def get_folder_files(folder_id: int):
    """フォルダ内のファイル一覧を取得"""
    all_files = []
    offset = 1

    while True:
        url = f"{BASE_URL}/cabinet/folder/files/get"
        params = {"folderId": folder_id, "limit": 100, "offset": offset}

        response = requests.get(url, headers=get_auth_header(), params=params)

        if response.status_code != 200:
            return all_files

        root = ET.fromstring(response.content)
        status = root.find(".//resultCode")
        # R-Cabinet APIでは resultCode=0 が成功
        if status is None or status.text not in ["0", "N000"]:
            break

        files = root.findall(".//file")
        if not files:
            break

        for file in files:
            # FileSizeは小数点を含む場合があるのでfloatで処理
            file_size_str = file.findtext('FileSize', '0')
            try:
                file_size = int(float(file_size_str))
            except ValueError:
                file_size = 0

            file_data = {
                'FileName': file.findtext('FileName', ''),
                'FileUrl': file.findtext('FileUrl', ''),
                'FileSize': file_size,
                'TimeStamp': file.findtext('TimeStamp', '')
            }
            all_files.append(file_data)

        if len(files) < 100:
            break

        offset += 1
        time.sleep(0.3)

    return all_files


def fetch_all_from_supabase(supabase: Client, table: str, columns: str = "*") -> list:
    """Supabaseから全件取得（ページネーション対応）"""
    all_data = []
    page_size = 1000
    offset = 0

    while True:
        response = supabase.table(table).select(columns).range(offset, offset + page_size - 1).execute()

        if not response.data:
            break

        all_data.extend(response.data)

        if len(response.data) < page_size:
            break

        offset += page_size

    return all_data


def sync_images_to_db(supabase: Client, images: list) -> dict:
    """画像一覧をDBに同期（upsert）"""
    try:
        # file_nameごとにグループ化（重複検出）
        file_dict = {}
        for img in images:
            file_name = img.get("FileName", "")
            folder_name = img.get("FolderName", "")
            folder_path = img.get("FolderPath", "")
            if file_name in file_dict:
                existing = file_dict[file_name]
                existing_folders = existing["folder_names"].split(", ")
                if folder_name not in existing_folders:
                    existing["folder_names"] += f", {folder_name}"
                try:
                    path_map = json.loads(existing["folder_path"]) if existing["folder_path"] else {}
                except (json.JSONDecodeError, TypeError):
                    path_map = {}
                path_map[folder_name] = folder_path
                existing["folder_path"] = json.dumps(path_map, ensure_ascii=False)
            else:
                file_dict[file_name] = {
                    "file_name": file_name,
                    "folder_names": folder_name,
                    "folder_path": json.dumps({folder_name: folder_path}, ensure_ascii=False),
                    "file_url": img.get("FileUrl", ""),
                    "file_size": img.get("FileSize", 0),
                    "file_timestamp": img.get("TimeStamp", "")
                }

        # 既存データを取得（ページネーション対応）
        existing_data = fetch_all_from_supabase(supabase, "rcabinet_images", "file_name, file_timestamp, folder_path")
        existing_dict = {row["file_name"]: (row.get("file_timestamp"), row.get("folder_path")) for row in existing_data}

        # 差分計算
        new_count = 0
        updated_count = 0
        duplicate_count = 0

        records_to_upsert = []
        for file_name, record in file_dict.items():
            if ", " in record["folder_names"]:
                duplicate_count += 1

            if file_name not in existing_dict:
                new_count += 1
                records_to_upsert.append(record)
            else:
                existing_ts, existing_path = existing_dict[file_name]
                # timestamp変化 or folder_path未設定（初回マイグレーション）で更新
                if existing_ts != record["file_timestamp"] or not existing_path:
                    updated_count += 1
                    records_to_upsert.append(record)

        # 削除済み検出
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
            "deleted": deleted_count,
            "total": len(file_dict)
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def update_sync_meta(supabase: Client, source: str, total_files: int = 0, is_full_sync: bool = False) -> bool:
    """rcabinet_sync_meta に最終同期時刻を記録（id=1 固定の単一行）"""
    try:
        now_iso = datetime.now(timezone.utc).isoformat()
        payload = {
            "id": 1,
            "last_sync_at": now_iso,
            "source": source,
            "total_files": total_files,
        }
        if is_full_sync:
            payload["last_full_sync_at"] = now_iso
        supabase.table("rcabinet_sync_meta").upsert(payload, on_conflict="id").execute()
        return True
    except Exception as e:
        print(f"  Warning: failed to update rcabinet_sync_meta: {e}")
        return False


def main():
    print("=" * 50)
    print("R-Cabinet Daily Sync Started")
    print("=" * 50)

    # 環境変数チェック（デバッグ用）
    print(f"SERVICE_SECRET set: {bool(SERVICE_SECRET)} (len={len(SERVICE_SECRET) if SERVICE_SECRET else 0})")
    print(f"LICENSE_KEY set: {bool(LICENSE_KEY)} (len={len(LICENSE_KEY) if LICENSE_KEY else 0})")
    print(f"SUPABASE_URL set: {bool(SUPABASE_URL)}")
    print(f"SUPABASE_KEY set: {bool(SUPABASE_KEY)} (len={len(SUPABASE_KEY) if SUPABASE_KEY else 0})")

    if not all([SERVICE_SECRET, LICENSE_KEY, SUPABASE_URL, SUPABASE_KEY]):
        print("Error: Missing environment variables")
        sys.exit(1)

    # Supabaseクライアント初期化
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    print("Supabase connected")

    # フォルダ一覧取得
    print("Fetching folders...")
    print(f"Allowed root prefixes: {ALLOWED_ROOT_PREFIXES}")
    folders = get_all_folders()
    print(f"Found {len(folders)} target folders (filtered)")

    if not folders:
        print("No folders found. Exiting.")
        sys.exit(1)

    # 全ファイル取得
    print("Fetching files from all folders...")
    all_files = []
    for i, folder in enumerate(folders):
        print(f"  [{i+1}/{len(folders)}] {folder['FolderName']} ({folder['FileCount']} files)")
        files = get_folder_files(int(folder['FolderId']))
        for f in files:
            f['FolderName'] = folder['FolderName']
            f['FolderPath'] = folder.get('FolderPath', '')
        all_files.extend(files)
        time.sleep(0.5)

    print(f"Total files fetched: {len(all_files)}")

    # DB同期
    print("Syncing to database...")
    result = sync_images_to_db(supabase, all_files)

    if result.get("success"):
        # 最終同期時刻をメタテーブルに記録（日次バッチは常にフル取得なので is_full_sync=True）
        meta_ok = update_sync_meta(
            supabase,
            source="daily_batch",
            total_files=result.get("total", 0),
            is_full_sync=True,
        )
        print("=" * 50)
        print("Sync completed successfully!")
        print(f"  Total: {result['total']}")
        print(f"  New: {result['new']}")
        print(f"  Updated: {result['updated']}")
        print(f"  Duplicate: {result['duplicate']}")
        print(f"  Deleted: {result['deleted']}")
        print(f"  Meta updated: {meta_ok}")
        print("=" * 50)
    else:
        print(f"Sync failed: {result.get('error')}")
        sys.exit(1)


if __name__ == "__main__":
    main()
