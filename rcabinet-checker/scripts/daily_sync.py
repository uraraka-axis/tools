"""
R-Cabinet 日次同期スクリプト
GitHub Actionsで毎日自動実行され、R-Cabinet APIからデータを取得してSupabaseに保存します。
"""

import os
import sys
import base64
import time
import xml.etree.ElementTree as ET
import requests
from supabase import create_client, Client

# 環境変数から認証情報を取得
SERVICE_SECRET = os.environ.get("RMS_SERVICE_SECRET", "")
LICENSE_KEY = os.environ.get("RMS_LICENSE_KEY", "")
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")

BASE_URL = "https://api.rms.rakuten.co.jp/es/1.0"


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
                'FileCount': int(folder.findtext('FileCount', '0'))
            }
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
            file_data = {
                'FileName': file.findtext('FileName', ''),
                'FileUrl': file.findtext('FileUrl', ''),
                'FileSize': int(file.findtext('FileSize', '0')),
                'TimeStamp': file.findtext('TimeStamp', '')
            }
            all_files.append(file_data)

        if len(files) < 100:
            break

        offset += 1
        time.sleep(0.3)

    return all_files


def sync_images_to_db(supabase: Client, images: list) -> dict:
    """画像一覧をDBに同期（upsert）"""
    try:
        # file_nameごとにグループ化（重複検出）
        file_dict = {}
        for img in images:
            file_name = img.get("FileName", "")
            folder_name = img.get("FolderName", "")
            if file_name in file_dict:
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

        # 既存データを取得
        existing = supabase.table("rcabinet_images").select("file_name, file_timestamp").execute()
        existing_dict = {row["file_name"]: row["file_timestamp"] for row in existing.data}

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
            elif existing_dict[file_name] != record["file_timestamp"]:
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
    folders = get_all_folders()
    print(f"Found {len(folders)} folders")

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
        all_files.extend(files)
        time.sleep(0.5)

    print(f"Total files fetched: {len(all_files)}")

    # DB同期
    print("Syncing to database...")
    result = sync_images_to_db(supabase, all_files)

    if result.get("success"):
        print("=" * 50)
        print("Sync completed successfully!")
        print(f"  Total: {result['total']}")
        print(f"  New: {result['new']}")
        print(f"  Updated: {result['updated']}")
        print(f"  Duplicate: {result['duplicate']}")
        print(f"  Deleted: {result['deleted']}")
        print("=" * 50)
    else:
        print(f"Sync failed: {result.get('error')}")
        sys.exit(1)


if __name__ == "__main__":
    main()
