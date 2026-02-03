"""
R-Cabinet 管理ツール
- フォルダ画像一覧：R-Cabinetのフォルダ毎に画像を一覧表示
- 画像存在チェック：コミックNoを入力して存在確認
"""

import streamlit as st
import requests
import base64
import xml.etree.ElementTree as ET
import pandas as pd
import time

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
def get_folder_files(folder_id: int):
    """指定フォルダ内の画像一覧を取得"""
    url = f"{BASE_URL}/cabinet/folder/files/get"
    headers = get_auth_header()

    all_files = []
    offset = 1  # 1始まり（ページ番号）
    limit = 100  # APIの上限は100件

    while True:
        params = {"folderId": folder_id, "offset": offset, "limit": limit}

        try:
            response = requests.get(url, headers=headers, params=params, timeout=30)
        except requests.exceptions.RequestException as e:
            return None, f"接続エラー: {str(e)}"

        if response.status_code != 200:
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

        if found_files:
            for f in found_files:
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

    st.markdown("####")  # 間隔調整

    mode = st.radio(
        "機能を選択",
        ["📂 画像一覧取得", "🔍 画像存在チェック"],
        label_visibility="collapsed"
    )

    st.markdown("####")  # 間隔調整
    st.divider()


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

    # ステップ3: 画像取得ボタン
    fetch_images_btn = st.button("📷 画像一覧を取得", type="primary")

    st.divider()

    # 画像取得処理
    if fetch_images_btn or st.session_state.images_loaded:
        if fetch_images_btn:
            st.session_state.images_loaded = False
            st.session_state.images_data = None

        if selected_folder_name == "📁 すべて（全フォルダ）":
            # 全フォルダの画像を取得
            if not st.session_state.images_loaded or fetch_images_btn:
                all_files = []
                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, folder in enumerate(folders):
                    status_text.text(f"取得中: {folder['FolderName']} ({i + 1}/{len(folders)})")
                    progress_bar.progress((i + 1) / len(folders))

                    files, err = get_folder_files(int(folder['FolderId']))
                    if files:
                        for f in files:
                            f['FolderName'] = folder['FolderName']
                        all_files.extend(files)

                progress_bar.empty()
                status_text.empty()

                st.session_state.images_data = all_files
                st.session_state.images_loaded = True

            all_files = st.session_state.images_data

            if all_files:
                st.success(f"📷 {len(all_files)} 件の画像（全フォルダ）")

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

                # CSVダウンロード
                csv_data = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 全データをCSVでダウンロード",
                    data=csv_data,
                    file_name="rcabinet_all_files.csv",
                    mime="text/csv"
                )
            else:
                st.warning("画像がありません。")

        else:
            # 特定フォルダの画像を取得
            selected_folder = folder_options[selected_folder_name]
            folder_id = int(selected_folder['FolderId'])

            with st.spinner(f"「{selected_folder['FolderName']}」の画像を取得中..."):
                files, error = get_folder_files(folder_id)

            if error:
                st.error(error)
            elif files:
                st.success(f"📷 {len(files)} 件の画像")

                # 検索フィルター
                search_term = st.text_input("🔍 ファイル名で絞り込み", placeholder="検索キーワード")

                if search_term:
                    files = [f for f in files if search_term.lower() in f['FileName'].lower()]
                    st.info(f"絞り込み結果: {len(files)} 件")

                # データフレーム表示
                df = pd.DataFrame(files)
                df = df[['FileName', 'FileUrl', 'FileSize', 'TimeStamp']]
                df.columns = ['ファイル名', 'URL', 'サイズ(KB)', '更新日時']

                st.dataframe(df, use_container_width=True, height=500)

                # CSVダウンロード
                csv_data = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 CSVでダウンロード",
                    data=csv_data,
                    file_name=f"rcabinet_{selected_folder['FolderName']}.csv",
                    mime="text/csv"
                )
            else:
                st.warning("このフォルダに画像はありません。")


elif mode == "🔍 画像存在チェック":
    st.title("🔍 画像存在チェック")
    st.markdown("コミックNoを入力して、R-Cabinetに画像が存在するか確認します。")

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

    # チェック実行
    if comic_numbers:
        check_button = st.button("🔍 チェック実行", type="primary")

        if check_button:
            st.markdown("### チェック結果")

            progress_bar = st.progress(0)
            status_text = st.empty()

            results = check_comic_images(comic_numbers, progress_bar, status_text)

            progress_bar.empty()
            status_text.empty()

            if results:
                df_results = pd.DataFrame(results)

                exists_count = len([r for r in results if r['存在'] == '✅ あり'])
                not_exists_count = len([r for r in results if r['存在'] == '❌ なし'])

                col1, col2, col3 = st.columns(3)
                col1.metric("総数", len(comic_numbers))
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

                csv_data = df_results.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 結果をCSVでダウンロード",
                    data=csv_data,
                    file_name="rcabinet_check_result.csv",
                    mime="text/csv"
                )

    else:
        st.warning("コミックNoを入力またはCSVをアップロードしてください。")
