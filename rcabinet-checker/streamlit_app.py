"""
R-Cabinet 画像存在チェックツール
コミックNoを入力してR-Cabinetに画像が存在するか確認する
"""

import streamlit as st
import requests
import base64
import xml.etree.ElementTree as ET
import pandas as pd
import time

# ページ設定
st.set_page_config(
    page_title="R-Cabinet 画像チェッカー",
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


@st.cache_data(ttl=300)  # 5分間キャッシュ
def get_all_folders():
    """R-Cabinetの全フォルダ一覧を取得"""
    url = f"{BASE_URL}/cabinet/folders/get"
    headers = get_auth_header()

    all_folders = []
    offset = 0
    limit = 100

    while True:
        params = {"offset": offset, "limit": limit}
        response = requests.get(url, headers=headers, params=params)

        if response.status_code != 200:
            st.error(f"フォルダ取得エラー: {response.status_code}")
            break

        root = ET.fromstring(response.text)
        folders = root.findall('.//folder')

        for folder in folders:
            all_folders.append({
                'FolderId': folder.findtext('FolderId', ''),
                'FolderName': folder.findtext('FolderName', ''),
                'FolderPath': folder.findtext('FolderPath', ''),
            })

        # 全件取得したかチェック
        folder_all_count = int(root.findtext('.//folderAllCount', '0'))
        if offset + limit >= folder_all_count:
            break
        offset += limit

    return all_folders


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

        # コミックNoで検索
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

        # APIレート制限対策（2-3リクエスト/秒）
        time.sleep(0.4)

    return results


# 認証情報チェック
if not SERVICE_SECRET or not LICENSE_KEY:
    st.error("⚠️ RMS API認証情報が設定されていません。Streamlit Secretsに設定してください。")
    st.code("""
# .streamlit/secrets.toml に以下を追加:
RMS_SERVICE_SECRET = "your_service_secret"
RMS_LICENSE_KEY = "your_license_key"
    """)
    st.stop()

# メインUI
st.title("🖼️ R-Cabinet 画像チェッカー")
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
        # 改行で分割し、空行を除去
        comic_numbers = [line.strip() for line in text_input.split('\n') if line.strip()]
        st.info(f"入力されたコミックNo: {len(comic_numbers)}件")

else:  # CSVアップロード
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

        # 列選択
        columns = df.columns.tolist()
        selected_column = st.selectbox(
            "コミックNo列を選択",
            columns,
            index=0
        )

        if selected_column:
            comic_numbers = df[selected_column].dropna().astype(str).tolist()
            st.info(f"読み込んだコミックNo: {len(comic_numbers)}件")

st.divider()

# チェック実行
if comic_numbers:
    col1, col2 = st.columns([1, 3])

    with col1:
        check_button = st.button("🔍 チェック実行", type="primary", use_container_width=True)

    if check_button:
        st.markdown("### チェック結果")

        progress_bar = st.progress(0)
        status_text = st.empty()

        with st.spinner("R-Cabinet APIに問い合わせ中..."):
            results = check_comic_images(comic_numbers, progress_bar, status_text)

        progress_bar.empty()
        status_text.empty()

        if results:
            df_results = pd.DataFrame(results)

            # サマリー表示
            exists_count = len([r for r in results if r['存在'] == '✅ あり'])
            not_exists_count = len([r for r in results if r['存在'] == '❌ なし'])

            col1, col2, col3 = st.columns(3)
            col1.metric("総数", len(comic_numbers))
            col2.metric("存在あり", exists_count)
            col3.metric("存在なし", not_exists_count)

            st.divider()

            # フィルター
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

            # CSVダウンロード
            csv_data = df_results.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 結果をCSVでダウンロード",
                data=csv_data,
                file_name="rcabinet_check_result.csv",
                mime="text/csv"
            )

else:
    st.warning("コミックNoを入力またはCSVをアップロードしてください。")

# サイドバー：フォルダ一覧
with st.sidebar:
    st.markdown("### R-Cabinet情報")

    if st.button("📂 フォルダ一覧を取得"):
        with st.spinner("取得中..."):
            folders = get_all_folders()

        if folders:
            st.success(f"フォルダ数: {len(folders)}")

            for folder in folders[:20]:
                st.markdown(f"- **{folder['FolderName']}** (`{folder['FolderPath']}`)")

            if len(folders) > 20:
                st.markdown(f"... 他 {len(folders) - 20} フォルダ")
