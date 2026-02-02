# Yahoo!ショッピング カテゴリ抽出ツール

Yahoo!ショッピングのカテゴリページから、サブカテゴリの階層構造を抽出してExcelファイルに出力するツールです。

## バージョン

### Web版（Streamlit）- 推奨
URLにアクセスするだけで使える、ブラウザベースのツールです。

### デスクトップ版（tkinter）
ローカルPCで実行するGUIツールです。

## 機能

- 指定したカテゴリURLから子カテゴリを再帰的に取得
- 取得階層数の指定が可能（1〜10階層）
- 進捗状況をリアルタイム表示
- Excelファイルへの出力（階層別集計付き）
- Bot検出回避のためのランダム待機機能
- パスワード認証（Web版）

## 必要環境

- Python 3.8以上

## インストール

```bash
pip install -r requirements.txt
```

---

## Web版の使い方（Streamlit Cloud）

### Streamlit Cloudへのデプロイ

1. GitHubリポジトリにこのフォルダをプッシュ

2. [Streamlit Cloud](https://streamlit.io/cloud) にアクセスしてログイン

3. 「New app」をクリック

4. 設定:
   - Repository: `uraraka-axis/tools`
   - Branch: `main`
   - Main file path: `scraping/yahoo-shopping-category-extractor/streamlit_app.py`

5. 「Advanced settings」→「Secrets」にパスワードを設定:
   ```toml
   password = "your-secure-password"
   ```

6. 「Deploy!」をクリック

### 使い方

1. デプロイされたURLにアクセス
2. パスワードを入力してログイン
3. カテゴリURLを入力
4. 取得階層数を設定
5. 「抽出開始」をクリック
6. 完了後、「Excelファイルをダウンロード」をクリック

---

## デスクトップ版の使い方

```bash
python yahoo_shopping_category_extractor.py
```

1. カテゴリURLを入力（例: `https://shopping.yahoo.co.jp/category/2517/list`）
2. 取得階層数を設定
3. 出力フォルダを選択
4. 「抽出開始」をクリック

---

## 出力形式

Excelファイルに以下の情報が出力されます：

| # | ジャンル1 | ジャンル2 | ... | カテゴリID | ページURL |
|---|----------|----------|-----|-----------|----------|
| 1 | ルート   | サブ1    | ... | 12345     | https://... |

## 注意事項

- サーバーへの負荷を考慮し、リクエスト間に1.5〜4秒のランダム待機が入ります
- 大量のカテゴリを取得する場合は時間がかかります
- Yahoo!ショッピングの利用規約を遵守してご利用ください

## 更新履歴

- 2026-02-02: Streamlit版（Web版）を追加
- 2026-02-02: Next.js対応（__NEXT_DATA__からJSON抽出方式に変更）
- 2025-12-25: 初版リリース
