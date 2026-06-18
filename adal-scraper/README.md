# ADAL ONLINE SHOP 商品スクレイパー

[https://adal-online.shop/](https://adal-online.shop/) の全カテゴリ・全商品をスクレイピングし、Excel に出力するツールです。
品番が複数ある場合は **1品番 = 1行** に展開します。

## GUI版（おすすめ） ※Python不要

`dist\adal_scraper.exe` を **ダブルクリック** すると操作画面が起動します。

1. 取得したい **カテゴリにチェック**（既定は全選択。全選択/全解除ボタンあり）
2. 必要なら **出力Excelファイル**（未指定なら自動命名）・**各カテゴリ最大件数**（0=全件、テスト時に便利）・**リクエスト間隔**（既定0.7秒）を設定
3. **「取得開始」** をクリック → 進捗バーと処理ログが表示されます
4. 途中でやめたいときは **「停止」**（停止時点までの取得分でExcelを出力）

※初回起動はexe展開のため数秒かかります。SmartScreen警告が出た場合は「詳細情報」→「実行」で起動できます。

## コマンドライン（exe）でも実行可能

引数を付けて起動するとCLIとして動きます（自動化・バッチ向け）。

```powershell
cd C:\Users\ssasa\tools\adal-scraper\dist

.\adal_scraper.exe --categories others  # 「その他」のみ
.\adal_scraper.exe --limit 5            # 各カテゴリ先頭5商品（テスト）
.\adal_scraper.exe --out 結果.xlsx --delay 1.0
.\adal_scraper.exe --gui                # 明示的にGUI起動
```

## ソース版（Python）

### セットアップ

```powershell
pip install requests beautifulsoup4 openpyxl
```

### 使い方

```powershell
# 引数なし → GUI起動
python adal_scraper.py

# 全カテゴリ巡回（フル取得・CLI）
python adal_scraper.py --categories chair,sofa,bench,kids,outdoors,table,others,outlet-option

# 動作確認: 「その他」カテゴリのみ
python adal_scraper.py --categories others

# 複数カテゴリ指定
python adal_scraper.py --categories chair,sofa

# 各カテゴリ先頭5商品だけ取得（テスト）
python adal_scraper.py --limit 5

# 出力ファイル名・リクエスト間隔(秒)を指定
python adal_scraper.py --out 結果.xlsx --delay 1.0
```

出力ファイル名を省略すると `adal_products_YYYYMMDD_HHMMSS.xlsx` で保存されます。

## 対象カテゴリ

| 表示名 | スラッグ | 件数(目安) |
|---|---|---|
| チェア | chair | 319 |
| ソファ | sofa | 225 |
| ベンチ | bench | 97 |
| キッズ | kids | 32 |
| アウトドア | outdoors | 44 |
| テーブル | table | 97 |
| その他 | others | 4 |
| クリアランスセール | outlet-option | 160 |

※フル取得は商品数が多い（合計約1,000商品）ため、`--delay`（既定0.7秒）でサーバに配慮しつつ巡回します。所要時間の目安は10〜20分程度です。

## 出力列（A〜P）

| 列 | 項目 | 備考 |
|---|---|---|
| A | No. | 出力行の連番 |
| B | カタログNo. | 例: ADAL Vol.28（無い商品は空欄） |
| C | 商品名 | |
| D | 品番 | 複数ある場合は行を分割 |
| E | カラー | 品番に対応する色／バリエーション名 |
| F | 材質 | |
| G | 重量 | |
| H | サイズ | |
| I | 保証期間 | |
| J | 配送 | 「お渡し方法について」等のリンク文言は除去 |
| K | お届け目安 | |
| L | カテゴリ | 巡回したカテゴリ名（例: チェア） |
| M | URL | 商品詳細ページのURL |
| N | カタログ価格 | |
| O | 法人会員価格 | 未ログイン表示値（ランク5基準）「〜」付き |
| P | 商品説明 | |

## 仕様メモ

- サイトはサーバーレンダリングのため `requests` + `BeautifulSoup` で取得（ブラウザ不要）。
- 一覧ページは `?pageno=N` でページ送り。1商品=複数リンクのため詳細IDで重複除去。
- カテゴリ列はパンくず末尾（商品ごとにサブ分類が異なる）ではなく、**巡回中のカテゴリ名**を採用。
- 品番欄の2レイアウトに対応:
  - A) 「カラー名 型番」が同一行（例: `ウォームグレー P3002-10JEC`）
  - B) 「ラベル行 → 型番行」が交互（例: `左テーブル` / `X4017-99LX`）
- 法人会員価格はログイン会員ランクで変動。本ツールは未ログイン状態の表示値（ランク5基準）を取得します。

## 注意

- 取得した法人会員価格・カタログ価格は表示時点のものです。
- サイト構造が変わるとセレクタの調整が必要になる場合があります（`adal_scraper.py` 内のCSSセレクタを参照）。
</content>
