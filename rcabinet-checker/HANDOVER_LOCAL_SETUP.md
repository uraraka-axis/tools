# 引き継ぎ手順書：rcabinet-checker を新PCでローカル実行する

R-Cabinet 管理ツール（楽天・ヤフー向け商品画像の不足特定→取得→加工→出品を支援する **Streamlitアプリ**）を、
新しいPC上で **ローカル実行（`streamlit run`）** できるようにするための手順書。

**Claude Code を使ってセットアップする前提**で書いています。
新PCに Claude Code を入れたら、このフォルダで開いて

> 「HANDOVER_LOCAL_SETUP.md の通りに、このPCで rcabinet-checker をローカル実行できるようにセットアップして」

と頼めば、git導入・Python導入・クローン・依存インストール・起動まで案内・実行してもらえます。
**人間にしかできないのは「秘密ファイル `secrets.toml` の受け渡し」だけ**です（後述）。

---

## 0. このアプリの構成（先に把握）

| 項目 | 内容 |
|---|---|
| 種類 | Streamlit Webアプリ（ローカルのブラウザ `localhost:8501` で動く） |
| 本体 | `streamlit_app.py`（単一ファイル・約5,200行） |
| リポジトリ | `https://github.com/uraraka-axis/tools.git`（**モノレポ**。本アプリは `rcabinet-checker/` 配下） |
| 正規の作業場所 | `C:\Users\ssasa\tools\rcabinet-checker\`（git管理下） |
| 言語/版 | Python 3.12（現行PCは 3.12.10 で動作） |
| ログイン | 起動後、アプリ画面でパスワード入力（`secrets.toml` の `password`） |
| 外部連携 | 楽天RMS API / Supabase(DB) / Google Gemini / Google Sheets / Yahoo API / GitHub |

> ⚠ 旧クローン `C:\Users\ssasa\Documents\GitHub\tools_ARCHIVE_20260422_DO_NOT_USE\` は**参照禁止**（過去に重複クローンで未コミット欠損事故あり）。作業場所は必ず `C:\Users\ssasa\tools\` 側だけにする。

---

## 1. 新PCに入れるソフト

```powershell
# Git
winget install Git.Git
# Python 3.12
winget install Python.Python.3.12
# （保守するなら）VSCode + Claude Code
winget install Microsoft.VisualStudioCode
```
- ブラウザ（Chrome等）は表示用に1つあればよい。
- インストール後、PowerShellを開き直して `git --version` / `python --version` が出ることを確認。

---

## 2. ソースの入手（クローン）

```powershell
# 置き場所を現行PCと合わせる（推奨）
New-Item -ItemType Directory -Force "C:\Users\<ユーザー名>\tools" | Out-Null
cd "C:\Users\<ユーザー名>\tools"
git clone https://github.com/uraraka-axis/tools.git .
```
- クローンには **GitHubの認証**が必要（プライベートリポジトリの場合）。
  - 後任の人のGitHubアカウントが `uraraka-axis` org の `tools` にアクセスできる必要がある。
  - 認証は GitHub CLI（`winget install GitHub.cli` → `gh auth login`）か、Personal Access Token を使う。
- これで `...\tools\rcabinet-checker\` ができる。以降の作業はこのフォルダ内。

> 補足：クローンせず**フォルダごとコピー**しても動くが、その場合は git 履歴と最新性の管理が崩れるのでクローン推奨。

---

## 3. 秘密ファイルの受け渡し（★人間の作業・最重要）

`.streamlit\secrets.toml` は **`.gitignore` で除外されており、クローンには含まれない**。
**これが無いとアプリは起動できない。** 現行PCの

```
C:\Users\ssasa\tools\rcabinet-checker\.streamlit\secrets.toml
```

を、新PCの同じ場所（`rcabinet-checker\.streamlit\secrets.toml`）へ**手動でコピー**する。

- 渡し方：**USBか社内共有で直接**。メール/チャットに平文添付しない。
- 中身（鍵の一覧。値は現行ファイルに入っている）：

```toml
password = "..."              # アプリのログインパスワード
RMS_SERVICE_SECRET = "..."    # 楽天RMS API
RMS_LICENSE_KEY = "..."       # 楽天RMS API
SUPABASE_URL = "..."          # Supabase(DB)
SUPABASE_KEY = "..."          # Supabase(DB)
GITHUB_TOKEN = "..."          # GitHub（CSV書き込み/Actions連携）
GEMINI_API_KEY = "..."        # Google Gemini

[yahoo]                       # Yahoo API
client_id = "..."
client_secret = "..."
refresh_token = "..."
seller_id = "..."

[google]                      # Google Sheets（OAuth refresh_token方式）
client_id = "..."
client_secret = "..."
refresh_token = "..."
spreadsheet_id = "..."
```

> Google/Yahoo は **refresh_token 方式**なので、このファイルをコピーするだけで認証が引き継がれる（ブラウザ再同意は不要）。
> 運用者（アカウント）ごと変えたい場合のみ、各サービスでキー/トークンを再発行して該当値を差し替える。

---

## 4. 依存パッケージのインストール

`rcabinet-checker\` フォルダ内で：

```powershell
cd "C:\Users\<ユーザー名>\tools\rcabinet-checker"

# （任意・推奨）専用の仮想環境を作る
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 依存をインストール
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

requirements.txt の中身（参考）：streamlit / requests / pandas / openpyxl / beautifulsoup4 / supabase / google-generativeai / google-api-python-client / google-auth / google-auth-oauthlib / Pillow

---

## 5. 起動

```powershell
cd "C:\Users\<ユーザー名>\tools\rcabinet-checker"
# 仮想環境を使った場合は先に .\.venv\Scripts\Activate.ps1
streamlit run streamlit_app.py
# streamlit コマンドが見つからない時は:
python -m streamlit run streamlit_app.py
```
- 自動でブラウザが開き `http://localhost:8501` が表示される。
- 最初に**パスワード入力**（`secrets.toml` の `password`）→ アプリ本体へ。
- 起動確認だけなら、画面が出てログインできればOK。

毎回ダブルクリックで起動したい場合は、フォルダに `起動.bat` を作っておくと楽：
```bat
@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python -m streamlit run streamlit_app.py
```

---

## 6. 動作確認チェックリスト

- [ ] `streamlit run` でブラウザが開く
- [ ] パスワードでログインできる（secrets.toml が正しく置けている証拠）
- [ ] R-Cabinetのフォルダ一覧が取得できる（楽天RMSキーが有効）
- [ ] 画像存在チェックが動く（Supabase接続が有効）
- [ ] 必要に応じて Gemini / ヤフー / Google Sheets 連携機能を一通り触る

どれかでエラーが出たら → 大抵は `secrets.toml` の該当キー不足/失効か、ネットワーク。画面のエラーメッセージを Claude Code に見せて切り分けてもらう。

---

## 7. 引き継ぎ時に決める/伝えること

- [ ] **GitHubアクセス**：後任の人が `uraraka-axis/tools` をクローン/プッシュできる権限を持っているか
- [ ] **secrets.toml を安全に渡したか**（USB/社内共有・平文添付NG）
- [ ] 各サービス（楽天RMS・Supabase・Gemini・Yahoo・Google）を**現行のキーのまま使う**か、**後任のアカウントで再発行**するか
- [ ] アプリのログイン**パスワード**を伝えたか
- [ ] 置き場所のパス（`C:\Users\<新ユーザー名>\tools\rcabinet-checker`）に合わせて手順内のパスを読み替えること

---

## 8. 注意・既知の前提

- **作業ディレクトリは1つに統一**：過去に重複クローンで未コミット欠損事故あり。`tools` のクローンは新PCでも1つだけにする。
- このアプリは **GitHub Actions 連携**（`scripts/daily_sync.py` 等で Supabase 同期 / comic-lister 連携）も持つが、それはクラウド側（GitHub）で動くものなので**新PCのローカル実行とは別**。ローカルで動かすのは Streamlit アプリ本体。
- 本体は単一の巨大ファイル（`streamlit_app.py`）。改修時は Claude Code で開いて該当関数を探すのが早い。R-Cabinet API の `fileName`/`filePath` 仕様や 20バイト制限など、ハマりやすい仕様はコミット履歴とコメントに記録あり。
