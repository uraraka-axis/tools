# manga-trend-radar

法人・店舗向けコミックレンタル「Smart Comic」のオウンドメディア
**『速報！マンガニュース』** の記事ネタを、毎朝（平日）自動で収集・選抜し
**Google Chat** に通知するツール。

```
収集(RSS) → 重複除去 → 自社目線スコアリング → 上位5件をChat通知 → 履歴保存
```

通知を見て採用する番号を返信 → Claude（Claude Code）が
**タイトル / ディスクリプション / URLスラッグ / 本文** を執筆する運用。

---

## 1. セットアップ

### 1-1. 依存インストール
```powershell
cd C:\Users\ssasa\tools\manga-trend-radar
pip install -r requirements.txt
```

### 1-2. .env を作成
`.env.example` を `.env` にコピーして値を入れる。

| 変数 | 必須 | 内容 |
|------|------|------|
| `GOOGLE_CHAT_WEBHOOK` | ✅ | Google Chat の Incoming Webhook URL |
| `ANTHROPIC_API_KEY`  | 任意 | 設定すると自社目線スコアリングがLLM化され精度UP。未設定でもキーワード辞書で動作 |

**Google Chat Webhook の取得**：通知したいスペース →「アプリと統合」→「Webhook」→ 作成 → URLをコピー。

### 1-3. 動作確認（Chat送信なし）
```powershell
python main.py --dry-run
```
土日でも実行され、選抜結果がコンソールに表示される。

---

## 2. 毎朝・平日の自動実行（タスクスケジューラ）

`run.bat` を平日 朝8:00 に実行するタスクを登録する（PowerShellで一度だけ）：

```powershell
schtasks /Create /TN "MangaTrendRadar" /SC WEEKLY /D MON,TUE,WED,THU,FRI ^
  /ST 08:00 /TR "C:\Users\ssasa\tools\manga-trend-radar\run.bat" /F
```

- 削除：`schtasks /Delete /TN "MangaTrendRadar" /F`
- 今すぐテスト実行：`schtasks /Run /TN "MangaTrendRadar"`
- 時刻変更：`/ST 09:30` のように指定して再登録（`/F`で上書き）

> 平日判定はスクリプト側（`WEEKDAYS_ONLY`）でも行うため、タスクを毎日実行に
> しても土日は自動スキップされる。

---

## 3. カスタマイズ（config.py）

情報源は2系統：
- **DIRECT_FEEDS** … 専門ニュースの直RSS（コミックナタリー/アニメ！アニメ！/ねとらぼ/はてブ アニメとゲーム/はてブ エンタメ）
  - はてブ2本は X/SNS発のバズ（広告ジャック・考察合戦等）を拾う用（2026-07-07追加）
- **AXIS_QUERIES** … Googleニュース検索RSSを「軸ごとのクエリ」で束ねて多軸収集（15軸）
  - 例: 新刊 / アニメ化 / 実写・映画化 / ドラマ化 / 配信(Netflix等) / 展覧会・原画展 /
    イベント・コラボ / 著名人×漫画 / 受賞・ランキング / 海外ヒット / 完結 / 聖地・地域 /
    SNS・ネット話題 / 広告ジャック・街頭 / 考察・反響（後ろ3軸は2026-07-07追加）

| 設定 | 既定 | 説明 |
|------|------|------|
| `AXIS_QUERIES` | 15軸 | **軸を足す/減らすのはここ**。`{axis, query, weight}` を1行追記。queryの`when:Nd`で期間指定 |
| `DIRECT_FEEDS` | 5本 | 専門ニュースの直RSS＋はてブ ホットエントリー2本 |
| `CANDIDATE_COUNT` | 20 | 1回に通知する候補数（上限超過時は自動でメッセージ分割） |
| `PER_AXIS_CAP` | 3 | 1軸あたり最大件数（多様性確保） |
| `COLLECT_HOURS` | 72 | 直RSSの対象時間（Gニュースは query の when: で制御） |
| `MAX_AGE_DAYS` | 14 | **全ソース共通の上限**。これより古い記事は日付があれば捨てる（鮮度確保） |
| `RECENCY_PENALTY` | 0.18 | 1日古いごとにスコアから引く量。大きいほど新しい記事を強く優先 |
| `POOL_LIMIT` | 130 | スコアリングに渡す候補プール上限 |
| `WEEKDAYS_ONLY` | True | 平日のみ通知 |
| `USE_LLM` | True | False で完全ヒューリスティック |
| `LLM_MODEL` | claude-haiku-4-5 | 精度を上げたいなら sonnet 系へ |
| `TYPE_KEYWORDS` / `FIT_KEYWORDS` / `NOISE_KEYWORDS` | — | スコアリングの語彙。自社目線の調整はここ |

**軸を1つ足す例**（config.py の AXIS_QUERIES に追記）：
```python
{"axis": "ゲーム化", "weight": 1.0, "query": "漫画原作 ゲーム化 when:30d"},
```

新しい直RSSを足すときは、まず単体で取得できるか確認：
```powershell
python -c "import feedparser,config; f=feedparser.parse('URL',agent=config.USER_AGENT); print(len(f.entries))"
```
（多くの日本のニュースサイトは User-Agent 無しだとブロックする点に注意）

> 補足: Googleニュースのリンクは `news.google.com/rss/articles/...` という転送URLです。
> クリックすると実際の記事へ飛びます（実URLへの変換は仕様上できないため転送URLのまま）。

---

## 4. ファイル構成

```
manga-trend-radar/
├── main.py        … オーケストレーション（これを実行）
├── collector.py   … RSS収集
├── scorer.py      … 自社目線スコアリング（LLM＋ヒューリスティック）
├── notifier.py    … Google Chat通知
├── storage.py     … 通知履歴（重複防止）
├── config.py      … 情報源・スコアリング語彙・各種設定
├── run.bat        … タスクスケジューラ用ランチャ
├── .env(.example) … 認証情報
└── history.json   … 自動生成（通知済み記録）
```

## 5. 鮮度（いつの記事を拾うか）

- 直RSS … 直近 `COLLECT_HOURS`（既定72時間）
- Googleニュース軸 … query の `when:Nd`（7〜14日）で取得
- **全ソース共通で `MAX_AGE_DAYS`（既定14日）より古い記事は破棄**
- さらにスコアリングで `RECENCY_PENALTY` により新しい記事を優先
- 「もっと新しいものだけ」にしたい場合は `MAX_AGE_DAYS` を 7 などへ下げる

## 6. 運用フロー

1. 毎朝（平日）Chatに候補20件が届く（上限超過時は自動分割）
2. 採用する番号をClaude Codeに伝える（例：「1,3,7で記事書いて」）
3. Claudeがタイトル/ディスクリプション/スラッグ/本文を執筆
4. 既存記事との重複は `history.json` で自動回避（45日保持）
