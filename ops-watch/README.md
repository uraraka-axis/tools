# OpsWatch - 業務自動化タスクの死活監視

タスクスケジューラの終了コードだけでなく「ログの中身」まで毎朝検証し、
結果を Google Chat に1通で通知する。exit 0 でも中身が失敗しているケース
（ログにエラー・完了マーカーなし・ログ未生成）を検出するのが目的。

## 仕組み（タスクごとに3層の述語）

1. タスクが存在・有効・最終結果コードが正常（0 / 実行中）
2. 最終実行が max_age_days 以内（週末PCオフを考慮して既定4日）
3. 最新ログに error_patterns がなく、success_patterns がある
   （＋タスクは走ったのにログが更新されていない、も検出）

## 運用

- タスク名: **OpsWatch**（毎日 10:00、StartWhenAvailable付き）
- 通知: Google Chat（pending-watcher と同じスペース。webhook は config.json）
- 全部正常でも毎日1通送る。**通知が来ない日 = OpsWatch 自体が死んでいる**
- クラッシュ時も Chat に通知（最上位例外フック）

## 手動実行

```
cd C:\Users\ssasa\tools\ops-watch
python ops_watch.py --dry-run   # 通知なしで結果確認
python ops_watch.py             # 本実行（Chat通知あり）
```

## 監視対象の追加・変更

config.json の targets に1エントリ追加する:

```json
{
  "task": "タスクスケジューラ上の名前",
  "desc": "説明 実行時刻",
  "max_age_days": 4,
  "log_dir": "run_*.log 形式のログフォルダ",   // または "log_file": "単一ログ"
  "success_patterns": ["完了マーカー"],         // 省略可
  "error_patterns": ["..."]                     // 省略時は default_error_patterns
}
```

ログなしのタスク（Recoru等）は task/desc/max_age_days のみでよい。

## 復旧手順

- タスクが消えた場合:
  `schtasks /create /tn "OpsWatch" /tr "C:\Users\ssasa\tools\ops-watch\daily_run.bat" /sc daily /st 10:00 /f`
  のあと PowerShell で StartWhenAvailable を付与（ops-check スキル参照）
- 誤検知が続く場合: 該当ツールのログ形式が変わった可能性。config.json の
  success_patterns / error_patterns を実ログに合わせて更新する

## ファイル

- ops_watch.py … 本体（Python 3.12 / 標準ライブラリのみ）
- _query_tasks.ps1 … タスク情報をJSONで返す読み取り専用ヘルパー（UTF-8 BOM）
- config.json … 監視対象とwebhook
- daily_run.bat … スケジューラ起動用（CRLF+ASCII）
- logs\run_*.log … 実行ログ（ops_check.ps1 が読む形式）
