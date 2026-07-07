# -*- coding: utf-8 -*-
"""
manga-trend-radar メイン処理

  収集 → 重複除去 → スコアリング → Google Chat 通知 → 履歴保存

毎朝（平日）タスクスケジューラから実行する想定。
手動確認は  python main.py --dry-run （Chat送信なし、内容だけ表示）。
手動送信は  python main.py --force   （曜日ガードを無視して送信）。

実行ログは log.txt に追記される（スケジューラ実行の調査用）。
"""
import sys
import time
from datetime import datetime
from pathlib import Path

# コンソールがcp932でも記事タイトルの特殊文字で落ちないようにする
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# .env を読み込む（python-dotenv があれば）
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import config
import collector
import scorer
import notifier
import storage
import cost_tracker

LOG_FILE = Path(__file__).parent / "log.txt"


def log(msg: str):
    """標準出力と log.txt の両方に、時刻付きで記録する。"""
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def collect_with_retry(tries: int = 4, wait: int = 30) -> list:
    """収集が0件（=起動直後などでネット未接続の可能性）ならリトライする。"""
    for i in range(1, tries + 1):
        items = collector.collect()
        if items:
            return items
        if i < tries:
            log(f"[warn] 収集0件（{i}/{tries}）。ネット接続待ちで{wait}秒後に再試行")
            time.sleep(wait)
    return []


def main():
    dry_run = "--dry-run" in sys.argv
    force = "--force" in sys.argv  # 手動実行：曜日に関係なく実行＆送信

    log(f"===== 実行開始 (dry_run={dry_run}, force={force}) =====")

    # 平日判定（月=0 ... 日=6）。--dry-run / --force なら曜日ガードを無視
    if config.WEEKDAYS_ONLY and datetime.now().weekday() >= 5 and not (dry_run or force):
        log("[info] 土日のためスキップ（WEEKDAYS_ONLY=True）")
        return

    # 1. 収集（0件ならリトライ）
    items = collect_with_retry()
    if not items:
        log("[error] 収集結果が0件（リトライしても取得できず）。ネットワーク不通の可能性。通知せず終了")
        return
    log(f"[info] 収集 {len(items)} 件")

    # 2. 重複除去（過去に通知済みのものを除外）
    history = storage.load_history()
    fresh = [it for it in items if not storage.is_seen(it, history)]
    log(f"[info] 新規候補 {len(fresh)} 件（既出 {len(items) - len(fresh)} 件を除外）")
    if not fresh:
        log("[info] 新規ネタなし。通知しません。")
        return

    # 3. スコアリング → 上位を選抜
    candidates = scorer.rank(fresh)
    if not candidates:
        log("[warn] 候補を選抜できませんでした。")
        return
    log(f"[info] 選抜 {len(candidates)} 件")

    # 3.5 LLM使用時はコストを記録（dry-runでもAPIは呼ばれるため必ず記録）
    cost_footer = None
    if scorer.last_usage:
        info = cost_tracker.record(scorer.last_usage)
        cost_footer = cost_tracker.build_footer(info)
        log(f"[info] {cost_footer.splitlines()[0]}")
        if info["alert"]:
            log(f"[warn] LLMコスト月累計 ¥{info['month_jpy']:.1f} がアラート閾値超過")

    # 4. 通知
    if dry_run:
        log("[info] DRY RUN（Chat送信なし）")
        today = datetime.now().strftime("%Y/%m/%d (%a)")
        for i, msg in enumerate(notifier.build_messages(candidates, today, cost_footer), 1):
            print(f"---- メッセージ {i} ----\n{msg}\n")
        return

    ok = notifier.notify(candidates, cost_footer)
    if ok:
        # 5. 送信成功時のみ履歴を保存（失敗時は次回再送させる）
        now_iso = datetime.now().isoformat()
        history = storage.mark_seen(candidates, history, now_iso)
        history = storage.prune(history, config.HISTORY_DAYS)
        storage.save_history(history)
        log("[info] 通知成功・履歴更新。実行完了")
    else:
        log("[error] Google Chat 通知に失敗。履歴は更新せず（次回再送します）")


if __name__ == "__main__":
    main()
