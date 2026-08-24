# -*- coding: utf-8 -*-
"""
ops_watch.py - 業務自動化タスクの死活監視（結果述語チェック + Google Chat通知）

タスクスケジューラの終了コードだけでなく「ログの中身」まで検証する。
  1. タスクが存在し有効で、最終結果コードが正常か
  2. 最終実行が新しいか（max_age_days 以内）
  3. 最新ログにエラーパターンがなく、完了マーカーがあるか
  4. タスクは走ったのにログが生成されていない、を検出

使い方:
  python ops_watch.py            # チェックして Chat 通知
  python ops_watch.py --dry-run  # 通知なし。結果を標準出力のみ
"""
import argparse
import datetime as dt
import glob
import json
import os
import re
import subprocess
import sys
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "config.json")
QUERY_PS1 = os.path.join(BASE, "_query_tasks.ps1")
LOG_DIR = os.path.join(BASE, "logs")

# 結果コードの意味（ops_check.ps1 と同じ）
RESULT_CODES = {
    0: None,
    267009: None,          # 0x41301 実行中
    267011: "NEVER_RUN",   # 0x41303 未実行
    1: "スクリプトが exit 1（.bat/Python側のエラー）",
    2147946720: "実行拒否 0x800710E0（未ログオン/スリープ中に時刻到来）",
    2147750687: "既に実行中のため開始せず 0x8004131F（ハングの可能性）",
}

_log_lines = []


def log(msg):
    line = "[%s] %s" % (dt.datetime.now().strftime("%H:%M:%S"), msg)
    _log_lines.append(line)
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode("cp932", errors="replace").decode("cp932"))


def flush_log():
    os.makedirs(LOG_DIR, exist_ok=True)
    path = os.path.join(LOG_DIR, "run_%s.log" % dt.datetime.now().strftime("%Y%m%d_%H%M%S"))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(_log_lines) + "\n")


def read_text(path):
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, "r", encoding="cp932", errors="replace") as f:
        return f.read()


def query_tasks():
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", QUERY_PS1],
        capture_output=True, timeout=120,
    )
    text = p.stdout.decode("utf-8", errors="replace").strip()
    if not text:
        raise RuntimeError("_query_tasks.ps1 の出力が空: " + p.stderr.decode("utf-8", errors="replace")[:300])
    data = json.loads(text)
    if isinstance(data, dict):
        data = [data]
    return {d["name"]: d for d in data}


def latest_log(target):
    """(path, mtime) を返す。対象ログがなければ (None, None)"""
    if target.get("log_file"):
        p = target["log_file"]
        if os.path.exists(p):
            return p, dt.datetime.fromtimestamp(os.path.getmtime(p))
        return None, None
    if target.get("log_dir"):
        files = glob.glob(os.path.join(target["log_dir"], "run_*.log"))
        if not files:
            return None, None
        p = max(files, key=os.path.getmtime)
        return p, dt.datetime.fromtimestamp(os.path.getmtime(p))
    return None, None


TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")


def recent_lines(lines, cutoff):
    """行頭タイムスタンプで cutoff 以降の行だけ残す（過去エラーの再警報防止）。
    タイムスタンプが1つもないログはそのまま返す（report.log 形式）。"""
    kept, in_window, found_ts = [], False, False
    for line in lines:
        m = TS_RE.match(line)
        if m:
            found_ts = True
            ts = dt.datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
            in_window = ts >= cutoff
        if in_window:
            kept.append(line)
    return kept if found_ts else lines


def check_log_content(target, defaults, now):
    """ログ内容の述語チェック。問題のリストを返す"""
    problems = []
    path, _ = latest_log(target)
    if path is None:
        problems.append("ログファイルが見つからない")
        return problems
    text = read_text(path)
    lines = text.splitlines()
    if target.get("log_file"):
        # 追記型: 完了マーカーは直近100行、エラー走査は直近25時間分に限定
        lines = lines[-100:]
        err_lines = recent_lines(lines, now - dt.timedelta(hours=25))
        text = "\n".join(lines)
    else:
        err_lines = lines
    err_patterns = target.get("error_patterns", defaults)
    for pat in err_patterns:
        for line in err_lines:
            if pat.lower() in line.lower():
                problems.append("ログにエラー: %s" % line.strip()[:120])
                break
        else:
            continue
        break  # 1件報告すれば十分
    for pat in target.get("success_patterns", []):
        if pat not in text:
            problems.append("完了マーカー「%s」がログにない" % pat)
    return problems


def check_target(target, info, defaults, now):
    """(status, details) status: OK / WARN / NG"""
    details = []
    if info is None or not info.get("found"):
        return "NG", ["タスクスケジューラに存在しない"]
    if info.get("state") == "Disabled":
        return "NG", ["タスクが無効化されている"]

    code = int(info.get("last_result", 0))
    code_meaning = RESULT_CODES.get(code, "要調査 (0x%08X)" % code)
    if code_meaning == "NEVER_RUN" or not info.get("last_run"):
        return "WARN", ["まだ一度も実行されていない"]
    if code_meaning is not None:
        details.append("結果コード %d: %s" % (code, code_meaning))

    last_run = dt.datetime.strptime(info["last_run"], "%Y-%m-%d %H:%M:%S")
    age_days = (now - last_run).total_seconds() / 86400.0
    if age_days > target.get("max_age_days", 4):
        details.append("最終実行が古い: %s（%.1f日前）" % (info["last_run"], age_days))

    if target.get("log_dir") or target.get("log_file"):
        _, log_mtime = latest_log(target)
        if log_mtime is not None and log_mtime < last_run - dt.timedelta(minutes=30):
            details.append("タスクは %s に走ったがログが更新されていない（最新ログ %s）"
                           % (info["last_run"], log_mtime.strftime("%m/%d %H:%M")))
        details.extend(check_log_content(target, defaults, now))

    return ("NG", details) if details else ("OK", [])


def send_chat(webhook, text):
    req = urllib.request.Request(
        webhook,
        data=json.dumps({"text": text}).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=UTF-8"},
    )
    with urllib.request.urlopen(req, timeout=30) as res:
        return res.status


def build_message(results, now):
    ok = [r for r in results if r[1] == "OK"]
    warn = [r for r in results if r[1] == "WARN"]
    ng = [r for r in results if r[1] == "NG"]
    head = "🩺 OpsWatch %s\n" % now.strftime("%Y-%m-%d %H:%M")
    head += "✅ 正常 %d / ⚠️ 注意 %d / 🚨 異常 %d\n" % (len(ok), len(warn), len(ng))
    lines = []
    for name, status, details, desc in ng:
        lines.append("🚨 %s（%s）" % (name, desc))
        for d in details:
            lines.append("　・" + d)
    for name, status, details, desc in warn:
        lines.append("⚠️ %s（%s）: %s" % (name, desc, "、".join(details)))
    if not lines:
        lines.append("全 %d タスク正常です。" % len(results))
    return head + "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="Chat通知せずログのみ")
    args = ap.parse_args()

    with open(CONFIG, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    defaults = cfg.get("default_error_patterns", ["[ERROR]", "Traceback"])
    now = dt.datetime.now()

    log("=== OpsWatch 開始 %s%s ===" % (now.strftime("%Y-%m-%d %H:%M"), " (dry-run)" if args.dry_run else ""))
    infos = query_tasks()

    results = []  # (task, status, details, desc)
    for t in cfg["targets"]:
        status, details = check_target(t, infos.get(t["task"]), defaults, now)
        results.append((t["task"], status, details, t.get("desc", "")))
        log("%-4s %s %s" % (status, t["task"], "; ".join(details) if details else ""))

    msg = build_message(results, now)
    if args.dry_run:
        log("--- dry-run: 通知はしない。以下が送信予定の本文 ---")
        for line in msg.splitlines():
            log("| " + line)
    else:
        code = send_chat(cfg["webhook_url"], msg)
        log("Chat通知 送信: HTTP %d" % code)

    log("=== 完了 ===")
    flush_log()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        log("クラッシュ: %r" % e)
        try:
            with open(CONFIG, "r", encoding="utf-8") as f:
                _cfg = json.load(f)
            if "--dry-run" not in sys.argv:
                send_chat(_cfg["webhook_url"], "🚨 OpsWatch 自体がクラッシュしました: %r" % e)
        except Exception:
            pass
        flush_log()
        sys.exit(1)
