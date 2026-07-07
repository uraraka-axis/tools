# -*- coding: utf-8 -*-
"""通知済み記事の履歴管理（重複通知の防止）。"""
import json
import hashlib
from datetime import datetime, timedelta
from pathlib import Path

HISTORY_FILE = Path(__file__).parent / "history.json"


def _key(item: dict) -> str:
    """URL優先、無ければタイトルでハッシュ化したキーを返す。"""
    base = (item.get("link") or item.get("title") or "").strip().lower()
    return hashlib.sha1(base.encode("utf-8")).hexdigest()


def load_history() -> dict:
    if HISTORY_FILE.exists():
        try:
            return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def is_seen(item: dict, history: dict) -> bool:
    return _key(item) in history


def mark_seen(items: list, history: dict, now_iso: str) -> dict:
    for it in items:
        history[_key(it)] = now_iso
    return history


def prune(history: dict, days: int) -> dict:
    """古い履歴を削除して肥大化を防ぐ。"""
    cutoff = datetime.now() - timedelta(days=days)
    kept = {}
    for k, v in history.items():
        try:
            if datetime.fromisoformat(v) >= cutoff:
                kept[k] = v
        except Exception:
            kept[k] = v
    return kept


def save_history(history: dict) -> None:
    HISTORY_FILE.write_text(
        json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8"
    )
