# -*- coding: utf-8 -*-
"""RSS/Atomフィードから直近のマンガ関連記事を収集する。"""
import time
from datetime import datetime, timedelta, timezone

import feedparser

import config


def _entry_datetime(entry) -> datetime | None:
    for attr in ("published_parsed", "updated_parsed"):
        t = entry.get(attr)
        if t:
            try:
                return datetime.fromtimestamp(time.mktime(t), tz=timezone.utc)
            except Exception:
                continue
    return None


def collect() -> list[dict]:
    """各ソースを巡回し、直近 COLLECT_HOURS 以内の記事を集めて返す。"""
    cutoff = datetime.now(timezone.utc) - timedelta(hours=config.COLLECT_HOURS)
    max_age_cutoff = datetime.now(timezone.utc) - timedelta(days=config.MAX_AGE_DAYS)
    items: list[dict] = []

    for src in config.SOURCES:
        try:
            feed = feedparser.parse(src["url"], agent=config.USER_AGENT)
        except Exception as e:
            print(f"[warn] {src['name']} の取得に失敗: {e}")
            continue

        if getattr(feed, "bozo", False) and not feed.entries:
            print(f"[warn] {src['name']} のフィードが空 or 解析エラー")
            continue

        # 直RSSは COLLECT_HOURS で絞る。Googleニュースは query側の when: に任せる
        do_filter = src.get("filter_recency", True)

        for e in feed.entries:
            dt = _entry_datetime(e)
            # 直RSSは COLLECT_HOURS、それ以外は MAX_AGE_DAYS を上限に古い記事を捨てる
            # （日付不明の記事はフィード次第で拾う）
            if dt is not None:
                if do_filter and dt < cutoff:
                    continue
                if dt < max_age_cutoff:
                    continue
            summary = (e.get("summary") or e.get("description") or "").strip()
            items.append({
                "title": (e.get("title") or "").strip(),
                "link": (e.get("link") or "").strip(),
                "summary": summary[:400],
                "source": src["name"],
                "axis": src.get("axis", ""),
                "weight": src["weight"],
                "published": dt.isoformat() if dt else "",
            })

    # 同一URLの重複を除去（フィード内重複対策）
    seen, uniq = set(), []
    for it in items:
        k = it["link"] or it["title"]
        if k in seen:
            continue
        seen.add(k)
        uniq.append(it)

    print(f"[info] 収集件数: {len(uniq)} 件（{len(config.SOURCES)} ソース）")
    return uniq
