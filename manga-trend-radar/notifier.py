# -*- coding: utf-8 -*-
"""Google Chat（Incoming Webhook）へ候補を通知する。

20件など多めでも、Google Chatの1メッセージ上限（約4096文字）に収まるよう
自動でメッセージを分割して順番に送信する。
"""
import os
import json
import time
import urllib.request
from datetime import datetime

import config

CHAT_LIMIT = 3800  # 1メッセージの安全上限（実上限4096に余裕を持たせる）


def _axis_rank(axis: str) -> int:
    try:
        return config.AXIS_ORDER.index(axis)
    except (ValueError, AttributeError):
        return len(getattr(config, "AXIS_ORDER", []))


def _item_block(idx: int, c: dict) -> str:
    kw = c.get("keywords", "")
    return (
        f"*{idx}. {c['proposed_title']}*\n"
        f"   軸: {c.get('type', c.get('axis','-'))} ｜ 優先度: {c.get('score','-')}\n"
        f"   理由: {c.get('reason','-')}\n"
        + (f"   想定KW: {kw}\n" if kw else "")
        + f"   出典: {c['source']}\n"
        f"   {c['link']}"
    )


def build_messages(candidates: list[dict], today: str,
                   footer: str | None = None) -> list[str]:
    """候補を軸順に並べ、上限内で分割したメッセージ本文のリストを返す。
    footer はコスト表示等をメッセージ末尾に付ける任意行。"""
    ordered = sorted(candidates, key=lambda c: _axis_rank(c.get("axis", "")))
    blocks = [_item_block(i, c) for i, c in enumerate(ordered, 1)]

    header = (f"*【{today}】速報！マンガニュース 記事候補 {len(blocks)}件*\n"
              "_採用するものの番号を返信してください（例: 1,3,7）_")

    messages, cur, cur_len = [], [header], len(header)
    for b in blocks:
        add = len(b) + 2
        if cur_len + add > CHAT_LIMIT and len(cur) > 1:
            messages.append("\n\n".join(cur))
            cur, cur_len = [], 0
        cur.append(b)
        cur_len += add
    if cur:
        messages.append("\n\n".join(cur))

    # コスト等のフッターを最終メッセージへ（入らなければ単独メッセージ）
    if footer:
        if len(messages[-1]) + len(footer) + 2 <= CHAT_LIMIT:
            messages[-1] += "\n\n" + footer
        else:
            messages.append(footer)

    # 2通目以降にページ表記を付与
    if len(messages) > 1:
        messages = [m + f"\n\n_({i}/{len(messages)})_"
                    for i, m in enumerate(messages, 1)]
    return messages


def _post(webhook: str, text: str) -> bool:
    payload = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(
        webhook, data=payload,
        headers={"Content-Type": "application/json; charset=UTF-8"},
    )
    with urllib.request.urlopen(req, timeout=20) as res:
        return res.status in (200, 204)


def notify(candidates: list[dict], footer: str | None = None) -> bool:
    webhook = os.environ.get("GOOGLE_CHAT_WEBHOOK")
    today = datetime.now().strftime("%Y/%m/%d (%a)")
    messages = build_messages(candidates, today, footer)

    if not webhook:
        print("[warn] GOOGLE_CHAT_WEBHOOK 未設定。以下を通知する予定でした:\n")
        print("\n\n---- (次メッセージ) ----\n\n".join(messages))
        return False

    ok_all = True
    for i, text in enumerate(messages):
        try:
            ok = _post(webhook, text)
            print(f"[info] Google Chat 通知 {i+1}/{len(messages)}: "
                  f"{'成功' if ok else '失敗'}")
            ok_all = ok_all and ok
        except Exception as e:
            print(f"[error] Google Chat 通知 {i+1}/{len(messages)} 失敗: {e}")
            ok_all = False
        if i < len(messages) - 1:
            time.sleep(1)  # 連投の間隔
    return ok_all
