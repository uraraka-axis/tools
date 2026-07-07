# -*- coding: utf-8 -*-
"""LLM選抜のAPIコストを日次で記録し、月累計とアラート判定を返す。

- 記録先: cost_log.json（月ごとにトークン数・概算円・実行回数を積算）
- 単価・為替・アラート閾値は config.py（LLM_PRICE_* / USD_JPY / MONTHLY_COST_ALERT_JPY）
- 表示は概算。正確な請求は Anthropic Console の Usage を正とする
"""
import json
from datetime import datetime
from pathlib import Path

import config

COST_FILE = Path(__file__).parent / "cost_log.json"


def _load() -> dict:
    if COST_FILE.exists():
        try:
            return json.loads(COST_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save(data: dict) -> None:
    COST_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                         encoding="utf-8")


def cost_jpy(input_tokens: int, output_tokens: int) -> float:
    usd = (input_tokens / 1_000_000 * config.LLM_PRICE_USD_PER_MTOK_IN
           + output_tokens / 1_000_000 * config.LLM_PRICE_USD_PER_MTOK_OUT)
    return usd * config.USD_JPY


def record(usage: dict) -> dict:
    """1回分のusage({'input':n,'output':n})を積算し、サマリーを返す。"""
    month = datetime.now().strftime("%Y-%m")
    data = _load()
    m = data.setdefault(month, {"input_tokens": 0, "output_tokens": 0,
                                "cost_jpy": 0.0, "runs": 0})
    run_jpy = cost_jpy(usage["input"], usage["output"])
    m["input_tokens"] += usage["input"]
    m["output_tokens"] += usage["output"]
    m["cost_jpy"] = round(m["cost_jpy"] + run_jpy, 2)
    m["runs"] += 1
    _save(data)
    return {
        "month": month,
        "run_jpy": run_jpy,
        "month_jpy": m["cost_jpy"],
        "runs": m["runs"],
        "alert": m["cost_jpy"] > config.MONTHLY_COST_ALERT_JPY,
    }


def build_footer(info: dict) -> str:
    """Chat通知の末尾に付けるコスト行。"""
    line = (f"💰 LLM選抜コスト: 今回 ¥{info['run_jpy']:.1f}"
            f" ／ {info['month']} 累計 ¥{info['month_jpy']:.1f}"
            f"（{info['runs']}回・概算）")
    if info["alert"]:
        line += (f"\n⚠️ 月累計がアラート閾値 ¥{config.MONTHLY_COST_ALERT_JPY} を超えました。"
                 "config.py の USE_LLM=False で停止できます")
    return line
