# -*- coding: utf-8 -*-
"""
候補記事を「Smart Comic（法人・店舗向けコミックレンタル）目線」でスコアリングし、
上位 CANDIDATE_COUNT 件を選んで、記事化のための仮タイトル等を付与する。

優先：Anthropic API（LLM）による文脈スコアリング。
APIキーが無い／USE_LLM=False の場合は、キーワード辞書ベースのヒューリスティックで代替。
"""
import os
import re
import json
from datetime import datetime, timezone

import config

# 直近のLLM呼び出しのトークン使用量（main.py がコスト記録に使う。未使用なら None）
last_usage: dict | None = None


def _age_days(item: dict) -> float | None:
    """記事の経過日数。published が無ければ None。"""
    p = item.get("published", "")
    if not p:
        return None
    try:
        dt = datetime.fromisoformat(p)
        return (datetime.now(timezone.utc) - dt).total_seconds() / 86400
    except Exception:
        return None


def _work_key(title: str) -> str:
    """タイトルから作品名（『』「』内）を抜き出し、重複判定キーにする。
    括弧が無ければ空文字（＝重複判定しない）。"""
    m = re.search(r"[『「]([^』」]{2,40})[』」]", title)
    return m.group(1).strip() if m else ""


# ---------------------------------------------------------------------------
# ヒューリスティック・スコアリング（フォールバック）
# ---------------------------------------------------------------------------
def _classify_type(text: str) -> str:
    best, best_hit = "その他", 0
    for typ, kws in config.TYPE_KEYWORDS.items():
        hit = sum(1 for k in kws if k in text)
        if hit > best_hit:
            best, best_hit = typ, hit
    return best


def _heuristic_score(item: dict) -> dict:
    text = f"{item['title']} {item['summary']}"
    score = 1.0 * item.get("weight", 1.0)

    # 種別ヒット数で加点
    typ = _classify_type(text)
    if typ != "その他":
        score += 1.5

    # 法人/店舗フィット加点
    score += 0.6 * sum(1 for k in config.FIT_KEYWORDS if k in text)

    # ノイズ減点
    score -= 0.8 * sum(1 for k in config.NOISE_KEYWORDS if k in text)

    # 新しさ加点（古いほど減点。日付不明は中庸の7日扱い）
    age = _age_days(item)
    score -= config.RECENCY_PENALTY * (age if age is not None else 7)

    axis = item.get("axis") or typ
    return {
        **item,
        "type": axis,
        "score": round(score, 2),
        "proposed_title": item["title"],
        "reason": f"{item['source']}より。{axis}の話題。",
        "keywords": "",
    }


def heuristic_rank(items: list[dict]) -> list[dict]:
    """スコア順に並べつつ、1軸 PER_AXIS_CAP 件までで多様性を確保して選抜。"""
    scored = sorted((_heuristic_score(it) for it in items),
                    key=lambda x: x["score"], reverse=True)

    picked, axis_count, seen_works = [], {}, set()
    # 1巡目：軸の上限＋同一作品の重複除去で多様性優先に集める
    for it in scored:
        if len(picked) >= config.CANDIDATE_COUNT:
            break
        ax = it.get("axis") or it["type"]
        if axis_count.get(ax, 0) >= config.PER_AXIS_CAP:
            continue
        wk = _work_key(it["title"])
        if wk and wk in seen_works:      # 同じ作品名は1本だけ
            continue
        if wk:
            seen_works.add(wk)
        axis_count[ax] = axis_count.get(ax, 0) + 1
        picked.append(it)

    # 2巡目：埋まらなければ重複作品だけ避けてスコア順で補充
    if len(picked) < config.CANDIDATE_COUNT:
        chosen = {id(p) for p in picked}
        for it in scored:
            if len(picked) >= config.CANDIDATE_COUNT:
                break
            if id(it) in chosen:
                continue
            wk = _work_key(it["title"])
            if wk and wk in seen_works:
                continue
            if wk:
                seen_works.add(wk)
            picked.append(it)
    return picked[: config.CANDIDATE_COUNT]


# ---------------------------------------------------------------------------
# LLMスコアリング（Anthropic API）
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """あなたは法人・店舗向けコミックレンタルサービス「Smart Comic」の
オウンドメディア『速報！マンガニュース』の編集者です。

【メディアの狙い】
温浴・宿泊施設、賃貸物件、介護施設、待合スペース等を持つ法人に向けて、
マンガの話題で集客し、自社のコミックレンタル導入へ自然に繋げる記事を量産する。

【良いネタの条件】（優先度順）
1. 話題性：今まさに検索・SNSで伸びている / 多くの人が知りたい
2. 法人・店舗フィット：施設にマンガを置く文脈と結びつけられる
3. 新しさ：「速報」メディアなので、できるだけ新しい記事を優先（古いネタは避ける）
4. SEO見込み：狙えるキーワードが明確で、競合が薄い
5. 自社導線：記事末尾でレンタル導入の訴求に自然につながる
6. 軸の多様性：以下の多様な軸からバランス良く選ぶ
   （総合ニュース / SNS・ネット話題 / 広告ジャック・街頭 / 考察・反響 /
     新刊・最新巻 / 完結・連載再開 / アニメ化 / 実写・映画化 /
     ドラマ化 / 配信(Netflix等) / 受賞・ランキング / 展覧会・原画展 /
     イベント・コラボ / 聖地・地域コラボ / 著名人×漫画 / 海外・世界的ヒット）
   特に「SNS・ネット話題」（Xでのバズ・広告ジャック・考察合戦など）は
   速報メディアとして価値が高いので、該当候補があれば必ず1〜3件含める。

【避けたいネタ】
単発グッズ・くじ・フィギュア等の物販告知、声優の私生活、極端にニッチな同人、
同一作品・同一話題の重複（似たネタは最良の1本に絞る）。

各候補には収集軸(axis)が付いています。なるべく多くの軸が並ぶように選んでください。
与えられた候補から、上記基準で最も記事にすべきものを選び、JSONで返してください。"""


def _build_user_prompt(items: list[dict], n: int) -> str:
    lines = ["# 本日の候補記事プール\n"]
    for i, it in enumerate(items):
        age = _age_days(it)
        age_s = f"{int(age)}日前" if age is not None else "日付不明"
        lines.append(
            f"[{i}] ({it.get('axis','-')}/{age_s}) {it['title']}\n"
            f"    source: {it['source']}\n"
            f"    summary: {it['summary'][:160]}\n"
        )
    lines.append(
        f"\n上記から最も記事化すべき {n} 件を選び、次のJSON配列だけを出力してください。"
        " 説明文やコードフェンスは不要です。\n"
        "[{\n"
        '  "index": 候補番号(整数),\n'
        '  "type": "その候補の軸(axis)をそのまま",\n'
        '  "proposed_title": "クリックされる日本語の仮タイトル(32文字前後)",\n'
        '  "reason": "なぜ今これを書くべきか・自社導線の作り方(80文字以内)",\n'
        '  "keywords": "想定検索キーワードをカンマ区切りで2〜4個",\n'
        '  "score": 1〜10の整数(記事化の優先度)\n'
        "}, ...]\n"
        f"必ず {n} 件、軸が偏らないようにし、scoreの高い順に並べてください。"
    )
    return "\n".join(lines)


def llm_rank(items: list[dict]) -> list[dict] | None:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("[info] ANTHROPIC_API_KEY 未設定 → ヒューリスティックで代替します")
        return None
    try:
        import anthropic
    except ImportError:
        print("[warn] anthropic 未インストール → ヒューリスティックで代替します")
        return None

    pool = items[: config.POOL_LIMIT]
    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model=config.LLM_MODEL,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{
                "role": "user",
                "content": _build_user_prompt(pool, config.CANDIDATE_COUNT),
            }],
        )
        global last_usage
        u = getattr(resp, "usage", None)
        if u is not None:
            last_usage = {"input": u.input_tokens, "output": u.output_tokens}
        raw = resp.content[0].text.strip()
        # コードフェンスが付いた場合の保険
        if raw.startswith("```"):
            raw = raw.split("```")[1].lstrip("json").strip()
        picks = json.loads(raw)
    except Exception as e:
        print(f"[warn] LLMスコアリング失敗: {e} → ヒューリスティックで代替します")
        return None

    result = []
    for p in picks[: config.CANDIDATE_COUNT]:
        try:
            src = pool[int(p["index"])]
        except (KeyError, ValueError, IndexError):
            continue
        result.append({
            **src,
            "type": p.get("type", "その他"),
            "proposed_title": p.get("proposed_title", src["title"]),
            "reason": p.get("reason", ""),
            "keywords": p.get("keywords", ""),
            "score": p.get("score", 0),
        })
    return result or None


def rank(items: list[dict]) -> list[dict]:
    if config.USE_LLM:
        llm = llm_rank(items)
        if llm:
            return llm
    return heuristic_rank(items)
