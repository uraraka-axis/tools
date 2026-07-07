# -*- coding: utf-8 -*-
"""
manga-trend-radar 設定ファイル

情報源は2系統：
  (A) 専門ニュースの直RSS（コミックナタリー等）……総合的なマンガニュース
  (B) Google ニュース検索RSS（軸ごとのクエリ）……多軸でネタを掘る

軸を増やしたいときは AXIS_QUERIES に1行足すだけ。
"""
import urllib.parse


# ---------------------------------------------------------------------------
# フィード取得時のUser-Agent（無いとブロックするサイトがあるため必須）
# ---------------------------------------------------------------------------
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")


def _gnews(query: str) -> str:
    """Google ニュースの検索RSS URLを組み立てる。
    query に `when:14d` 等を入れると期間で絞れる。"""
    q = urllib.parse.quote(query)
    return f"https://news.google.com/rss/search?q={q}&hl=ja&gl=JP&ceid=JP:ja"


# ---------------------------------------------------------------------------
# (A) 専門ニュースの直RSS（動作確認済み）
#   filter_recency=True … COLLECT_HOURS で直近のみに絞る
# ---------------------------------------------------------------------------
DIRECT_FEEDS = [
    {"name": "コミックナタリー", "axis": "総合ニュース", "weight": 1.3,
     "url": "https://natalie.mu/comic/feed/news", "filter_recency": True},
    {"name": "アニメ！アニメ！", "axis": "総合ニュース", "weight": 1.0,
     "url": "https://animeanime.jp/rss/index.rdf", "filter_recency": True},
    {"name": "ITmedia ねとらぼ", "axis": "ネット話題", "weight": 1.0,
     "url": "https://rss.itmedia.co.jp/rss/2.0/netlab.xml", "filter_recency": True},
    # はてなブックマーク ホットエントリー … X/SNS発のバズ（広告ジャック・考察合戦等）が
    # ブックマークを集めて浮上する場所。マンガ以外のノイズはスコアリングで沈む前提
    {"name": "はてブ アニメとゲーム", "axis": "SNS・ネット話題", "weight": 1.1,
     "url": "https://b.hatena.ne.jp/hotentry/game.rss", "filter_recency": True},
    {"name": "はてブ エンタメ", "axis": "SNS・ネット話題", "weight": 1.0,
     "url": "https://b.hatena.ne.jp/hotentry/entertainment.rss", "filter_recency": True},
]

# ---------------------------------------------------------------------------
# (B) Google ニュース 軸別クエリ
#   query 末尾の when:Nd で期間を指定（gnews側で絞るので filter_recency=False）
#   軸を足す/減らす/言い回しを変えるのはここだけ触ればOK
# ---------------------------------------------------------------------------
AXIS_QUERIES = [
    {"axis": "新刊・最新巻",   "weight": 1.0, "query": "人気漫画 最新刊 OR 新刊 when:7d"},
    {"axis": "アニメ化",       "weight": 1.1, "query": "漫画 アニメ化 決定 when:10d"},
    {"axis": "実写・映画化",   "weight": 1.1, "query": "漫画原作 実写化 OR 映画化 when:14d"},
    {"axis": "ドラマ化",       "weight": 1.1, "query": "漫画原作 ドラマ化 when:14d"},
    {"axis": "配信(Netflix等)", "weight": 1.2, "query": "Netflix OR ディズニープラス 漫画 OR アニメ 配信 when:14d"},
    {"axis": "展覧会・原画展", "weight": 1.2, "query": "漫画 原画展 OR 展覧会 OR 美術館 when:14d"},
    {"axis": "イベント・コラボ", "weight": 1.0, "query": "漫画 OR アニメ コラボ OR イベント 開催 when:10d"},
    {"axis": "著名人×漫画",    "weight": 1.2, "query": "俳優 OR 芸能人 OR スポーツ選手 好きな漫画 OR 愛読書 when:14d"},
    {"axis": "受賞・ランキング", "weight": 1.1, "query": "マンガ大賞 OR 漫画賞 OR 漫画 ランキング 発表 when:14d"},
    {"axis": "海外・世界的ヒット", "weight": 1.0, "query": "漫画 OR アニメ 海外 人気 OR 世界的ヒット when:14d"},
    {"axis": "完結・連載再開",  "weight": 1.0, "query": "人気漫画 完結 OR 連載再開 when:14d"},
    {"axis": "聖地・地域コラボ", "weight": 1.1, "query": "アニメ OR 漫画 聖地 OR 自治体 OR ふるさと コラボ when:14d"},
    {"axis": "SNS・ネット話題",  "weight": 1.3, "query": "漫画 OR マンガ Xで話題 OR SNSで話題 OR トレンド入り OR バズ when:7d"},
    {"axis": "広告ジャック・街頭", "weight": 1.2, "query": "漫画 OR アニメ 広告ジャック OR 駅広告 OR 渋谷ジャック OR 交通広告 when:14d"},
    {"axis": "考察・反響",       "weight": 1.1, "query": "漫画 最新刊 OR 最終回 考察 OR 難解 OR 衝撃 話題 when:14d"},
]

# 上記2系統を結合した最終ソース一覧（collector はこれを読む）
SOURCES = DIRECT_FEEDS + [
    {"name": f"Gニュース:{a['axis']}", "axis": a["axis"], "weight": a["weight"],
     "url": _gnews(a["query"]), "filter_recency": False}
    for a in AXIS_QUERIES
]


# ---------------------------------------------------------------------------
# 自社（Smart Comic：法人・店舗向けコミックレンタル）目線のキーワード辞書
#   ヒューリスティック・スコアリング（APIキー未設定時のフォールバック）で使う
# ---------------------------------------------------------------------------
TYPE_KEYWORDS = {
    "新刊・最新巻": ["新刊", "最新刊", "発売", "巻", "完結", "連載再開", "重版"],
    "賞・ランキング": ["大賞", "受賞", "ランキング", "1位", "賞", "ノミネート", "発表"],
    "メディア化": ["アニメ化", "実写化", "映画化", "ドラマ化", "舞台化", "配信", "Netflix", "放送"],
    "時事×マンガ": ["コラボ", "空港", "駅", "聖地", "展", "原画展", "美術館", "イベント", "自治体", "地域"],
    "著名人×マンガ": ["俳優", "声優", "芸能人", "選手", "好きな漫画", "愛読"],
    "研究・データ": ["研究", "調査", "データ", "脳", "効果", "統計", "実験", "大学"],
    "SNS・ネット話題": ["Xで話題", "SNSで話題", "トレンド入り", "バズ", "ジャック", "広告",
                     "考察", "難解", "反響", "衝撃", "ツイート", "ポスト"],
}

# 法人/店舗レンタルへ導線を作りやすい＝加点ワード
FIT_KEYWORDS = [
    "施設", "店舗", "待合", "ホテル", "宿泊", "温浴", "サウナ", "スーパー銭湯",
    "ネットカフェ", "漫画喫茶", "賃貸", "介護", "病院", "クリニック", "カフェ",
    "話題", "人気", "売上", "ヒット", "社会現象", "ブーム", "行列", "世界",
]

# ノイズ（基本的に除外したい）＝減点ワード
NOISE_KEYWORDS = [
    "グッズ", "フィギュア", "プライズ", "くじ", "アクリル", "缶バッジ",
    "ガチャ", "カプセル", "コスプレ", "チケット先行", "通販予約",
]

# 軸の表示順（通知でこの順に並べると見やすい）
AXIS_ORDER = ["総合ニュース", "SNS・ネット話題", "広告ジャック・街頭", "考察・反響",
              "新刊・最新巻", "完結・連載再開", "アニメ化",
              "実写・映画化", "ドラマ化", "配信(Netflix等)", "受賞・ランキング",
              "展覧会・原画展", "イベント・コラボ", "聖地・地域コラボ",
              "著名人×漫画", "海外・世界的ヒット", "ネット話題"]


# ---------------------------------------------------------------------------
# 動作設定
# ---------------------------------------------------------------------------
CANDIDATE_COUNT = 20         # 通知する候補数
COLLECT_HOURS = 72           # 直RSS(filter_recency=True)の対象時間
MAX_AGE_DAYS = 14            # 全ソース共通の上限。これより古い記事は日付があれば捨てる
RECENCY_PENALTY = 0.18       # 1日古いごとにスコアから引く量（新しい記事を優先）
POOL_LIMIT = 130             # スコアリングに渡す候補プールの上限
PER_AXIS_CAP = 3             # 1つの軸から選ぶ最大件数（多様性確保／ヒューリスティック時）
HISTORY_DAYS = 45            # 重複判定に使う履歴の保持日数
WEEKDAYS_ONLY = True         # 平日のみ通知（土日はスキップ）

# LLMスコアリング（Anthropic API）
USE_LLM = True               # False で完全ヒューリスティックのみ
LLM_MODEL = "claude-haiku-4-5-20251001"   # 安価・高速。必要なら sonnet 系へ

# LLM選抜のコスト管理（単価は Haiku 4.5。LLM_MODEL を変えたらここも更新）
LLM_PRICE_USD_PER_MTOK_IN = 1.0    # $/100万入力トークン
LLM_PRICE_USD_PER_MTOK_OUT = 5.0   # $/100万出力トークン
USD_JPY = 155                      # 表示用の概算レート（請求はUSD）
MONTHLY_COST_ALERT_JPY = 300       # 月累計がこれを超えたら通知に⚠️を出す
