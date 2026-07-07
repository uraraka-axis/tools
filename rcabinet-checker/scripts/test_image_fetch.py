"""
画像取得ロジックの検証スクリプト（実ネットワークアクセスあり）

streamlit_app.py の誤画像事故対策（タイトル照合ロジック）を検証する。
- normalize_title_for_match / title_matches の単体テスト
- get_openbd_info / get_ndl_thumbnail のISBN直引きテスト
- get_bookoff_image / get_amazon_image / get_rakuten_image の照合付き取得テスト
- _workflow_process_one_image の統合テスト（事故ケース・正常ケース・存在しないJAN）

実行:
    python scripts/test_image_fetch.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Windows のコンソール(cp932)では絵文字が出力できずクラッシュするため、
# 標準出力をUTF-8に強制（テスト結果の欠落を防ぐ）
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import requests

# streamlit_app.py はモジュールレベルでStreamlit UIコードを実行するが、
# 「bare mode」（streamlit runを介さない直接import）でも動作することを確認済み
# （st.set_page_config / st.secrets / st.sidebar 等は bare mode で無害化される）
import streamlit_app as app


def section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def test_normalize_and_match():
    section("1. normalize_title_for_match / title_matches 単体テスト")

    cases = [
        # (expected_titles, candidate, expect_match, ラベル)
        (["センセ。"], "それ犯罪かもしれない図鑑", False, "事故ケース: センセ。 vs それ犯罪かもしれない図鑑"),
        (["ONE PIECE"], "ONE PIECE 1 (ジャンプコミックス)", True, "ONE PIECE vs ONE PIECE 1 (ジャンプコミックス)"),
        (["カブのイサキ"], "ヨコハマ買い出し紀行(1)", False, "事故ケース: カブのイサキ vs ヨコハマ買い出し紀行(1)"),
        (["センセ。"], "センセ。 12", True, "センセ。 vs センセ。12（正しい候補）"),
        (["ONE PIECE"], "ワンピース naruto", False, "ONE PIECE vs 無関係タイトル"),
    ]

    all_ok = True
    for expected_titles, candidate, expect, label in cases:
        result = app.title_matches(expected_titles, candidate)
        ok = (result == expect)
        all_ok = all_ok and ok
        print(f"  [{'OK' if ok else 'NG'}] {label} -> got={result} expect={expect}")

    print(f"  normalize_title_for_match('センセ。') = {app.normalize_title_for_match('センセ。')!r}")
    print(f"  normalize_title_for_match('') = {app.normalize_title_for_match('')!r}")
    print(f"  title_matches([], '何か') = {app.title_matches([], '何か')} (expected_titles空 -> False)")

    print(f"\n  総合結果: {'ALL OK' if all_ok else 'FAILURES EXIST'}")
    return all_ok


def test_openbd_ndl():
    section("2. get_openbd_info / get_ndl_thumbnail テスト")
    session = requests.Session()

    cases = [
        ("9784253151146", "センセ。12巻（事故ケース）"),
        ("9784088725093", "ONE PIECE 1巻"),
        ("9784999999999", "存在しないJAN"),
    ]

    for jan, label in cases:
        print(f"\n  -- {label} (JAN: {jan}) --")
        info = app.get_openbd_info(jan, session)
        print(f"  openBD: {info}")
        ndl_url = app.get_ndl_thumbnail(jan, session)
        print(f"  NDL thumbnail: {ndl_url}")


def test_scrapers():
    section("3. get_bookoff_image / get_amazon_image / get_rakuten_image テスト")
    session = requests.Session()

    cases = [
        ("9784253151146", ["センセ。", "センセ。 12"], "センセ。12巻（事故ケース）"),
        ("9784088725093", ["ONE PIECE"], "ONE PIECE 1巻"),
        ("9784999999999", ["存在しないタイトル"], "存在しないJAN"),
    ]

    accident_free = True

    for jan, expected_titles, label in cases:
        print(f"\n  -- {label} (JAN: {jan}) expected={expected_titles} --")

        try:
            bookoff = app.get_bookoff_image(jan, session, expected_titles)
        except Exception as e:
            bookoff = f"EXCEPTION: {e}"
        print(f"  bookoff : {bookoff}")

        try:
            amazon = app.get_amazon_image(jan, session, expected_titles)
        except Exception as e:
            amazon = f"EXCEPTION: {e}"
        print(f"  amazon  : {amazon}")
        if isinstance(amazon, tuple):
            url, title = amazon
            if "犯罪" in (title or "") or "それ犯罪かもしれない図鑑" in (title or ""):
                accident_free = False
                print("  !!!! 事故再発: 「それ犯罪かもしれない図鑑」が返された !!!!")

        try:
            rakuten = app.get_rakuten_image(jan, session, expected_titles)
        except Exception as e:
            rakuten = f"EXCEPTION: {e}"
        print(f"  rakuten : {rakuten}")

    print(f"\n  事故再発チェック: {'PASS（事故なし）' if accident_free else 'FAIL（事故再発）'}")
    return accident_free


def test_workflow_integration():
    section("4. _workflow_process_one_image 統合テスト")

    badge_path = os.path.join(os.path.dirname(app.__file__), "images", "badge_free_shipping.jpg")
    session = requests.Session()

    cases = [
        {
            "comic_no": "TEST_SENSE",
            "first_jan": "9784253151146",
            "type": "tanpin",
            "is_tanpin": True,
            "genre": "", "publisher": "", "series": "",
            "title": "センセ。",
        },
        {
            "comic_no": "TEST_ONEPIECE",
            "first_jan": "9784088725093",
            "type": "tanpin",
            "is_tanpin": True,
            "genre": "", "publisher": "", "series": "",
            "title": "ONE PIECE",
        },
        {
            "comic_no": "TEST_NOTEXIST",
            "first_jan": "9784999999999",
            "type": "tanpin",
            "is_tanpin": True,
            "genre": "", "publisher": "", "series": "",
            "title": "存在しないタイトル",
        },
    ]

    accident_free = True

    for data in cases:
        print(f"\n  -- {data['comic_no']} (JAN: {data['first_jan']}, title={data['title']!r}) --")
        try:
            result = app._workflow_process_one_image(data, session, badge_path)
        except Exception as e:
            print(f"  EXCEPTION: {e}")
            continue

        print(f"  success={result['success']} source={result['source']}")
        print(f"  log: {result['log']}")

        if data["comic_no"] == "TEST_SENSE" and result["success"]:
            # 「それ犯罪かもしれない図鑑」の画像が採用されていないことを確認
            # (source が bookoff/amazon/rakuten の場合はimage_dictにmatched_titleは
            #  含まれないが、logに照合タイトルが出るのでそちらを確認する)
            if "犯罪" in result["log"]:
                accident_free = False
                print("  !!!! 事故再発: ログに「犯罪」を含むタイトルが採用された !!!!")

        if data["comic_no"] == "TEST_NOTEXIST" and result["success"]:
            print("  !!!! 想定外: 存在しないJANで成功してしまった（openBD/NDLで別の本が誤ヒットした可能性） !!!!")

    print(f"\n  事故再発チェック: {'PASS（事故なし）' if accident_free else 'FAIL（事故再発）'}")
    return accident_free


def main():
    results = {}
    results["normalize_and_match"] = test_normalize_and_match()

    test_openbd_ndl()

    try:
        results["scrapers"] = test_scrapers()
    except Exception as e:
        print(f"\n  [scrapers] 予期しないエラー: {e}")
        results["scrapers"] = None

    try:
        results["workflow_integration"] = test_workflow_integration()
    except Exception as e:
        print(f"\n  [workflow_integration] 予期しないエラー: {e}")
        results["workflow_integration"] = None

    section("最終サマリ")
    for k, v in results.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
