# -*- coding: utf-8 -*-
"""画像取得のプレビューCLI（誤画像対策の動作確認用）

使い方:
  python scripts/fetch_preview.py <JAN> <期待タイトル> [set|tanpin|yoyaku] [--out DIR]

例:
  python scripts/fetch_preview.py 9784253151146 センセ。 tanpin
  python scripts/fetch_preview.py 9784063145342 カブのイサキ set

実際のワークフロー（_workflow_process_one_image）をそのまま通すため、
タイトル照合・ソース順・バッジ合成/リサイズ加工まで本番同等の結果が得られる。
出力: <out>/<JAN>_<type>.jpg と 取得ログ（標準出力）
"""
import sys
import os
import io
import json

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import requests
import streamlit_app as app


def run_case(jan, title, ctype='tanpin', out_dir='.'):
    session = requests.Session()
    data = {
        'comic_no': f'PREVIEW_{jan}',
        'first_jan': jan,
        'title': title,
        'series': '',
        'type': ctype,
        'is_tanpin': (ctype == 'tanpin'),
    }
    badge_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                              'images', 'badge_free_shipping.jpg')
    result = app._workflow_process_one_image(data, session, badge_path)

    out = {
        'jan': jan,
        'expected_title': title,
        'type': ctype,
        'success': result['success'],
        'source': result.get('source'),
        'log': result['log'],
        'file': None,
    }
    if result['success'] and result.get('image'):
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, f'{jan}_{ctype}.jpg')
        with open(path, 'wb') as f:
            f.write(result['image']['image_data'])
        out['file'] = path
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    out_dir = '.'
    for i, a in enumerate(sys.argv):
        if a == '--out' and i + 1 < len(sys.argv):
            out_dir = sys.argv[i + 1]

    if len(args) < 2:
        print(__doc__)
        sys.exit(1)

    jan = args[0]
    title = args[1]
    ctype = args[2] if len(args) > 2 else 'tanpin'

    result = run_case(jan, title, ctype, out_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
