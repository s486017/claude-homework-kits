# -*- coding: utf-8 -*-
"""2단계 — GitHub 저장소의 .ipynb 를 내려받아 셀 단위 텍스트(.ipynb.txt)로 저장한다.

.ipynb 원본(JSON)은 사람이 읽기도, Claude 가 읽기도 나쁘다. 셀 경계를 살린
텍스트로 바꿔 두면 몇 번 셀에서 무슨 일이 났는지 그대로 지목할 수 있다.

사용: python fetch_notebooks.py [--repo URL] [--path notebooks] [--only 이름 ...]
"""
import argparse
import json
import re
import sys
from pathlib import Path
from urllib.request import urlopen, Request

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kit  # noqa: E402

API = 'https://api.github.com/repos/{owner}/{repo}/contents/{path}?ref={ref}'


def parse_repo(url):
    m = re.search(r'github\.com/([^/]+)/([^/#?]+?)(?:\.git)?(?:/tree/([^/]+)(?:/(.*))?)?/?$', url.strip())
    if not m:
        raise SystemExit('GitHub 저장소 주소를 알아볼 수 없습니다: %s' % url)
    return m.group(1), m.group(2), m.group(3) or 'main', m.group(4)


def get(url):
    req = Request(url, headers={'User-Agent': 'ts-homework-kit', 'Accept': 'application/vnd.github+json'})
    with urlopen(req, timeout=60) as r:
        return r.read()


def to_text(nb_json, source_url, raw_url):
    nb = json.loads(nb_json)
    out = ['# Source: %s' % source_url, '# Raw: %s' % raw_url,
           '# 출처 슬라이드: ', '=' * 70, '']
    for i, c in enumerate(nb.get('cells', []), 1):
        src = ''.join(c.get('source', []))
        out.append('##### [Cell %d] %s #####' % (i, c.get('cell_type', 'code')))
        out.append(src.rstrip())
        out.append('')
    return '\n'.join(out)


def main():
    cfg = kit.load()
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', default=cfg['notebook_repo'])
    ap.add_argument('--path', default=cfg['notebook_path'])
    ap.add_argument('--only', nargs='*', default=None, help='특정 노트북만 (이름 일부)')
    a = ap.parse_args()
    if not a.repo:
        raise SystemExit('노트북 저장소가 설정되지 않았습니다.\n'
                         '  python setup_env.py --repo https://github.com/<계정>/<저장소>')

    owner, repo, ref, sub = parse_repo(a.repo)
    path = sub or a.path
    dest = kit.path_for(cfg, 'notebooks')
    items = json.loads(get(API.format(owner=owner, repo=repo, path=path, ref=ref)))
    if isinstance(items, dict):
        raise SystemExit('저장소 경로를 찾을 수 없습니다: %s/%s/%s' % (owner, repo, path))
    nbs = [i for i in items if i['name'].endswith('.ipynb')]
    if a.only:
        nbs = [i for i in nbs if any(k in i['name'] for k in a.only)]
    if not nbs:
        raise SystemExit('노트북(.ipynb)이 없습니다: %s/%s/%s' % (owner, repo, path))

    print('저장소 %s/%s (%s) — 노트북 %d개' % (owner, repo, ref, len(nbs)))
    for i in nbs:
        name = i['name'][:-6]
        raw = i['download_url']
        page = 'https://github.com/%s/%s/blob/%s/%s/%s' % (owner, repo, ref, path, i['name'])
        txt = to_text(get(raw), page, raw)
        f = dest / (name + '.ipynb.txt')
        f.write_text(txt, encoding='utf-8')
        cells = txt.count('##### [Cell ')
        print('  %-42s 셀 %3d  → %s' % (name, cells, f.name))
    print()
    print('저장 위치: %s' % dest)
    print('다음 단계: python run_notebooks.py')


if __name__ == '__main__':
    main()
