# -*- coding: utf-8 -*-
"""2단계 — 노트북(.ipynb)을 가져와 셀 단위 텍스트(.ipynb.txt)로 저장한다.

.ipynb 원본(JSON)은 사람이 읽기도, Claude 가 읽기도 나쁘다. 셀 경계를 살린
텍스트로 바꿔 두면 몇 번 셀에서 무슨 일이 났는지 그대로 지목할 수 있다.

가져오는 방법 세 가지 — **GitHub 계정이 없어도 된다.**

  1. GitHub 저장소 전체 (강의 노트북이 공개 저장소에 있을 때)
     python fetch_notebooks.py --repo https://github.com/<계정>/<저장소>

  2. 내 PC 의 파일·폴더 (메일·LMS 로 .ipynb 를 직접 받았을 때)
     python fetch_notebooks.py --from ~/Downloads/과제노트북
     python fetch_notebooks.py --from a.ipynb b.ipynb

  3. 노트북 하나의 주소 (Colab·GitHub 링크를 그대로 붙여넣기)
     python fetch_notebooks.py --url https://colab.research.google.com/github/.../x.ipynb
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


def save(dest, name, txt):
    f = dest / (name + '.ipynb.txt')
    f.write_text(txt, encoding='utf-8')
    print('  %-42s 셀 %3d  → %s' % (name, txt.count('##### [Cell '), f.name))


def from_local(paths, dest, only):
    """내 PC 의 .ipynb 파일이나 폴더에서 가져온다 — 네트워크도 계정도 필요 없다."""
    files = []
    for p in paths:
        p = Path(p).expanduser()
        if p.is_dir():
            files += sorted(p.rglob('*.ipynb'))
        elif p.suffix == '.ipynb':
            files.append(p)
        else:
            print('  [건너뜀] .ipynb 가 아닙니다: %s' % p)
    files = [f for f in files if '.ipynb_checkpoints' not in str(f)]
    if only:
        files = [f for f in files if any(k in f.name for k in only)]
    if not files:
        raise SystemExit('가져올 .ipynb 를 찾지 못했습니다: %s' % ', '.join(str(p) for p in paths))
    print('내 PC 에서 노트북 %d개' % len(files))
    for f in files:
        save(dest, f.stem, to_text(f.read_bytes(), str(f), ''))
    return len(files)


def from_url(url, dest):
    """Colab / GitHub 링크 하나를 그대로 받는다."""
    raw = url
    m = re.search(r'colab\.research\.google\.com/github/(.+)$', url)
    if m:
        raw = 'https://raw.githubusercontent.com/' + m.group(1).replace('/blob/', '/')
    elif 'github.com' in url and '/blob/' in url:
        raw = url.replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/')
    name = raw.rstrip('/').split('/')[-1].replace('.ipynb', '') or 'notebook'
    print('주소에서 노트북 1개')
    save(dest, name, to_text(get(raw), url, raw))
    return 1


def main():
    cfg = kit.load()
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', default=cfg['notebook_repo'], help='GitHub 저장소 주소')
    ap.add_argument('--path', default=cfg['notebook_path'], help='저장소 안의 노트북 폴더')
    ap.add_argument('--from', dest='local', nargs='*', default=None,
                    help='내 PC 의 .ipynb 파일 또는 폴더')
    ap.add_argument('--url', default=None, help='노트북 하나의 주소 (Colab/GitHub 링크)')
    ap.add_argument('--only', nargs='*', default=None, help='특정 노트북만 (이름 일부)')
    a = ap.parse_args()

    dest = kit.path_for(cfg, 'notebooks')
    if a.local:
        n = from_local(a.local, dest, a.only)
        print('\n저장 위치: %s\n다음 단계: python run_notebooks.py' % dest)
        return
    if a.url:
        from_url(a.url, dest)
        print('\n저장 위치: %s\n다음 단계: python run_notebooks.py' % dest)
        return
    if not a.repo:
        raise SystemExit(
            '노트북을 어디서 가져올지 정하지 않았습니다. 셋 중 하나를 쓰세요.\n'
            '  1) 공개 저장소에 있는 경우 : fetch_notebooks.py --repo https://github.com/<계정>/<저장소>\n'
            '  2) 파일로 받은 경우       : fetch_notebooks.py --from <폴더 또는 .ipynb 경로>\n'
            '  3) Colab 링크만 있는 경우  : fetch_notebooks.py --url <노트북 주소>\n'
            '※ 본인 GitHub 계정이나 저장소는 필요 없습니다. 노트북을 "가져올 곳"을 알려 주는 것뿐입니다.')

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
        raw = i['download_url']
        page = 'https://github.com/%s/%s/blob/%s/%s/%s' % (owner, repo, ref, path, i['name'])
        save(dest, i['name'][:-6], to_text(get(raw), page, raw))
    print()
    print('저장 위치: %s' % dest)
    print('다음 단계: python run_notebooks.py')


if __name__ == '__main__':
    main()
