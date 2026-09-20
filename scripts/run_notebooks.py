# -*- coding: utf-8 -*-
"""3단계 — 노트북의 코드 셀을 순서대로 실행하고, 셀별 출력·그림을 저장한다.

- 출력   : 실행결과/<노트북명>/cellNN.txt
- 그림   : 실행결과/<노트북명>/cellNN_figK.png
- 요약   : 실행결과/<노트북명>/summary.md  (코드 + 출력 + 그림 + 수정 내역)

`%pip` 같은 주피터 매직은 건너뛰고, 주피터 전용 `display()` 는 print 로 대체한다.
오류가 나도 멈추지 않고 끝까지 간 뒤 summary.md 에 ERROR 로 남긴다 — 무엇이
깨졌는지 한눈에 보려면 전부 돌려 보는 편이 빠르기 때문이다.

사용: python run_notebooks.py [이름 ...] [--list]
"""
import argparse
import contextlib
import importlib.util
import io
import re
import sys
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kit  # noqa: E402

import matplotlib  # noqa: E402
matplotlib.use('Agg')
import matplotlib.pyplot as plt  # noqa: E402

_font = kit.korean_font()
if _font:
    plt.rcParams['font.family'] = _font
plt.rcParams['axes.unicode_minus'] = False

CELL_RE = re.compile(r'^##### \[Cell (\d+)\] (code|markdown) #####\s*$', re.M)


def load_fixes(ws):
    """작업 폴더의 fixes.py 에서 노트북별 수정 목록을 읽는다."""
    f = ws / 'fixes.py'
    if not f.exists():
        return {}
    spec = importlib.util.spec_from_file_location('ts_kit_fixes', f)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return getattr(mod, 'FIXES', {})


def parse_cells(text):
    marks = list(CELL_RE.finditer(text))
    cells = []
    for i, m in enumerate(marks):
        body = text[m.end(): marks[i + 1].start() if i + 1 < len(marks) else len(text)]
        body = body.split('\n--- [output]', 1)[0]
        cells.append((int(m.group(1)), m.group(2), body.strip('\n')))
    return cells


def clean_code(src):
    """주피터 매직(%, !)을 실행 가능한 형태로 중화한다."""
    lines = []
    for line in src.splitlines():
        s = line.strip()
        if s.startswith('%%'):
            continue
        if s.startswith(('%', '!')):
            lines.append(line[:len(line) - len(line.lstrip())] + 'pass  # [skipped] ' + s)
            continue
        lines.append(line)
    return '\n'.join(lines)


def run_one(src_path, out_root, fixes_all):
    name = src_path.name.replace('.ipynb.txt', '')
    out_dir = out_root / name
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob('cell*'):
        old.unlink()

    text = src_path.read_text(encoding='utf-8')
    fixes = fixes_all.get(name, [])

    def display(*objs, **_k):
        for o in objs:
            print(o.to_string() if hasattr(o, 'to_string') else o)

    ns = {'__name__': '__main__', 'display': display}
    st = {'cell': 0, 'k': 0, 'files': []}

    def show(*_a, **_k):
        for num in plt.get_fignums():
            st['k'] += 1
            f = out_dir / ('cell%02d_fig%d.png' % (st['cell'], st['k']))
            plt.figure(num).savefig(f, dpi=110, bbox_inches='tight')
            st['files'].append(f.name)
        plt.close('all')

    plt.show = show

    summary = ['# %s 실행 결과\n' % name]
    if 'display(' in text:
        summary.append('> 참고: 주피터 전용 `display()` 는 실행기에서 `print(df.to_string())` 로 대체함.\n')
    if fixes:
        summary.append('## 오류 및 수정 내역\n')
        for cnum, old, new, why in fixes:
            if old == new:
                summary.append('- **Cell %s**: %s' % (cnum, why))
            else:
                summary.append('- **Cell %s**: %s\n  - 원본: `%s`\n  - 수정: `%s`'
                               % (cnum, why, old.strip(), new.strip()))
        summary.append('')

    ok = err = figs = 0
    for num, kind, body in parse_cells(text):
        if kind != 'code':
            continue
        for cnum, old, new, _why in fixes:
            if cnum == num:
                if old and old not in body:
                    print('  [Cell %d] 주의: 수정 대상 문자열을 못 찾음: %s' % (num, old[:60]))
                if old:
                    body = body.replace(old, new)
        st.update(cell=num, k=0, files=[])
        buf = io.StringIO()
        status = 'OK'
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                exec(compile(clean_code(body), '<cell %d>' % num, 'exec'), ns)
            except Exception:
                status = 'ERROR'
                traceback.print_exc()
        show()
        out = buf.getvalue()
        (out_dir / ('cell%02d.txt' % num)).write_text(out, encoding='utf-8')
        summary.append('## Cell %d — %s\n```python\n%s\n```\n' % (num, status, body.strip()))
        if out.strip():
            summary.append('출력:\n```\n%s\n```\n' % out.rstrip())
        for f in st['files']:
            summary.append('![%s](%s)\n' % (f, f))
        figs += len(st['files'])
        ok += status == 'OK'
        err += status == 'ERROR'
        if status == 'ERROR':
            print('  [Cell %d] ERROR\n%s' % (num, '\n'.join('    ' + l for l in out.strip().splitlines()[-6:])))
    (out_dir / 'summary.md').write_text('\n'.join(summary), encoding='utf-8')
    return {'name': name, 'ok': ok, 'err': err, 'figs': figs, 'fixes': len(fixes)}


def main():
    cfg = kit.load()
    ap = argparse.ArgumentParser()
    ap.add_argument('names', nargs='*', help='특정 노트북만 (이름 일부)')
    ap.add_argument('--list', action='store_true', help='목록만 보기')
    a = ap.parse_args()

    src = kit.path_for(cfg, 'notebooks')
    out = kit.path_for(cfg, 'results')
    files = sorted(src.glob('*.ipynb.txt'))
    if a.names:
        files = [f for f in files if any(k in f.name for k in a.names)]
    if not files:
        raise SystemExit('노트북이 없습니다: %s\n  먼저 python fetch_notebooks.py' % src)
    if a.list:
        for f in files:
            print(f.name)
        return

    if not _font:
        print('[주의] 한글 폰트를 못 찾아 그래프의 한글이 깨질 수 있습니다.')
    rows = []
    for i, f in enumerate(files, 1):
        print('[%d/%d] %s' % (i, len(files), f.name.replace('.ipynb.txt', '')))
        rows.append(run_one(f, out, load_fixes(cfg['_workspace'])))

    print()
    print('%-42s %5s %5s %5s %s' % ('노트북', '정상', '오류', '그림', '상태'))
    for r in rows:
        state = '오류 남음' if r['err'] else ('수정 후 정상' if r['fixes'] else '정상')
        print('%-42s %5d %5d %5d %s' % (r['name'], r['ok'], r['err'], r['figs'], state))
    bad = [r['name'] for r in rows if r['err']]
    print()
    print('노트북 %d개 | 셀 %d개 실행 | 오류 %d개 | 그림 %d장'
          % (len(rows), sum(r['ok'] + r['err'] for r in rows),
             sum(r['err'] for r in rows), sum(r['figs'] for r in rows)))
    if bad:
        print('오류가 남은 노트북:', ', '.join(bad))
        print('→ 원인을 확인하고 %s 에 수정을 추가한 뒤 다시 실행하세요.' % (cfg['_workspace'] / 'fixes.py'))
    else:
        print('다음 단계: 보고서 작성 (Claude 에게 요청) → python batch_docx.py')


if __name__ == '__main__':
    main()
