# -*- coding: utf-8 -*-
"""4단계 — 보고서 폴더의 .md 를 전부 .docx 로 변환한다 (병렬).

사용: python batch_docx.py [--workers 8] [이름 ...]
"""
import argparse
import os
import sys
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kit  # noqa: E402


def job(args):
    md, out_dir = args
    from md2docx import convert
    t0 = time.time()
    name = Path(md).stem
    out = os.path.join(out_dir, name + '.docx')
    try:
        st = convert(md, out)
        st.update(file=name, ok=True, sec=round(time.time() - t0, 2),
                  kb=round(os.path.getsize(out) / 1024.0))
        return st
    except Exception:
        return {'file': name, 'ok': False, 'sec': round(time.time() - t0, 2),
                'error': traceback.format_exc()[-800:]}


def main():
    cfg = kit.load()
    ap = argparse.ArgumentParser()
    ap.add_argument('names', nargs='*', help='특정 보고서만 (이름 일부)')
    ap.add_argument('--workers', type=int, default=min(8, (os.cpu_count() or 4)))
    a = ap.parse_args()

    src = kit.path_for(cfg, 'reports')
    dst = kit.path_for(cfg, 'docx')
    files = sorted(src.glob('*.md'))
    if a.names:
        files = [f for f in files if any(k in f.name for k in a.names)]
    if not files:
        raise SystemExit('변환할 보고서(.md)가 없습니다: %s' % src)

    t0 = time.time()
    rows = []
    with ProcessPoolExecutor(max_workers=a.workers) as ex:
        futs = [ex.submit(job, (str(f), str(dst))) for f in files]
        for i, fut in enumerate(as_completed(futs), 1):
            r = fut.result()
            rows.append(r)
            print('[%d/%d] %-40s %s' % (i, len(files), r['file'], 'OK' if r['ok'] else '실패'))

    fail = [r for r in rows if not r['ok']]
    print()
    print('%d개 변환 | %.1f초 | 워커 %d개' % (len(rows), time.time() - t0, a.workers))
    print('그림 %d장 | 표 %d개 | 코드블록 %d개 | 그림 누락 %d건'
          % (sum(r.get('img_ok', 0) for r in rows), sum(r.get('tables', 0) for r in rows),
             sum(r.get('code', 0) for r in rows), sum(r.get('img_missing', 0) for r in rows)))
    for r in fail:
        print('[실패] %s\n%s' % (r['file'], r.get('error', '')))
    print('저장 위치:', dst)
    if not fail:
        print('다음 단계: python verify_docx.py')
    sys.exit(1 if fail else 0)


if __name__ == '__main__':
    main()
