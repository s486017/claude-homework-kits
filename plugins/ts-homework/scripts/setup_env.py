# -*- coding: utf-8 -*-
"""1단계 — 작업 폴더와 파이썬 환경을 준비한다.

사용: python setup_env.py [--name 이름] [--id 학번] [--repo <github url>]
"""
import argparse
import json
import platform
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kit  # noqa: E402

# (pip 이름, import 이름)
BASE = [('numpy', 'numpy'), ('pandas', 'pandas'), ('matplotlib', 'matplotlib'),
        ('statsmodels', 'statsmodels'), ('scikit-learn', 'sklearn'),
        ('yfinance', 'yfinance'), ('requests', 'requests'),
        ('python-docx', 'docx'), ('pillow', 'PIL')]

FIXES_TEMPLATE = '''# -*- coding: utf-8 -*-
"""노트북별 코드 수정 목록.

{노트북명: [(셀번호, 원본문자열, 바꿀문자열, 이유), ...]}

실행 중 오류가 나면 여기에 한 줄 추가하고 다시 돌린다.
원본 노트북은 건드리지 않으므로, 무엇을 왜 고쳤는지가 그대로 기록으로 남는다.
old 와 new 를 같게 두면 "코드 수정 없이 환경만 조치했다"는 기록이 된다.
"""

FIXES = {}
'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--name', default=None, help='이름 (보고서 표지에 들어감)')
    ap.add_argument('--id', dest='sid', default=None, help='학번')
    ap.add_argument('--course', default=None, help='과목명')
    ap.add_argument('--assignment', type=int, default=None, help='과제 번호 (기본 1)')
    ap.add_argument('--assignment-title', default=None, help='과제 제목')
    ap.add_argument('--no-prefix', action='store_true',
                    help='산출물 이름에 과제 번호를 붙이지 않음')
    ap.add_argument('--repo', default=None, help='노트북이 있는 GitHub 저장소 URL')
    ap.add_argument('--notebook-path', default=None, help='저장소 안의 노트북 폴더 (기본 notebooks)')
    ap.add_argument('--skip-install', action='store_true')
    a = ap.parse_args()

    ws = kit.workspace()
    f = ws / kit.CONFIG_NAME
    cfg = json.loads(f.read_text(encoding='utf-8')) if f.exists() else {}
    for key, val in (('student_name', a.name), ('student_id', a.sid),
                     ('course', a.course), ('notebook_repo', a.repo),
                     ('notebook_path', a.notebook_path)):
        if val:
            cfg[key] = val
    asg = dict(kit.DEFAULTS['assignment'], **cfg.get('assignment', {}))
    if a.assignment is not None:
        asg['number'] = a.assignment
    if a.assignment_title:
        asg['title'] = a.assignment_title
    if a.no_prefix:
        asg['prefix'] = False
    cfg['assignment'] = asg
    cfg.setdefault('dirs', kit.DEFAULTS['dirs'])
    cfg.setdefault('docx', kit.DEFAULTS['docx'])
    f.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')

    full = kit.load()
    for k in full['dirs']:
        kit.path_for(full, k)
    fixes = ws / 'fixes.py'
    if not fixes.exists():
        fixes.write_text(FIXES_TEMPLATE, encoding='utf-8')

    print('작업 폴더 :', ws)
    print('파이썬    :', sys.version.split()[0], '|', sys.executable)
    print('실행 환경 :', platform.system(), platform.release())
    venv = sys.prefix != getattr(sys, 'base_prefix', sys.prefix)
    print('가상환경  :', '예' if venv else '아니오 (권장: python -m venv .venv 후 활성화)')

    if a.skip_install:
        print('패키지    : 설치 건너뜀')
    else:
        installed = kit.ensure_packages(BASE)
        print('패키지    :', ('새로 설치 ' + ', '.join(installed)) if installed else '이미 전부 설치됨')

    font = kit.korean_font()
    print('한글 폰트 :', font or '없음 — 그래프 한글이 깨집니다 (맥: AppleGothic, 리눅스: NanumGothic 설치)')
    missing = [k for k in ('student_name', 'student_id') if not cfg.get(k)]
    if missing:
        print()
        print('[확인 필요] 이름/학번이 비어 있습니다. 보고서 표지에 들어가므로 채워야 합니다:')
        print('  python setup_env.py --name "홍길동" --id "20260001"')
    print()
    print('다음 단계: python fetch_notebooks.py')


if __name__ == '__main__':
    main()
