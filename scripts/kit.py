# -*- coding: utf-8 -*-
"""공통 설정 — 작업 폴더(ts_kit.config.json)를 읽고 경로·폰트를 정한다.

모든 스크립트는 **현재 작업 폴더**를 기준으로 동작한다. 플러그인이 어디에
설치돼 있든, 운영체제가 무엇이든 상관없이 돌아가야 하므로 경로를 하드코딩하지 않는다.
"""
import json
import os
import sys
from pathlib import Path

CONFIG_NAME = 'ts_kit.config.json'

DEFAULTS = {
    'student_name': '',
    'student_id': '',
    'course': '',
    'assignment': {
        'number': 1,        # 과제 번호 — 산출물 이름 앞에 '과제1_' 로 붙는다
        'title': '',        # 과제 제목 (보고서 표지에 들어감)
        'prefix': True,     # False 면 번호를 파일 이름에 붙이지 않는다
    },
    'notebook_repo': '',
    'notebook_path': 'notebooks',
    'dirs': {
        'notebooks': '노트북',
        'results': '실행결과',
        'reports': '보고서',
        'docx': '보고서_docx',
    },
    'docx': {
        'font': '맑은 고딕',
        'mono_font': 'Consolas',
        'body_size': 10.5,
        'max_image_width_in': 6.0,
        'page_break_per_section': False,
    },
}


def _merge(base, over):
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = _merge(base[k], v) if isinstance(v, dict) and isinstance(base.get(k), dict) else v
    return out


def workspace(start=None):
    """ts_kit.config.json 이 있는 폴더를 위로 올라가며 찾는다. 없으면 현재 폴더."""
    p = Path(start or os.getcwd()).resolve()
    for d in [p, *p.parents]:
        if (d / CONFIG_NAME).exists():
            return d
    return p


def load(start=None):
    ws = workspace(start)
    f = ws / CONFIG_NAME
    cfg = _merge(DEFAULTS, json.loads(f.read_text(encoding='utf-8')) if f.exists() else {})
    cfg['_workspace'] = ws
    cfg['_config_exists'] = f.exists()
    return cfg


def prefix(cfg):
    """산출물 이름 앞에 붙일 과제 번호 — 예: '과제1_'. 끄면 빈 문자열."""
    a = cfg.get('assignment') or {}
    if not a.get('prefix', True) or not a.get('number'):
        return ''
    return '과제%s_' % a['number']


def label(cfg):
    """보고서 표지에 쓸 과제 이름 — 예: '과제 1 — 시계열 모델 실습'."""
    a = cfg.get('assignment') or {}
    if not a.get('number'):
        return a.get('title', '')
    return '과제 %s%s' % (a['number'], (' — ' + a['title']) if a.get('title') else '')


def path_for(cfg, key):
    """설정된 하위 폴더의 절대 경로 (없으면 만든다)."""
    d = cfg['_workspace'] / cfg['dirs'][key]
    d.mkdir(parents=True, exist_ok=True)
    return d


def korean_font():
    """설치된 한글 폰트 중 matplotlib이 쓸 수 있는 것을 고른다 (윈도우/맥/리눅스)."""
    try:
        from matplotlib import font_manager
    except ImportError:
        return None
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in ('Malgun Gothic', 'AppleGothic', 'NanumGothic', 'NanumBarunGothic',
                 'Noto Sans CJK KR', 'Noto Sans KR', 'Gulim', 'Batang'):
        if name in installed:
            return name
    return None


def ensure_packages(pkgs, quiet=True):
    """빠진 패키지만 현재 파이썬에 설치한다. (venv 안이면 venv에 들어간다)"""
    import importlib
    import subprocess
    missing = []
    for pkg, mod in pkgs:
        try:
            importlib.import_module(mod)
        except ImportError:
            missing.append(pkg)
    if missing:
        cmd = [sys.executable, '-m', 'pip', 'install',
               '--index-url', 'https://pypi.org/simple', *missing]
        if quiet:
            cmd.append('-q')
        subprocess.run(cmd, check=True)
    return missing
