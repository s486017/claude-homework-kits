# 노트북 과제 키트 (ts-homework)

강의에서 받은 주피터 노트북을 **직접 실행하고, 그 결과로 Word 보고서까지 만드는** 과제를
Claude 와 함께 처리하는 도구 모음입니다.

```
① 준비            ② 실행                 ③ 보고서            ④ 문서화
환경·설정      →  노트북 내려받기·실행 →  마크다운 작성   →  .docx 변환·검증
ts_kit.config    실행결과/               보고서/            보고서_docx/
```

## 설치

### Claude Code 를 쓰는 경우

```
/plugin marketplace add s486017/claude-homework-kits
/plugin install ts-homework@homework-kits
```

### Claude 데스크톱 앱 / 웹을 쓰는 경우

`plugins/ts-homework/skills/ts-homework/` 폴더를 zip 으로 압축해 **설정 → Capabilities → Skills** 에 올립니다.
(스크립트 실행이 필요한 단계는 Claude Code 쪽이 편합니다.)

## 쓰는 법

과제 폴더를 하나 만들고 그 안에서 Claude 에게 이렇게 말하면 됩니다.

```
/과제-1-준비   이름 홍길동, 학번 20260001, 1번 과제
             (노트북은 어디서 받았는지만 알려 주면 됩니다 — 아래 참고)
/과제-2-실행
/과제-3-보고서
/과제-4-문서화
```

명령을 외울 필요는 없습니다. "노트북 과제 좀 도와줘" 라고 해도 스킬이 알아서 잡습니다.

### GitHub 계정이 없어도 됩니다

**본인 저장소는 필요 없습니다.** 노트북을 가져올 방법만 셋 중 하나 고르면 됩니다.

| 상황 | 방법 |
|---|---|
| 강의 노트북이 공개 저장소에 있음 | `--repo https://github.com/<계정>/<저장소>` |
| 메일·LMS 로 `.ipynb` 파일을 받음 | `--from <폴더 또는 파일 경로>` (네트워크도 불필요) |
| Colab 링크만 있음 | `--url <노트북 주소>` |

## 단계별로 하는 일

| 단계 | 스크립트 | 결과물 |
|---|---|---|
| ① 준비 | `setup_env.py` | 작업 폴더, `ts_kit.config.json`, `fixes.py`, 패키지 설치, 한글 폰트 확인 |
| ② 실행 | `fetch_notebooks.py` → `run_notebooks.py` | `실행결과/<노트북>/` 에 셀별 출력·그림·요약 |
| ③ 보고서 | (Claude 가 작성) | `보고서/과제1_*.md` — 코드 설명 → 실행 결과 → 결과 정리 |
| ④ 문서화 | `batch_docx.py` → `verify_docx.py` | `보고서_docx/과제1_*.docx` + 검증 결과 |

## 설계에서 신경 쓴 것

- **원본 노트북은 고치지 않습니다.** 최신 라이브러리에서 깨지는 셀은 작업 폴더의 `fixes.py` 에
  `(셀번호, 원본, 수정, 이유)` 로 기록하고 실행할 때 적용합니다. 그 기록이 그대로 보고서의
  "오류 및 수정 내역"이 됩니다 — 과제에서 점수가 되는 부분입니다.
- **오류가 나도 끝까지 실행합니다.** 무엇이 깨졌는지 한 번에 보는 편이 빠릅니다.
- **제출 전에 검증합니다.** 그림 수, 목차 링크 연결, 링크 대상 파일 존재, 본문 유실을 대조하고,
  Word 나 LibreOffice 가 있으면 실제로 열어 페이지 수와 표시된 그림까지 확인합니다.
- **경로를 하드코딩하지 않습니다.** 작업 폴더 기준으로 동작하고, 윈도우·맥·리눅스에서 같이 돕니다.
- **과제마다 번호가 붙습니다.** 보고서와 docx 는 `과제1_...` 로 저장되고, 다음 과제는 번호만
  올리면 같은 폴더에서 이어서 할 수 있습니다. 노트북 실행 결과는 공유해서 다시 쓰지 않아도 됩니다.

## 설정 (`ts_kit.config.json`)

```json
{
  "student_name": "홍길동",
  "student_id": "20260001",
  "assignment": { "number": 1, "title": "시계열 모델 실습", "prefix": true },
  "notebook_repo": "https://github.com/<계정>/<저장소>",
  "notebook_path": "notebooks",
  "dirs": { "notebooks": "노트북", "results": "실행결과", "reports": "보고서", "docx": "보고서_docx" },
  "docx": {
    "font": "맑은 고딕",
    "mono_font": "Consolas",
    "body_size": 10.5,
    "max_image_width_in": 6.0,
    "page_break_per_section": false
  }
}
```

## 준비물

- Python 3.10 이상 (가상환경 권장)
- 패키지는 `setup_env.py` 가 자동 설치: numpy, pandas, matplotlib, statsmodels, scikit-learn,
  yfinance, requests, python-docx, pillow
- 한글 폰트 — 윈도우는 기본 설치(맑은 고딕), 맥은 AppleGothic, 리눅스는 `fonts-nanum` 필요
- (선택) 실제 렌더링 검증용: 윈도우 + MS Word + `pywin32`, 또는 LibreOffice. 둘 다 `pymupdf` 필요

## 주의

- **강의안 PDF 등 저작권이 있는 자료는 저장소에 올리지 마세요.** 이 키트는 공개된 GitHub
  노트북을 내려받아 쓰는 것을 전제로 합니다.
- 이름·학번은 각자 `setup_env.py` 로 설정합니다. 설정 파일을 그대로 복사해 쓰지 마세요.
- 보고서 내용은 각자 실행 결과를 보고 쓰는 것입니다. 같은 노트북이라도 데이터를 받는 시점에
  따라 숫자가 달라집니다.

## 라이선스

MIT. 마음껏 쓰고 고치고 나눠 주세요. 고친 게 쓸 만하면 PR 도 환영합니다.
