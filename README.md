# WBS 서버 공유 방식: GitHub 원본 저장소

로컬 zip 전달이 아니라, GitHub 저장소를 원본 서버로 두고 모든 PC에서 같은 파일을 기준으로 작업하는 방식입니다.

## 권장 구조

GitHub 저장소 하나를 만듭니다.

현재 WBS 관리 저장소:

```text
https://github.com/sehuniotech-ctrl/IoTech-info
```

SEMS 작업 참고 저장소:

```text
https://github.com/sehuniotech-ctrl/Iotech_sems
```

저장소 안에는 아래처럼 둡니다.

```text
IoTech-info/
  README.md
  PROMPT_FOR_GITHUB_WBS_SYNC.md
  data/
    wbs_items.json
  scripts/
    build_pure_wbs_workbook.py
  files/
    WBS_latest.xlsx
    WBS_업무정리_순수WBS.xlsx
  archive/
    2026-05-26_WBS_latest.xlsx
    2026-05-26_WBS_업무정리_순수WBS.xlsx
```

## 역할

```text
files/WBS_업무정리_순수WBS.xlsx
```

현재 최신 WBS 파일입니다. 다른 PC에서는 이 파일을 내려받아 열면 됩니다.

```text
files/WBS_latest.xlsx
```

한글 파일명 다운로드 문제가 있을 때 쓰는 동일한 최신 WBS 파일입니다.

```text
archive/
```

날짜별 백업본을 보관합니다. 실수로 파일이 망가졌을 때 이전 버전으로 돌아가기 쉽습니다.

```text
scripts/build_pure_wbs_workbook.py
```

Excel 파일을 다시 생성하는 스크립트입니다.

```text
data/wbs_items.json
```

나중에 WBS 항목을 코드/자동화로 관리하고 싶을 때 원본 데이터 역할을 합니다.

```text
PROMPT_FOR_GITHUB_WBS_SYNC.md
```

다른 컴퓨터의 Codex/ChatGPT에게 줄 작업 지시문입니다.

## 작업 흐름

### 1. 메인 PC에서 수정

1. Excel 파일 수정
2. 저장
3. Git commit
4. GitHub push

```powershell
git add files/WBS_업무정리_순수WBS.xlsx scripts/build_pure_wbs_workbook.py data/wbs_items.json PROMPT_FOR_GITHUB_WBS_SYNC.md README.md
git commit -m "Update WBS workbook"
git push
```

### 2. 다른 PC에서 이어 작업

```powershell
git pull
```

그 다음:

```text
files/WBS_업무정리_순수WBS.xlsx
```

파일을 열어 이어 작업합니다.

### 3. 다른 PC에서 수정 후 다시 서버에 반영

```powershell
git add files/WBS_업무정리_순수WBS.xlsx
git commit -m "Update WBS from other PC"
git push
```

## Notion의 역할

Notion은 파일 원본 저장소가 아니라, 아래 링크를 모아두는 허브로 쓰는 것이 좋습니다.

- 최신 Excel 파일 GitHub 링크
- GitHub 저장소 링크
- WBS 진행 요약 페이지
- 각 작업 상세 페이지

## 주의

- 같은 파일을 여러 PC에서 동시에 수정하면 충돌이 날 수 있습니다.
- 작업 전에는 항상 `git pull`을 먼저 합니다.
- 작업 후에는 바로 `git commit`과 `git push`를 합니다.
- Excel 파일은 바이너리라 충돌이 나면 자동 병합이 어렵습니다. 한 번에 한 PC에서만 수정하는 방식이 좋습니다.
