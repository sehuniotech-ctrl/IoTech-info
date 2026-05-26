아래 프롬프트를 다른 PC의 Codex/ChatGPT에 그대로 붙여넣으세요.

---

나는 GitHub 저장소를 원본 서버로 두고 WBS 업무관리 파일을 관리하고 있다.

원본 저장소:

- WBS 관리 저장소: https://github.com/sehuniotech-ctrl/IoTech-info
- SEMS 작업 참고 저장소: https://github.com/sehuniotech-ctrl/Iotech_sems

목표:

1. GitHub 저장소의 `files/WBS_업무정리_순수WBS.xlsx`가 최신 원본이다.
2. 작업 전에는 반드시 `git pull`로 최신 상태를 받는다.
3. 작업 후에는 `git add`, `git commit`, `git push`로 서버에 반영한다.
4. Excel 파일은 병합 충돌이 어려우므로, 수정 전후로 현재 branch와 변경 파일을 확인한다.

WBS 파일 구조:

- `WBS목록` 시트
  - 1Depth
  - 2Depth
  - 3Depth
  - 상태
  - 시작일
  - 종료일
  - 소요일
  - 다음 행동
  - Notion

- `일정표` 시트
  - 왼쪽: 1Depth / 2Depth / 3Depth
  - 오른쪽: 날짜별 일정 막대
  - 같은 1Depth/2Depth는 셀 병합
  - 날짜 칸마다 세로줄
  - 월 경계는 굵은 세로줄

상태값:

- 미진행
- 진행중
- 검토중
- 완료

색상:

- 미진행: 노란색
- 진행중: 파란색
- 검토중: 초록색
- 완료: 회색

주의사항:

- Excel 표(Table) 기능은 쓰지 않는다. 병합 셀과 충돌해서 Excel 복구 경고가 날 수 있다.
- Notion 열의 `열기` 링크는 유지한다.
- 수정 후 파일이 Excel에서 복구 경고 없이 열리는지 확인한다.
- 가능하면 `scripts/build_pure_wbs_workbook.py`도 함께 관리한다.

먼저 현재 저장소 상태를 확인하고, 내가 요청하는 변경만 반영해줘.

---
