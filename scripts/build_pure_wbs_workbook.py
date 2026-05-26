from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


output_dir = Path("outputs/wbs-simple-20260526")
output_dir.mkdir(parents=True, exist_ok=True)
output_path = output_dir / "WBS_업무정리_순수WBS.xlsx"

rows = [
    ["아이오테크 SEMS SL 개발", "FW 개발", "LCU보드 데모 준비", "미진행", date(2026, 5, 26), date(2026, 6, 5), "LCU보드 데모에 필요한 기능과 현재 자료 위치 확인", "https://www.notion.so/36c5c9ca0fd3811db6facdf61c7f349e"],
    ["아이오테크 SEMS SL 개발", "HW 개발", "회로도", "미진행", date(2026, 5, 26), date(2026, 5, 29), "최신 회로도 파일 위치 확인", "https://www.notion.so/36c5c9ca0fd381d985f1ce78fe292504"],
    ["아이오테크 SEMS SL 개발", "HW 개발", "PCB", "미진행", date(2026, 6, 1), date(2026, 6, 12), "PCB 관련 자료와 진행 상태 확인", "https://www.notion.so/36c5c9ca0fd381e6b45df9803e9e36c8"],
    ["아이오테크 SEMS SL 개발", "HW 개발", "아트웍", "미진행", date(2026, 6, 8), date(2026, 6, 19), "아트웍 진행 파일과 수정사항 확인", "https://www.notion.so/36c5c9ca0fd381b09e0cc1f1a7485102"],
    ["아이오테크 SEMS SL 개발", "통합/데모", "FW-HW 통합 테스트", "미진행", date(2026, 6, 15), date(2026, 6, 26), "데모 성공 기준과 테스트 항목 정리", "https://www.notion.so/36c5c9ca0fd381bcaddfe6762f1a4c84"],
    ["아이오테크 SEMS SL 개발", "단발 요청", "임시 요청/확인사항", "미진행", date(2026, 5, 26), date(2026, 6, 30), "들어오는 요청은 여기 먼저 적고 나중에 제자리로 이동", "https://www.notion.so/36c5c9ca0fd3817fba1adbaf6b8880a1"],
    ["위웨이크 앱 개발", "운동 앱", "기능/운영 정리", "미진행", date(2026, 5, 26), date(2026, 6, 12), "운동 앱의 현재 기능과 수정사항 정리", "https://www.notion.so/36c5c9ca0fd3812ba480d43f622e10a5"],
    ["위웨이크 앱 개발", "정부정책 앱", "기능/운영 정리", "미진행", date(2026, 5, 26), date(2026, 6, 12), "정부정책 앱의 현재 기능과 수정사항 정리", "https://www.notion.so/36c5c9ca0fd381b28a3ecd5ca9833abf"],
    ["개인 관리", "건강", "해야 할 일 정리", "미진행", date(2026, 5, 26), date(2026, 6, 2), "생각나는 일을 먼저 적기", "https://www.notion.so/36c5c9ca0fd381cd9214e2e65c659fee"],
    ["개인 관리", "돈/생활", "해야 할 일 정리", "미진행", date(2026, 5, 26), date(2026, 6, 2), "생각나는 일을 먼저 적기", "https://www.notion.so/36c5c9ca0fd38189b6b4dbe582dc9937"],
]

wb = Workbook()
ws = wb.active
ws.title = "WBS목록"
ws.sheet_view.showGridLines = False

headers = ["1Depth", "2Depth", "3Depth", "상태", "시작일", "종료일", "소요일", "다음 행동", "Notion"]
ws.append(headers)
for item in rows:
    ws.append(item[:6] + [None, item[6], "열기"])

for row in range(2, len(rows) + 2):
    ws.cell(row=row, column=7).value = f"=IF(OR(E{row}=\"\",F{row}=\"\"),\"\",F{row}-E{row}+1)"
    ws.cell(row=row, column=9).hyperlink = rows[row - 2][7]
    ws.cell(row=row, column=9).style = "Hyperlink"

header_fill = PatternFill("solid", fgColor="1F4E79")
light_fill = PatternFill("solid", fgColor="EAF3F8")
stripe_fill = PatternFill("solid", fgColor="F7FBFD")
thin_gray = Side(style="thin", color="7F7F7F")
blue_line = Side(style="thin", color="9CCFE5")

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

for row in ws.iter_rows(min_row=2, max_row=len(rows)+1, min_col=1, max_col=9):
    for cell in row:
        cell.font = Font(name="맑은 고딕", size=10)
        cell.alignment = Alignment(horizontal="center" if cell.column <= 7 else "left", vertical="center", wrap_text=True)
        cell.border = Border(left=thin_gray, right=thin_gray, top=blue_line, bottom=blue_line)
        if cell.column <= 3:
            cell.fill = light_fill
        elif cell.row % 2 == 0:
            cell.fill = stripe_fill

for col, width in {
    "A": 24,
    "B": 16,
    "C": 24,
    "D": 10,
    "E": 12,
    "F": 12,
    "G": 9,
    "H": 38,
    "I": 12,
}.items():
    ws.column_dimensions[col].width = width

for row in range(2, len(rows) + 2):
    ws.row_dimensions[row].height = 34
for col in ["E", "F"]:
    for row in range(2, len(rows) + 2):
        ws[f"{col}{row}"].number_format = "yyyy-mm-dd"

def merge_runs(sheet, col, start_row, values):
    run_start = 0
    for idx in range(1, len(values) + 1):
        if idx == len(values) or values[idx] != values[run_start]:
            if idx - run_start > 1:
                sheet.merge_cells(start_row=start_row + run_start, start_column=col, end_row=start_row + idx - 1, end_column=col)
                sheet.cell(row=start_row + run_start, column=col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            run_start = idx

merge_runs(ws, 1, 2, [r[0] for r in rows])
merge_runs(ws, 2, 2, [(r[0], r[1]) for r in rows])
ws.freeze_panes = "D2"

timeline = wb.create_sheet("일정표")
timeline.sheet_view.showGridLines = False

first = min(r[4] for r in rows).replace(day=1)
last_raw = max(r[5] for r in rows)
last = date(last_raw.year + (last_raw.month // 12), (last_raw.month % 12) + 1, 1) - timedelta(days=1)
dates = []
cursor = first
while cursor <= last:
    dates.append(cursor)
    cursor += timedelta(days=1)

last_col = 4 + len(dates) - 1
last_col_letter = get_column_letter(last_col)

timeline.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
timeline["A2"] = "■ WBS 작업 일정"
timeline["A2"].font = Font(name="맑은 고딕", size=12)
timeline["A2"].alignment = Alignment(horizontal="left", vertical="center")

timeline.merge_cells("A3:B4")
timeline["A3"] = "일정계획"
timeline.merge_cells("C3:C4")
timeline["C3"] = "3Depth 작업"

for cell_ref in ["A3", "C3"]:
    cell = timeline[cell_ref]
    cell.fill = PatternFill("solid", fgColor="EDEDED")
    cell.font = Font(name="맑은 고딕", bold=True, size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center")

month_start = 0
for idx, day in enumerate(dates + [None]):
    if idx == len(dates) or day.month != dates[month_start].month:
        start_col = 4 + month_start
        end_col = 4 + idx - 1
        timeline.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
        cell = timeline.cell(row=3, column=start_col)
        cell.value = f"{dates[month_start].month}월"
        cell.fill = PatternFill("solid", fgColor="EDEDED")
        cell.font = Font(name="맑은 고딕", bold=True, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        month_start = idx

for idx, day in enumerate(dates):
    cell = timeline.cell(row=4, column=4 + idx)
    cell.value = day
    cell.number_format = "d"
    cell.fill = PatternFill("solid", fgColor="EDEDED")
    cell.font = Font(name="맑은 고딕", bold=True, size=8)
    cell.alignment = Alignment(horizontal="center", vertical="center")

start_row = 5
for idx, item in enumerate(rows):
    row = start_row + idx
    timeline.cell(row=row, column=1).value = item[0]
    timeline.cell(row=row, column=2).value = item[1]
    timeline.cell(row=row, column=3).value = item[2]
    for col in range(1, last_col + 1):
        cell = timeline.cell(row=row, column=col)
        cell.font = Font(name="맑은 고딕", size=9)
        cell.alignment = Alignment(horizontal="center" if col <= 2 else "left", vertical="center", wrap_text=True)
        if col >= 4 and dates[col - 4].weekday() >= 5:
            cell.fill = PatternFill("solid", fgColor="D9D9D9")
        else:
            cell.fill = PatternFill("solid", fgColor="FFFFFF")

merge_runs(timeline, 1, start_row, [r[0] for r in rows])
merge_runs(timeline, 2, start_row, [(r[0], r[1]) for r in rows])

thin = Side(style="thin", color="000000")
date_grid = Side(style="thin", color="B7B7B7")
month_grid = Side(style="medium", color="000000")
month_start_cols = {4}
for idx in range(1, len(dates)):
    if dates[idx].month != dates[idx - 1].month:
        month_start_cols.add(4 + idx)

for row in range(3, start_row + len(rows)):
    for col in range(1, last_col + 1):
        cell = timeline.cell(row=row, column=col)
        if col >= 4:
            left = month_grid if col in month_start_cols else date_grid
            cell.border = Border(left=left, right=date_grid, top=thin, bottom=thin)
        else:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

bar_range = f"D{start_row}:{last_col_letter}{start_row + len(rows) - 1}"
for status, color in [("미진행", "FFC000"), ("진행중", "2F5597"), ("검토중", "92D050"), ("완료", "A6A6A6")]:
    timeline.conditional_formatting.add(
        bar_range,
        FormulaRule(
            formula=[f'=AND(D$4>=WBS목록!$E2,D$4<=WBS목록!$F2,WBS목록!$D2="{status}")'],
            fill=PatternFill("solid", fgColor=color),
        ),
    )

timeline.column_dimensions["A"].width = 24
timeline.column_dimensions["B"].width = 16
timeline.column_dimensions["C"].width = 31
for col in range(4, last_col + 1):
    timeline.column_dimensions[get_column_letter(col)].width = 3
timeline.row_dimensions[2].height = 20
timeline.row_dimensions[3].height = 18
timeline.row_dimensions[4].height = 18
for row in range(start_row, start_row + len(rows)):
    timeline.row_dimensions[row].height = 16.5
timeline.freeze_panes = "D5"

wb.save(output_path)
print(output_path)
