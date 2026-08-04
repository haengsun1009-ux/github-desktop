# -*- coding: utf-8 -*-
"""
28.1 실행절차 4번 / 25. Self QA 항목7,8,10 — 수식 참조 감사 + Pipeline Parity + 값 타당성 검증
LibreOffice로 재계산된 파일(recalc/*.xlsx)의 캐시된 값과, 원본 파일(수식 포함)을 함께 대조한다.
"""
import openpyxl
import re

ORIG = "/workspace/kdis_output/KDIS_괄사_20260804.xlsx"
CALC = "/workspace/kdis_output/recalc/KDIS_괄사_20260804.xlsx"

wb_formula = openpyxl.load_workbook(ORIG, data_only=False)
wb_value = openpyxl.load_workbook(CALC, data_only=True)

errors = []
warnings = []

# ------------------------------------------------------------------
# 1) 수식 참조 감사: 모든 셀의 수식에서 '시트명'! 패턴을 추출해 실제 존재 여부 확인
# ------------------------------------------------------------------
sheet_names = set(wb_formula.sheetnames)
sheet_ref_pattern = re.compile(r"'([^']+)'!")
checked_formulas = 0
bad_refs = []
for ws in wb_formula.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                checked_formulas += 1
                for m in sheet_ref_pattern.finditer(cell.value):
                    ref_sheet = m.group(1)
                    if ref_sheet not in sheet_names:
                        bad_refs.append((ws.title, cell.coordinate, ref_sheet, cell.value))

print(f"[수식 참조 감사] 전수 검사한 수식 셀 수: {checked_formulas}")
if bad_refs:
    for b in bad_refs:
        errors.append(f"존재하지 않는 시트명 참조: {b}")
else:
    print("  -> 존재하지 않는 시트명 참조: 0건 (PASS)")

# ------------------------------------------------------------------
# 2) 명시적 오류값(#REF!, #VALUE!, #N/A 등) 전수 스캔 — 재계산된 파일 기준
# ------------------------------------------------------------------
ERROR_TOKENS = ("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!")
explicit_errors = []
for ws in wb_value.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and any(tok in cell.value for tok in ERROR_TOKENS):
                explicit_errors.append((ws.title, cell.coordinate, cell.value))

print(f"[명시적 수식오류 스캔] 발견 건수: {len(explicit_errors)}")
for e in explicit_errors:
    print("   ", e)
    errors.append(f"명시적 수식오류: {e}")

# ------------------------------------------------------------------
# 3) 값 타당성 샘플 검증: 02 탭 5개 항목 합 == 점수(자동) 열, 범위(0~100) 확인
# ------------------------------------------------------------------
ws2 = wb_value["02_Buyer_Score_Detail"]
row = 4
score_mismatch = []
range_issues = []
n_checked = 0
while ws2.cell(row=row, column=1).value not in (None, ""):
    n_checked += 1
    vals = [ws2.cell(row=row, column=c).value for c in range(2,7)]  # B..F (5 sub scores)
    total = ws2.cell(row=row, column=7).value  # G
    grade = ws2.cell(row=row, column=8).value  # H
    name = ws2.cell(row=row, column=1).value
    if vals and total is not None:
        if sum(vals) != total:
            score_mismatch.append((name, vals, total))
    if total is None or not (0 <= total <= 100):
        range_issues.append((name, total))
    # 등급 정합성
    expected_grade = "A+" if total>=90 else "A" if total>=80 else "B+" if total>=65 else "B" if total>=50 else "C"
    if grade != expected_grade:
        range_issues.append((name, total, grade, "등급불일치, 기대값:", expected_grade))
    row += 1

print(f"[값 타당성 검증] 02탭 대조 대상: {n_checked}건, 5항목합-점수 불일치: {len(score_mismatch)}건, 범위/등급 이상: {len(range_issues)}건")
for m in score_mismatch:
    errors.append(f"점수 합계 불일치: {m}")
for r_ in range_issues:
    errors.append(f"점수/등급 범위 이상: {r_}")

# ------------------------------------------------------------------
# 4) Pipeline Parity: 00_Summary 확보Buyer 수 == 01/02/05 실제 행수(값 대조는 재계산본 기준)
# ------------------------------------------------------------------
def count_nonempty(ws, col, start_row):
    n = 0
    r = start_row
    while True:
        v = ws.cell(row=r, column=col).value
        if v in (None, ""):
            break
        n += 1
        r += 1
    return n

ws1v = wb_value["01_TOP_Buyer_List"]
ws2v = wb_value["02_Buyer_Score_Detail"]
ws5v = wb_value["05_CRM_Import"]
ws0v = wb_value["00_Summary"]

n01 = count_nonempty(ws1v, 2, 3)   # B열(회사명), 3행부터
n02 = count_nonempty(ws2v, 1, 4)   # A열, 4행부터
n05 = count_nonempty(ws5v, 1, 4)   # A열, 4행부터

summary_buyer_count = None
for r in range(1, 30):
    lbl = ws0v.cell(row=r, column=1).value
    if lbl and "확보 Buyer" in str(lbl):
        summary_buyer_count = ws0v.cell(row=r, column=2).value
        break

print(f"[Pipeline Parity] 00_Summary 확보Buyer={summary_buyer_count} / 01행수={n01} / 02행수={n02} / 05행수={n05}")
if not (summary_buyer_count == n01 == n02 == n05):
    errors.append(f"Pipeline Parity 불일치: summary={summary_buyer_count}, 01={n01}, 02={n02}, 05={n05}")
else:
    print("  -> 100% 일치 (PASS)")

# ------------------------------------------------------------------
# 5) 01탭 점수 열 내림차순 정렬 확인 (정렬 붕괴 = 참조 오류 신호)
# ------------------------------------------------------------------
scores_col = []
r = 3
while True:
    v = ws1v.cell(row=r, column=6).value  # F열=점수(자동)... wait column 6 is 점수(자동)? header: 순위1,회사명2,채널유형3,셀프/MD4,등급5,점수6,연락처7...
    name = ws1v.cell(row=r, column=2).value
    if name in (None, ""):
        break
    scores_col.append(v)
    r += 1
is_sorted = all(scores_col[i] >= scores_col[i+1] for i in range(len(scores_col)-1))
print(f"[정렬 검증] 01탭 점수 내림차순 정렬 여부: {is_sorted} (총 {len(scores_col)}건)")
if not is_sorted:
    errors.append("01_TOP_Buyer_List 점수 내림차순 정렬 붕괴 감지")

# ------------------------------------------------------------------
# 6) 검증상태(확인됨/부분확인) 건수가 03/01/02/05 탭 간 일치하는지("8. 부분확인 포함여부")
# ------------------------------------------------------------------
def count_status(ws, col, start_row, status):
    n = 0
    r = start_row
    while True:
        name = ws.cell(row=r, column=2).value if ws.title=="01_TOP_Buyer_List" else ws.cell(row=r, column=1).value
        if name in (None, ""):
            break
        val = ws.cell(row=r, column=col).value
        if val == status:
            n += 1
        r += 1
    return n

n_confirmed_01 = count_status(ws1v, 8, 3, "확인됨")
n_partial_01 = count_status(ws1v, 8, 3, "부분확인")
n_confirmed_05 = count_status(ws5v, 9, 4, "확인됨")
n_partial_05 = count_status(ws5v, 9, 4, "부분확인")
print(f"[검증등급 일치] 01: 확인됨{n_confirmed_01}/부분확인{n_partial_01}  |  05: 확인됨{n_confirmed_05}/부분확인{n_partial_05}")
if not (n_confirmed_01 == n_confirmed_05 and n_partial_01 == n_partial_05):
    errors.append("01/05 탭 검증상태(확인됨/부분확인) 건수 불일치")

# ------------------------------------------------------------------
# 결과 출력
# ------------------------------------------------------------------
print("\n================ 감사 결과 ================")
if errors:
    print(f"FAIL — {len(errors)}건 오류 발견:")
    for e in errors:
        print(" -", e)
else:
    print("ALL PASS — 수식참조/오류값/값타당성/Pipeline Parity/정렬/검증등급 전항목 이상 없음")
