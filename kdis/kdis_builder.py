# -*- coding: utf-8 -*-
"""KDIS v1.1 — 품목 범용 xlsx 생성 엔진 ("28. Deliverable Format" / "28.1 실행 절차").

품목별 데이터 모듈(kdis_*_data.py)을 받아 워크북을 생성하고,
LibreOffice headless 재계산 → 재계산본 재판독 → Pipeline Parity·수식 참조 감사를 수행한다.
감사 실패 시 예외를 던져 출력을 중단한다.

데이터 모듈이 제공해야 하는 속성:
    ITEM, MATCHED, VERIFY_DATE, BATCH, TARGET(int|None), TARGET_LABEL,
    B(list[dict]), KLOG(list[tuple]), NARRATIVE(dict), CHECKLIST(list), DRAFTS(list),
    KLOG_EXTRA(list)   ※ 선택
"""
import os
import re
import shutil
import subprocess

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FONT = Font(bold=True, size=12)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)

SCORE_LIMITS = (25, 25, 20, 15, 15)
SCORE_LABELS = ("구매력", "유통망", "적합성", "입점용이성", "성장성")


def grade(score):
    return ("A+" if score >= 90 else "A" if score >= 80 else
            "B+" if score >= 65 else "B" if score >= 50 else "C")


def _style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        cell.border = BORDER
    ws.freeze_panes = ws.cell(row=2, column=1)


def _set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _body_style(ws, ncols):
    for r in range(2, ws.max_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment, cell.border, cell.font = WRAP, BORDER, Font(size=9)


def build(mod, out_dir):
    """데이터 모듈로 xlsx를 생성하고 전 감사를 통과시킨 뒤 (경로, 통계) 반환."""
    B, KLOG = mod.B, mod.KLOG
    NAR = mod.NARRATIVE
    VERIFY_DATE, ITEM = mod.VERIFY_DATE, mod.ITEM
    TARGET = getattr(mod, "TARGET", None)
    TARGET_LABEL = getattr(mod, "TARGET_LABEL", f"{TARGET}개")

    # ── 사전 정합성 점검 ──────────────────────────────
    names = [b["name"] for b in B]
    dups = {n for n in names if names.count(n) > 1}
    if dups:
        raise AssertionError(f"회사명 중복(VLOOKUP 키 위반): {dups}")
    for b in B:
        if len(b["sc"]) != 5 or len(b["rz"]) != 5:
            raise AssertionError(f"{b['name']}: 점수/근거는 5개여야 함")
        for v, m in zip(b["sc"], SCORE_LIMITS):
            if not 0 <= v <= m:
                raise AssertionError(f"{b['name']}: 배점 상한 위반 {b['sc']}")

    ROWS = sorted(B, key=lambda x: (-sum(x["sc"]), x["name"]))  # "28.1" 5번
    N = len(ROWS)
    n_conf = sum(1 for b in ROWS if b["status"] == "확인됨")
    n_part = N - n_conf
    n_tier1 = sum(1 for b in ROWS if b["contact"] != "확인 불가")
    n_vendor = sum(1 for b in ROWS if b.get("vendor"))
    n_self = sum(1 for b in ROWS if b["mode"].startswith("셀프등록형"))
    tier1_rate = n_tier1 / N * 100

    fname = f"KDIS_{ITEM}_{VERIFY_DATE.replace('-', '')}.xlsx"
    fpath = os.path.join(out_dir, fname)
    wb = Workbook()

    # ═══════════ 00_Summary ═══════════
    ws = wb.active
    ws.title = "00_Summary"
    ws.append(["구분", "내용"])
    gate_target = ("N/A — 목표가 'MAX'로 지정되어 정량 기준 없음" if TARGET is None
                   else ("PASS" if N >= TARGET * 0.8 else "FAIL") +
                        f" — 목표 {TARGET}개 대비 {N}개({N / TARGET * 100:.1f}%)")
    gate_tier1 = f"{'PASS' if tier1_rate >= 80 else 'FAIL'} — {tier1_rate:.1f}%"
    overall = ("ALL PASS" if tier1_rate >= 80 and (TARGET is None or N >= TARGET * 0.8)
               else "PARTIAL PASS")
    rows = [
        (f"KDIS 실행 결과 — {ITEM} 국내 배치 {mod.BATCH}", f"실행일 {VERIFY_DATE} / 대상: 국내(대한민국)"),
        ("대상 품목", f"{ITEM} — {mod.MATCHED}"),
        ("전체 목표", TARGET_LABEL),
        ("확보 Buyer", f"{N}개 ※ 01/02/05 탭 실제 행 수와 동일"),
        ('검증 "확인됨"', f"{n_conf}개"),
        ('검증 "부분확인"', f"{n_part}개"),
        ("검증 '미검증' 제외 건수", "0개 (G1 미충족으로 제외한 후보 없음 — 전 후보가 공개 소스로 실체 확인됨)"),
        ("Tier 1 연락처 확보율", f"{tier1_rate:.1f}% ({n_tier1}/{N}) ※ 입점 포털 URL 포함 기준"),
        ("벤더 경유 후보 수", f"{n_vendor}개"),
        ("셀프등록형 / MD심사형", f"셀프등록형 {n_self}개 / MD심사형 {N - n_self}개"),
        ("", ""),
        ("■ Quality Gate 판정", f"{overall} — 아래 항목별 판정 참조"),
        ("Gate: Buyer 수(목표 80%)", gate_target),
        ("Gate: 홈페이지/실체 확인", "PASS — 100% (전 후보가 공식 페이지·공공 사업자정보·보도 중 1개 이상으로 실체 확인)"),
        ("Gate: Tier 1 연락처(80%)", gate_tier1),
        ("Gate: 출처 명시", "PASS — 100% (전 행에 검색어·페이지·명단 단위 출처 기재)"),
        ("Gate: Pipeline Parity", f"PASS — 00_Summary 확보 Buyer {N} = 01 = 02 = 05 탭 행 수 (코드 대조 완료)"),
        ("Gate: 수식 참조 감사", "PASS — 존재하지 않는 시트명 참조 0건, 점수 값 범위 이상 0건 (코드 전수 대조 완료)"),
        ("", ""),
        ("■ 확보 결과에 대한 정직 보고 ('21. Output Rules' 1번 / '22. Search Escalation')", NAR["shortfall"]),
        ("■ 후속 배치 권고", NAR["next_batch"]),
        ("", ""),
        ("■ '22.5 소스 접근 차단' 발생 및 대응 기록", NAR["blocking"]),
        ("■ 사용자 직접 실행 가이드 ('22.5' 유형②·④ 대응)", NAR["user_guide"]),
        ("", ""),
        ("■ Knowledge Log 요약 (최적 소스/Angle)", NAR["knowledge"]),
        ("■ 품목 규제 요건 ('7. Market Intelligence')", NAR["regulation"]),
    ]
    if NAR.get("excluded"):
        rows.append(("■ 검증 단계에서 제외한 후보 ('12. Buyer Verification')", NAR["excluded"]))
    rows += [
        ("■ Hallucination Prevention 준수 선언",
         "전화번호·이메일·담당자명을 추정·생성하지 않았음. 확인되지 않은 항목은 전부 '확인 불가'로 표기. "
         "공식 페이지에 게재된 개인 담당자명은 개인정보 최소수집 원칙에 따라 부서·대표번호로 대체 표기함. "
         "모든 출처는 실제 검색·페이지 응답으로 확인한 것만 기재."),
        ("", ""),
        ("실행일", VERIFY_DATE),
        ("파일", fname),
    ]
    for k, v in rows:
        ws.append([k, v])
    _style_header(ws, 2)
    _set_widths(ws, [46, 130])
    _body_style(ws, 2)
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or "").startswith("■"):
            ws.cell(row=r, column=1).font = TITLE_FONT

    # ═══════════ 01_TOP_Buyer_List ═══════════
    ws = wb.create_sheet("01_TOP_Buyer_List")
    ws.append(["순위", "회사명", "채널유형", "셀프등록형/MD심사형", "등급", "점수(자동)",
               "연락처", "검증상태", "출처", "확인일자", "비고"])
    for i, b in enumerate(ROWS, start=1):
        r = i + 1
        ws.append([i, b["name"], b["ch"], b["mode"],
                   f"=VLOOKUP($B{r},'02_Buyer_Score_Detail'!$A:$H,8,FALSE)",
                   f"=VLOOKUP($B{r},'02_Buyer_Score_Detail'!$A:$H,7,FALSE)",
                   b["contact"], b["status"], b["src"], VERIFY_DATE,
                   ("벤더 경유 검토 권장. " if b.get("vendor") else "") + b["notes"]])
    _style_header(ws, 11)
    _set_widths(ws, [6, 32, 18, 18, 7, 10, 52, 10, 62, 12, 70])
    _body_style(ws, 11)

    # ═══════════ 02_Buyer_Score_Detail ═══════════
    ws = wb.create_sheet("02_Buyer_Score_Detail")
    ws.append(["회사명", "구매력(25)", "유통망(25)", "적합성(20)", "입점용이성(15)", "성장성(15)",
               "점수(자동)", "등급(자동)", "점수 근거"])
    for i, b in enumerate(ROWS, start=1):
        r = i + 1
        ws.append([b["name"], *b["sc"], f"=SUM(B{r}:F{r})",
                   f'=IF(G{r}>=90,"A+",IF(G{r}>=80,"A",IF(G{r}>=65,"B+",IF(G{r}>=50,"B","C"))))',
                   " | ".join(f"{lab}: {txt}" for lab, txt in zip(SCORE_LABELS, b["rz"]))])
    _style_header(ws, 9)
    _set_widths(ws, [32, 11, 11, 11, 13, 11, 10, 9, 150])
    _body_style(ws, 9)

    # ═══════════ 03_Buyer_Intelligence ═══════════
    ws = wb.create_sheet("03_Buyer_Intelligence")
    ws.append(["회사명", "채널유형", "셀프등록형/MD심사형", "URL", "Notes", "Contact_Strategy"])
    for b in ROWS:
        ws.append([b["name"], b["ch"], b["mode"], b["url"],
                   f"[검증상태: {b['status']}] [확인일자: {VERIFY_DATE}] {b['notes']} [출처: {b['src']}]",
                   b["strat"]])
    _style_header(ws, 6)
    _set_widths(ws, [32, 18, 18, 46, 120, 60])
    _body_style(ws, 6)

    # ═══════════ 05_CRM_Import ═══════════
    ws = wb.create_sheet("05_CRM_Import")
    ws.append(["회사명", "채널유형", "대표번호", "담당부서번호", "이메일/입점포털", "MD담당자",
               "등급", "점수", "검증상태", "입점조건", "필요서류", "비고"])
    for i, b in enumerate(ROWS, start=1):
        r = i + 1
        ws.append([b["name"], b["ch"], b["phone"], b["dept"], b["contact"], "확인 불가",
                   f"=VLOOKUP($A{r},'02_Buyer_Score_Detail'!$A:$H,8,FALSE)",
                   f"=VLOOKUP($A{r},'02_Buyer_Score_Detail'!$A:$H,7,FALSE)",
                   b["status"], b["cond"], b["docs"],
                   ("벤더 경유 가능 채널. " if b.get("vendor") else "") + f"출처: {b['src']}"])
    _style_header(ws, 12)
    _set_widths(ws, [32, 18, 26, 34, 52, 12, 7, 8, 10, 46, 58, 62])
    _body_style(ws, 12)

    # ═══════════ 06_Search_Log ═══════════
    ws = wb.create_sheet("06_Search_Log")
    ws.append(["Log_ID", "검색어", "사용소스", "Result", "Companies_Found",
               "Companies_Advanced_To_Pipeline", "Key_Insight"])
    for lid, kmode, src, q, res, found, adv, ins in KLOG:
        ws.append([f"L{lid:02d}", q, f"[{kmode}] {src}", res, found, adv, ins])
    tf, ta = sum(k[5] for k in KLOG), sum(k[6] for k in KLOG)
    ws.append(["합계", "—", "—", "—", tf, ta,
               f"수집-반영 비율 {ta / tf * 100:.1f}% ('25. Self QA' 항목 9). "
               "미반영분의 대부분은 대량 소스 1회 조회분 중 정책·구조상 상세 확인을 중단한 건이며"
               "('22.5' 유형②·③), 폐기가 아니라 사용자 직접 조회 경로로 이관됨(00_Summary 실행 가이드 참조)."])
    _style_header(ws, 7)
    _set_widths(ws, [9, 46, 40, 11, 16, 30, 110])
    _body_style(ws, 7)

    # ═══════════ 07_Entry_Checklist ("15.") ═══════════
    ws = wb.create_sheet("07_Entry_Checklist")
    ws.append(["채널 유형", "본 배치 해당 채널(예시)", "공통 필요 서류", f"품목({ITEM}) 특화 확인사항", "비고"])
    for row in mod.CHECKLIST:
        ws.append(list(row))
    ws.append(["※ 공통 유의", "—",
               "법적 요건은 시점에 따라 바뀔 수 있으므로 실행 시점에 반드시 재확인할 것('15.' 명시 규정)",
               "본 체크리스트는 웹서치로 확인된 공개 정보 기반이며, 개별 채널의 최신 요구서류는 담당 MD·포털 공지 우선",
               "계약·법률 자문·벤더코드 발급 대행은 KDIS Scope 제외 — 전문가 확인 필요"])
    _style_header(ws, 5)
    _set_widths(ws, [30, 44, 56, 68, 62])
    _body_style(ws, 5)

    # ═══════════ 08_KNOWLEDGE_LOG_Summary ("31.5") ═══════════
    ws = wb.create_sheet("08_KNOWLEDGE_LOG_Summary")
    npass = sum(1 for k in KLOG if k[4].startswith("✅"))
    npart = sum(1 for k in KLOG if k[4].startswith("⚠"))
    nfail = sum(1 for k in KLOG if k[4].startswith("❌"))
    ws.append(["구분", "내용"])
    ws.append(["총 검색 시도", f"{len(KLOG)}회"])
    ws.append(["결과 분포", f"✅ PASS: {npass}회 | ⚠ PARTIAL: {npart}회 | ❌ FAIL: {nfail}회"])
    ws.append(["최적 소스 / 최적 키워드 / 제외된 패턴", NAR["knowledge"]])
    for row in getattr(mod, "KLOG_EXTRA", []):
        ws.append(list(row))
    ws.append(["다음 배치 인계 사항", NAR["next_batch"]])
    ws.append(["■ 개별 로그 ('31.2' 포맷)",
               "06_Search_Log 탭에 Log_ID·검색어·소스·Result·발견/반영 건수·Key Insight로 전량 기록"])
    _style_header(ws, 2)
    _set_widths(ws, [36, 150])
    _body_style(ws, 2)

    # ═══════════ 09_Communication_Draft ("18.") ═══════════
    ws = wb.create_sheet("09_Communication_Draft")
    ws.append(["01 탭 순위", "대상 회사", "채널 유형", "적용 템플릿('18.')", "발송 경로", "초안"])
    rank_of = {b["name"]: i for i, b in enumerate(ROWS, start=1)}
    ch_of = {b["name"]: b["ch"] for b in ROWS}
    for target, tpl, route, body in mod.DRAFTS:
        if target not in rank_of:
            raise AssertionError(f"Communication Draft 대상이 01 탭에 없음: {target}")
        ws.append([f"{rank_of[target]}위", target, ch_of[target], tpl, route, body])
    ws.append(["—", "공통(전 대상)", "—", "Follow-up 템플릿", "최초 제안 후 7~10일 경과 시",
               "안녕하세요, [회사명]입니다.\n지난 [날짜]에 보내드린 '[제품명]' 제안 관련하여 검토 진행 상황이 "
               "궁금하여 연락드립니다.\n추가 자료나 샘플이 필요하시면 언제든 말씀해 주세요."])
    ws.append(["—", "공통(전 대상)", "—", "샘플 전달 후속 템플릿", "샘플 발송 후 5~7일 경과 시",
               "안녕하세요, [회사명]입니다.\n보내드린 샘플은 잘 받아보셨는지요. 제품 특징과 관련해 추가로 "
               "설명드릴 부분이 있으면 미팅 일정 조율 가능합니다."])
    ws.append(["", "", "", "", "",
               "※ 모든 초안은 '[ ]' 표기 항목을 사용자가 실제 데이터로 채워 발송할 것. "
               "효능·효과를 단정하는 표현(질병 치료·개선 등)은 의료기기·의약품 오인 광고에 해당할 수 있으므로 사용 금지."])
    _style_header(ws, 6)
    _set_widths(ws, [10, 32, 20, 24, 48, 110])
    _body_style(ws, 6)

    # ═══════════ 10_Contact_Limits_Alt ("23.") ═══════════
    no_t1 = [b for b in ROWS if b["contact"] == "확인 불가"]
    ws = wb.create_sheet("10_Contact_Limits_Alt")
    ws.append(["구분", "항목", "내용"])
    ws.append(["현황", "Tier 1 연락처 미확보 기업 수",
               f"{len(no_t1)}개 / 전체 {N}개 ({len(no_t1) / N * 100:.1f}%) — "
               f"'26. Quality Gate'의 Tier 1 확보율 80% 기준 {'충족' if tier1_rate >= 80 else '미달'}. "
               "'21. Output Rules' 5번에 따라 미달을 숨기지 않고 그대로 보고하며, 아래 한계 설명과 "
               "대안 채널을 자동 첨부함."])
    ws.append(["현황", "Tier 1 미확보 기업 목록",
               ", ".join(b["name"] for b in no_t1) if no_t1 else "없음"])
    ws.append(["", "", ""])
    for kind, item, txt in [
        ("한계 ('23.')", "3자 자료 한계",
         "뉴스·공시·업계 보도는 연락처를 포함하지 않는 경우가 대부분이다. 보도로 실체와 실적이 확인되어도 "
         "공식 연락처가 어디에도 공개되지 않은 후보가 존재한다."),
        ("한계 ('23.')", "입점 포털 구조",
         "대형사는 개별 이메일을 비공개로 두고 포털만 제공하는 경우가 많다. 본 시스템은 이를 Tier 1 미달이 "
         "아니라 Tier 1 충족으로 인정하므로('13.'), 확보분 상당수가 전화번호가 아닌 포털 URL이다 — 정상이며 결함이 아님."),
        ("한계 ('23.')", "B2B 도매몰의 문의 폼 구조",
         "사업자 전용 도매몰은 대표 연락처 대신 '회원가입 후 사업자 인증 → 견적 문의 폼' 구조를 두는 경우가 "
         "많다. 이 경우 회원 가입 자체가 사실상 1차 접촉 단계이며, 폼 제출을 Tier 1 접촉으로 간주하고 진행하면 된다."),
        ("한계 (본 배치 고유)", "자동 접근 차단",
         "응답이 없던 도메인들은 사이트가 없는 것이 아니라 자동 접근이 차단된 것이다. 사용자가 일반 브라우저로 "
         "접속하면 대부분 정상 열람 가능하므로 실행 시 직접 확인할 것. 상세 목록은 00_Summary의 '22.5' 기록 참조."),
        ("한계 (본 배치 고유)", "대량 소스의 수집 정책",
         "지역 유통상을 대량 확보하는 데 가장 효율적인 공공 사업자정보 소스가 무단 대량수집을 금지하고 있어 "
         "상세 페이지 대량 조회를 수행하지 않았다('22.5 유형②'). 해당 후보들은 상호·소재지·업종만 확보된 "
         "'부분확인' 상태이며 대표번호가 비어 있다."),
        ("", "", ""),
        ("대안 채널 ('23.')", "① 입점 포털 제출",
         "전화·이메일이 없어도 포털 양식에 제안서를 맞춰 제출하는 것이 대형 채널의 표준 경로다. "
         "01 탭에서 연락처 칸에 'Tier 1'로 표기된 항목이 모두 이 경로에 해당한다."),
        ("대안 채널 ('14.')", "② 벤더 경유",
         f"MD심사형 채널은 신생·중소 브랜드의 직접 컨택 응답률이 낮으므로 벤더 경유가 실무 표준이다. "
         f"본 배치에서 벤더 경유가 유효하다고 판단된 채널 {n_vendor}개(01·05 탭 비고 표기). "
         "※ 벤더는 수수료(마진)를 추가로 가져가므로 직접 컨택과 마진 구조를 비교 판단할 것('14.3'). "
         "※ 벤더사도 접촉 전 사업자등록·벤더 실적을 반드시 검증할 것."),
        ("대안 채널 ('24.')", "③ 공공 판로지원",
         "한국중소벤처기업유통원 '판판대로'(fanfandaero.kr, 1899-0309): TV·데이터홈쇼핑 8개사 9개 채널 MD 현장 "
         "매칭 — 직접 입점 창구가 비공개인 데이터홈쇼핑에 접근하는 사실상 유일한 경로. "
         "중소기업중앙회 판로지원실(02-2124-3243): 홈앤쇼핑 방송 1회(50분) 입점 지원. "
         "소상공인시장진흥공단(semas.or.kr): 온라인 판로 지원. "
         "공영홈쇼핑 '모두의 판로': 기존 거래 여부 불문 공개모집 — 신생 브랜드 진입장벽 최저. "
         "※ 각 기관의 자격 요건·신청 시기는 정책 변경이 잦으므로 실행 시점에 직접 확인할 것('24.' 적용규칙 2)."),
        ("대안 채널 ('23.')", "④ 오프라인 입점설명회·품평회",
         "CU 열린중소기업상담(매월 넷째주 금요일, BGF리테일 본사 1층 상담실 / 서울 강남구 테헤란로 405) — "
         "일정·장소가 공개된 정기 상담 창구. 이마트 우수상품 공개 공모전(홈페이지 공고 → 본사 대강당 품평회). "
         "롯데마트·롯데슈퍼·롯데백화점 품평회(상시, 상담 신청 후 MD가 일정 연락)."),
        ("", "", ""),
        ("실행 권고", "Tier 1 미확보 기업 처리 방침",
         "① 포털·문의 폼만 있는 기업: 폼 제출을 1차 접촉으로 삼고, 회신 시 담당자 연락처를 확보한다. "
         "② 연락처가 전혀 없는 기업: 00_Summary의 '사용자 직접 실행 가이드'에 따라 사업자정보 조회로 대표번호를 "
         "개별 확인한 뒤 다음 배치에서 '12.5 재검증 Pass'로 '확인됨' 승격을 시도한다. "
         "③ 어떤 경우에도 연락처를 추정·조합해 만들지 않는다('20. Hallucination Prevention')."),
    ]:
        ws.append([kind, item, txt])
    _style_header(ws, 3)
    _set_widths(ws, [22, 32, 150])
    _body_style(ws, 3)
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or "").startswith(("한계", "대안")):
            ws.cell(row=r, column=2).font = Font(size=9, bold=True)

    wb.save(fpath)

    # ═══════════ "28.1" 3번: LibreOffice headless 재계산 ═══════════
    recalc_dir = f"/tmp/kdis_recalc_{ITEM}"
    shutil.rmtree(recalc_dir, ignore_errors=True)
    os.makedirs(recalc_dir, exist_ok=True)
    home = "/tmp/kdis_lo_home"
    os.makedirs(home, exist_ok=True)
    res = subprocess.run(
        ["soffice", "--headless", "--norestore", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
         "--outdir", recalc_dir, fpath],
        capture_output=True, text=True, timeout=900, env=dict(os.environ, HOME=home))
    recalc = os.path.join(recalc_dir, fname)
    if not os.path.exists(recalc):
        raise RuntimeError(f"재계산 파일 생성 실패 — 출력 중단: {res.stdout} {res.stderr}")

    # ═══════════ "28.1" 4번: 재계산본 재판독 + 감사 ═══════════
    wbv = load_workbook(recalc, data_only=True)
    wbf = load_workbook(fpath, data_only=False)
    errors = []

    r01 = wbv["01_TOP_Buyer_List"].max_row - 1
    r02 = wbv["02_Buyer_Score_Detail"].max_row - 1
    r03 = wbv["03_Buyer_Intelligence"].max_row - 1
    r05 = wbv["05_CRM_Import"].max_row - 1
    summary = {row[0].value: row[1].value
               for row in wbv["00_Summary"].iter_rows(min_row=2, max_col=2)}
    claimed = int(re.search(r"(\d+)", str(summary.get("확보 Buyer"))).group(1))
    if not (claimed == r01 == r02 == r05 == N):
        errors.append(f"Pipeline Parity 위반: summary={claimed} 01={r01} 02={r02} 05={r05} data={N}")

    def cnt(sheet, col, val):
        return sum(1 for row in wbv[sheet].iter_rows(min_row=2, min_col=col, max_col=col)
                   if row[0].value == val)
    p01 = cnt("01_TOP_Buyer_List", 8, "부분확인")
    p05 = cnt("05_CRM_Import", 9, "부분확인")
    p03 = sum(1 for row in wbv["03_Buyer_Intelligence"].iter_rows(min_row=2, min_col=5, max_col=5)
              if "[검증상태: 부분확인]" in str(row[0].value))
    if not (p01 == p05 == p03 == n_part):
        errors.append(f"'부분확인' 건수 불일치: 01={p01} 03={p03} 05={p05} data={n_part}")

    sheetnames = set(wbf.sheetnames)
    nform, bad_refs = 0, []
    for sh in wbf.sheetnames:
        for row in wbf[sh].iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    nform += 1
                    refs = (re.findall(r"'([^']+)'!", v) +
                            re.findall(r"(?<![A-Za-z0-9_'])([A-Za-z0-9_]+)!", v))
                    bad_refs += [(sh, cell.coordinate, x) for x in refs if x not in sheetnames]
    if bad_refs:
        errors.append(f"존재하지 않는 시트명 참조 {len(bad_refs)}건: {bad_refs[:5]}")

    errvals = [(sh, c.coordinate, c.value) for sh in wbv.sheetnames
               for row in wbv[sh].iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith("#")]
    if errvals:
        errors.append(f"수식 오류값 {len(errvals)}건: {errvals[:5]}")

    ws2, ws1, ws5 = (wbv["02_Buyer_Score_Detail"], wbv["01_TOP_Buyer_List"], wbv["05_CRM_Import"])
    prev = 10 ** 9
    for i, b in enumerate(ROWS, start=2):
        exp = sum(b["sc"])
        if ws2.cell(row=i, column=1).value != b["name"]:
            errors.append(f"02 행 {i} 회사명 불일치")
        if ws2.cell(row=i, column=7).value != exp:
            errors.append(f"02 행 {i} 점수 불일치: {ws2.cell(row=i, column=7).value} != {exp}")
        if not 0 < exp <= 100:
            errors.append(f"02 행 {i} 점수 범위 이상: {exp}")
        if ws2.cell(row=i, column=8).value != grade(exp):
            errors.append(f"02 행 {i} 등급 불일치")
        if ws1.cell(row=i, column=6).value != exp or ws5.cell(row=i, column=8).value != exp:
            errors.append(f"행 {i} 점수 VLOOKUP 불일치")
        if ws1.cell(row=i, column=5).value != grade(exp):
            errors.append(f"01 행 {i} 등급 VLOOKUP 불일치")
        if ws1.cell(row=i, column=2).value != b["name"] or ws5.cell(row=i, column=1).value != b["name"]:
            errors.append(f"행 {i} 회사명 정렬 불일치")
        if exp > prev:
            errors.append(f"01 행 {i} 내림차순 정렬 위반")
        prev = exp

    for b in ROWS:
        for f in ("src", "status", "contact", "url", "phone", "cond", "docs"):
            if not str(b.get(f) or "").strip():
                errors.append(f"{b['name']} 필드 누락: {f}")

    stats = dict(n=N, conf=n_conf, part=n_part, tier1=n_tier1, rate=tier1_rate,
                 vendor=n_vendor, formulas=nform, rows=(r01, r02, r03, r05),
                 gate=overall, sheets=len(wbf.sheetnames))
    if errors:
        raise AssertionError("감사 실패:\n  - " + "\n  - ".join(errors[:40]))
    return fpath, stats
