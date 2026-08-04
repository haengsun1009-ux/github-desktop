# KDIS v1.1 실행 결과 — 품목: 괄사 / 대상: 국내(대한민국)

이 폴더는 "KDIS(Korean Domestic Buyer Intelligence System) v1.1" 마스터 프롬프트에 따라
Phase 1~3(Market Intelligence → Buyer Discovery → Verification/Scoring/xlsx 생성)을 실제로
실행한 결과물입니다.

## 산출물

- **`KDIS_괄사_20260804.xlsx`** — 최종 산출물(9개 탭). 모든 점수/등급은 엑셀 수식(SUM/IF/VLOOKUP)으로
  계산되며, LibreOffice headless로 재계산 후 수식 참조 감사를 통과한 상태입니다.
- `build_kdis.py` — 36개 바이어 후보 원본 데이터(출처·근거 포함)
- `search_log.py` — 이번 세션에서 실제로 수행한 24건의 web_search/web_fetch 시도 로그
- `generate_workbook.py` — xlsx 생성 스크립트(openpyxl)
- `audit.py` — 수식 참조 감사 + Pipeline Parity + 값 타당성 검증 스크립트(전항목 PASS 확인됨)
- `companies.json`, `search_log.json` — 원본 데이터(재현/재사용 가능)

## ⚠ 목표 수량 대비 결과 (반드시 읽어주세요)

- **요청 목표: 1,000개 / 실제 확보: 36개(확인됨 15, 부분확인 21) — 달성률 3.6%**
- KDIS의 핵심 원칙("2. Core Principles" 원칙3 정확성>수량, "20. Hallucination Prevention")에 따라,
  실제 web_search/web_fetch로 확인되지 않는 바이어를 지어내 1,000개를 채우지 않았습니다.
- **구조적 사유**: '괄사'라는 단일 품목을 취급하는 국내 "기관형 바이어"(유통채널·도매상·벤더 단위)의
  절대 수 자체가 제한적입니다(대형마트·홈쇼핑·오픈마켓·H&B·에스테틱 도매·한방용품 도매·공공조달 등을
  다 합쳐도 실질 모집단은 수십~백여 개 수준). 1,000개 규모는 채널이 아니라 "개별 최종 사용 사업장"
  (전국 에스테틱샵·스파·한의원 등)까지 확장해야 하며, 이는 로컬 사업체 디렉토리 기반의 별도 대량 조사
  과제로, 이번 세션의 웹서치 도구만으로는 실명·연락처를 정확히 검증하며 완결할 수 없어 무리하게
  채우지 않았습니다.
- 상세 사유·시도한 라운드(Search Escalation Round 1~5)·대안 채널 안내는 xlsx의
  **`00_Summary` 탭 [목표 미달 보고]** 섹션에 전문이 기재되어 있습니다.
- 다음 배치를 진행하려면 `01_TOP_Buyer_List`의 회사명 전체를 "기존 확보 리스트"로 넘겨 중복을
  제거해야 합니다(KDIS "2." 원칙7).

## Self QA / Quality Gate 결과 요약

| 항목 | 결과 |
|---|---|
| Buyer 목표 수량(80%↑) | **FAIL** (3.6%) — 위 사유 참조 |
| 중복 제거 | PASS (36개 전건 유일) |
| 홈페이지/실체 확인률(TOP리스트 100%) | PASS |
| Tier1 연락처 확인률(80%↑) | **FAIL** (72.2%, 10개사 연락처 확인 불가로 명시) |
| 출처 명시율(100%) | PASS |
| Hallucination 징후 | PASS (없음 — 미확인 항목은 전부 "확인 불가"로 정직 표기) |
| Pipeline Parity(00=01=02=05) | **PASS** (36=36=36=36, `audit.py` 자동 검증) |
| 부분확인 포함 여부 | PASS (01/05 탭 확인됨15/부분확인21 일치) |
| 수집-반영 비율 | Found 106 / Advanced 38 — 큰 격차 사유는 06_Search_Log Key_Insight에 기재 |
| 수식 참조 감사 | **PASS** (존재하지 않는 시트 참조 0건, 명시적 오류값 0건, 5항목합=점수 불일치 0건) |

Quality Gate는 Buyer 수량과 Tier1 연락처 확인률 두 항목에서 기준(80%) 미달로 **PARTIAL PASS**
처리되었으며, 이는 문서 "21. Output Rules" 5번(기준 미달을 숨기거나 임의로 통과 처리하지 않는다)에
따라 그대로 보고합니다.
