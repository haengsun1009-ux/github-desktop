# -*- coding: utf-8 -*-
"""06_Search_Log 데이터 — 이번 세션에서 실제로 수행한 web_search / web_fetch 시도 기록"""

search_log = [
    dict(id="L01", query="괄사 시장 규모 유통 채널 한국 2026", source="Google/Naver 종합검색(web_search)",
         result="PASS", found=5, advanced=0,
         insight="괄사 단독 통계는 없으나 한국 뷰티·퍼스널케어 시장(2025년 약 136.6억 USD, 온라인채널 CAGR 7.57%) 확인 — Market Intelligence(7.)에 반영, 버이어 후보는 아님"),
    dict(id="L02", query="괄사 KC인증 규제 화장품 미용기기 신고", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=0,
         insight="일반 괄사(비전동)는 화장품/미용기기 인증 대상 아님, 전동식·의학적 효능 표방 시에만 KC/식약처 인증 필요 — 15.체크리스트에 반영"),
    dict(id="L03", query="괄사 도매 총판 site:co.kr", source="Google 종합검색(web_search)",
         result="PASS", found=6, advanced=3,
         insight="도매꾹/도매매, 도매리스트닷컴(랭킹), 현대교역(괄사 직접판매) 확보 — '입점문의/총판' 키워드가 도매 채널 발굴에 적중률 높음"),
    dict(id="L04", query="가샤 스크래핑 뷰티툴 협회 대한화장품협회 회원사", source="Google 종합검색(web_search)",
         result="⚠ PARTIAL", found=1, advanced=0,
         insight="대한화장품협회 회원사 명단은 사이트 내 실시간 검색 UI로만 제공되어 특정 회원사(괄사 취급) 식별 불가 — 협회 경로는 이 품목에는 낮은 우선순위로 하향 조정"),
    dict(id="L05", query="에스테틱 소모품 도매 유통사 괄사 스톤 판매", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=3,
         insight="현대교역, 웰뷰((주)원미코리아), (주)아미인터내셔널 확보 — '에스테틱 소모품 도매' 키워드가 괄사 전문 유통사 발굴에 최적중"),
    dict(id="L06", query="뷰티 소품 수입 유통 회사 괄사 마사지 도구 총판 모집", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=0,
         insight="현대교역 총판/대리점 모집 공고 재확인(중복) — 해외 OEM 제조사(킨웡, SHEFMON 등) 다수 노출되나 이들은 '국내 바이어'가 아닌 해외 공급자로 제외"),
    dict(id="L07", query="올리브영 입점문의 뷰티소품 MD 제안 절차", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=1,
         insight="CJ올리브영 공식 사업제휴제안 페이지(corp.oliveyoung.com) 확보, MD상담→품평회→신용평가→계약 절차 확인"),
    dict(id="L08", query="홈쇼핑 뷰티기기 협력사 등록 괄사 마사지기", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=1,
         insight="CJ온스타일 파트너스 시스템 및 공식 입점안내서(PDF) 확보 — 뷰티 디바이스 카테고리 명시 확인"),
    dict(id="L09", query="에스테틱 소모품 도매몰 목록 피부관리 재료상 사이트", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=4,
         insight="초이스코스메틱, 유니코스메틱, 딸라스파, 조은뷰티도매 확보 — 다만 다수가 괄사 직접판매 미확인(부분확인 처리)"),
    dict(id="L10", query="한방 재료 도매 괄사 지압봉 한의원 용품 유통", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=5,
         insight="정음바이오, 한의사몰, 안진팜메디, 렉스모아(괄사전문), PGO(괄사전문) 확보 — 한방/경락 계열 검색이 괄사 전문업체 발굴에 매우 효과적(신규 Angle)"),
    dict(id="L11", query="다이소 입점 문의 절차 협력사", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=1,
         insight="(주)아성다이소 공식 납품문의 채널(daiso.co.kr/cs/supply) 확보"),
    dict(id="L12", query="스포츠마사지 물리치료 용품 도매 괄사 근막이완 도구", source="Google 종합검색(web_search)",
         result="⚠ PARTIAL", found=5, advanced=0,
         insight="검색결과 대부분 해외 OEM 제조사(킨웡·Poleview 등)로 국내 바이어 아님 — 이 키워드는 '해외 공급자 발굴'에는 유효하나 '국내 바이어 발굴'에는 낮은 적중률로 판단, 다음 라운드에서 우선순위 하향"),
    dict(id="L13", query="뷰티 유통대행사 벤더 입점대행 올리브영 마트 리스트", source="Google 종합검색(web_search)",
         result="⚠ PARTIAL", found=5, advanced=1,
         insight="에스엘라이프·모스트·실리콘투 등 벤더사 존재는 뉴스로 확인되나 공식 홈페이지/연락처 미확보 — '23. Contact Limits' 해당, 벤더 발굴은 업계 네트워킹 필요성 확인"),
    dict(id="L14", query="이지마켓B2B 도매토피아 JC커머스 온채널 오너클랜 사업자 등록 정보", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=2,
         insight="이지마켓B2B((주)커넥트어센드, 사업자정보 완전확인), 온채널(이메일만 확인) 확보 — 도매토피아/JC커머스/오너클랜은 사업자정보 미확인으로 이번 배치 제외(향후 재검증 대상)"),
    dict(id="L15", query="GS샵 NS홈쇼핑 공영쇼핑 입점 제안 협력사 등록 홈페이지", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=2,
         insight="GS샵(대표번호 확인), 공영홈쇼핑(8단계 절차 확인) 확보 — NS홈쇼핑은 이번 세션에서 공식 URL 재확인 못해 제외(향후 재검증 대상)"),
    dict(id="L16", query="쿠팡윙 스마트스토어 셀러 등록 G마켓 파트너 11번가 셀러오피스 공식 URL", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=4,
         insight="4대 오픈마켓 셀러센터 공식 URL 전부 확보 — 셀프등록형 채널은 공식 URL 검증만으로 G1~G4 충족 용이"),
    dict(id="L17", query="이마트 파트너스 롯데마트 홈플러스 납품 문의 공식 사이트 URL", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=3,
         insight="이마트(동반성장 페이지), 롯데마트(입점상담 시스템), 홈플러스(조인어스파트너스) 확보 — 대형마트 3사 모두 공식 창구는 확인되나 개별 MD 직접 연락처는 비공개(G4 부분충족)"),
    dict(id="L18", query="마켓컬리 오늘의집 입점 문의 뷰티 카테고리 절차", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=2,
         insight="마켓컬리(구글폼 접수), 오늘의집(공식 입점제안서 PDF) 확보"),
    dict(id="L19", query="롯데홈쇼핑 협력사 등록 현대홈쇼핑 파트너스 입점 제안 홈페이지", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=2,
         insight="롯데홈쇼핑(파트너시스템 URL), 현대홈쇼핑(협력사시스템·동반성장 페이지) 확보"),
    dict(id="L20", query="홈플러스 상품 공급 협력사 등록 문의 사업자 정보", source="Google 종합검색(web_search)",
         result="PASS", found=5, advanced=0,
         insight="홈플러스 '조인어스 파트너스' 상세 절차 재확인(중복 보강) — 기업회생 절차 진행 중 리스크 확인, 00_Summary 비고에 반영"),
    dict(id="F01", query="WebFetch: rexmoa.co.kr (http/https 각 1회)", source="공식 홈페이지 직접 접속(web_fetch)",
         result="❌ FAIL", found=0, advanced=0,
         insight="http는 500 오류, https는 'NHN호스팅 서비스 반영 중' 안내페이지만 노출 — 로그인/인증키 필요형(22.5 유형④)은 아니고 기술적 접속 불가. 동일 세션 내 2회 시도 후 중단, '부분확인'으로 유지하고 사용자에게 재시도 안내"),
    dict(id="F02", query="WebFetch: amiint.com / jebio.co.kr / shop-open.co.kr(현대교역)", source="공식 홈페이지 직접 접속(web_fetch)",
         result="PASS", found=3, advanced=0,
         insight="현대교역 대표번호(02-2617-7000), 정음바이오 고객센터(070-8282-3570)/이메일 확보 — 사이트 푸터의 사업자정보 직접 대조가 Tier1 연락처 확보에 효과적(신규 Angle)"),
    dict(id="F03", query="WebFetch: withgs.gsshop.com, wellbeau.co.kr(스니펫)", source="공식 홈페이지 직접 접속(web_fetch)",
         result="PASS", found=2, advanced=0,
         insight="GS샵 대표 입점문의 전화(02-2007-4434) 확보, 웰뷰는 검색 스니펫에서 사업자등록번호·통신판매업신고번호까지 확인되어 '확인됨' 등급 부여"),
    dict(id="L21", query="[참조] 나라장터·중소기업유통센터·소진공·동반성장위원회 — KDIS 데이터소스맵(9.2) 기 등재 공식기관", source="문서 내장 Data Source Map(9.2) 참조, 이번 세션 재검색 안함",
         result="⚠ PARTIAL", found=4, advanced=4,
         insight="공식기관 자체의 실존은 의심 여지 없으나(공공기관), 이번 세션에서 최신 연락처·담당부서를 재검색하지 않아 '부분확인'으로 보수적 처리 — 실행 시점 URL 및 담당부서 재확인 필수(00_Summary 비고 반영)"),
]

total_found = sum(x["found"] for x in search_log)
total_advanced = sum(x["advanced"] for x in search_log)
print(f"검색 시도: {len(search_log)}건 / Found 합계: {total_found} / Advanced 합계: {total_advanced}")
pass_n = sum(1 for x in search_log if x["result"]=="PASS")
partial_n = sum(1 for x in search_log if "PARTIAL" in x["result"])
fail_n = sum(1 for x in search_log if "FAIL" in x["result"])
print(f"PASS={pass_n} PARTIAL={partial_n} FAIL={fail_n}")

import json
with open('/workspace/kdis_output/search_log.json', 'w', encoding='utf-8') as f:
    json.dump(search_log, f, ensure_ascii=False, indent=2)
