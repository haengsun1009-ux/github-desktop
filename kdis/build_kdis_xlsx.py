# -*- coding: utf-8 -*-
"""KDIS v1.1 산출물 생성 CLI.

사용법:
    python3 build_kdis_xlsx.py                       # 등록된 전 품목 생성
    python3 build_kdis_xlsx.py kdis_showertowel_data # 특정 품목만 생성

각 데이터 모듈은 "28.1" 절차(생성 → LibreOffice 재계산 → 재판독 → Pipeline Parity·
수식 참조 감사)를 통과해야만 파일이 최종 산출물로 인정된다. 감사 실패 시 종료 코드 1.
"""
import importlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kdis_builder  # noqa: E402

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
DATASETS = ["kdis_guasha_data"]


def main(names):
    failed = []
    for name in names:
        mod = importlib.import_module(name)
        print(f"\n━━━ {mod.ITEM} ({name}) ━━━")
        try:
            path, s = kdis_builder.build(mod, OUT_DIR)
        except Exception as exc:  # 감사 실패 → 출력 중단
            print(f"[FAIL] {exc}")
            failed.append(name)
            continue
        print(f"  탭 {s['sheets']}개 / 수식 {s['formulas']}개 / "
              f"행 수 01={s['rows'][0]} 02={s['rows'][1]} 03={s['rows'][2]} 05={s['rows'][3]}")
        print(f"  ALL CHECKS PASS — 확보 {s['n']} / 확인됨 {s['conf']} / 부분확인 {s['part']} / "
              f"Tier1 {s['tier1']}({s['rate']:.1f}%) / 벤더경유 {s['vendor']} / Gate {s['gate']}")
        print(f"  파일: {path}")
    if failed:
        sys.exit(f"\n감사 실패 품목: {', '.join(failed)}")
    print("\n전 품목 감사 통과.")


if __name__ == "__main__":
    main(sys.argv[1:] or DATASETS)
