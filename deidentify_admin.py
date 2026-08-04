#!/usr/bin/env python3
"""어드민 case-metrics CSV에서 개인식별정보를 제거해 대시보드용 파일을 만든다.

이 repo는 공개다. 어드민 원본에는 고객 실명·나이·성별이 행 단위로 들어 있어
그대로 커밋할 수 없다. 대시보드가 실제로 쓰는 컬럼만 화이트리스트로 남긴다.

사용법:
    python3 deidentify_admin.py ~/Downloads/case-metrics-20260803-1127.csv
    → data/admin_case_metrics.csv 생성

화이트리스트 방식(신규 컬럼이 생기면 자동으로 빠진다)이므로,
어드민이 컬럼을 추가해도 실수로 개인정보가 흘러나가지 않는다.
"""

import os
import sys

import pandas as pd

# 대시보드가 쓰는 컬럼만 유지
KEEP = [
    'row_kind',                # PRIMARY / COMPANION / CONSULTATION_ONLY (건수 vs 인원 구분)
    'first_consult_date',      # 채팅접수일자 + 내원일 없는 행의 집계월 보완
    'channel',                 # ONLINE / OFFLINE
    'language',                # 사용언어
    'hospital_name',           # 예약클리닉
    'hospital_region',         # 지역
    'treatment_type',          # 시술 / 수술
    'workflow_status',         # 완료 / 취소 / 노쇼 …
    'reservation_bucket',
    'booked_date',             # 내원일
    'booked_time',
    'estimated_amount',
    'actual_amount',           # 매출
    'paid',
    'consulter_nationality',   # 국적 (ISO2)
    'candidate_nationality',
    'group_size',
    'consulter_form',
    'treatments',              # 시술/수술명
    'interpreter_needed',
]

# 명시적으로 제거되는 개인식별정보 — 문서화 목적
DROPPED_PII = [
    'consulter_name', 'candidate_name',      # 고객 실명
    'consulter_age', 'candidate_age',        # 나이
    'consulter_gender', 'candidate_gender',  # 성별
    'consultation_code', 'case_id_hash',     # 케이스 식별자
    'consultant_name',                       # 상담사 (대시보드 미사용)
    'consultation_duration_min', 'has_notes', 'booked_hospital_count',
]


def main():
    if len(sys.argv) < 2:
        sys.exit(f"사용법: python3 {os.path.basename(__file__)} <case-metrics-*.csv>")

    src = os.path.expanduser(sys.argv[1])
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       'data', 'admin_case_metrics.csv')

    df = pd.read_csv(src, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    keep = [c for c in KEEP if c in df.columns]
    removed = [c for c in df.columns if c not in keep]

    os.makedirs(os.path.dirname(out), exist_ok=True)
    df[keep].to_csv(out, index=False, encoding='utf-8-sig')

    print(f"입력 : {src}")
    print(f"출력 : {out}")
    print(f"행   : {len(df):,}")
    print(f"유지 {len(keep)}컬럼: {', '.join(keep)}")
    print(f"제거 {len(removed)}컬럼: {', '.join(removed)}")

    leaked = [c for c in removed if c not in DROPPED_PII]
    if leaked:
        print(f"\n※ 문서에 없는 신규 컬럼이 제거됨 (확인 권장): {', '.join(leaked)}")


if __name__ == '__main__':
    main()
