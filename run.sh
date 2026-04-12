#!/bin/bash
# 언니가이드 월간 트렌드 리포트 생성
#
# 사용법:
#   ./run.sh                          → 최신 Excel, 최신 월 자동 감지
#   ./run.sh 2026-04                  → 최신 Excel, 특정 월 지정
#   ./run.sh 2026-04 ~/파일.xlsx      → 특정 Excel, 특정 월 지정
#
# 매달 루틴:
#   1. Excel에 해당 월 예약 + 정산 데이터 추가
#   2. 터미널에서 ./run.sh 실행
#   3. ~/Documents/unniguide-report/ 에서 HTML 확인

cd "$(dirname "$0")"

MONTH="${1:-}"
EXCEL="${2:-}"

if [ -n "$EXCEL" ] && [ -n "$MONTH" ]; then
    python3 generate_report.py "$EXCEL" "$MONTH"
elif [ -n "$MONTH" ]; then
    python3 generate_report.py "" "$MONTH"
else
    python3 generate_report.py
fi

# 생성 후 공통 리포트 자동 열기
LATEST=$(ls -t unniguide_report_*.html 2>/dev/null | head -1)
if [ -n "$LATEST" ]; then
    echo ""
    echo "🌐 브라우저에서 열기: open $LATEST"
    open "$LATEST"
fi
