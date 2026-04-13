#!/usr/bin/env python3
"""
언니가이드 인터랙티브 대시보드 (Streamlit) v2
- 운영 트렌드 Excel + 내부리포트 Excel 2개 소스
- 취소/노쇼 트래킹 포함
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import re
import os
import io
import glob
from collections import defaultdict
from datetime import datetime

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(
    page_title="언니가이드 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 브랜드 스타일
# ============================================================
BRAND_ORANGE = '#FF6A3B'
BRAND_PLUM = '#330C2E'
BRAND_IVORY = '#FBF9F1'
BRAND_GREEN = '#00B894'
BRAND_RED = '#E74C3C'

CHART_COLORS = [
    '#FF6A3B', '#330C2E', '#00B894', '#FDCB6E', '#0984E3',
    '#E17055', '#00CEC9', '#A29BFE', '#FD79A8', '#55A3E8',
    '#F39C12', '#2ECC71', '#E74C3C', '#9B59B6', '#1ABC9C',
]

st.markdown("""
<style>
    .block-container { padding-top: 1rem; }
    [data-testid="stMetric"] {
        background: white; border: 1px solid #E9ECEF;
        border-radius: 12px; padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    [data-testid="stMetricLabel"] { font-size: 13px !important; }
    [data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 800 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { padding: 8px 20px; font-weight: 600; }
    div[data-testid="stSidebarContent"] { background: #FBF9F1; }
    h1, h2, h3 { color: #330C2E !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 상수
# ============================================================
NAME_NORMALIZE = {
    '사적인아름다운지유의원': '사적인아름다움지유의원',
    '루호성형외과': '루호성형외과의원',
    '우리성형외과': '우리성형외과의원',
    '테이아 의원': '테이아의원',
    '티에스성형외과의원': '티에스성형외과',
    '톡스앤필-시논현': '톡스앤필의원-신논현점',
    '톡스앤필 - 신논': '톡스앤필의원-신논현점',
    '톡스앤필 - 신논현': '톡스앤필의원-신논현점',
    '제이필 - 홍대': '제이필의원-홍대점',
    '제이필 - 강남': '제이필의원-강남점',
    '플래저성형외과': '플레저성형외과의원',
    '플래너성형외과': '플래너성형외과의원',
    '유픽의원 홍대': '유픽의원-홍대점',
    '유픽의원-홍대': '유픽의원-홍대점',
    '유픽의원-강남': '유픽의원-강남점',
    '디어 청담의원': '청담디어의원',
    '홍대셀레나': '홍대셀레나의원',
    '히트성형외과': '히트성형외과의원',
}

COUNTRY_FLAG = {
    '태국': '🇹🇭', '대만': '🇹🇼', '중국': '🇨🇳', '미국': '🇺🇸',
    '호주': '🇦🇺', '일본': '🇯🇵', '홍콩': '🇭🇰', '싱가포르': '🇸🇬',
    '베트남': '🇻🇳', '필리핀': '🇵🇭', '말레이시아': '🇲🇾', '인도네시아': '🇮🇩',
    '영국': '🇬🇧', '프랑스': '🇫🇷', '독일': '🇩🇪', '캐나다': '🇨🇦',
    '인도': '🇮🇳', '러시아': '🇷🇺', '몽골': '🇲🇳', '폴란드': '🇵🇱',
    '캄보디아': '🇰🇭', '아일랜드': '🇮🇪', '키르기스스탄': '🇰🇬',
    '부탄': '🇧🇹', '스페인': '🇪🇸', '뉴질랜드': '🇳🇿',
}


# 시술 키워드 매핑 (자유 텍스트 → 표준 시술 카테고리)
PROCEDURE_KEYWORDS = [
    ('울쎄라', ['울쎄라', 'ulthera', '울쎼라']),
    ('보톡스', ['보톡스', 'botox', '보톡', '더마톡신']),
    ('포텐자', ['포텐자', 'potenza']),
    ('써마지', ['써마지', 'thermage']),
    ('슈링크', ['슈링크', 'shrink']),
    ('리쥬란', ['리쥬란', 'rejuran']),
    ('쥬베룩', ['쥬베룩', 'juvelook']),
    ('필러', ['필러', 'filler']),
    ('리프팅', ['리프팅', 'lifting', '거상', '실리프팅']),
    ('지방분해주사', ['지방분해', '윤곽주사', '커팅주사', 'fat dissolv']),
    ('레이저', ['레이저', 'laser', '피코', 'pico', 'BBL']),
    ('올리지오', ['올리지오', 'oligio']),
    ('온다', ['온다', 'onda']),
    ('물광주사', ['물광', '더마샤인', 'skinbooster']),
    ('스킨보톡스', ['스킨보톡스', '스킨보', 'skinbtx']),
    ('인모드', ['인모드', 'inmode']),
    ('눈수술', ['눈수술', '눈매교정', '쌍꺼풀', '상안검', '눈 수술', '눈재수술']),
    ('코수술', ['코수술', '코끝', '콧대', '코 수술', '코 첫수술']),
    ('소프웨이브', ['소프웨이브', 'sofwave']),
    ('엑소좀', ['엑소좀', 'exosome']),
    ('아쿠아필', ['아쿠아필', 'aquapeel', '아쿠아 필']),
    ('셀르디엠', ['셀르디엠', 'cellrdm']),
    ('수액', ['수액']),
    ('모공치료', ['모공']),
    ('여드름치료', ['여드름']),
]


def extract_procedures(text):
    """자유 텍스트에서 시술 키워드를 추출하여 리스트로 반환"""
    if not text or str(text).strip() in ('', 'nan'):
        return []
    text_lower = str(text).lower()
    found = []
    for label, keywords in PROCEDURE_KEYWORDS:
        for kw in keywords:
            if kw.lower() in text_lower:
                found.append(label)
                break
    return found if found else [str(text).strip()[:30]]


def normalize_hospital(name):
    if not name or str(name).strip() in ('', 'nan', 'None'):
        return None
    name = str(name).strip()
    return NAME_NORMALIZE.get(name, name)


def format_krw(amount):
    if pd.isna(amount) or amount == 0:
        return '0원'
    amount = float(amount)
    if amount >= 100_000_000:
        return f"{amount / 100_000_000:.1f}억원"
    elif amount >= 10_000:
        return f"{amount / 10_000:,.0f}만원"
    else:
        return f"{amount:,.0f}원"


# ============================================================
# 데이터 로딩
# ============================================================
@st.cache_data(show_spinner=False, ttl=600)
def load_operation_excel(file_path):
    """운영 트렌드 Excel → 예약완료 df + 정산 df + 전체예약 df"""

    # URL인 경우 다운로드
    if str(file_path).startswith('http'):
        import urllib.request
        tmp = '/tmp/unniguide_gsheet_op.xlsx'
        urllib.request.urlretrieve(file_path, tmp)
        file_path = tmp

    # 예약 시트
    # 시트 이름 호환: 로컬 Excel vs Google Sheets
    xls = pd.ExcelFile(file_path)
    res_sheet = None
    for name in xls.sheet_names:
        if '예약확정' in name:
            res_sheet = name
            break
    if res_sheet is None:
        raise ValueError(f"예약확정 시트를 찾을 수 없습니다. 시트 목록: {xls.sheet_names}")
    df_res = pd.read_excel(file_path, sheet_name=res_sheet, header=1)
    expected_cols = [
        'NO', '채팅접수일자', '예약확정일', '담당자', '고객명', '그룹여부',
        '고객국적', '사용언어', '예약상태', '통역서비스요청', '종류',
        '시술수술명', '추천클리닉', '예약클리닉', '내원일', '시간',
        '예상금액', '실제금액', '금액확인', '설문발송여부',
        '후기작성여부', '캐시백지급대상자', '캐시백지급여부', '캐시백금액',
        '캐시백지급일자', 'Remark', '시술수술확정항목',
    ]
    actual_cols = expected_cols + [f'extra_{i}' for i in range(max(0, len(df_res.columns) - len(expected_cols)))]
    df_res.columns = actual_cols[:len(df_res.columns)]
    df_res['병원명'] = df_res['예약클리닉'].apply(normalize_hospital)
    df_res['내원일'] = pd.to_datetime(df_res['내원일'], errors='coerce')
    df_res['월'] = df_res['내원일'].dt.to_period('M').astype(str)
    df_res['실제금액'] = pd.to_numeric(df_res['실제금액'], errors='coerce').fillna(0)
    df_res['고객국적'] = df_res['고객국적'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
    df_res['종류'] = df_res['종류'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
    df_res['시술수술명'] = df_res['시술수술명'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
    df_res['예약상태'] = df_res['예약상태'].apply(lambda x: str(x).strip() if pd.notna(x) else '')

    df_completed = df_res[df_res['예약상태'] == '시/수술 완료'].copy()
    df_all = df_res.copy()  # 전체 (취소/노쇼 포함)

    # 정산 시트 (openpyxl)
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    # 시트 이름 호환: 로컬 Excel vs Google Sheets
    settle_sheet = None
    for name in wb.sheetnames:
        if '정산' in name:
            settle_sheet = name
            break
    if settle_sheet is None:
        raise ValueError(f"정산 시트를 찾을 수 없습니다. 시트 목록: {wb.sheetnames}")
    ws2 = wb[settle_sheet]
    settlement_records = []
    current_month = None
    for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, values_only=False):
        vals = [c.value for c in row]
        a_val = str(vals[0]).strip() if vals[0] else ''
        if '정산 내역' in a_val:
            parts = a_val.replace('년', '-').replace('월', '').replace('정산 내역', '').strip()
            try:
                year, month = parts.split('-')[:2]
                current_month = f"{year.strip()}-{int(month.strip()):02d}"
            except Exception:
                pass
            continue
        if a_val in ('NO', 'NO.', '', 'None', '재무팀 정산 요청 내역') or '정산 요청일' in a_val:
            continue
        if current_month and vals[1]:
            hospital = str(vals[1]).strip()
            if hospital in ('병원명', ''):
                continue
            try:
                settlement_records.append({
                    '정산월': current_month,
                    '병원명': normalize_hospital(hospital),
                    '고객명': str(vals[2]).strip() if vals[2] else '',
                    '국적': str(vals[4]).strip() if vals[4] else '',
                    '구분': str(vals[6]).strip() if vals[6] else '',
                    '시술금액': float(vals[7]) if vals[7] else 0,
                    '수수료금액': float(vals[8]) if vals[8] else 0,
                })
            except (ValueError, TypeError):
                pass
    wb.close()
    df_settle = pd.DataFrame(settlement_records)

    return df_completed, df_settle, df_all


@st.cache_data(show_spinner=False, ttl=600)
def load_internal_report(file_path):
    """내부리포트 Excel → 월별트렌드 df, 병원별성과 df, 취소노쇼 요약 df, 취소노쇼 상세 df"""

    if str(file_path).startswith('http'):
        import urllib.request
        tmp = '/tmp/unniguide_gsheet_ir.xlsx'
        urllib.request.urlretrieve(file_path, tmp)
        file_path = tmp

    # 월별 트렌드
    df_monthly = pd.read_excel(file_path, sheet_name='월별 트렌드', header=None, skiprows=3)
    df_monthly.columns = ['월', '완료건수', '시술건수', '수술건수', '시수술금액', '수수료매출', '평균객단가', '취소노쇼']
    df_monthly = df_monthly.dropna(subset=['월'])
    # 헤더 행 제거 (숫자로 변환 불가능한 행)
    df_monthly = df_monthly[pd.to_numeric(df_monthly['완료건수'], errors='coerce').notna()].copy()
    for col in ['완료건수', '시술건수', '수술건수', '시수술금액', '수수료매출', '평균객단가', '취소노쇼']:
        df_monthly[col] = pd.to_numeric(df_monthly[col], errors='coerce').fillna(0)

    # 병원별 성과
    df_hosp = pd.read_excel(file_path, sheet_name='병원별 성과', header=None, skiprows=3)
    df_hosp.columns = ['순위', '병원명', '누적건수', '누적시수술금액', '누적수수료', '최신월건수', '최신월금액', '전월대비']
    df_hosp = df_hosp.dropna(subset=['병원명'])
    df_hosp = df_hosp[pd.to_numeric(df_hosp['누적건수'], errors='coerce').notna()].copy()
    for col in ['누적건수', '누적시수술금액', '누적수수료', '최신월건수', '최신월금액', '전월대비']:
        df_hosp[col] = pd.to_numeric(df_hosp[col], errors='coerce').fillna(0)
    df_hosp['병원명'] = df_hosp['병원명'].apply(normalize_hospital)

    # 취소/노쇼 트래킹
    df_cancel_raw = pd.read_excel(file_path, sheet_name='취소 노쇼 트래킹', header=None)

    # 전체 요약 (row 4)
    total_cancel = df_cancel_raw.iloc[4, 1] if len(df_cancel_raw) > 4 else 0
    total_noshow = df_cancel_raw.iloc[4, 3] if len(df_cancel_raw) > 4 else 0
    cancel_rate = df_cancel_raw.iloc[4, 5] if len(df_cancel_raw) > 4 else ''

    cancel_summary = {
        'total_cancel': int(total_cancel) if pd.notna(total_cancel) else 0,
        'total_noshow': int(total_noshow) if pd.notna(total_noshow) else 0,
        'cancel_rate': str(cancel_rate),
    }

    # 병원별 취소/노쇼 (row 10~)
    cancel_hospital_rows = []
    for idx in range(11, len(df_cancel_raw)):
        row = df_cancel_raw.iloc[idx]
        hospital = row[0]
        if pd.isna(hospital) or str(hospital).strip() == '':
            break
        cancel_hospital_rows.append({
            '병원명': normalize_hospital(str(hospital).strip()),
            '전체예약': int(row[1]) if pd.notna(row[1]) else 0,
            '취소': int(row[2]) if pd.notna(row[2]) else 0,
            'No-show': int(row[3]) if pd.notna(row[3]) else 0,
            '취소노쇼율': float(row[4]) if pd.notna(row[4]) else 0,
        })
    df_cancel_hospital = pd.DataFrame(cancel_hospital_rows)

    # 상세 내역 (row 36~)
    detail_start = None
    for idx in range(30, len(df_cancel_raw)):
        val = df_cancel_raw.iloc[idx, 0]
        if pd.notna(val) and '월' == str(val).strip():
            detail_start = idx + 1
            break

    cancel_details = []
    if detail_start:
        for idx in range(detail_start, len(df_cancel_raw)):
            row = df_cancel_raw.iloc[idx]
            if pd.isna(row[0]):
                continue
            cancel_details.append({
                '월': str(row[0]).strip(),
                '상태': str(row[1]).strip() if pd.notna(row[1]) else '',
                '병원명': normalize_hospital(str(row[2]).strip()) if pd.notna(row[2]) else '',
                '국적': str(row[3]).strip() if pd.notna(row[3]) else '',
                '고객명': str(row[4]).strip() if pd.notna(row[4]) else '',
                '종류': str(row[5]).strip() if pd.notna(row[5]) else '',
                '시술수술명': str(row[6]).strip() if pd.notna(row[6]) else '',
            })
    df_cancel_detail = pd.DataFrame(cancel_details)

    return df_monthly, df_hosp, cancel_summary, df_cancel_hospital, df_cancel_detail


# ============================================================
# Google Sheets 설정
# ============================================================
GSHEET_ID_MAIN = "1pNQiaK67nz6FhxssxgWvoiQr6YwT-5MCwDVvqUZW1SY"
GSHEET_ID_REPORT = "16xOwlg8nptwbdM3uvbhr012v77xECiUIgy6wKjT5QYI"
GID_RESERVATION = 123775075   # 예약확정 시트
GID_SETTLEMENT = 622724794    # 내부리포트 (정산 포함)
GID_OFFLINE = 1126075757      # 오프라인 데일리
GID_DASHBOARD = 1029704191    # 가공 대시보드


def gsheet_xlsx_url(sheet_id):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"


def gsheet_csv_url(sheet_id, gid):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


# ============================================================
# 사이드바
# ============================================================
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 12px 0 20px;">
        <span style="font-size:28px; font-weight:800; color:{BRAND_PLUM};">UNNI</span>
        <span style="font-size:28px; font-weight:400; color:{BRAND_PLUM};"> GUIDE</span>
        <br><span style="font-size:13px; color:{BRAND_ORANGE}; font-weight:600;">운영 대시보드</span>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    data_source = st.radio("데이터 소스", ["Google Sheets (자동)", "로컬 파일"], index=0)

    df_completed = df_settle = df_all = None
    df_monthly = df_hosp_perf = cancel_summary = df_cancel_hospital = df_cancel_detail = None

    if data_source == "Google Sheets (자동)":
        st.caption("Google Sheets에서 자동으로 데이터를 불러옵니다.")
        if st.button("데이터 새로고침", type="primary"):
            st.cache_data.clear()

        try:
            with st.spinner("운영 데이터 로딩 중..."):
                xlsx_url = gsheet_xlsx_url(GSHEET_ID_MAIN)
                df_completed, df_settle, df_all = load_operation_excel(xlsx_url)
            st.success("운영 데이터 로드 완료")
        except Exception as e:
            st.error(f"운영 데이터 로드 실패: {str(e)[:50]}")

        try:
            with st.spinner("내부리포트 로딩 중..."):
                xlsx_url2 = gsheet_xlsx_url(GSHEET_ID_REPORT)
                df_monthly, df_hosp_perf, cancel_summary, df_cancel_hospital, df_cancel_detail = load_internal_report(xlsx_url2)
            st.success("내부리포트 로드 완료")
        except Exception as e:
            st.warning(f"내부리포트 로드 실패: {str(e)[:50]}")

    else:
        # 로컬 파일 모드 (기존 방식)
        st.markdown("**1. 운영 트렌드 데이터**")
        pattern1 = os.path.expanduser('~/Downloads/언니가이드 운영 트렌드 데이터_*.xlsx')
        pattern1b = os.path.expanduser('~/Desktop/언니가이드_리포트/언니가이드 운영 트렌드 데이터_*.xlsx')
        op_candidates = sorted(glob.glob(pattern1) + glob.glob(pattern1b), key=os.path.getmtime, reverse=True)

        if op_candidates:
            op_file = st.selectbox("운영 Excel", op_candidates, format_func=os.path.basename, key="op")
            with st.spinner("운영 데이터 로딩..."):
                df_completed, df_settle, df_all = load_operation_excel(op_file)
        else:
            op_upload = st.file_uploader("운영 Excel 업로드", type=['xlsx'], key="op_up")
            if op_upload:
                tmp = "/tmp/unniguide_op.xlsx"
                with open(tmp, "wb") as f:
                    f.write(op_upload.getvalue())
                with st.spinner("운영 데이터 로딩..."):
                    df_completed, df_settle, df_all = load_operation_excel(tmp)

        st.markdown("**2. 내부 리포트**")
        pattern2 = os.path.expanduser('~/Downloads/언니가이드_내부리포트_*.xlsx')
        pattern2b = os.path.expanduser('~/Desktop/언니가이드_리포트/언니가이드_내부리포트_*.xlsx')
        ir_candidates = sorted(glob.glob(pattern2) + glob.glob(pattern2b), key=os.path.getmtime, reverse=True)

        if ir_candidates:
            ir_file = st.selectbox("내부리포트 Excel", ir_candidates, format_func=os.path.basename, key="ir")
            with st.spinner("내부리포트 로딩..."):
                df_monthly, df_hosp_perf, cancel_summary, df_cancel_hospital, df_cancel_detail = load_internal_report(ir_file)
        else:
            ir_upload = st.file_uploader("내부리포트 Excel 업로드", type=['xlsx'], key="ir_up")
            if ir_upload:
                tmp2 = "/tmp/unniguide_ir.xlsx"
                with open(tmp2, "wb") as f:
                    f.write(ir_upload.getvalue())
                with st.spinner("내부리포트 로딩..."):
                    df_monthly, df_hosp_perf, cancel_summary, df_cancel_hospital, df_cancel_detail = load_internal_report(tmp2)

# ============================================================
# 데이터 없으면 안내
# ============================================================
if df_completed is None:
    st.markdown(f"""
    <div style="text-align:center; padding:80px 0;">
        <div style="font-size:48px; margin-bottom:16px;">📊</div>
        <h2>언니가이드 운영 대시보드</h2>
        <p style="color:#636E72; font-size:16px; margin-top:8px;">
            왼쪽 사이드바에서 데이터 파일을 선택하거나 업로드해주세요.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ============================================================
# 필터
# ============================================================
all_months_res = sorted(df_completed['월'].dropna().unique())
all_months_settle = sorted(df_settle['정산월'].unique()) if len(df_settle) > 0 else []
all_months = sorted(set(all_months_res + all_months_settle))

if not all_months:
    st.error("데이터에 유효한 월 정보가 없습니다.")
    st.stop()

with st.sidebar:
    st.divider()
    st.subheader("필터")
    if len(all_months) >= 2:
        # 정산 데이터가 있는 최신월을 기본 끝점으로 (진행 중인 월 제외)
        default_end = all_months_settle[-1] if all_months_settle else all_months[-1]
        if default_end not in all_months:
            default_end = all_months[-1]
        month_range = st.select_slider("기간 선택", options=all_months, value=(all_months[0], default_end))
        selected_months = [m for m in all_months if month_range[0] <= m <= month_range[1]]
    else:
        selected_months = all_months
        month_range = (all_months[0], all_months[0])

    all_nationalities = sorted([n for n in df_completed['고객국적'].unique() if n and n != 'nan'])
    selected_nationalities = st.multiselect("국적", all_nationalities, default=[], placeholder="전체 국적")

    all_hospitals = sorted(df_completed['병원명'].dropna().unique())
    selected_hospitals = st.multiselect("병원", all_hospitals, default=[], placeholder="전체 병원")
    st.divider()
    st.caption(f"예약완료 {len(df_completed):,}건 | 정산 {len(df_settle):,}건")

# 필터 적용
mask_res = df_completed['월'].isin(selected_months)
mask_set = df_settle['정산월'].isin(selected_months) if len(df_settle) > 0 else pd.Series(dtype=bool)
if selected_nationalities:
    mask_res = mask_res & df_completed['고객국적'].isin(selected_nationalities)
    if len(df_settle) > 0:
        mask_set = mask_set & df_settle['국적'].isin(selected_nationalities)
if selected_hospitals:
    mask_res = mask_res & df_completed['병원명'].isin(selected_hospitals)
    if len(df_settle) > 0:
        mask_set = mask_set & df_settle['병원명'].isin(selected_hospitals)

filtered_res = df_completed[mask_res].copy()
filtered_set = df_settle[mask_set].copy() if len(df_settle) > 0 else pd.DataFrame()

# ============================================================
# 헤더
# ============================================================
period_label = f"{month_range[0]} ~ {month_range[1]}" if month_range[0] != month_range[1] else month_range[0]
st.markdown(f"""
<div style="background: linear-gradient(135deg, {BRAND_ORANGE} 0%, #E8551F 100%);
     color: white; padding: 28px 32px; border-radius: 0 0 20px 20px; margin: -1rem -1rem 24px -1rem;">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <div style="font-size:14px; opacity:0.85; margin-bottom:4px;">UNNI GUIDE 운영 대시보드</div>
            <div style="font-size:24px; font-weight:800;">{period_label} 데이터</div>
        </div>
        <div style="font-size:13px; opacity:0.75;">예약 {len(filtered_res):,}건 · 정산 {len(filtered_set):,}건</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# 탭
# ============================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📊 Overview", "🌍 국적 분석", "🏥 병원 분석", "💉 시술 트렌드", "⚠️ 취소/No-show", "📋 원본 데이터", "📦 리포트 생성"
])

# 공통: 최신월/전월
latest_month = selected_months[-1] if selected_months else None
prev_candidates = [m for m in all_months if m < latest_month] if latest_month else []
prev_month = prev_candidates[-1] if prev_candidates else None

# ============================================================
# Tab 1: Overview
# ============================================================
with tab1:
    latest_res = filtered_res[filtered_res['월'] == latest_month] if latest_month else filtered_res
    prev_res = df_completed[df_completed['월'] == prev_month] if prev_month else pd.DataFrame()
    latest_set = filtered_set[filtered_set['정산월'] == latest_month] if latest_month and len(filtered_set) > 0 else filtered_set
    prev_set = df_settle[df_settle['정산월'] == prev_month] if prev_month and len(df_settle) > 0 else pd.DataFrame()

    if selected_nationalities:
        prev_res = prev_res[prev_res['고객국적'].isin(selected_nationalities)]
        if len(prev_set) > 0:
            prev_set = prev_set[prev_set['국적'].isin(selected_nationalities)]
    if selected_hospitals:
        prev_res = prev_res[prev_res['병원명'].isin(selected_hospitals)]
        if len(prev_set) > 0:
            prev_set = prev_set[prev_set['병원명'].isin(selected_hospitals)]

    # --- KPI ---
    st.subheader(f"핵심 지표 ({latest_month})" if latest_month else "핵심 지표")
    col1, col2, col3, col4, col5 = st.columns(5)

    cur_cnt = len(latest_set) if len(latest_set) > 0 else len(latest_res)
    prev_cnt = len(prev_set) if len(prev_set) > 0 else len(prev_res)
    cur_rev = latest_set['시술금액'].sum() if len(latest_set) > 0 else 0
    prev_rev = prev_set['시술금액'].sum() if len(prev_set) > 0 else 0
    cur_comm = latest_set['수수료금액'].sum() if len(latest_set) > 0 else 0
    prev_comm = prev_set['수수료금액'].sum() if len(prev_set) > 0 else 0
    avg_price = cur_rev / cur_cnt if cur_cnt > 0 else 0
    prev_avg = prev_rev / prev_cnt if prev_cnt > 0 else 0

    def pct_delta(cur, prev):
        return f"{(cur - prev) / max(prev, 1) * 100:+.1f}%" if prev > 0 else None

    with col1:
        st.metric("시/수술 완료", f"{cur_cnt:,}건", pct_delta(cur_cnt, prev_cnt))
    with col2:
        st.metric("정산 매출", format_krw(cur_rev), pct_delta(cur_rev, prev_rev))
    with col3:
        st.metric("수수료 매출", format_krw(cur_comm), pct_delta(cur_comm, prev_comm))
    with col4:
        st.metric("인당 객단가", format_krw(avg_price), pct_delta(avg_price, prev_avg))
    with col5:
        nat_count = latest_set['국적'].nunique() if len(latest_set) > 0 else latest_res['고객국적'].nunique()
        st.metric("참여 국적", f"{nat_count}개국")

    # --- 운영 현황 (내부리포트 기반) ---
    if df_monthly is not None and latest_month:
        m_row = df_monthly[df_monthly['월'] == latest_month]
        if len(m_row) > 0:
            r = m_row.iloc[0]
            # 전체 예약 접수 계산 (완료 + 취소+노쇼 + 기타)
            all_month = df_all[df_all['월'] == latest_month] if df_all is not None else pd.DataFrame()
            total_접수 = len(all_month)
            cancel_cnt = len(all_month[all_month['예약상태'] == '예약 취소']) if len(all_month) > 0 else 0
            noshow_cnt = len(all_month[all_month['예약상태'].str.lower().str.contains('no-show|no show|noshow', na=False)]) if len(all_month) > 0 else 0

            st.markdown("")
            st.subheader(f"{latest_month} 운영 현황")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("전체 예약 접수", f"{total_접수}건")
            c2.metric("시/수술 완료", f"{int(r['완료건수'])}건")
            c3.metric("예약 취소", f"{cancel_cnt}건")
            c4.metric("No-show", f"{noshow_cnt}건")

    st.markdown("")

    # --- 월별 트렌드 차트 (내부리포트 우선) ---
    st.subheader("월별 트렌드")
    col_c1, col_c2 = st.columns(2)

    with col_c1:
        if df_monthly is not None:
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=df_monthly['월'], y=df_monthly['완료건수'],
                name='완료 건수', marker_color=BRAND_ORANGE, opacity=0.8,
                text=df_monthly['완료건수'].apply(lambda x: int(x) if pd.notna(x) else 0), textposition='outside',
            ))
            fig1.add_trace(go.Scatter(
                x=df_monthly['월'], y=df_monthly['완료건수'].cumsum(),
                name='누적 건수', line=dict(color=BRAND_PLUM, width=2.5),
                mode='lines+markers', yaxis='y2',
            ))
            fig1.update_layout(
                title="월별 완료 건수 & 누적",
                yaxis=dict(title="월별 건수"), yaxis2=dict(title="누적", overlaying='y', side='right'),
                legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
                height=420, margin=dict(t=50, b=80),
            )
            st.plotly_chart(fig1, use_container_width=True)
        else:
            monthly_res = df_completed.groupby('월').size().reset_index(name='건수').sort_values('월')
            monthly_res['누적'] = monthly_res['건수'].cumsum()
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(x=monthly_res['월'], y=monthly_res['건수'], name='완료', marker_color=BRAND_ORANGE, text=monthly_res['건수'], textposition='outside'))
            fig1.add_trace(go.Scatter(x=monthly_res['월'], y=monthly_res['누적'], name='누적', line=dict(color=BRAND_PLUM, width=2.5), mode='lines+markers', yaxis='y2'))
            fig1.update_layout(title="월별 완료 건수", yaxis=dict(title="건수"), yaxis2=dict(title="누적", overlaying='y', side='right'), height=420, margin=dict(t=50, b=80), legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5))
            st.plotly_chart(fig1, use_container_width=True)

    with col_c2:
        if df_monthly is not None:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                x=df_monthly['월'], y=df_monthly['시수술금액'],
                name='시/수술 금액', marker_color=BRAND_ORANGE, opacity=0.8,
                text=[format_krw(v) for v in df_monthly['시수술금액']], textposition='outside',
            ))
            fig2.add_trace(go.Bar(
                x=df_monthly['월'], y=df_monthly['수수료매출'],
                name='수수료 매출', marker_color=BRAND_GREEN, opacity=0.7,
            ))
            fig2.update_layout(
                title="월별 매출 & 수수료",
                yaxis=dict(title="금액 (원)"), barmode='group',
                legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
                height=420, margin=dict(t=50, b=80),
            )
            st.plotly_chart(fig2, use_container_width=True)
        elif len(df_settle) > 0:
            ms = df_settle.groupby('정산월').agg(매출=('시술금액','sum'), 수수료=('수수료금액','sum')).reset_index().sort_values('정산월')
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(x=ms['정산월'], y=ms['매출'], name='매출', marker_color=BRAND_ORANGE, text=[format_krw(v) for v in ms['매출']], textposition='outside'))
            fig2.add_trace(go.Bar(x=ms['정산월'], y=ms['수수료'], name='수수료', marker_color=BRAND_GREEN))
            fig2.update_layout(title="월별 매출", barmode='group', height=420, margin=dict(t=50, b=80), legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5))
            st.plotly_chart(fig2, use_container_width=True)

    # --- 월별 성장률 테이블 ---
    st.subheader("월별 성장률 (MoM)")
    if df_monthly is not None:
        display_m = df_monthly.copy()
        display_m['시수술금액_표시'] = display_m['시수술금액'].apply(format_krw)
        display_m['수수료매출_표시'] = display_m['수수료매출'].apply(format_krw)
        display_m['평균객단가_표시'] = display_m['평균객단가'].apply(format_krw)
        display_m['건수MoM'] = display_m['완료건수'].pct_change().apply(lambda x: f"{x*100:+.1f}%" if pd.notna(x) else '-')
        display_m['매출MoM'] = display_m['시수술금액'].pct_change().apply(lambda x: f"{x*100:+.1f}%" if pd.notna(x) else '-')
        st.dataframe(
            display_m[['월', '완료건수', '시수술금액_표시', '수수료매출_표시', '평균객단가_표시', '취소노쇼', '건수MoM', '매출MoM']].rename(columns={
                '시수술금액_표시': '시/수술 금액', '수수료매출_표시': '수수료', '평균객단가_표시': '객단가',
                '취소노쇼': '취소+노쇼', '건수MoM': '건수 MoM', '매출MoM': '매출 MoM',
            }),
            use_container_width=True, hide_index=True,
        )


# ============================================================
# Tab 2: 국적 분석
# ============================================================
with tab2:
    st.subheader("국적별 고객 분석")
    st.caption(f"데이터 기간: {period_label} | 정산 데이터 기준")

    if len(filtered_set) > 0:
        nat_data = filtered_set.groupby('국적').agg(
            건수=('시술금액', 'count'), 매출=('시술금액', 'sum'), 수수료=('수수료금액', 'sum'),
        ).sort_values('매출', ascending=False).reset_index()
        nat_data['비중'] = (nat_data['건수'] / nat_data['건수'].sum() * 100).round(1)
        nat_data['객단가'] = (nat_data['매출'] / nat_data['건수']).round(0)
        nat_data['국기'] = nat_data['국적'].map(COUNTRY_FLAG).fillna('🌍')
        nat_data['국적표시'] = nat_data['국기'] + ' ' + nat_data['국적']
        label_col = '국적표시'
    else:
        nat_data = filtered_res.groupby('고객국적').agg(
            건수=('고객국적', 'count'), 매출=('실제금액', 'sum'),
        ).sort_values('건수', ascending=False).reset_index()
        nat_data['비중'] = (nat_data['건수'] / nat_data['건수'].sum() * 100).round(1)
        nat_data['객단가'] = (nat_data['매출'] / nat_data['건수']).round(0)
        nat_data['국적표시'] = nat_data['고객국적']
        label_col = '국적표시'

    # 도넛 차트용: 상위 8개 + 기타
    TOP_N = 8
    pie_data = nat_data.head(TOP_N).copy()
    if len(nat_data) > TOP_N:
        etc = nat_data.iloc[TOP_N:]
        etc_row = pd.DataFrame([{
            label_col: '🌐 기타 ' + str(len(etc)) + '개국',
            '건수': etc['건수'].sum(),
            '매출': etc['매출'].sum(),
        }])
        pie_data = pd.concat([pie_data, etc_row], ignore_index=True)

    col1, col2 = st.columns(2)
    with col1:
        fig_d = px.pie(pie_data, values='건수', names=label_col, color_discrete_sequence=CHART_COLORS, hole=0.45)
        fig_d.update_layout(
            title="국적별 예약 비중", height=450,
            legend=dict(font=dict(size=13), orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05),
            margin=dict(r=160),
        )
        fig_d.update_traces(textinfo='percent', textfont_size=13, insidetextorientation='horizontal')
        st.plotly_chart(fig_d, use_container_width=True)
    with col2:
        top_n = nat_data.head(10).sort_values('객단가')
        fig_b = px.bar(top_n, x='객단가', y=label_col, orientation='h', color_discrete_sequence=[BRAND_ORANGE],
                       text=top_n['객단가'].apply(format_krw))
        fig_b.update_layout(title="국적별 인당 객단가", height=420)
        fig_b.update_traces(textposition='outside')
        st.plotly_chart(fig_b, use_container_width=True)

    # 상세 테이블
    st.subheader("국적별 상세 데이터")
    disp = nat_data[[label_col, '건수', '비중', '매출', '객단가']].copy()
    if '수수료' in nat_data.columns:
        disp['수수료'] = nat_data['수수료'].apply(format_krw)
    disp.columns = ['국적', '건수', '비중(%)', '매출', '객단가'] + (['수수료'] if '수수료' in disp.columns else [])
    disp['매출'] = nat_data['매출'].apply(format_krw)
    disp['객단가'] = nat_data['객단가'].apply(format_krw)
    st.dataframe(disp, use_container_width=True, hide_index=True)

    # 월별 국적 추이
    st.subheader("월별 국적 인입 트렌드")
    nat_col = '국적' if len(filtered_set) > 0 else '고객국적'
    month_col = '정산월' if len(filtered_set) > 0 else '월'
    src = filtered_set if len(filtered_set) > 0 else filtered_res
    top5 = nat_data.head(5)[nat_data.columns[0]].tolist() if '국적' not in nat_data.columns else nat_data.head(5)['국적'].tolist()
    if len(src) > 0:
        trend = src[src[nat_col].isin(top5)].groupby([month_col, nat_col]).size().reset_index(name='건수')
        if len(trend) > 0:
            fig_t = px.line(trend, x=month_col, y='건수', color=nat_col, color_discrete_sequence=CHART_COLORS, markers=True)
            fig_t.update_layout(height=420, margin=dict(t=20, b=80), legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, font=dict(size=12)))
            st.plotly_chart(fig_t, use_container_width=True)


# ============================================================
# Tab 3: 병원 분석
# ============================================================
with tab3:
    st.subheader("병원별 성과")
    st.caption(f"데이터 기간: {period_label} | 정산 데이터 기준")

    # 내부리포트 병원별 성과가 있으면 우선 사용
    if df_hosp_perf is not None and len(df_hosp_perf) > 0:
        st.markdown("*정산 기준 누적 데이터*")
        # 정규화된 이름 기준으로 합산 (같은 병원 다른 표기 합치기)
        hosp_agg = df_hosp_perf.groupby('병원명').agg(
            누적건수=('누적건수', 'sum'), 누적시수술금액=('누적시수술금액', 'sum'),
            누적수수료=('누적수수료', 'sum'), 최신월건수=('최신월건수', 'sum'), 최신월금액=('최신월금액', 'sum'),
        ).sort_values('누적시수술금액', ascending=False).reset_index()
        hosp_agg['순위'] = range(1, len(hosp_agg) + 1)

        fig_h = px.bar(
            hosp_agg.head(15).sort_values('누적시수술금액'), x='누적시수술금액', y='병원명',
            orientation='h', color_discrete_sequence=[BRAND_ORANGE],
            text=[format_krw(v) for v in hosp_agg.head(15).sort_values('누적시수술금액')['누적시수술금액']],
        )
        fig_h.update_layout(title=f"병원 누적 매출 순위 TOP 15", height=max(400, 15 * 35), margin=dict(l=180, t=50))
        fig_h.update_traces(textposition='outside')
        st.plotly_chart(fig_h, use_container_width=True)

        # 최신월 수수료 계산 (정산 데이터에서)
        latest_settle_month = all_months_settle[-1] if all_months_settle else None
        if latest_settle_month and len(df_settle) > 0:
            latest_comm = df_settle[df_settle['정산월'] == latest_settle_month].groupby('병원명')['수수료금액'].sum().reset_index()
            latest_comm.columns = ['병원명', '최신월수수료']
            hosp_agg = hosp_agg.merge(latest_comm, on='병원명', how='left')
            hosp_agg['최신월수수료'] = hosp_agg['최신월수수료'].fillna(0)
        else:
            hosp_agg['최신월수수료'] = 0

        disp_h = hosp_agg[['순위', '병원명', '누적건수', '누적시수술금액', '누적수수료', '최신월건수', '최신월금액', '최신월수수료']].copy()
        disp_h['누적시수술금액'] = disp_h['누적시수술금액'].apply(format_krw)
        disp_h['누적수수료'] = disp_h['누적수수료'].apply(format_krw)
        disp_h['최신월금액'] = disp_h['최신월금액'].apply(format_krw)
        disp_h['최신월수수료'] = disp_h['최신월수수료'].apply(format_krw)
        month_label = latest_settle_month if latest_settle_month else '최신월'
        disp_h.columns = ['순위', '병원명', '누적건수', '누적매출', '누적수수료', f'{month_label} 건수', f'{month_label} 매출', f'{month_label} 수수료']
        st.dataframe(disp_h, use_container_width=True, hide_index=True)

    elif len(filtered_set) > 0:
        hosp_s = filtered_set.groupby('병원명').agg(
            건수=('시술금액', 'count'), 매출=('시술금액', 'sum'), 수수료=('수수료금액', 'sum'),
        ).sort_values('매출', ascending=False).reset_index()
        hosp_s['객단가'] = (hosp_s['매출'] / hosp_s['건수']).round(0)
        hosp_s['순위'] = range(1, len(hosp_s) + 1)

        fig_h = px.bar(hosp_s.head(15).sort_values('매출'), x='매출', y='병원명', orientation='h',
                       color_discrete_sequence=[BRAND_ORANGE], text=[format_krw(v) for v in hosp_s.head(15).sort_values('매출')['매출']])
        fig_h.update_layout(title="병원 매출 순위 TOP 15", height=max(400, 15*35), margin=dict(l=180, t=50))
        fig_h.update_traces(textposition='outside')
        st.plotly_chart(fig_h, use_container_width=True)

        disp_h = hosp_s[['순위', '병원명', '건수', '매출', '수수료', '객단가']].copy()
        disp_h['매출'] = disp_h['매출'].apply(format_krw)
        disp_h['수수료'] = disp_h['수수료'].apply(format_krw)
        disp_h['객단가'] = disp_h['객단가'].apply(format_krw)
        st.dataframe(disp_h, use_container_width=True, hide_index=True)

    # 병원별 월별 추이
    st.subheader("병원별 월별 매출 추이")
    if len(df_settle) > 0:
        top_h = list(filtered_set.groupby('병원명')['시술금액'].sum().sort_values(ascending=False).head(10).index) if len(filtered_set) > 0 else []
        sel_h = st.multiselect("병원 선택", options=top_h + [h for h in all_hospitals if h not in top_h], default=top_h[:5], max_selections=10, key="hosp_trend")
        if sel_h:
            hm = df_settle[df_settle['병원명'].isin(sel_h)].groupby(['정산월', '병원명'])['시술금액'].sum().reset_index()
            fig_ht = px.line(hm, x='정산월', y='시술금액', color='병원명', color_discrete_sequence=CHART_COLORS, markers=True)
            fig_ht.update_layout(title="선택 병원 월별 매출", yaxis_title="정산 매출", height=480, margin=dict(t=40, b=100), legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, font=dict(size=11)))
            st.plotly_chart(fig_ht, use_container_width=True)

    # 월별 1위
    st.subheader("월별 매출 1위 병원")
    if len(df_settle) > 0:
        top1 = []
        for m in sorted(all_months_settle):
            md = df_settle[df_settle['정산월'] == m]
            if len(md) > 0:
                hr = md.groupby('병원명')['시술금액'].sum().sort_values(ascending=False)
                top1.append({'월': m, '1위 병원': hr.index[0], '매출': format_krw(hr.values[0]), '건수': len(md[md['병원명'] == hr.index[0]])})
        if top1:
            st.dataframe(pd.DataFrame(top1), use_container_width=True, hide_index=True)


# ============================================================
# Tab 4: 시술 트렌드
# ============================================================
with tab4:
    st.subheader("시술 트렌드")
    st.caption(f"데이터 기간: {period_label} | 전체 예약 데이터 기준 (시/수술 완료 + 예약확정 포함)")

    # 시술 데이터: 완료건뿐 아니라 전체 예약 중 시술명이 있는 건 활용
    if df_all is not None:
        proc_source = df_all[df_all['월'].isin(selected_months)].copy()
        if selected_nationalities:
            proc_source = proc_source[proc_source['고객국적'].isin(selected_nationalities)]
        if selected_hospitals:
            proc_source = proc_source[proc_source['병원명'].isin(selected_hospitals)]
    else:
        proc_source = filtered_res.copy()

    proc_source_valid = proc_source[proc_source['시술수술명'].str.strip() != ''].copy()

    col1, col2 = st.columns(2)

    with col1:
        # 시술/수술 비중 - 내부리포트 or 정산 기준
        if df_monthly is not None:
            m_filtered = df_monthly[df_monthly['월'].isin(selected_months)] if len(df_monthly) > 0 else df_monthly
            total_시술 = m_filtered['시술건수'].sum()
            total_수술 = m_filtered['수술건수'].sum()
            type_df = pd.DataFrame({'구분': ['시술', '수술'], '건수': [int(total_시술), int(total_수술)]})
        elif len(filtered_set) > 0:
            tc = filtered_set['구분'].value_counts().reset_index()
            tc.columns = ['구분', '건수']
            type_df = tc[tc['구분'].str.strip() != '']
        else:
            tc = proc_source['종류'].value_counts().reset_index()
            tc.columns = ['구분', '건수']
            type_df = tc[tc['구분'].str.strip() != '']

        if len(type_df) > 0:
            fig_ty = px.pie(type_df, values='건수', names='구분', color_discrete_sequence=[BRAND_ORANGE, BRAND_PLUM, BRAND_GREEN], hole=0.45)
            fig_ty.update_layout(title="시술 vs 수술 비중 (정산 기준)", height=380)
            fig_ty.update_traces(textinfo='percent+label')
            st.plotly_chart(fig_ty, use_container_width=True)

    with col2:
        # 키워드 기반 시술 카테고리 집계
        all_proc_rows = []
        for _, row in proc_source_valid.iterrows():
            procs = extract_procedures(row['시술수술명'])
            for p in procs:
                all_proc_rows.append({'시술카테고리': p, '매출': row['실제금액']})
        df_proc_kw = pd.DataFrame(all_proc_rows)

        if len(df_proc_kw) > 0:
            top_kw = df_proc_kw.groupby('시술카테고리').agg(
                건수=('시술카테고리', 'count'), 매출=('매출', 'sum'),
            ).sort_values('건수', ascending=False).head(15).reset_index()
            fig_p = px.bar(
                top_kw.sort_values('건수'), x='건수', y='시술카테고리',
                orientation='h', color_discrete_sequence=[BRAND_ORANGE], text='건수',
            )
            fig_p.update_layout(title=f"시술 카테고리 TOP 15 ({period_label})", height=480, margin=dict(l=160, t=50))
            fig_p.update_traces(textposition='outside')
            st.plotly_chart(fig_p, use_container_width=True)

    # 시술 카테고리 상세 테이블
    st.subheader("시술 카테고리별 상세 데이터")
    if len(df_proc_kw) > 0:
        top_kw_full = df_proc_kw.groupby('시술카테고리').agg(
            건수=('시술카테고리', 'count'), 매출=('매출', 'sum'),
        ).sort_values('건수', ascending=False).reset_index()
        top_kw_full['매출표시'] = top_kw_full['매출'].apply(format_krw)
        top_kw_full['객단가'] = (top_kw_full['매출'] / top_kw_full['건수']).apply(format_krw)
        st.dataframe(
            top_kw_full[['시술카테고리', '건수', '매출표시', '객단가']].rename(columns={'시술카테고리': '시술 카테고리', '매출표시': '총 매출'}),
            use_container_width=True, hide_index=True,
        )
        st.caption("* 하나의 예약에 여러 시술이 포함된 경우 각각 카운트됩니다.")

    # 국적별 선호 시술 (키워드 기반)
    st.subheader("국적별 선호 시술 TOP 5")
    top_c = nat_data.head(5)['국적'].tolist() if '국적' in nat_data.columns else nat_data.head(5)['고객국적'].tolist() if '고객국적' in nat_data.columns else []
    if top_c:
        cols_p = st.columns(min(len(top_c), 3))
        for i, country in enumerate(top_c):
            with cols_p[i % 3]:
                flag = COUNTRY_FLAG.get(country, '🌍')
                c_procs = proc_source_valid[proc_source_valid['고객국적'] == country]['시술수술명']
                kw_list = []
                for txt in c_procs:
                    kw_list.extend(extract_procedures(txt))
                cp = pd.Series(kw_list).value_counts().head(5)
                if len(cp) > 0:
                    st.markdown(f"**{flag} {country}**")
                    for proc, cnt in cp.items():
                        st.markdown(f"- {proc[:35]} ({cnt}건)")
                    st.markdown("")


# ============================================================
# Tab 5: 취소/No-show
# ============================================================
with tab5:
    st.subheader("취소 / No-show 트래킹")
    st.caption("데이터 기간: 전체 누적 | 내부리포트 기준")

    if cancel_summary is not None:
        # 전체 요약
        c1, c2, c3 = st.columns(3)
        c1.metric("총 취소", f"{cancel_summary['total_cancel']}건")
        c2.metric("총 No-show", f"{cancel_summary['total_noshow']}건")
        c3.metric("취소+노쇼율", cancel_summary['cancel_rate'])

        st.markdown("")

        # 월별 취소+노쇼 추이
        if df_monthly is not None:
            st.subheader("월별 취소+노쇼 추이")
            fig_cn = go.Figure()
            fig_cn.add_trace(go.Bar(
                x=df_monthly['월'], y=df_monthly['취소노쇼'].apply(lambda x: int(x) if pd.notna(x) else 0),
                name='취소+노쇼', marker_color=BRAND_RED, opacity=0.8,
                text=df_monthly['취소노쇼'].apply(lambda x: int(x) if pd.notna(x) else 0), textposition='outside',
            ))
            cancel_rate_monthly = (df_monthly['취소노쇼'] / (df_monthly['완료건수'] + df_monthly['취소노쇼']) * 100).round(1)
            fig_cn.add_trace(go.Scatter(
                x=df_monthly['월'], y=cancel_rate_monthly,
                name='취소+노쇼율(%)', line=dict(color=BRAND_PLUM, width=2.5),
                mode='lines+markers+text', yaxis='y2',
                text=[f"{v}%" for v in cancel_rate_monthly], textposition='top center',
            ))
            fig_cn.update_layout(
                yaxis=dict(title="건수"), yaxis2=dict(title="비율(%)", overlaying='y', side='right'),
                height=360, legend=dict(orientation="h", yanchor="bottom", y=1.08),
            )
            st.plotly_chart(fig_cn, use_container_width=True)

        # 병원별 취소/노쇼 현황
        if df_cancel_hospital is not None and len(df_cancel_hospital) > 0:
            st.subheader("병원별 취소/No-show 현황")

            fig_ch = go.Figure()
            df_ch = df_cancel_hospital.sort_values('취소노쇼율', ascending=True)
            fig_ch.add_trace(go.Bar(x=df_ch['취소'], y=df_ch['병원명'], name='취소', orientation='h', marker_color=BRAND_ORANGE))
            fig_ch.add_trace(go.Bar(x=df_ch['No-show'], y=df_ch['병원명'], name='No-show', orientation='h', marker_color=BRAND_RED))
            fig_ch.update_layout(
                title="병원별 취소 & No-show", barmode='stack',
                height=max(400, len(df_ch) * 28), margin=dict(l=180, t=50),
                legend=dict(orientation="h", yanchor="bottom", y=1.08),
            )
            st.plotly_chart(fig_ch, use_container_width=True)

            # 테이블
            disp_cn = df_cancel_hospital.sort_values('취소노쇼율', ascending=False).copy()
            disp_cn['취소노쇼율'] = (disp_cn['취소노쇼율'] * 100).round(1).astype(str) + '%'
            st.dataframe(disp_cn, use_container_width=True, hide_index=True)

        # 상세 내역
        if df_cancel_detail is not None and len(df_cancel_detail) > 0:
            st.subheader("취소/No-show 상세 내역")
            # 필터
            status_filter = st.multiselect("상태", df_cancel_detail['상태'].unique().tolist(), default=df_cancel_detail['상태'].unique().tolist(), key="cn_status")
            filtered_cn = df_cancel_detail[df_cancel_detail['상태'].isin(status_filter)]
            if selected_hospitals:
                filtered_cn = filtered_cn[filtered_cn['병원명'].isin(selected_hospitals)]
            st.dataframe(filtered_cn.sort_values('월', ascending=False), use_container_width=True, hide_index=True, height=400)

    else:
        st.info("내부리포트 Excel을 업로드하면 취소/No-show 데이터를 볼 수 있습니다.")
        # 운영 데이터에서 기본 취소/노쇼 추출
        if df_all is not None:
            cancel_df = df_all[df_all['예약상태'].isin(['예약 취소'])].copy()
            noshow_df = df_all[df_all['예약상태'].str.lower().str.contains('no-show|no show|noshow', na=False)].copy()
            c1, c2 = st.columns(2)
            c1.metric("예약 취소 (전체)", f"{len(cancel_df)}건")
            c2.metric("No-show (전체)", f"{len(noshow_df)}건")


# ============================================================
# Tab 6: 원본 데이터
# ============================================================
with tab6:
    st.subheader("원본 데이터 조회 및 다운로드")
    data_type = st.radio("데이터 선택", ["예약 완료 데이터", "정산 데이터", "전체 예약 (취소/노쇼 포함)"], horizontal=True)

    if data_type == "예약 완료 데이터":
        cols = ['월', '병원명', '고객국적', '종류', '시술수술명', '실제금액']
        avail = [c for c in cols if c in filtered_res.columns]
        df_d = filtered_res[avail].sort_values('월', ascending=False)
        st.dataframe(df_d, use_container_width=True, hide_index=True, height=500)
        st.download_button("CSV 다운로드", df_d.to_csv(index=False).encode('utf-8-sig'), "예약완료_필터.csv", "text/csv")

    elif data_type == "정산 데이터":
        if len(filtered_set) > 0:
            cols = ['정산월', '병원명', '국적', '구분', '시술금액', '수수료금액']
            avail = [c for c in cols if c in filtered_set.columns]
            df_d = filtered_set[avail].sort_values('정산월', ascending=False)
            st.dataframe(df_d, use_container_width=True, hide_index=True, height=500)
            st.download_button("CSV 다운로드", df_d.to_csv(index=False).encode('utf-8-sig'), "정산_필터.csv", "text/csv")
        else:
            st.info("정산 데이터가 없습니다.")

    else:
        if df_all is not None:
            cols = ['월', '병원명', '고객국적', '예약상태', '종류', '시술수술명', '실제금액']
            avail = [c for c in cols if c in df_all.columns]
            mask = df_all['월'].isin(selected_months)
            if selected_nationalities:
                mask = mask & df_all['고객국적'].isin(selected_nationalities)
            if selected_hospitals:
                mask = mask & df_all['병원명'].isin(selected_hospitals)
            df_d = df_all[mask][avail].sort_values('월', ascending=False)
            st.dataframe(df_d, use_container_width=True, hide_index=True, height=500)
            st.download_button("CSV 다운로드", df_d.to_csv(index=False).encode('utf-8-sig'), "전체예약_필터.csv", "text/csv")


# ============================================================
# Tab 7: 리포트 생성 (팀원 누구나 사용 가능)
# ============================================================
with tab7:
    st.subheader("📦 병원용 HTML 리포트 생성")
    st.caption("버튼 클릭 한 번으로 공통 리포트 + 병원별 34개 리포트를 생성하여 ZIP 파일로 다운로드합니다.")

    st.markdown("**생성 기준월 선택**")
    report_month = st.selectbox(
        "리포트 월",
        options=all_months,
        index=len(all_months) - 1 if all_months else 0,
        format_func=lambda x: f"{x} ({datetime.strptime(x + '-01', '%Y-%m-%d').strftime('%Y년 %m월')})",
        key="report_month_sel",
    )

    st.markdown("")
    st.info("""
**생성 프로세스:**
1. 아래 버튼 클릭 → 스크립트 실행 (약 10-30초 소요)
2. ZIP 파일 자동 다운로드
3. 압축 풀면 공통 리포트 + 병원별 34개 HTML 파일
4. 각 병원에 카톡/이메일로 개별 전달
    """)

    if st.button("🚀 리포트 일괄 생성 + ZIP 다운로드", type="primary", use_container_width=True):
        import subprocess
        import zipfile
        import tempfile
        import shutil

        with st.spinner(f"{report_month} 리포트 생성 중..."):
            try:
                # Google Sheets URL에서 임시 xlsx 다운로드
                import urllib.request
                tmp_xlsx = "/tmp/unniguide_report_input.xlsx"
                urllib.request.urlretrieve(
                    f"https://docs.google.com/spreadsheets/d/{GSHEET_ID_MAIN}/export?format=xlsx",
                    tmp_xlsx,
                )

                # generate_report.py 실행
                script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generate_report.py")
                result = subprocess.run(
                    ["python3", script_path, tmp_xlsx, report_month],
                    capture_output=True, text=True, timeout=300,
                )

                if result.returncode != 0:
                    st.error(f"리포트 생성 실패: {result.stderr[:500]}")
                else:
                    st.success("리포트 생성 완료!")

                    # ZIP 만들기
                    output_dir = os.path.dirname(os.path.abspath(__file__))
                    month_str = report_month.replace("-", "")
                    common_html = os.path.join(output_dir, f"unniguide_report_{month_str}.html")
                    hospital_dir = os.path.join(output_dir, "hospitals")

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip:
                        zip_path = tmp_zip.name

                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                        if os.path.exists(common_html):
                            zf.write(common_html, f"00_공통_트렌드_리포트_{month_str}.html")
                        if os.path.exists(hospital_dir):
                            for fname in os.listdir(hospital_dir):
                                if fname.endswith(f"_{month_str}.html"):
                                    zf.write(
                                        os.path.join(hospital_dir, fname),
                                        f"병원별/{fname}",
                                    )

                    with open(zip_path, 'rb') as f:
                        zip_bytes = f.read()

                    st.download_button(
                        label=f"📥 언니가이드_리포트_{month_str}.zip 다운로드",
                        data=zip_bytes,
                        file_name=f"언니가이드_리포트_{month_str}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

                    # 실행 로그
                    with st.expander("실행 로그 보기"):
                        st.code(result.stdout)

            except Exception as e:
                st.error(f"오류 발생: {str(e)}")

    st.divider()

    st.subheader("📄 제휴사용 센터 리포트 생성")
    st.caption("아모레퍼시픽 등 외부 제휴 브랜드 공유용 원페이지 HTML 리포트 (A4 PDF 인쇄 최적화)")

    if st.button("🏢 제휴사용 리포트 생성", use_container_width=True):
        import subprocess
        with st.spinner("제휴사용 리포트 생성 중..."):
            try:
                script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generate_partner_report.py")
                result = subprocess.run(
                    ["python3", script_path],
                    capture_output=True, text=True, timeout=120,
                )

                if result.returncode != 0:
                    st.error(f"생성 실패: {result.stderr[:500]}")
                else:
                    output_dir = os.path.dirname(os.path.abspath(__file__))
                    html_path = os.path.join(output_dir, "unniguide_center_report_202603.html")
                    if os.path.exists(html_path):
                        with open(html_path, 'r', encoding='utf-8') as f:
                            html_content = f.read()
                        st.success("제휴사용 리포트 생성 완료!")
                        st.download_button(
                            label="📥 제휴사용_센터_리포트.html 다운로드",
                            data=html_content.encode('utf-8'),
                            file_name="언니가이드_센터_리포트.html",
                            mime="text/html",
                            use_container_width=True,
                        )
                        st.caption("💡 다운로드한 HTML을 브라우저에서 열고 Cmd+P → PDF로 저장하면 제휴사 공유용 PDF가 됩니다.")
                    else:
                        st.error("리포트 파일을 찾을 수 없습니다.")
            except Exception as e:
                st.error(f"오류: {str(e)}")
