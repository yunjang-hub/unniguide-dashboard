#!/usr/bin/env python3
"""
언니가이드 오프라인 센터 제휴사 공유용 대시보드
- 아모레퍼시픽 등 제휴 브랜드에 공유하는 센터 운영 데이터
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(page_title="언니가이드 센터 리포트", page_icon="🏢", layout="wide")

BRAND_ORANGE = '#FF6A3B'
BRAND_PLUM = '#330C2E'
BRAND_GREEN = '#00B894'
BRAND_IVORY = '#FBF9F1'

st.markdown("""
<style>
    .block-container { padding-top: 1rem; }
    [data-testid="stMetric"] {
        background: white; border: 1px solid #E9ECEF;
        border-radius: 12px; padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    [data-testid="stMetricLabel"] { font-size: 13px !important; }
    [data-testid="stMetricValue"] { font-size: 26px !important; font-weight: 800 !important; }
    h1, h2, h3 { color: #330C2E !important; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 데이터 로딩
# ============================================================
def fix_enc(s):
    if pd.isna(s):
        return s
    try:
        return str(s).encode('latin1').decode('euc-kr')
    except Exception:
        try:
            return str(s).encode('latin1').decode('cp949')
        except Exception:
            return str(s)


@st.cache_data(show_spinner=False)
def load_offline_csv(path):
    df = pd.read_csv(path, encoding='latin1', header=None)

    # 메타 행: row 1=year, row 2=month, row 3=week, row 4=date
    months_row = df.iloc[2].tolist()
    dates_row = df.iloc[4].tolist()

    # 데이터 행 (row 5~) 라벨 디코딩
    labels = {}
    for i in range(5, len(df)):
        raw = df.iloc[i, 0]
        label = fix_enc(raw) if pd.notna(raw) else ''
        labels[i] = str(label).strip()

    # 월별 컬럼 인덱스
    month_cols = {}
    for i, m in enumerate(months_row):
        ms = str(m).strip()
        if ms.isdigit():
            mi = int(ms)
            if mi not in month_cols:
                month_cols[mi] = []
            month_cols[mi].append(i)

    def get_row_sum(row_idx, col_indices):
        total = 0
        for c in col_indices:
            v = df.iloc[row_idx, c]
            if pd.isna(v) or str(v).strip() in ('-', '', 'nan'):
                continue
            try:
                total += float(str(v).replace(',', ''))
            except ValueError:
                pass
        return int(total)

    def get_daily_values(row_idx, col_indices):
        vals = []
        for c in col_indices:
            v = df.iloc[row_idx, c]
            d = str(dates_row[c])[:10] if c < len(dates_row) else ''
            if pd.isna(v) or str(v).strip() in ('-', '', 'nan'):
                vals.append((d, 0))
            else:
                try:
                    vals.append((d, int(float(str(v).replace(',', '')))))
                except ValueError:
                    vals.append((d, 0))
        return vals

    # 행 인덱스 매핑 (라벨로 찾기)
    def find_row(keyword):
        for i, label in labels.items():
            if keyword in label:
                return i
        return None

    return df, labels, month_cols, dates_row, get_row_sum, get_daily_values, find_row


# ============================================================
# 사이드바
# ============================================================
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 12px 0 20px;">
        <span style="font-size:26px; font-weight:800; color:{BRAND_PLUM};">UNNI</span>
        <span style="font-size:26px; font-weight:400; color:{BRAND_PLUM};"> GUIDE</span>
        <br><span style="font-size:13px; color:{BRAND_ORANGE}; font-weight:600;">오프라인 센터 리포트</span>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    pattern = os.path.expanduser("~/Downloads/*Offline*Daily*.csv")
    candidates = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)
    data_loaded = False

    if candidates:
        csv_file = st.selectbox("오프라인 데이터 CSV", candidates, format_func=os.path.basename)
        data_loaded = True
    else:
        uploaded = st.file_uploader("오프라인 CSV 업로드", type=['csv'])
        if uploaded:
            tmp = "/tmp/offline_daily.csv"
            with open(tmp, "wb") as f:
                f.write(uploaded.getvalue())
            csv_file = tmp
            data_loaded = True

if not data_loaded:
    st.markdown("""
    <div style="text-align:center; padding:80px 0;">
        <div style="font-size:48px;">🏢</div>
        <h2>언니가이드 오프라인 센터 리포트</h2>
        <p style="color:#636E72;">사이드바에서 오프라인 데이터 CSV를 선택해주세요.</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

with st.spinner("데이터 로딩 중..."):
    df, labels, month_cols, dates_row, get_row_sum, get_daily_values, find_row = load_offline_csv(csv_file)

# 월 선택
available_months = sorted(month_cols.keys())
with st.sidebar:
    st.divider()
    selected_month = st.selectbox("월 선택", available_months, index=len(available_months) - 1 if available_months else 0, format_func=lambda x: f"2026년 {x}월")

cols = month_cols.get(selected_month, [])
if not cols:
    st.error("선택한 월에 데이터가 없습니다.")
    st.stop()

month_label = f"2026년 {selected_month}월"

# ============================================================
# 주요 행 인덱스
# ============================================================
ROW_예약건수 = find_row('센터방문 예약 건수')
ROW_방문예정 = find_row('센터방문 예정수')
ROW_상담건수 = find_row('센터 상담 건수')
ROW_예약자상담 = find_row('[예약자] 실제 센터 상담수')
ROW_워크인상담 = find_row('[워크인] 실제 센터 상담수')
ROW_병원예약 = find_row('센터 상담 후 병원 예약수')
ROW_노쇼 = find_row('노쇼')
ROW_취소 = find_row('취소')
ROW_미진행 = find_row('상담 미진행')

# 방한/재한 구분
ROW_예약_방한 = None
ROW_예약_재한 = None
for i, label in labels.items():
    if '방한 외국인' in label and i > (ROW_예약건수 or 0) and i < (ROW_방문예정 or 999):
        ROW_예약_방한 = i
    if '재한 외국인' in label and i > (ROW_예약건수 or 0) and i < (ROW_방문예정 or 999):
        ROW_예약_재한 = i

# ============================================================
# 헤더
# ============================================================
st.markdown(f"""
<div style="background: linear-gradient(135deg, {BRAND_ORANGE} 0%, #E8551F 100%);
     color: white; padding: 28px 32px; border-radius: 0 0 20px 20px; margin: -1rem -1rem 24px -1rem;">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <div style="font-size:14px; opacity:0.85; margin-bottom:4px;">UNNI GUIDE 오프라인 센터</div>
            <div style="font-size:24px; font-weight:800;">{month_label} 운영 리포트</div>
        </div>
        <div style="font-size:13px; opacity:0.75;">제휴사 공유용</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# KPI
# ============================================================
st.subheader(f"{month_label} 핵심 지표")

예약건수 = get_row_sum(ROW_예약건수, cols) if ROW_예약건수 else 0
상담건수 = get_row_sum(ROW_상담건수, cols) if ROW_상담건수 else 0
예약자상담 = get_row_sum(ROW_예약자상담, cols) if ROW_예약자상담 else 0
워크인상담 = get_row_sum(ROW_워크인상담, cols) if ROW_워크인상담 else 0
병원예약 = get_row_sum(ROW_병원예약, cols) if ROW_병원예약 else 0
노쇼 = get_row_sum(ROW_노쇼, cols) if ROW_노쇼 else 0
취소 = get_row_sum(ROW_취소, cols) if ROW_취소 else 0
방한 = get_row_sum(ROW_예약_방한, cols) if ROW_예약_방한 else 0
재한 = get_row_sum(ROW_예약_재한, cols) if ROW_예약_재한 else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("센터방문 예약", f"{예약건수:,}건")
c2.metric("실제 상담", f"{상담건수:,}건")
c3.metric("병원 예약 전환", f"{병원예약:,}건")
c4.metric("전환율", f"{병원예약 / max(상담건수, 1) * 100:.1f}%")
c5.metric("일평균 방문예약", f"{예약건수 / len(cols):.1f}건")

st.markdown("")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("예약자 상담", f"{예약자상담:,}건")
c2.metric("워크인 상담", f"{워크인상담:,}건")
c3.metric("방한 외국인", f"{방한:,}명")
c4.metric("재한 외국인", f"{재한:,}명")
visit_rate = 상담건수 / max(예약건수, 1) * 100
c5.metric("예약 대비 상담율", f"{visit_rate:.1f}%")

# ============================================================
# 퍼널 차트
# ============================================================
st.markdown("")
st.subheader("고객 퍼널")
funnel_data = pd.DataFrame({
    '단계': ['센터방문 예약', '실제 상담', '병원 예약 전환'],
    '건수': [예약건수, 상담건수, 병원예약],
})
fig_funnel = go.Figure(go.Funnel(
    y=funnel_data['단계'], x=funnel_data['건수'],
    textinfo="value+percent initial",
    marker=dict(color=[BRAND_ORANGE, BRAND_PLUM, BRAND_GREEN]),
    textfont=dict(size=16),
))
fig_funnel.update_layout(height=300, margin=dict(t=20, b=20, l=20, r=20))
st.plotly_chart(fig_funnel, use_container_width=True)

# ============================================================
# 일별 추이
# ============================================================
st.subheader("일별 추이")

if ROW_예약건수:
    daily_booking = get_daily_values(ROW_예약건수, cols)
    daily_consult = get_daily_values(ROW_상담건수, cols) if ROW_상담건수 else [(d, 0) for d, _ in daily_booking]

    df_daily = pd.DataFrame({
        '날짜': [d for d, _ in daily_booking],
        '센터방문 예약': [v for _, v in daily_booking],
        '실제 상담': [v for _, v in daily_consult],
    })

    fig_daily = go.Figure()
    fig_daily.add_trace(go.Bar(
        x=df_daily['날짜'], y=df_daily['센터방문 예약'],
        name='센터방문 예약', marker_color=BRAND_ORANGE, opacity=0.7,
    ))
    fig_daily.add_trace(go.Scatter(
        x=df_daily['날짜'], y=df_daily['실제 상담'],
        name='실제 상담', line=dict(color=BRAND_PLUM, width=2.5),
        mode='lines+markers',
    ))
    fig_daily.update_layout(
        height=380, margin=dict(t=20, b=80),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
        xaxis=dict(tickangle=-45),
    )
    st.plotly_chart(fig_daily, use_container_width=True)

# ============================================================
# 언어별 분포
# ============================================================
st.subheader("언어별 고객 분포")

# 예약 건수 기준 언어별 (row 7=영어, 8=중국어, 9=태국어 - 예약건수 하위)
lang_rows = {}
for i in range(ROW_예약건수 + 1 if ROW_예약건수 else 6, ROW_예약건수 + 4 if ROW_예약건수 else 9):
    label = labels.get(i, '')
    if '영어' in label:
        lang_rows['영어'] = i
    elif '중국어' in label:
        lang_rows['중국어'] = i
    elif '태국어' in label:
        lang_rows['태국어'] = i

col1, col2 = st.columns(2)

with col1:
    lang_data = []
    for lang, row_i in lang_rows.items():
        cnt = get_row_sum(row_i, cols)
        lang_data.append({'언어': lang, '건수': cnt})
    df_lang = pd.DataFrame(lang_data)
    if len(df_lang) > 0:
        fig_lang = px.pie(df_lang, values='건수', names='언어',
                          color_discrete_sequence=[BRAND_ORANGE, BRAND_PLUM, BRAND_GREEN], hole=0.45)
        fig_lang.update_layout(title="센터 예약 - 언어별 비중", height=380)
        fig_lang.update_traces(textinfo='percent+label+value', textfont_size=14)
        st.plotly_chart(fig_lang, use_container_width=True)

with col2:
    # 방한 vs 재한
    visitor_type = pd.DataFrame({
        '구분': ['방한 외국인', '재한 외국인'],
        '건수': [방한, 재한],
    })
    fig_vt = px.pie(visitor_type, values='건수', names='구분',
                    color_discrete_sequence=[BRAND_PLUM, BRAND_GREEN], hole=0.45)
    fig_vt.update_layout(title="방한 vs 재한 외국인", height=380)
    fig_vt.update_traces(textinfo='percent+label+value', textfont_size=14)
    st.plotly_chart(fig_vt, use_container_width=True)

# ============================================================
# 취소/노쇼
# ============================================================
st.subheader("취소 & No-show")
c1, c2, c3 = st.columns(3)
c1.metric("노쇼", f"{노쇼:,}건", f"{노쇼 / max(예약건수, 1) * 100:.1f}%")
c2.metric("방문 전 취소", f"{취소:,}건", f"{취소 / max(예약건수, 1) * 100:.1f}%")
미진행 = get_row_sum(ROW_미진행, cols) if ROW_미진행 else 0
c3.metric("예약 후 상담 미진행", f"{미진행:,}건")

# ============================================================
# 월별 비교 (가용한 월)
# ============================================================
if len(available_months) > 1:
    st.subheader("월별 비교")
    monthly_summary = []
    for m in available_months:
        mc = month_cols[m]
        row = {
            '월': f"{m}월",
            '센터예약': get_row_sum(ROW_예약건수, mc) if ROW_예약건수 else 0,
            '실제상담': get_row_sum(ROW_상담건수, mc) if ROW_상담건수 else 0,
            '병원예약': get_row_sum(ROW_병원예약, mc) if ROW_병원예약 else 0,
            '노쇼': get_row_sum(ROW_노쇼, mc) if ROW_노쇼 else 0,
            '취소': get_row_sum(ROW_취소, mc) if ROW_취소 else 0,
        }
        row['전환율'] = f"{row['병원예약'] / max(row['실제상담'], 1) * 100:.1f}%"
        monthly_summary.append(row)

    df_ms = pd.DataFrame(monthly_summary)

    fig_ms = go.Figure()
    fig_ms.add_trace(go.Bar(x=df_ms['월'], y=df_ms['센터예약'], name='센터예약', marker_color=BRAND_ORANGE, opacity=0.8, text=df_ms['센터예약'], textposition='outside'))
    fig_ms.add_trace(go.Bar(x=df_ms['월'], y=df_ms['실제상담'], name='실제상담', marker_color=BRAND_PLUM, opacity=0.8))
    fig_ms.add_trace(go.Bar(x=df_ms['월'], y=df_ms['병원예약'], name='병원예약', marker_color=BRAND_GREEN, opacity=0.8))
    fig_ms.update_layout(
        barmode='group', height=380, margin=dict(t=20, b=80),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
    )
    st.plotly_chart(fig_ms, use_container_width=True)

    st.dataframe(df_ms, use_container_width=True, hide_index=True)

# ============================================================
# 푸터
# ============================================================
st.markdown(f"""
<div style="text-align:center; padding:32px 0; color:#636E72; font-size:13px;
     border-top:1px solid #E9ECEF; margin-top:40px;">
    <p><strong>UNNI GUIDE</strong> | 강남언니 언니가이드 오프라인 센터</p>
    <p style="margin-top:4px;">본 리포트는 제휴 브랜드 전용 자료입니다.</p>
</div>
""", unsafe_allow_html=True)
