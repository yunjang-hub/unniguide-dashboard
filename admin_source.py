"""언니가이드 어드민 case-metrics CSV → 운영시트 스키마 변환.

2026년 7월부터 운영을 구글시트에서 어드민으로 이관하면서 원천이 둘로 갈렸다.
  ~ 2026-06 내원 : 운영시트(예약확정 리스트)가 정확  (어드민 커버율 4월 14% / 5월 87% / 6월 94%)
    2026-07 내원 ~ : 어드민이 정확              (운영시트는 2026-07-07에서 입력 중단)

이 모듈은 어드민 export(case-metrics-*.csv)를 운영시트와 동일한 컬럼 이름/값 체계로
번역만 담당한다. 병합 컷오프와 집계는 app.py가 처리한다.
"""

import pandas as pd

# 어드민 export가 first_consult_date 기준으로 잘려 나오기 때문에,
# 컷오프 월보다 최소 3개월 앞선 상담분까지 포함된 파일을 받아야 내원일 기준 집계가 온전해진다.
ADMIN_CUTOFF_MONTH = '2026-07'

# 어드민 raw를 담는 구글시트 탭 이름(키워드 매칭). 고객 실명·금액이 들어 있어
# 공개 GitHub repo에는 커밋하지 않고, 기존 운영시트와 같은 접근통제 아래 둔다.
ADMIN_SHEET_KEYWORD = '어드민 케이스'


# ------------------------------------------------------------------
# 값 매핑
# ------------------------------------------------------------------
# 어드민 workflow_status → 운영시트 '예약 상태'
# app.py가 '시/수술 완료'(완료 판정) / '예약 취소'(취소) / no-show 문자열(노쇼)에 의존한다.
STATUS_MAP = {
    'CLOSED': '시/수술 완료',
    'CANCELLED': '예약 취소',
    'NO_SHOW': 'No-show',
    'BOOKED': '예약 확정',
    'SURGERY_DATE_CONFIRMED': '수술일 확정',
    'RECOMMENDED': '병원 추천',
    'NOT_PROGRESSED': '미진행',
    'CONSULT_ONLY': '상담만',
}

TREATMENT_TYPE_MAP = {'PROCEDURE': '시술', 'SURGERY': '수술'}

LANGUAGE_MAP = {
    'en': '영어', 'mandarin': '중국어', 'th': '태국어',
    'ru': '러시아어', 'ja': '일본어', 'ko': '한국어',
}

# 어드민은 ISO 3166-1 alpha-2, 운영시트는 한글 국가명을 쓴다.
# 한글 표기는 운영시트에 이미 쓰인 표기를 그대로 따랐다(예: '투니지아', '로마니아', '사우디 아라비아').
ISO2_TO_KR = {
    'TH': '태국', 'TW': '대만', 'US': '미국', 'AU': '호주', 'HK': '홍콩',
    'CN': '중국', 'SG': '싱가포르', 'FR': '프랑스', 'CA': '캐나다', 'MN': '몽골',
    'MY': '말레이시아', 'PH': '필리핀', 'DE': '독일', 'GB': '영국', 'IT': '이탈리아',
    'IN': '인도', 'ID': '인도네시아', 'ES': '스페인', 'NL': '네덜란드', 'IL': '이스라엘',
    'RU': '러시아', 'MX': '멕시코', 'BR': '브라질', 'SE': '스웨덴', 'JP': '일본',
    'NZ': '뉴질랜드', 'PL': '폴란드', 'CH': '스위스', 'TR': '튀르키예', 'VN': '베트남',
    'SA': '사우디 아라비아', 'DK': '덴마크', 'PT': '포르투갈', 'CO': '콜롬비아',
    'AT': '오스트리아', 'CL': '칠레', 'NO': '노르웨이', 'PE': '페루', 'LT': '리투아니아',
    'CZ': '체코', 'SK': '슬로바키아', 'KG': '키르기스스탄', 'OM': '오만', 'EC': '에콰도르',
    'FI': '핀란드', 'BE': '벨기에', 'UA': '우크라이나', 'TN': '투니지아', 'BN': '브루나이',
    'ZA': '남아프리카', 'PR': '푸에르토리코', 'AE': '아랍 에미레이트', 'LA': '라오스',
    'MA': '모로코', 'IE': '아일랜드', 'PK': '파키스탄', 'HU': '헝가리', 'RO': '로마니아',
    'NG': '나이지리아', 'EG': '이집트', 'LV': '라트비아', 'VE': '베네수엘라',
    'MM': '미얀마', 'HR': '크로아티아', 'AR': '아르헨티나', 'KZ': '카자흐스탄',
    'GR': '그리스', 'QA': '카타르', 'CR': '코스타리카', 'PA': '파나마', 'KH': '캄보디아',
    'GU': '괌', 'EE': '에스토니아', 'GT': '과테말라', 'PY': '파라과이',
    'TL': '동티모르', 'XK': '코소보', 'NP': '네팔', 'IO': '영국령 인도양 지역',
    'TJ': '타지키스탄', 'BT': '부탄', 'BG': '불가리아', 'AF': '아프가니스탄',
    'KE': '케냐', 'SD': '수단', 'DO': '도미니카 공화국', 'HN': '온두라스',
    'LB': '레바논', 'RE': '레위니옹', 'MV': '몰디브', 'KR': '한국', 'MO': '마카오',
    'GE': '조지아', 'IQ': '이라크', 'BS': '바하마', 'BY': '벨라루스',
    'UZ': '우즈베키스탄', 'LK': '스리랑카', 'BD': '방글라데시', 'JO': '요르단',
    'KW': '쿠웨이트', 'BH': '바레인', 'CY': '키프로스', 'SI': '슬로베니아',
    'RS': '세르비아', 'IS': '아이슬란드', 'LU': '룩셈부르크', 'MT': '몰타',
    'UY': '우루과이', 'BO': '볼리비아', 'GH': '가나', 'ET': '에티오피아',
    'TZ': '탄자니아', 'MU': '모리셔스', 'FJ': '피지', 'PG': '파푸아뉴기니',
}

# 어드민 병원 표기 → 운영시트 표기. app.py의 NAME_NORMALIZE에 합쳐서 쓴다.
ADMIN_HOSPITAL_ALIASES = {
    'BLS(비엘에스)의원-명동점': 'BLS의원 명동',
    'TU치과의원(티유치과)': '티유치과의원',
    '디에이성형외과': '디에이성형외과의원',
    '우아성형외과': '우아성형외과의원',
    '유픽의원 - 강남': '유픽의원-강남점',
    '유픽의원 - 홍대': '유픽의원-홍대점',
    '제이필의원 강남': '제이필의원-강남점',
    '제이필의원 홍대점': '제이필의원-홍대점',
    '테이아의원- 강남본점': '테이아의원',
    '테이아의원- 명동글로벌': '테이아의원 명동점',
    '톡스앤필신논현점': '톡스앤필의원-신논현점',
    '플레저성형외과': '플레저성형외과의원',
    '허쉬성형외과': '허쉬성형외과의원',
}

# 운영시트 안에서도 같은 병원이 두 표기로 들어가 있어 함께 정리한다.
SHEET_HOSPITAL_ALIASES = {
    '릴리브 의원': '릴리브의원',
    '톡스앤필-홍대점': '톡스앤필의원-홍대점',
    '티유 치과의원': '티유치과의원',
}


def _clean_dates(s):
    """날짜 파싱 + 연도 오타 보정.

    어드민에 booked_date='2006-06-01'처럼 연도를 두 자리 잘못 입력한 건이 있다.
    2025년 이전 날짜는 실운영상 존재할 수 없으므로 +20년으로 되돌린다.
    """
    d = pd.to_datetime(s, errors='coerce')
    bad = d.notna() & (d.dt.year < 2025)
    if bad.any():
        d.loc[bad] = d.loc[bad] + pd.DateOffset(years=20)
    return d


def load_admin_case_metrics(path):
    """어드민 case-metrics CSV 파일 → 운영시트 스키마 DataFrame."""
    return normalize_admin_frame(pd.read_csv(path, dtype=str))


def normalize_admin_frame(df):
    """어드민 case-metrics 원본(CSV/시트탭 무관) → 운영시트 스키마 DataFrame.

    반환 컬럼은 app.py의 load_operation_excel이 만드는 것과 같은 표준명을 쓴다.
    (병원명 정규화 / '월' 파생은 app.py가 공통 처리)
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna('')

    required = ['workflow_status', 'hospital_name', 'booked_date',
                'actual_amount', 'treatment_type', 'consulter_nationality']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"어드민 CSV에서 필수 컬럼을 찾을 수 없습니다: {missing}")

    def _s(col):
        return df[col].astype(str).str.strip() if col in df.columns else pd.Series([''] * len(df))

    out = pd.DataFrame(index=df.index)

    # 국적: 상담자 기준, 비어있으면 시술 대상자(candidate)로 보완
    nat = _s('consulter_nationality').replace('', pd.NA).fillna(_s('candidate_nationality'))
    out['고객국적'] = nat.map(lambda c: ISO2_TO_KR.get(c, c) if c else '')

    # 고객명: 상담자 기준, 비어있으면 시술 대상자
    out['고객명'] = _s('consulter_name').replace('', pd.NA).fillna(_s('candidate_name')).fillna('')

    out['예약상태'] = _s('workflow_status').map(lambda v: STATUS_MAP.get(v, '상담만' if v == '' else v))
    out['종류'] = _s('treatment_type').map(lambda v: TREATMENT_TYPE_MAP.get(v, ''))
    out['사용언어'] = _s('language').map(lambda v: LANGUAGE_MAP.get(v, v))
    out['시술수술명'] = _s('treatments')
    out['예약클리닉'] = _s('hospital_name')
    out['추천클리닉'] = _s('hospital_name')
    out['상담사'] = _s('consultant_name')
    out['시간'] = _s('booked_time')

    out['내원일'] = _clean_dates(_s('booked_date'))
    out['채팅접수일자'] = _clean_dates(_s('first_consult_date'))
    # 어드민에는 '예약 확정일'이 없다. app.py가 내원일이 빈 행의 집계월을 예약확정일로
    # 보완하므로, 취소/미내원 건이 통째로 빠지지 않도록 최초 상담일을 대체값으로 넣는다.
    out['예약확정일'] = out['채팅접수일자']

    out['실제금액'] = pd.to_numeric(_s('actual_amount'), errors='coerce').fillna(0)
    out['예상금액'] = pd.to_numeric(_s('estimated_amount'), errors='coerce').fillna(0)

    # 운영시트의 '그룹 여부'(예: '2명') 재현
    gs = pd.to_numeric(_s('group_size'), errors='coerce')
    out['그룹 여부'] = gs.map(lambda n: f"{int(n)}명" if pd.notna(n) and n > 1 else '')

    # 어드민 고유 필드 — 동반자/채널 구분은 운영시트에 없던 정보라 살려둔다
    out['row_kind'] = _s('row_kind')
    out['channel'] = _s('channel')
    out['hospital_region'] = _s('hospital_region')
    out['통역 서비스 요청'] = _s('interpreter_needed').map({'Y': '통역 요청', 'N': ''}).fillna('')
    out['_source'] = '어드민'

    return out.reset_index(drop=True)
