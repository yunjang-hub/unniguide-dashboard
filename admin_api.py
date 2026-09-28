"""언니가이드 어드민 Public Read API → case-metrics 스키마 DataFrame (메모리 전용).

어드민 CSV export를 손으로 받던 절차를 대체한다. 결과는 파일로 쓰지 않고
app.py가 st.cache_data로 들고 있다가 TTL이 지나면 다시 부른다(= 매일 자동 갱신).

  - export는 first_consult_date 4개월 창으로 잘려 나와 리드타임 긴 건이 빠졌다.
    /cases는 날짜 파라미터를 생략하면 전 기간 전량을 주므로 그 함정이 없다.
  - 키는 Streamlit secrets(UG_PUBLIC_API_KEY)에만 둔다. 이 repo는 공개다.

API와 export의 차이:
  - 행 단위: API는 케이스×병원 1행 + companion_count. export는 동반자도 1행씩이었다.
    대시보드는 행 수 = 인원으로 세므로 동반자 수만큼 COMPANION 행(금액 0)을 덧붙인다.
    API actual_amount는 동반자 몫을 이미 포함한다(2026-07 export 대조로 확인).
  - treatments(시술명)·language·hospital_region: API에 없다 → 공란.
    시술명은 개발팀 필드 추가 대기. 지역은 app.py HOSPITAL_GU가 병원명으로 판별한다.
"""

import datetime as dt
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

BASE = 'https://admin.unniguide.com/api/public/v1'
PAGE_SIZE = 200
KST = dt.timezone(dt.timedelta(hours=9))  # 조회 시각 표시용

# admin_source.normalize_admin_frame이 읽는 case-metrics 컬럼
COLUMNS = [
    'row_kind', 'first_consult_date', 'channel', 'language', 'hospital_name',
    'hospital_region', 'treatment_type', 'workflow_status', 'reservation_bucket',
    'booked_date', 'booked_time', 'estimated_amount', 'actual_amount', 'paid',
    'consulter_nationality', 'candidate_nationality', 'group_size', 'consulter_form',
    'treatments', 'interpreter_needed',
]


class AdminApiError(RuntimeError):
    pass


def _call(key, path, params, retries=8):
    url = f"{BASE}{path}?{urllib.parse.urlencode(params)}"
    for attempt in range(retries):
        req = urllib.request.Request(url, headers={'Authorization': f'Bearer {key}'})
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                return json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                wait = e.headers.get('Retry-After')
                time.sleep(float(wait) if wait and wait.replace('.', '', 1).isdigit()
                           else min(2 ** attempt, 30))
                continue
            # 응답 본문에 고객 데이터가 섞일 수 있어 상태코드만 올린다
            hint = ' (키 만료·스코프 누락 확인: read:cases)' if e.code in (401, 403, 404) else ''
            raise AdminApiError(f'{path} HTTP {e.code}{hint}') from None
        except urllib.error.URLError:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            raise AdminApiError(f'{path} 네트워크 오류') from None


def fetch_cases(key, workers=3):
    """/cases 전량(include=hospitals). 첫 페이지로 total_count를 알고 나머지는 병렬로."""
    params = {'include': 'hospitals', 'page_size': PAGE_SIZE}
    first = _call(key, '/cases', {**params, 'page': 0})
    total = first['meta']['total_count']
    pages = range(1, -(-total // PAGE_SIZE))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        rest = list(ex.map(lambda p: _call(key, '/cases', {**params, 'page': p})['data'], pages))
    cases = first['data'] + [c for chunk in rest for c in chunk]
    # 페이징 도중 신규 케이스가 생기면 경계에서 중복될 수 있다
    uniq = {c['id']: c for c in cases}
    return list(uniq.values())


def _kst_date(ts):
    if not ts:
        return ''
    # 소수점 자릿수가 들쭉날쭉(5자리 등)해 3.10 이하 fromisoformat이 못 읽는다 → pandas로 파싱
    t = pd.to_datetime(ts, utc=True, errors='coerce')
    return '' if pd.isna(t) else t.tz_convert('Asia/Seoul').date().isoformat()


def cases_to_frame(cases):
    """케이스 JSON → case-metrics 스키마. 개인식별정보는 옮기지 않는다."""
    out = []
    for c in cases:
        cust = c.get('customer') or {}
        base = dict.fromkeys(COLUMNS, '')
        base.update({
            'first_consult_date': _kst_date(c.get('chat_started_at') or c.get('created_at')),
            'channel': c.get('channel') or '',
            'consulter_nationality': cust.get('nationality') or '',
            'interpreter_needed': 'Y' if c.get('needs_interpreter') else 'N',
        })
        hospitals = c.get('hospitals') or []
        if not hospitals:
            out.append({**base, 'row_kind': 'CONSULTATION_ONLY', 'group_size': '1'})
            continue
        for h in hospitals:
            comp = int(h.get('companion_count') or 0)
            row = {**base,
                   'row_kind': 'PRIMARY',
                   'hospital_name': h.get('hospital_name') or '',
                   'treatment_type': h.get('treatment_type') or '',
                   'workflow_status': h.get('workflow_status') or '',
                   'booked_date': h.get('visit_date') or '',
                   'booked_time': h.get('visit_time') or '',
                   'estimated_amount': '' if h.get('estimated_amount') is None else str(h['estimated_amount']),
                   'actual_amount': '' if h.get('actual_amount') is None else str(h['actual_amount']),
                   'group_size': str(comp + 1)}
            out.append(row)
            for _ in range(comp):
                out.append({**row, 'row_kind': 'COMPANION',
                            'estimated_amount': '0', 'actual_amount': '0'})
    return pd.DataFrame(out, columns=COLUMNS).astype(str)


def load_admin_api_frame(key):
    """API → case-metrics 스키마 DataFrame + 조회 시각(KST)."""
    df = cases_to_frame(fetch_cases(key))
    return df, dt.datetime.now(KST).strftime('%Y-%m-%d %H:%M')
