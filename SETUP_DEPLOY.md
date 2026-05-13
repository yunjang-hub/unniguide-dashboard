# 언니가이드 대시보드 배포 가이드 (Streamlit Community Cloud)

## 왜 Vercel은 안 됐는가
Vercel은 Python 파일을 서버리스 함수(WSGI/ASGI handler)로 실행한다. Streamlit은 자체 웹서버를 띄우고 WebSocket으로 세션 상태를 유지하는 모델이라 구조적으로 호환되지 않는다. 어떤 설정으로도 동작시킬 수 없으므로 **Vercel 프로젝트는 삭제 권장**한다.

→ Vercel Dashboard → Healingpaper → `unniguide-dashboard` 프로젝트 → Settings → Delete Project

---

## 배포 구조 (확정)
- **호스팅**: Streamlit Community Cloud (`share.streamlit.io`) — 무료, GitHub repo 자동 연동
- **데이터 인증**: GCP 서비스 계정 JSON → Streamlit Secrets
- **슬립 방지**: UptimeRobot 5분 핑

---

## STEP 1. GCP 서비스 계정 발급 (10분, Yun 직접)

1. https://console.cloud.google.com 접속
2. 상단 프로젝트 선택 → **새 프로젝트** 만들기 (이름: `unniguide-dashboard`)
3. 좌측 메뉴 → **API 및 서비스 → 라이브러리** → 다음 2개 검색해서 "사용 설정":
   - Google Sheets API
   - Google Drive API
4. **API 및 서비스 → 사용자 인증 정보 → 사용자 인증 정보 만들기 → 서비스 계정**
   - 이름: `unniguide-reader`
   - 역할: 비워두고 완료 (시트 권한은 시트 쪽에서 부여)
5. 생성된 서비스 계정 클릭 → **키 → 키 추가 → 새 키 만들기 → JSON** → 다운로드
6. 서비스 계정 이메일 복사 (예: `unniguide-reader@unniguide-dashboard.iam.gserviceaccount.com`)

## STEP 2. Google Sheets 권한 부여 (2분, Yun 직접)

대상 시트 2개에 위에서 복사한 서비스 계정 이메일을 **뷰어**로 공유:
- `1pNQiaK67nz6FhxssxgWvoiQr6YwT-5MCwDVvqUZW1SY` (운영 트렌드)
- `16xOwlg8nptwbdM3uvbhr012v77xECiUIgy6wKjT5QYI` (내부 리포트)

이후 "링크 있는 모든 사용자" 권한은 **해제**해도 된다(서비스 계정만으로 동작).

## STEP 3. Streamlit Community Cloud 배포 (5분, Yun 직접)

1. https://share.streamlit.io 접속 → GitHub `yunjang-hub` 계정으로 로그인
2. **New app** 클릭
3. 입력값:
   - Repository: `yunjang-hub/unniguide-dashboard`
   - Branch: `main`
   - Main file path: `py/app.py`
   - App URL (선택): `unniguide-dashboard` 또는 원하는 이름
4. **Advanced settings → Secrets** 에 아래 형식으로 STEP 1에서 받은 JSON을 변환해서 붙여넣기:

```toml
[gcp_service_account]
type = "service_account"
project_id = "unniguide-dashboard"
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...여러 줄...
-----END PRIVATE KEY-----
"""
client_email = "unniguide-reader@unniguide-dashboard.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

> 변환 팁: JSON에서 `\n`이 들어간 `private_key`는 위처럼 triple-quote `"""..."""`로 감싸고 실제 줄바꿈으로 풀어 넣는다.

5. **Deploy** → 2~3분 후 `https://<앱이름>.streamlit.app` 공개 URL 발급

## STEP 4. 슬립 방지 — UptimeRobot (5분, Yun 직접)

Streamlit Community Cloud는 7일 무방문 시 슬립. UptimeRobot 무료 핑으로 살려둔다.

1. https://uptimerobot.com 가입 (무료)
2. **+ New Monitor**
   - Monitor Type: HTTP(s)
   - Friendly Name: `unniguide-dashboard`
   - URL: STEP 3의 Streamlit 앱 URL
   - Monitoring Interval: **5 minutes**
3. Save → 5분마다 GET 요청이 자동으로 가서 슬립 방지

## STEP 5. 사내 접근 제어 (선택)

`@healingpaper.com` 계정만 접근 가능하게 하려면:
- Streamlit Cloud → 앱 Settings → **Sharing** → "Only specific people" → 도메인 또는 이메일 화이트리스트 등록
- 단, 무료 플랜은 **Viewer 비공개 제한이 일부만 지원**됨. 완전 제한이 필요하면 유료(Teams $250/월) 또는 Render/Cloud Run으로 이동 고려

---

## 로컬 개발 (참고)
```bash
mkdir -p ~/.streamlit
cp ~/Downloads/unniguide-dashboard-xxxxxx.json ~/.streamlit/service_account.json
cd ~/Documents/Unniguide/unniguide-report
pip install -r requirements.txt
streamlit run py/app.py
```

코드의 `_get_gcp_authed_session()`이 자동으로 위 경로의 JSON을 잡는다. 파일이 없으면 기존 공개 URL 방식으로 fallback.

---

## 트러블슈팅
- **`PERMISSION_DENIED`**: STEP 2의 시트 공유를 건너뜀. 서비스 계정 이메일을 시트에 뷰어로 추가.
- **`Could not import google.auth`**: requirements.txt에 `google-auth` 누락. Streamlit Cloud는 `requirements.txt`를 자동 설치.
- **앱 슬립**: UptimeRobot 모니터가 멈췄거나 URL 오타. UptimeRobot 대시보드에서 마지막 체크 시각 확인.
