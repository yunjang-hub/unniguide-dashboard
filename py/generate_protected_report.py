#!/usr/bin/env python3
"""
제휴사별 비밀번호 보호 리포트 생성기

사용법:
  python3 generate_protected_report.py "아모레퍼시픽" AMORE2026
  → aparterned/amorepacific_202603.html 생성 (비번 AMORE2026)

  python3 generate_protected_report.py "파마리서치" PHARMA2026
  → partner/pharmaresearch_202603.html 생성 (비번 PHARMA2026)
"""
import sys
import os
import hashlib
import re
from datetime import datetime

if len(sys.argv) < 3:
    print("사용법: python3 generate_protected_report.py <제휴사명> <비밀번호> [원본HTML]")
    print("예시: python3 generate_protected_report.py '아모레퍼시픽' AMORE2026")
    sys.exit(1)

partner_name = sys.argv[1]
password = sys.argv[2]
source_html = sys.argv[3] if len(sys.argv) > 3 else os.path.expanduser(
    '~/Documents/Unniguide/unniguide-report/unniguide_center_report_web_202603.html'
)

OUTPUT_DIR = os.path.expanduser('~/Documents/Unniguide/unniguide-report/partner/')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 비밀번호 해시
pw_hash = hashlib.sha256(password.encode()).hexdigest()

# 파일명 안전화
def safe_name(name):
    return re.sub(r'[^\w가-힣]', '_', name).lower()

partner_slug = safe_name(partner_name)
output_file = os.path.join(OUTPUT_DIR, f'{partner_slug}_202603.html')

# 원본 HTML 읽기
with open(source_html, 'r', encoding='utf-8') as f:
    html = f.read()

# 비밀번호 게이트 CSS + HTML + JS 삽입
gate_css = """
<style>
#pw-gate {
  position: fixed; top: 0; left: 0; right: 0; bottom: 0;
  background: linear-gradient(135deg, #330C2E 0%, #5C2D54 100%);
  display: flex; align-items: center; justify-content: center;
  z-index: 99999; font-family: -apple-system,BlinkMacSystemFont,'Segoe UI','Noto Sans KR',sans-serif;
}
#pw-gate-box {
  background: white; border-radius: 20px; padding: 48px 40px;
  max-width: 440px; width: 90%; text-align: center;
  box-shadow: 0 20px 60px rgba(0,0,0,0.3);
}
#pw-gate-box .logo { font-size: 14px; font-weight: 800; letter-spacing: 2px; color: #FF6A3B; margin-bottom: 12px; }
#pw-gate-box h1 { font-size: 22px; font-weight: 800; color: #330C2E; margin-bottom: 8px; }
#pw-gate-box .subtitle { font-size: 14px; color: #636E72; margin-bottom: 8px; }
#pw-gate-box .partner-label { font-size: 13px; color: #FF6A3B; font-weight: 700; margin-bottom: 28px; padding: 6px 14px; background: #FFF0EB; border-radius: 20px; display: inline-block; }
#pw-gate-box .lock { font-size: 40px; margin-bottom: 16px; }
#pw-gate-box input {
  width: 100%; padding: 14px 18px; font-size: 16px;
  border: 2px solid #E9ECEF; border-radius: 12px;
  outline: none; transition: border 0.15s; margin-bottom: 12px;
  text-align: center; letter-spacing: 2px;
}
#pw-gate-box input:focus { border-color: #FF6A3B; }
#pw-gate-box button {
  width: 100%; padding: 14px; font-size: 15px; font-weight: 700;
  background: #FF6A3B; color: white; border: none; border-radius: 12px;
  cursor: pointer; transition: background 0.15s;
}
#pw-gate-box button:hover { background: #E8551F; }
#pw-gate-box .error { color: #E74C3C; font-size: 13px; margin-top: 10px; min-height: 18px; }
#pw-gate-box .hint { font-size: 12px; color: #999; margin-top: 20px; }
body.pw-locked { overflow: hidden; }
body.pw-locked > *:not(#pw-gate) { filter: blur(8px); pointer-events: none; user-select: none; }
</style>
"""

gate_html = f"""
<div id="pw-gate">
  <div id="pw-gate-box">
    <div class="lock">🔐</div>
    <div class="logo">UNNI GUIDE</div>
    <h1>제휴사 전용 리포트</h1>
    <div class="subtitle">본 리포트는 비밀번호 보호된 자료입니다.</div>
    <div class="partner-label">For. {partner_name}</div>
    <input type="password" id="pw-input" placeholder="비밀번호 입력" autofocus />
    <button onclick="checkPassword()">열람하기</button>
    <div class="error" id="pw-error"></div>
    <div class="hint">비밀번호는 담당자에게 문의해 주세요.</div>
  </div>
</div>
<script>
async function sha256(str) {{
  const buf = new TextEncoder().encode(str);
  const hash = await crypto.subtle.digest('SHA-256', buf);
  return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
}}
async function checkPassword() {{
  const input = document.getElementById('pw-input').value;
  const expected = '{pw_hash}';
  const hash = await sha256(input);
  if (hash === expected) {{
    document.getElementById('pw-gate').style.display = 'none';
    document.body.classList.remove('pw-locked');
    sessionStorage.setItem('ug_unlock_{partner_slug}', '1');
  }} else {{
    document.getElementById('pw-error').textContent = '비밀번호가 올바르지 않습니다.';
    document.getElementById('pw-input').value = '';
  }}
}}
document.getElementById('pw-input').addEventListener('keypress', e => {{
  if (e.key === 'Enter') checkPassword();
}});
// 세션 내 재접속 시 통과
if (sessionStorage.getItem('ug_unlock_{partner_slug}') === '1') {{
  document.getElementById('pw-gate').style.display = 'none';
  document.body.classList.remove('pw-locked');
}} else {{
  document.body.classList.add('pw-locked');
}}
</script>
"""

# </head> 직전에 CSS 삽입, <body> 직후에 gate HTML 삽입
html = html.replace('</head>', gate_css + '</head>')
html = html.replace('<body>', '<body>' + gate_html, 1)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"✅ 제휴사 비밀번호 리포트 생성 완료")
print(f"   파일: {output_file}")
print(f"   제휴사: {partner_name}")
print(f"   비밀번호: {password}")
print(f"   배포 URL (GitHub Pages 업로드 후): https://yunjang-hub.github.io/unniguide-dashboard/partner/{partner_slug}_202603.html")
