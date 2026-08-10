"""DART 중계 회귀 테스트.

가짜 DART 서버와 중계 함수를 실제로 띄워 요청을 끝까지 태운다.
주소 변환만 확인하면 ZIP 깨짐이나 오류코드 유실을 놓치기 때문이다.
"""

import importlib
import io
import sys
import threading
import zipfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse

import requests

import dart_relay as R

sys.path.insert(0, 'vercel_relay/api')
import relay as RELAY

PASS, FAIL = [], []


def check(name, cond, detail=''):
    (PASS if cond else FAIL).append(name)
    print(f"  [{'OK ' if cond else 'FAIL'}] {name} {detail}")


def serve(handler_cls, host='127.0.0.1'):
    """빈 포트에 서버를 띄우고 (주소, 종료함수) 반환.

    가짜 DART와 중계를 같은 호스트명으로 띄우면 중계가 자기 자신을 DART로
    오인해 무한 루프에 빠진다. 이름을 다르게 준다 (둘 다 같은 기기를 가리킨다).
    """
    srv = HTTPServer(('127.0.0.1', 0), handler_cls)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return f'http://{host}:{srv.server_port}', srv.shutdown


# ── 가짜 DART ──────────────────────────────────────────────────────
ZIP_BODY = io.BytesIO()
with zipfile.ZipFile(ZIP_BODY, 'w') as z:
    z.writestr('CORPCODE.xml', '<result><list><corp_code>00126380</corp_code></list></result>')
ZIP_BYTES = ZIP_BODY.getvalue()

received = []


class FakeDart(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urlparse(self.path)
        received.append(self.path)
        q = parse_qs(u.query)

        if u.path == '/api/corpCode.xml':
            if q.get('crtfc_key', [''])[0] == 'good':
                body, ctype, code = ZIP_BYTES, 'application/zip', 200
            else:
                body = b'<result><status>010</status><message>\xeb\x93\xb1\xeb\xa1\x9d\xec\x95\x88\xeb\x90\xa8</message></result>'
                ctype, code = 'application/xml', 200
        elif u.path == '/api/fail.json':
            body, ctype, code = b'{"status":"800","message":"system error"}', 'application/json', 500
        elif u.path == '/dsaf001/main.do':
            body, ctype, code = b'<html>viewer</html>', 'text/html', 200
        else:
            body, ctype, code = b'{"status":"000","list":[]}', 'application/json', 200

        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)


dart_url, stop_dart = serve(FakeDart, host='localhost')
DART_HOST = dart_url.replace('http://', '')

# 가짜 DART를 허용 목록에 넣고, 중계가 그쪽을 보도록 한다
RELAY.ALLOWED_HOSTS = {DART_HOST.split(':')[0]}
relay_url, stop_relay = serve(RELAY.handler, host='127.0.0.1')

print("=== 1. 허용 호스트 검사 (공개 프록시 방지) ===")
ok, why = RELAY._check('https://evil.example.com/steal')
check('허용 안 된 호스트 거부', not ok, f'→ {why}')
ok, why = RELAY._check('file:///etc/passwd')
check('file:// 거부', not ok, f'→ {why}')
ok, why = RELAY._check(None)
check('빈 요청 거부', not ok, f'→ {why}')
RELAY.ALLOWED_HOSTS = {'opendart.fss.or.kr', 'dart.fss.or.kr'}
ok, _ = RELAY._check('https://opendart.fss.or.kr/api/x.json')
check('opendart 허용', ok)
ok, _ = RELAY._check('http://dart.fss.or.kr/dsaf001/main.do?rcpNo=1')
check('dart(원문 뷰어) 허용', ok)
RELAY.ALLOWED_HOSTS = {DART_HOST.split(':')[0]}

print("\n=== 2. 중계 왕복 — 응답이 그대로 오는가 ===")
target = f'{dart_url}/api/corpCode.xml?crtfc_key=preflight'
r = requests.get(R.build_relay_url(relay_url, target), timeout=10)
check('상태코드 200', r.status_code == 200)
check('DART 응답 본문 전달', b'<status>010</status>' in r.content, f'→ {r.content[:40]}')
check('Content-Type 유지', 'xml' in r.headers.get('Content-Type', ''),
      f"→ {r.headers.get('Content-Type')}")

print("\n=== 3. ZIP(회사목록)이 깨지지 않는가 ===")
r = requests.get(R.build_relay_url(relay_url, f'{dart_url}/api/corpCode.xml?crtfc_key=good'),
                 timeout=10)
check('ZIP 바이트 동일', r.content == ZIP_BYTES, f'({len(r.content)}바이트)')
try:
    z = zipfile.ZipFile(io.BytesIO(r.content))
    check('ZIP 열림', '00126380' in z.read('CORPCODE.xml').decode())
except Exception as e:
    check('ZIP 열림', False, f'→ {e}')

print("\n=== 4. DART 오류를 삼키지 않는가 ===")
r = requests.get(R.build_relay_url(relay_url, f'{dart_url}/api/fail.json'), timeout=10)
check('오류 상태코드 전달', r.status_code == 500, f'→ {r.status_code}')
check('오류 본문 전달', b'system error' in r.content)

print("\n=== 5. 공시 원문 뷰어(dart.fss.or.kr)도 중계되는가 ===")
r = requests.get(R.build_relay_url(relay_url, f'{dart_url}/dsaf001/main.do?rcpNo=123'), timeout=10)
check('뷰어 응답 전달', b'viewer' in r.content)

print("\n=== 6. 토큰 검사 ===")
import os
os.environ['RELAY_TOKEN'] = 'secret1'
r = requests.get(R.build_relay_url(relay_url, target), timeout=10)
check('토큰 없으면 401', r.status_code == 401, f'→ {r.status_code}')
r = requests.get(R.build_relay_url(relay_url, target),
                 headers={'X-Relay-Token': 'wrong'}, timeout=10)
check('틀린 토큰 401', r.status_code == 401)
r = requests.get(R.build_relay_url(relay_url, target),
                 headers={'X-Relay-Token': 'secret1'}, timeout=10)
check('맞는 토큰 통과', r.status_code == 200)
del os.environ['RELAY_TOKEN']
r = requests.get(R.build_relay_url(relay_url, target), timeout=10)
check('토큰 미설정 시 검사 안 함', r.status_code == 200)

print("\n=== 7. requests.get 가로채기 ===")
R = importlib.reload(R)          # _installed 초기화
R.DART_HOSTS = (DART_HOST.split(':')[0],)
before = requests.get
check('설치 성공', R.install(relay_url) is True)
check('두 번째 설치는 무시', R.install(relay_url) is False)
check('requests.get 교체됨', requests.get is not before)

received.clear()
# 라이브러리가 하는 것과 같은 형태로 호출 (params 분리)
resp = requests.get(f'{dart_url}/api/test.json', params={'crtfc_key': 'k', 'corp_code': '00126380'})
check('중계 경유 성공', resp.status_code == 200)
check('질의문자열 보존', received and 'crtfc_key=k' in received[0] and 'corp_code=00126380' in received[0],
      f'→ {received[0] if received else "없음"}')

received.clear()
requests.get(f'{relay_url}/', timeout=5)   # DART 호스트가 아님 → 중계 안 함
check('DART 외 호스트는 우회 안 함', not received)

# 앱은 주식수 조회에 Session 을 쓴다. 여기가 안 감싸이면 그 요청만 중계를 우회한다.
received.clear()
sess = requests.Session()
resp = sess.get(f'{dart_url}/api/stockTotqySttus.json', params={'crtfc_key': 'k'})
check('Session.get 도 중계 경유', resp.status_code == 200 and received
      and 'crtfc_key=k' in received[0], f'→ {received[0] if received else "없음"}')
received.clear()
sess.get(f'{relay_url}/', timeout=5)
check('Session 도 DART 외에는 우회 안 함', not received)

print("\n=== 8. 설정 읽기 ===")
os.environ['DART_RELAY_URL'] = 'https://x.vercel.app/'
os.environ['DART_RELAY_TOKEN'] = 'tok'
u, t = R.relay_config()
check('환경변수에서 읽음', u == 'https://x.vercel.app' and t == 'tok', f'→ {u}, {t}')
check('끝 슬래시 제거', not u.endswith('/'))


class FakeSecrets(dict):
    pass


u, t = R.relay_config(FakeSecrets({'DART_RELAY_URL': 'https://s.vercel.app'}))
check('secrets 우선', u == 'https://s.vercel.app')
del os.environ['DART_RELAY_URL'], os.environ['DART_RELAY_TOKEN']
u, t = R.relay_config()
check('설정 없으면 None', u is None and t is None)

print("\n=== 9. 중계 상태 점검(probe) 문구 ===")
ok, msg = R.probe(None)
check('주소 없으면 실패', not ok and '설정' in msg)
ok, msg = R.probe('http://127.0.0.1:1', timeout=3)
check('죽은 주소 → 연결 안내', not ok and '연결되지 않' in msg, f'→ {msg}')

requests.get = before
stop_relay()
stop_dart()

print("\n" + "=" * 56)
print(f"통과 {len(PASS)}건 / 실패 {len(FAIL)}건")
if FAIL:
    for f in FAIL:
        print(f"  - {f}")
    sys.exit(1)
print("전부 통과")
