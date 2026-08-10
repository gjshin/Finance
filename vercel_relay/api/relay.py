"""DART 중계 함수 (Vercel 서울 리전).

해외 서버에서 DART로 직접 가는 요청을 대신 받아, 서울에서 DART로 보내고
응답을 그대로 돌려준다.

    Streamlit Cloud (해외) → 이 함수 (서울) → DART

의존성 없이 표준 라이브러리만 쓴다. 설치할 것이 없어 배포가 단순하고 빠르다.
"""

import gzip
import io
import os
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler

# 아무 주소로나 중계하면 공개 프록시가 된다. DART 두 호스트만 허용한다.
ALLOWED_HOSTS = {'opendart.fss.or.kr', 'dart.fss.or.kr'}

UPSTREAM_TIMEOUT = 25
USER_AGENT = 'Mozilla/5.0 (compatible; dart-relay/1.0)'


def _check(target):
    """중계해도 되는 주소인지. (허용여부, 사유)"""
    if not target:
        return False, 'u 파라미터가 없습니다.'
    try:
        p = urllib.parse.urlparse(target)
    except Exception:
        return False, '주소를 해석할 수 없습니다.'
    if p.scheme not in ('http', 'https'):
        return False, 'http/https 만 중계합니다.'
    if p.hostname not in ALLOWED_HOSTS:
        return False, f'허용되지 않은 주소입니다: {p.hostname}'
    return True, None


class handler(BaseHTTPRequestHandler):

    def _send(self, code, body, ctype='text/plain; charset=utf-8'):
        data = body if isinstance(body, bytes) else str(body).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        # 토큰이 설정돼 있으면 요구한다 (설정 안 하면 검사하지 않는다)
        expected = (os.environ.get('RELAY_TOKEN') or '').strip()
        if expected and self.headers.get('X-Relay-Token', '') != expected:
            return self._send(401, '중계 토큰이 올바르지 않습니다.')

        query = urllib.parse.urlparse(self.path).query
        target = (urllib.parse.parse_qs(query).get('u') or [None])[0]

        ok, why = _check(target)
        if not ok:
            return self._send(403, why)

        req = urllib.request.Request(target, headers={
            'User-Agent': USER_AGENT,
            'Accept-Encoding': 'gzip',
        })
        try:
            with urllib.request.urlopen(req, timeout=UPSTREAM_TIMEOUT) as r:
                raw = r.read()
                if r.headers.get('Content-Encoding') == 'gzip':
                    raw = gzip.GzipFile(fileobj=io.BytesIO(raw)).read()
                ctype = r.headers.get('Content-Type', 'application/octet-stream')
                return self._send(r.status, raw, ctype)
        except urllib.error.HTTPError as e:
            # DART가 낸 오류는 그대로 전달한다 (키 오류 등은 클라이언트가 읽어야 한다)
            body = e.read() if hasattr(e, 'read') else b''
            return self._send(e.code, body or f'DART 오류 {e.code}',
                              e.headers.get('Content-Type', 'text/plain'))
        except Exception as e:
            # 502 = 중계는 살아 있으나 DART까지 못 갔다는 신호
            return self._send(502, f'DART 연결 실패: {e}')
