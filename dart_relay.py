"""DART 중계 서버 경유 설정.

Streamlit Cloud 같은 해외 서버에서는 DART에 직접 닿지 않는다.
서울 리전에 올린 중계 함수를 거치도록 요청 주소를 바꿔치기한다.

    Streamlit Cloud (해외) → Vercel 함수 (서울) → DART

OpenDartReader는 내부에서 requests.get 을 직접 부르고 중계를 지원하지 않는다.
그래서 requests.get 을 한 번 감싸는 방식을 쓴다. 라이브러리의 모든 모듈이
같은 requests 모듈 객체를 공유하므로 한 곳만 바꾸면 전부 적용된다.

중계 서버가 설정되지 않았으면 아무것도 하지 않는다 — 국내 PC에서는 직접 호출이 빠르다.
"""

import os
from urllib.parse import quote

import requests

# 중계 대상 호스트. API와 공시원문 뷰어가 서로 다른 호스트를 쓴다.
DART_HOSTS = ('opendart.fss.or.kr', 'dart.fss.or.kr')

# 라이브러리가 timeout 을 지정하지 않아 접속 불가 시 수 분간 멈춘다. 기본값을 준다.
DEFAULT_TIMEOUT = 30

_installed = False


def relay_config(secrets=None):
    """중계 설정을 읽는다. Streamlit secrets → 환경변수 순.

    반환: (중계주소, 토큰). 설정이 없으면 (None, None).
    """
    url = token = None
    if secrets is not None:
        try:
            url = secrets.get('DART_RELAY_URL')
            token = secrets.get('DART_RELAY_TOKEN')
        except Exception:
            pass
    url = (url or os.environ.get('DART_RELAY_URL') or '').strip().rstrip('/')
    token = (token or os.environ.get('DART_RELAY_TOKEN') or '').strip()
    return (url or None), (token or None)


def build_relay_url(base, target_url):
    """중계 서버가 이해하는 주소로 바꾼다.

    대상 주소를 통째로 인코딩해 넘기므로 질의문자열이 뒤엉키지 않는다.
    """
    return f"{base}/api/relay?u={quote(target_url, safe='')}"


def needs_relay(url):
    """DART 호스트로 가는 요청인지"""
    return any(h in str(url) for h in DART_HOSTS)


def _rewrite(base, token, timeout, url, params, kwargs):
    """중계용 주소와 호출 인자를 만든다."""
    full = requests.Request('GET', url, params=params).prepare().url
    kwargs.setdefault('timeout', timeout)
    if token:
        headers = dict(kwargs.get('headers') or {})
        headers['X-Relay-Token'] = token
        kwargs['headers'] = headers
    return build_relay_url(base, full), kwargs


def install(base, token=None, timeout=DEFAULT_TIMEOUT):
    """DART 요청만 중계로 돌린다. 중복 호출은 무시한다.

    requests.get 과 Session.get 을 모두 감싼다. 앱이 주식수 조회에는 Session 을
    쓰기 때문에 한쪽만 감싸면 그 요청만 중계를 우회한다.
    """
    global _installed
    if _installed or not base:
        return False

    base = base.rstrip('/')
    orig_get = requests.get
    orig_session_get = requests.Session.get

    def relayed_get(url, params=None, **kwargs):
        if not needs_relay(url):
            return orig_get(url, params=params, **kwargs)
        new_url, kwargs = _rewrite(base, token, timeout, url, params, kwargs)
        return orig_get(new_url, **kwargs)

    def relayed_session_get(self, url, **kwargs):
        if not needs_relay(url):
            return orig_session_get(self, url, **kwargs)
        params = kwargs.pop('params', None)
        new_url, kwargs = _rewrite(base, token, timeout, url, params, kwargs)
        return orig_session_get(self, new_url, **kwargs)

    requests.get = relayed_get
    requests.Session.get = relayed_session_get
    _installed = True
    return True


def is_installed():
    return _installed


def probe(base, token=None, timeout=20):
    """중계 서버가 살아 있고 DART까지 닿는지 확인. (성공여부, 설명)"""
    if not base:
        return False, '중계 서버 주소가 설정되지 않았습니다.'
    target = 'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=preflight'
    headers = {'X-Relay-Token': token} if token else None
    try:
        r = requests.request('GET', build_relay_url(base, target),
                             headers=headers, timeout=timeout)
    except requests.exceptions.Timeout:
        return False, f'중계 서버가 {timeout}초 안에 응답하지 않았습니다.'
    except requests.exceptions.ConnectionError:
        return False, '중계 서버에 연결되지 않습니다. 주소를 확인해주세요.'
    except Exception as e:
        return False, f'중계 서버 확인 중 오류: {e}'

    if r.status_code == 401:
        return False, '중계 서버가 토큰을 거부했습니다. DART_RELAY_TOKEN 을 확인해주세요.'
    if r.status_code == 403:
        return False, '중계 서버가 이 주소로의 요청을 거부했습니다.'
    if r.status_code == 502:
        return False, ('중계 서버는 살아 있지만 거기서 DART로 연결되지 않습니다. '
                       '중계 서버 리전이 서울(icn1)인지 확인해주세요.')
    if r.status_code != 200:
        return False, f'중계 서버가 예상 못 한 응답을 보냈습니다 (HTTP {r.status_code}).'

    head = r.content[:300].decode('utf-8', errors='replace').lstrip().lower()
    if '<status>' in head or '"status"' in head or head.startswith('pk'):
        return True, '중계 서버를 통해 DART에 연결됩니다.'
    if '<html' in head or '<!doctype html' in head:
        return False, '중계는 됐지만 DART가 아니라 HTML 페이지가 돌아왔습니다.'
    return False, '중계 서버 응답 형태를 알 수 없습니다.'
