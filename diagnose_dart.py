"""DART 접속 진단 도구.

같은 파일을 두 곳에서 돌려 결과를 비교하는 것이 목적이다.
  1) 회계사님 PC (한국 IP)  → 기준선
  2) Streamlit Cloud (해외) → 비교 대상

DART API는 잘못된 키에도 정상 JSON을 돌려주므로, API 키 없이도 도달 여부를 판정할 수 있다.

실행:
    streamlit run diagnose_dart.py
"""

import json
import socket
import time

import requests
import streamlit as st

DART_API = "https://opendart.fss.or.kr/api/corpCode.xml"
DART_HOST = "opendart.fss.or.kr"
TIMEOUT = 15


# ── 판정 로직 (화면과 분리 — 테스트 가능) ────────────────────────────

def classify(probe):
    """조회 결과 → (판정, 설명). 세 갈래로 나뉜다."""
    if probe.get('error_kind') == 'timeout':
        return '연결실패', f'{TIMEOUT}초 안에 응답이 없습니다. 차단이거나 네트워크가 막혀 있습니다.'
    if probe.get('error_kind') == 'unreachable':
        return '연결실패', '서버에 연결되지 않습니다. 차단이거나 네트워크가 막혀 있습니다.'
    if probe.get('error_kind'):
        return '연결실패', f"조회 중 오류: {probe.get('error')}"

    body = probe.get('body') or ''
    head = body.lstrip()[:400].lower()

    # DART API는 XML 또는 JSON으로 status 코드를 돌려준다.
    # 그게 왔다면 요청이 DART까지 도달했다는 뜻이다.
    if '<status>' in head or '"status"' in head or head.startswith('pk'):
        return '정상', 'DART가 정상 응답했습니다. 이 위치에서는 차단되지 않았습니다.'

    if '<html' in head or '<!doctype html' in head:
        return '차단의심', ('DART가 아니라 HTML 페이지가 돌아왔습니다. '
                        '차단 페이지이거나 중간에 다른 장비가 가로챈 것입니다.')

    return '판단불가', '응답은 왔지만 형태를 알 수 없습니다. 아래 원문을 확인해주세요.'


def probe_dart(api_key='preflight', timeout=TIMEOUT):
    """DART API에 1건 요청. 예외를 종류별로 구분해 담아 돌려준다."""
    out = {'url': DART_API, 'elapsed': None, 'status_code': None,
           'body': None, 'error': None, 'error_kind': None}
    t0 = time.time()
    try:
        r = requests.get(DART_API, params={'crtfc_key': api_key}, timeout=timeout)
        out['elapsed'] = round(time.time() - t0, 2)
        out['status_code'] = r.status_code
        out['content_type'] = r.headers.get('Content-Type', '')
        # 응답이 ZIP(정상 corpCode)일 수 있어 바이트 앞부분만 안전하게 문자열화
        out['body'] = r.content[:600].decode('utf-8', errors='replace')
    except requests.exceptions.Timeout:
        out['elapsed'] = round(time.time() - t0, 2)
        out['error_kind'] = 'timeout'
        out['error'] = 'timeout'
    except requests.exceptions.ConnectionError as e:
        out['elapsed'] = round(time.time() - t0, 2)
        out['error_kind'] = 'unreachable'
        out['error'] = str(e)[:300]
    except Exception as e:
        out['elapsed'] = round(time.time() - t0, 2)
        out['error_kind'] = 'other'
        out['error'] = str(e)[:300]
    return out


def resolve_host(host=DART_HOST):
    try:
        return socket.gethostbyname(host), None
    except Exception as e:
        return None, str(e)[:200]


# 국가 정보가 두 실행 결과를 비교하는 핵심 지표라, 한 곳이 막혀도 다른 곳으로 시도한다.
IP_SERVICES = [
    ('https://ipinfo.io/json', 'ip', 'country', 'org'),
    ('https://ipapi.co/json/', 'ip', 'country_code', 'org'),
    ('https://api.ipify.org?format=json', 'ip', None, None),
]


def where_am_i(timeout=8):
    """실행 위치의 공인 IP와 국가. 실패해도 진단은 계속한다."""
    last = None
    for url, k_ip, k_country, k_org in IP_SERVICES:
        try:
            j = requests.get(url, timeout=timeout).json()
            return {'ip': j.get(k_ip),
                    'country': j.get(k_country) if k_country else None,
                    'org': j.get(k_org) if k_org else None}, None
        except Exception as e:
            last = type(e).__name__
    return None, f'{len(IP_SERVICES)}곳 모두 실패 ({last})'


def build_report(loc, loc_err, host_ip, host_err, probe, verdict, detail):
    """복사해서 붙일 수 있는 텍스트."""
    lines = [
        "===== DART 접속 진단 =====",
        f"판정      : {verdict} — {detail}",
        "",
        f"공인 IP   : {loc.get('ip') if loc else f'조회 실패 ({loc_err})'}",
        f"국가      : {loc.get('country') if loc else '-'}",
        f"통신사/호스팅: {loc.get('org') if loc else '-'}",
        "",
        f"DNS 해석  : {host_ip or f'실패 ({host_err})'}",
        f"요청 주소 : {probe['url']}",
        f"소요 시간 : {probe['elapsed']}초",
        f"HTTP 코드 : {probe['status_code']}",
        f"Content-Type: {probe.get('content_type', '-')}",
        f"오류      : {probe['error'] or '없음'}",
        "",
        "--- 응답 앞부분 ---",
        (probe['body'] or '(없음)')[:600],
        "==========================",
    ]
    return "\n".join(lines)


# ── 화면 ──────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="DART 접속 진단", page_icon="🔎")
    st.title("🔎 DART 접속 진단")
    st.caption(
        "이 페이지를 **두 곳에서** 실행하고 결과를 비교합니다.\n\n"
        "① 회계사님 PC (한국) → 기준선  ②  Streamlit Cloud (해외) → 비교 대상"
    )

    with st.expander("API 키 (없어도 진단됩니다)"):
        api_key = st.text_input(
            "OpenDART API 키", type="password",
            help="비워두면 임시 키로 도달 여부만 확인합니다. "
                 "DART는 잘못된 키에도 정상 응답을 돌려주기 때문에 판정에는 문제가 없습니다.")

    if not st.button("진단 시작", type="primary"):
        st.stop()

    with st.spinner("확인 중..."):
        loc, loc_err = where_am_i()
        host_ip, host_err = resolve_host()
        probe = probe_dart(api_key.strip() or 'preflight')
        verdict, detail = classify(probe)

    if verdict == '정상':
        st.success(f"**{verdict}** — {detail}")
    elif verdict == '연결실패':
        st.error(f"**{verdict}** — {detail}")
    else:
        st.warning(f"**{verdict}** — {detail}")

    c1, c2, c3 = st.columns(3)
    c1.metric("국가", (loc or {}).get('country') or '?')
    c2.metric("소요 시간", f"{probe['elapsed']}초" if probe['elapsed'] is not None else '-')
    c3.metric("HTTP 코드", probe['status_code'] or '-')

    if loc:
        st.write(f"공인 IP `{loc.get('ip')}` · {loc.get('org') or ''}")
        if loc.get('country') and loc['country'] != 'KR':
            st.info("한국 밖에서 실행 중입니다. 이 결과가 '해외' 쪽 비교 대상입니다.")
    else:
        st.write(f"공인 IP 조회 실패: {loc_err}")

    st.divider()
    report = build_report(loc, loc_err, host_ip, host_err, probe, verdict, detail)
    st.subheader("결과 (전체 복사해서 보내주세요)")
    st.code(report, language="text")


if __name__ == "__main__":
    main()
