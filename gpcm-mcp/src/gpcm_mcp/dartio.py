"""DART 호출 래퍼. gpcm_kr.py 에서 그대로 옮겼다.

get_dart_reader 는 원본의 @st.cache_resource 대신 지연 싱글턴으로 바꿨다.
API 키를 캐시 키로 쓰면 그게 곧 비밀 보관이 된다.
"""

import os
import time

import pandas as pd
import requests
import OpenDartReader


# --- DART PL Fetch Functions (Need dart instance) ---
def _safe_dart_call(fn, *args, max_retry=2, **kwargs):
    last_err = None
    for _ in range(max_retry + 1):
        try:
            df = fn(*args, **kwargs)
            return df, None
        except Exception as e:
            last_err = e
            time.sleep(0.4)
    return None, last_err

def safe_finstate(dart_instance, corp_code, year, reprt_code, max_retry=2):
    # OpenDartReader.finstate 는 fs_div 인자를 받지 않는다 (finstate_all 에만 존재)
    return _safe_dart_call(dart_instance.finstate, corp_code, year, max_retry=max_retry, reprt_code=reprt_code)

def safe_finstate_all(dart_instance, corp_code, year, reprt_code, fs_div=None, max_retry=2):
    kwargs = {'reprt_code': reprt_code}
    if fs_div is not None:
        kwargs['fs_div'] = fs_div
    return _safe_dart_call(dart_instance.finstate_all, corp_code, year, max_retry=max_retry, **kwargs)

def fetch_pl_df(dart_instance, corp_code, year, reprt_code):
    df, err = safe_finstate(dart_instance, corp_code, year, reprt_code)
    if df is not None and not df.empty: return df, 'finstate', None

    for fs in ['CFS', 'OFS']:
        df, err = safe_finstate_all(dart_instance, corp_code, year, reprt_code, fs_div=fs)
        if df is not None and not df.empty: return df, f'finstate_all|{fs}', None

    df, err = safe_finstate_all(dart_instance, corp_code, year, reprt_code, fs_div=None)
    if df is not None and not df.empty: return df, 'finstate_all|no_fs_div', None
    
    return None, 'N/A', 'NO_DATA'

def filter_income_statement(df: pd.DataFrame):
    if df is None or df.empty: return df
    if 'sj_div' in df.columns:
        df2 = df[df['sj_div'].astype(str) == 'IS'].copy()
        if not df2.empty: return df2
    if 'sj_nm' in df.columns:
        df2 = df[df['sj_nm'].astype(str).str.contains('손익|포괄손익', na=False)].copy()
        return df2
    return df

def check_dart_reachable(timeout=10):
    """OpenDartReader 는 timeout 을 지정하지 않아 접속 불가 시 수 분간 멈춘다.
    실제 조회 전에 짧은 timeout 으로 도달 가능 여부만 먼저 확인한다."""
    try:
        requests.get('https://opendart.fss.or.kr/api/corpCode.xml',
                     params={'crtfc_key': 'preflight'}, timeout=timeout)
        return True, None
    except requests.exceptions.Timeout:
        return False, 'timeout'
    except requests.exceptions.ConnectionError:
        return False, 'unreachable'
    except Exception as e:
        return False, str(e)

API_KEY_ENV = 'OPENDART_API_KEY'

_reader = None


def get_dart_reader(api_key=None):
    """OpenDartReader 를 한 번만 만들어 재사용한다.

    API 키는 환경변수에서만 읽는다. 도구 인자로 받으면 키가 대화 기록과 세션 로그에
    영구히 남는다. (api_key 인자는 테스트에서 가짜를 넣을 때만 쓴다.)

    원본은 @st.cache_resource 로 키마다 리더를 캐시했는데, 키를 캐시 키로 쓰는 것이
    곧 비밀을 메모리에 쌓아 두는 것이다. 프로세스 하나에 키 하나면 그 문제가 없다.
    """
    global _reader
    if api_key is not None:
        return OpenDartReader(api_key)
    if _reader is None:
        key = os.environ.get(API_KEY_ENV, '').strip()
        if not key:
            raise RuntimeError(
                f"{API_KEY_ENV} 환경변수가 없습니다. "
                "https://opendart.fss.or.kr 에서 무료로 발급받아 설정하세요. "
                "각자 본인 키를 쓰셔야 합니다 — 하루 조회 한도가 키마다 주어집니다."
            )
        _reader = OpenDartReader(key)
    return _reader


def api_key_configured():
    """키가 설정돼 있는지만 알려준다. 키 자체도, 앞자리도 내보내지 않는다."""
    return bool(os.environ.get(API_KEY_ENV, '').strip())
