"""발행·유통주식수 조회. gpcm_kr.py 에서 그대로 옮겼다."""

import pandas as pd
import requests


def _to_int(x):
    try:
        if x is None:
            return None
        s = str(x).strip().replace(',', '')
        if s == '' or s.lower() == 'nan':
            return None
        return int(float(s))
    except Exception:
        return None

# --- DART 유통주식수 ---
DART_STOCKTOTQY_URL = "https://opendart.fss.or.kr/api/stockTotqySttus.json"

# 매 호출마다 새 TLS 연결을 맺지 않도록 세션 재사용
_DART_SESSION = requests.Session()

def fetch_dart_distb_shares(api_key, corp_code: str, bsns_year: int, reprt_code: str, cache=None):
    """같은 (회사, 연도, 보고서)를 기간별로 반복 조회하므로 결과를 캐시한다.
    '데이터 없음' 응답도 캐시해야 fallback 탐색의 재조회가 사라진다.
    다만 네트워크 오류(ERR)는 캐시하지 않아 일시적 실패가 고정되지 않도록 한다."""
    ck = ('shares', corp_code, int(bsns_year), str(reprt_code))
    if cache is not None and ck in cache:
        return cache[ck]

    shares, meta = _fetch_dart_distb_shares(api_key, corp_code, bsns_year, reprt_code)

    if cache is not None and meta.get('status') != 'ERR':
        cache[ck] = (shares, meta)
    return shares, meta


def _fetch_dart_distb_shares(api_key, corp_code: str, bsns_year: int, reprt_code: str):
    meta = {'shares': None, 'rcept_no': None, 'stlm_dt': None, 'se': None, 'status': None, 'message': None}
    try:
        params = {
            'crtfc_key': api_key,
            'corp_code': corp_code,
            'bsns_year': str(bsns_year),
            'reprt_code': str(reprt_code),
        }
        resp = _DART_SESSION.get(DART_STOCKTOTQY_URL, params=params, timeout=10)
        resp.raise_for_status()
        js = resp.json()

        meta['status'] = js.get('status')
        meta['message'] = js.get('message')

        if js.get('status') != '000':
            return None, meta

        df = pd.DataFrame(js.get('list', []))
        if df.empty:
            return None, meta

        if 'se' in df.columns:
            c1 = df[df['se'].astype(str).str.contains('보통', na=False)]
            c2 = df[df['se'].astype(str).str.contains('합계', na=False)]
            pick = c1 if not c1.empty else (c2 if not c2.empty else df)
        else:
            pick = df

        row = pick.iloc[0].to_dict()
        meta['rcept_no'] = row.get('rcept_no')
        meta['stlm_dt'] = row.get('stlm_dt')
        meta['se'] = row.get('se')

        shares = _to_int(row.get('distb_stock_co'))
        if shares is None:
            istc = _to_int(row.get('istc_totqy'))
            tes = _to_int(row.get('tesstk_co'))
            if istc is not None and tes is not None:
                shares = istc - tes

        meta['shares'] = shares
        return shares, meta

    except Exception as e:
        meta['status'] = meta['status'] or 'ERR'
        meta['message'] = str(e)
        return None, meta

def get_outstanding_shares(api_key, corp_code: str, ticker: str, bsns_year: int, reprt_code: str, df_krx: pd.DataFrame, cache=None):
    # 1. DART API 조회 (요청한 기준년도/분기)
    shares, meta = fetch_dart_distb_shares(api_key, corp_code, bsns_year, reprt_code, cache=cache)
    if shares is not None and shares > 0:
        return shares, f"DART({reprt_code})", meta

    # 2. 직전 보고서들에서 주식수 조회 시도 (요청 분기에 주식수 누락 시)
    # 시간순: 11013 (1Q), 11012 (반기), 11014 (3Q), 11011 (사업보고서)
    order = ['11013', '11012', '11014', '11011']
    try:
        current_idx = order.index(reprt_code)
    except Exception:
        current_idx = 3 # 기본 사업보고서 매핑

    cy = bsns_year
    ci = current_idx - 1
    
    # 최근 8개 분기(약 2년치)를 역순으로 훑어 가장 최근 공시된 주식총수를 찾음
    for _ in range(8):
        if ci < 0:
            cy -= 1
            ci = 3
        
        fb_code = order[ci]
        fb_shares, fb_meta = fetch_dart_distb_shares(api_key, corp_code, cy, fb_code, cache=cache)
        
        if fb_shares is not None and fb_shares > 0:
            # 과거 분기 정보를 찾았을 경우, 해당 출처(년도와 보고서 코드) 명시하여 반환
            return fb_shares, f"DART(Fallback:{cy}-{fb_code})", fb_meta
            
        ci -= 1

    # 3. KRX 캐시 조회 (작동 안 할 확률 높음)
    try:
        row = df_krx[df_krx['Code'] == ticker]
        if not row.empty:
            shares_krx = _to_int(row.iloc[0].get('Stocks'))
            if shares_krx is not None and shares_krx > 0:
                meta_f = dict(meta)
                meta_f['shares'] = shares_krx
                return shares_krx, 'KRX', meta_f
    except Exception:
        pass

    return None, 'N/A', meta
