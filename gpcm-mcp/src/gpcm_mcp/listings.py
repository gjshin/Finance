"""KRX 상장·업종 목록과 회사 식별. gpcm_kr.py 에서 그대로 옮겼다.

원본의 @st.cache_resource(ttl=3600) 두 개를 ttl_cached 로 바꿨다. lru_cache 는
TTL 이 없어서, 오래 떠 있는 MCP 서버가 낡은 상장목록을 영원히 물고 있게 된다.
"""

import time

import FinanceDataReader as fdr
import pandas as pd

from .cache import ttl_cached


@ttl_cached(3600)
def get_krx_listing():
    """KRX 상장종목 목록 조회 - 재시도 및 fallback 포함 (Streamlit Cloud 에러 대응)"""
    # 1차 시도: KRX 전체 (최대 3번)
    for attempt in range(3):
        try:
            df = fdr.StockListing('KRX')
            if df is not None and not df.empty:
                return df
        except Exception:
            if attempt < 2:
                time.sleep(1.0)

    # 2차 시도: KOSPI + KOSDAQ 개별 조회 후 병합
    frames = []
    for mkt in ['KOSPI', 'KOSDAQ']:
        for attempt in range(2):
            try:
                df_mkt = fdr.StockListing(mkt)
                if df_mkt is not None and not df_mkt.empty:
                    frames.append(df_mkt)
                    break
            except Exception:
                if attempt < 1:
                    time.sleep(1.0)
    if frames:
        return pd.concat(frames, ignore_index=True).drop_duplicates(subset=['Code'])

    # 최후 fallback: 빈 DataFrame 반환 (코드만으로 진행 가능하도록)
    return pd.DataFrame(columns=['Code', 'Name', 'Stocks'])

@ttl_cached(3600)
def get_krx_industry_listing():
    """업종·주요제품이 붙은 상장사 목록 (KRX 상장회사목록).

    get_krx_listing()이 쓰는 'KRX'는 가격·시총 목록이라 업종이 없다. 업종으로
    Peer 후보를 추리려면 이쪽이 필요하다. 인증키는 필요 없다.

    반환: Code, Name, Market, Sector(업종), Industry(주요제품), SettleMonth(결산월)
    실패하면 빈 DataFrame — 호출부는 종목코드 직접 입력으로 되돌아간다.
    """
    for attempt in range(2):
        try:
            df = fdr.StockListing('KRX-DESC')
            if df is not None and not df.empty and 'Sector' in df.columns:
                return df
        except Exception:
            if attempt < 1:
                time.sleep(1.0)
    return pd.DataFrame(columns=['Code', 'Name', 'Market', 'Sector', 'Industry', 'SettleMonth'])


def peer_candidate_rows(df_ind, sector: str):
    """업종 하나에 속한 상장사를 종목코드 순으로 정리한다.

    결산월이 12월이 아니면 비교기간이 어긋나므로 골라내기 전에 보이게 표시한다.
    """
    if df_ind is None or df_ind.empty or not sector:
        return []
    sub = df_ind[df_ind['Sector'] == sector].sort_values('Code')
    rows = []
    for _, r in sub.iterrows():
        settle = str(r.get('SettleMonth') or '').strip()
        rows.append({
            'Code': str(r.get('Code') or '').zfill(6),
            'Name': str(r.get('Name') or ''),
            'Market': str(r.get('Market') or ''),
            'Product': str(r.get('Industry') or '').strip(),
            'SettleMonth': settle,
            'FiscalNot12': bool(settle) and '12' not in settle,
        })
    return rows


def resolve_company_info(dart_instance, ticker: str):
    df_krx = get_krx_listing()
    rows = df_krx[df_krx['Code'] == ticker]
    krx_name = rows.iloc[0]['Name'] if not rows.empty else None

    # DART 내장 corp_codes 로 직접 이름 검색 (KRX 실패 대비)
    if krx_name is None:
        try:
            dart_rows = dart_instance.corp_codes[dart_instance.corp_codes['stock_code'] == ticker]
            if not dart_rows.empty:
                krx_name = dart_rows.iloc[0]['corp_name']
        except Exception:
            pass

    corp_code = None
    try:
        corp_code = dart_instance.find_corp_code(ticker)
    except Exception:
        corp_code = None

    if not corp_code and krx_name:
        try:
            corp_code = dart_instance.find_corp_code(krx_name)
        except Exception:
            corp_code = None

    return corp_code, krx_name
