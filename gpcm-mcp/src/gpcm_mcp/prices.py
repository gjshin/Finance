"""주가·지수 시계열. gpcm_kr.py 에서 그대로 옮겼다."""

from datetime import datetime, timedelta

import FinanceDataReader as fdr
import pandas as pd
import yfinance as yf


def _to_price_series(df, col='Close'):
    """yf.download / fdr.DataReader 결과에서 1-D 가격 Series 추출."""
    if isinstance(df, pd.Series):
        return df
    if col in df.columns:
        s = df[col]
    else:
        s = df.iloc[:, 0]
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]
    return s

def _slice_from(df, start_date_str):
    """이미 받아둔 넓은 구간 시계열에서 start_date 이후만 잘라낸다."""
    if df is None or len(df) == 0:
        return df
    try:
        idx = df.index
        if not isinstance(idx, pd.DatetimeIndex):
            idx = pd.to_datetime(idx)
        start = pd.to_datetime(start_date_str)
        if getattr(idx, 'tz', None) is not None:
            start = start.tz_localize(idx.tz)
        return df[idx >= start]
    except Exception:
        return df


def _get_market_index_data(market_idx, start, end, cache):
    """시장지수는 전 종목 공통이므로 (지수, 시작, 종료) 기준으로 1회만 조회한다."""
    key = (market_idx, start, end)
    if key in cache:
        return cache[key]
    if market_idx.startswith('^'):
        data = yf.download(market_idx, start=start, end=end, progress=False)
    else:
        data = fdr.DataReader(market_idx, start, end)
    cache[key] = data
    return data


def get_market_index(ticker):
    """
    티커 기반으로 거래소 및 시장지수 코드 반환 (한국 종목만 지원)
    Returns: (exchange_name, index_symbol)
    """
    # 한국 종목 - FinanceDataReader 기준
    # KS11 (KOSPI)가 fdr에서 실패하는 경우가 많아 yfinance 심볼(^KS11)로 대체
    return 'KRX', '^KS11'  # 기본값: KOSPI
def get_stock_price(ticker: str, date_str: str):
    try:
        td = pd.to_datetime(date_str)
        if td > datetime.now():
            return None, None
        df = fdr.DataReader(ticker, td - timedelta(days=10), td)
        if df is not None and not df.empty:
            return float(df.iloc[-1]['Close']), df.index[-1].strftime('%Y-%m-%d')
        return None, None
    except Exception:
        return None, None
