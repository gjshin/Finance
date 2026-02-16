import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import io
import numpy as np
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.chart.axis import ChartLines
import FinanceDataReader as fdr
from scipy import stats
import requests
from bs4 import BeautifulSoup

# ==========================================
# 1. Helper Functions (v17 Logic + Beta Calculation)
# ==========================================

def get_market_index(ticker):
    """
    티커 기반으로 거래소 및 시장지수 코드 반환
    Returns: (exchange_name, index_symbol)
    """
    ticker_upper = ticker.upper()

    # 아시아
    if ticker_upper.endswith('.KS'):
        return 'KOSPI', 'KS11'
    elif ticker_upper.endswith('.KQ'):
        return 'KOSDAQ', 'KQ11'
    elif ticker_upper.endswith('.T'):
        return 'TSE', '^N225'  # Nikkei 225
    elif ticker_upper.endswith('.SS'):
        return 'SSE', '000001.SS'  # Shanghai Composite
    elif ticker_upper.endswith('.SZ'):
        return 'SZSE', '399001.SZ'  # Shenzhen Component
    elif ticker_upper.endswith('.HK'):
        return 'HKEX', '^HSI'  # Hang Seng
    elif ticker_upper.endswith('.SI'):
        return 'SGX', '^STI'  # Straits Times Index
    elif ticker_upper.endswith('.TW'):
        return 'TWSE', '^TWII'  # Taiwan Weighted
    elif ticker_upper.endswith('.BO') or ticker_upper.endswith('.NS'):
        return 'NSE', '^NSEI'  # Nifty 50

    # 북미
    elif ticker_upper.endswith('.TO') or ticker_upper.endswith('.V'):
        return 'TSX', '^GSPTSE'  # S&P/TSX Composite
    elif ticker_upper.endswith('.MX'):
        return 'BMV', '^MXX'  # IPC Mexico

    # 유럽
    elif ticker_upper.endswith('.F') or ticker_upper.endswith('.DE') or ticker_upper.endswith('.BE'):
        return 'XETRA', '^GDAXI'  # DAX
    elif ticker_upper.endswith('.PA'):
        return 'Euronext', '^FCHI'  # CAC 40
    elif ticker_upper.endswith('.L'):
        return 'LSE', '^FTSE'  # FTSE 100
    elif ticker_upper.endswith('.MI'):
        return 'Borsa', 'FTSEMIB.MI'  # FTSE MIB
    elif ticker_upper.endswith('.MC'):
        return 'BME', '^IBEX'  # IBEX 35
    elif ticker_upper.endswith('.AS'):
        return 'Euronext', '^AEX'  # AEX Amsterdam
    elif ticker_upper.endswith('.SW') or ticker_upper.endswith('.VX'):
        return 'SIX', '^SSMI'  # Swiss Market Index
    elif ticker_upper.endswith('.VI'):
        return 'VSE', '^ATX'  # ATX Vienna
    elif ticker_upper.endswith('.BR'):
        return 'Euronext', '^BFX'  # BEL 20
    elif ticker_upper.endswith('.ST'):
        return 'OMX', '^OMX'  # OMX Stockholm 30
    elif ticker_upper.endswith('.OL'):
        return 'OSE', 'OSEBX.OL'  # Oslo Børs Index
    elif ticker_upper.endswith('.CO'):
        return 'OMX', '^OMXC20'  # OMX Copenhagen 20
    elif ticker_upper.endswith('.HE'):
        return 'OMX', '^OMXH25'  # OMX Helsinki 25
    elif ticker_upper.endswith('.IR'):
        return 'Euronext', '^ISEQ'  # ISEQ Overall

    # 오세아니아
    elif ticker_upper.endswith('.AX'):
        return 'ASX', '^AORD'  # All Ordinaries
    elif ticker_upper.endswith('.NZ'):
        return 'NZX', '^NZ50'  # NZX 50

    # 기타
    elif ticker_upper.endswith('.SA'):
        return 'B3', '^BVSP'  # Bovespa
    elif ticker_upper.endswith('.ME'):
        return 'MOEX', 'IMOEX.ME'  # MOEX Russia
    elif ticker_upper.endswith('.JO'):
        return 'JSE', 'J203.JO'  # FTSE/JSE Top 40

    # 기본값: 미국 S&P 500
    else:
        return 'US', '^GSPC'  # S&P 500

def calculate_beta(stock_returns, market_returns, min_periods=20):
    """
    주식 수익률과 시장 수익률로부터 베타 계산
    Returns: raw_beta, adjusted_beta (None if invalid)
    """
    try:
        # [수정됨] Timezone 문제 해결: yfinance(tz-aware)와 FDR(tz-naive) 간 인덱스 통일
        if stock_returns.index.tz is not None:
            stock_returns.index = stock_returns.index.tz_localize(None)
        if market_returns.index.tz is not None:
            market_returns.index = market_returns.index.tz_localize(None)

        if len(stock_returns) < min_periods or len(market_returns) < min_periods:
            return None, None

        # 공통 인덱스로 정렬
        common_idx = stock_returns.index.intersection(market_returns.index)
        if len(common_idx) < min_periods:
            return None, None

        stock_ret = stock_returns.loc[common_idx].dropna()
        market_ret = market_returns.loc[common_idx].dropna()

        common_idx2 = stock_ret.index.intersection(market_ret.index)
        if len(common_idx2) < min_periods:
            return None, None

        stock_ret = stock_ret.loc[common_idx2]
        market_ret = market_ret.loc[common_idx2]

        # 선형회귀로 베타 계산
        slope, intercept, r_value, p_value, std_err = stats.linregress(market_ret, stock_ret)
        raw_beta = slope

        # 조정 베타: 2/3 * Raw Beta + 1/3 * 1.0 (Bloomberg 방식)
        adjusted_beta = (2/3) * raw_beta + (1/3) * 1.0

        # 값 검증: NaN, inf, 비정상 값 체크
        if not np.isfinite(raw_beta) or not np.isfinite(adjusted_beta):
            return None, None

        # 극단적 값 필터링 (베타가 -10 ~ 10 범위를 벗어나면 이상)
        if abs(raw_beta) > 10 or abs(adjusted_beta) > 10:
            return None, None

        return float(raw_beta), float(adjusted_beta)

    except Exception as e:
        # 계산 중 에러 발생 시 None 반환
        return None, None

    return raw_beta, adjusted_beta

@st.cache_data(ttl=86400)  # 24시간 캐시
def get_corporate_tax_rates_from_wikipedia():
    """
    국가별 법인세율 조회 (Wikipedia 크롤링 + 견고한 기본값)
    Returns: dict {country_code: tax_rate}
    """
    # [1] 견고한 기본값 설정 (2025년 기준, 지방세 포함)
    tax_rates = {
        # 아시아
        'KR': 0.231,  # 한국: 21% + 지방세 2.1% = 23.1% (중간 구간)
        'JP': 0.304,  # 일본: 23.2% (국세) + 지방세 ~7% = 30.4%
        'CN': 0.25,   # 중국: 25%
        'HK': 0.165,  # 홍콩: 16.5%
        'SG': 0.17,   # 싱가포르: 17%
        'TW': 0.20,   # 대만: 20%
        'IN': 0.304,  # 인도: 25.17% + 할증세 = 30.4%

        # 북미
        'US': 0.21,   # 미국: 21% (연방세, 주세 별도)
        'CA': 0.265,  # 캐나다: 15% (연방) + 11.5% (평균 주세) = 26.5%
        'MX': 0.30,   # 멕시코: 30%

        # 유럽
        'DE': 0.30,   # 독일: 15% + 연대세 + 영업세 = ~30%
        'FR': 0.25,   # 프랑스: 25%
        'GB': 0.25,   # 영국: 25%
        'IT': 0.24,   # 이탈리아: 24%
        'ES': 0.25,   # 스페인: 25%
        'NL': 0.256,  # 네덜란드: 25.8%
        'CH': 0.148,  # 스위스: 14.8% (평균)
        'AT': 0.24,   # 오스트리아: 24%
        'BE': 0.25,   # 벨기에: 25%
        'SE': 0.206,  # 스웨덴: 20.6%
        'NO': 0.22,   # 노르웨이: 22%
        'DK': 0.22,   # 덴마크: 22%
        'FI': 0.20,   # 핀란드: 20%
        'IE': 0.125,  # 아일랜드: 12.5%
        'LU': 0.245,  # 룩셈부르크: 24.5%

        # 오세아니아
        'AU': 0.30,   # 호주: 30%
        'NZ': 0.28,   # 뉴질랜드: 28%

        # 기타
        'BR': 0.34,   # 브라질: 34%
        'RU': 0.20,   # 러시아: 20%
        'ZA': 0.27,   # 남아공: 27%
    }

    # [2] Wikipedia에서 최신 세율 업데이트 시도
    try:
        url = "https://en.wikipedia.org/wiki/List_of_countries_by_tax_rates"
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.content, 'html.parser')

        # Wikipedia 테이블에서 법인세율 추출
        tables = soup.find_all('table', {'class': 'wikitable'})
        country_map = {
            'United States': 'US', 'Japan': 'JP', 'Canada': 'CA', 'Germany': 'DE',
            'Austria': 'AT', 'South Korea': 'KR', 'China': 'CN', 'Hong Kong': 'HK',
            'Singapore': 'SG', 'Taiwan': 'TW', 'India': 'IN', 'Mexico': 'MX',
            'France': 'FR', 'United Kingdom': 'GB', 'Italy': 'IT', 'Spain': 'ES',
            'Netherlands': 'NL', 'Switzerland': 'CH', 'Belgium': 'BE', 'Sweden': 'SE',
            'Norway': 'NO', 'Denmark': 'DK', 'Finland': 'FI', 'Ireland': 'IE',
            'Luxembourg': 'LU', 'Australia': 'AU', 'New Zealand': 'NZ', 'Brazil': 'BR',
            'Russia': 'RU', 'South Africa': 'ZA',
        }

        for table in tables:
            rows = table.find_all('tr')
            for row in rows[1:]:  # 헤더 제외
                cols = row.find_all('td')
                if len(cols) >= 2:
                    country = cols[0].get_text(strip=True)
                    tax_text = cols[1].get_text(strip=True) if len(cols) > 1 else ''

                    # 숫자만 추출 (예: "21%" → 0.21)
                    import re
                    match = re.search(r'(\d+\.?\d*)', tax_text)
                    if match:
                        rate = float(match.group(1)) / 100

                        # 국가명 매칭
                        for name, code in country_map.items():
                            if name.lower() in country.lower():
                                tax_rates[code] = rate
                                break

        # Wikipedia 크롤링 성공 시 알림
        # st.info(f"✅ Wikipedia에서 {len(tax_rates)}개 국가 법인세율 업데이트 완료")

    except Exception as e:
        # 크롤링 실패 시 기본값 사용 (경고 제거 - 너무 많이 뜸)
        pass

    return tax_rates

def get_korean_marginal_tax_rate(pretax_income_millions):
    """
    한국 법인세 한계세율 산출 (2025년 기준, 지방세 포함)
    과세표준 기준 (단위: 백만원)
    - 2억 이하: 9% (국세) + 0.9% (지방세 10%) = 9.9%
    - 2억 ~ 200억: 19% + 1.9% = 20.9%
    - 200억 ~ 3,000억: 21% + 2.1% = 23.1%
    - 3,000억 초과: 24% + 2.4% = 26.4%
    """
    if pd.isna(pretax_income_millions) or pretax_income_millions == 0:
        return 0.231  # 기본값 (중간 구간)

    # 백만원 단위로 들어온 값
    if pretax_income_millions <= 200:
        return 0.099
    elif pretax_income_millions <= 20000:
        return 0.209
    elif pretax_income_millions <= 300000:
        return 0.231
    else:
        return 0.264

def get_country_from_ticker(ticker):
    """
    티커로부터 국가 코드 추출 (거래소 suffix 기반)
    참고: https://help.yahoo.com/kb/SLN2310.html
    """
    ticker_upper = ticker.upper()

    # 아시아
    if ticker_upper.endswith('.KS') or ticker_upper.endswith('.KQ'):
        return 'KR'  # 한국 (KOSPI, KOSDAQ)
    elif ticker_upper.endswith('.T'):
        return 'JP'  # 일본 (Tokyo)
    elif ticker_upper.endswith('.SS'):
        return 'CN'  # 중국 (Shanghai)
    elif ticker_upper.endswith('.SZ'):
        return 'CN'  # 중국 (Shenzhen)
    elif ticker_upper.endswith('.HK'):
        return 'HK'  # 홍콩
    elif ticker_upper.endswith('.SI'):
        return 'SG'  # 싱가포르
    elif ticker_upper.endswith('.TW'):
        return 'TW'  # 대만
    elif ticker_upper.endswith('.BO') or ticker_upper.endswith('.NS'):
        return 'IN'  # 인도 (Bombay, National Stock Exchange)

    # 북미
    elif ticker_upper.endswith('.TO') or ticker_upper.endswith('.V'):
        return 'CA'  # 캐나다 (Toronto, Vancouver)
    elif ticker_upper.endswith('.MX'):
        return 'MX'  # 멕시코

    # 유럽
    elif ticker_upper.endswith('.F') or ticker_upper.endswith('.DE') or ticker_upper.endswith('.BE'):
        return 'DE'  # 독일 (Frankfurt, Xetra, Berlin)
    elif ticker_upper.endswith('.PA'):
        return 'FR'  # 프랑스 (Paris)
    elif ticker_upper.endswith('.L'):
        return 'GB'  # 영국 (London)
    elif ticker_upper.endswith('.MI'):
        return 'IT'  # 이탈리아 (Milan)
    elif ticker_upper.endswith('.MC'):
        return 'ES'  # 스페인 (Madrid)
    elif ticker_upper.endswith('.AS'):
        return 'NL'  # 네덜란드 (Amsterdam)
    elif ticker_upper.endswith('.SW') or ticker_upper.endswith('.VX'):
        return 'CH'  # 스위스 (Swiss Exchange, Virt-X)
    elif ticker_upper.endswith('.VI'):
        return 'AT'  # 오스트리아 (Vienna)
    elif ticker_upper.endswith('.BR'):
        return 'BE'  # 벨기에 (Brussels)
    elif ticker_upper.endswith('.ST'):
        return 'SE'  # 스웨덴 (Stockholm)
    elif ticker_upper.endswith('.OL'):
        return 'NO'  # 노르웨이 (Oslo)
    elif ticker_upper.endswith('.CO'):
        return 'DK'  # 덴마크 (Copenhagen)
    elif ticker_upper.endswith('.HE'):
        return 'FI'  # 핀란드 (Helsinki)
    elif ticker_upper.endswith('.IR'):
        return 'IE'  # 아일랜드 (Irish)

    # 오세아니아
    elif ticker_upper.endswith('.AX'):
        return 'AU'  # 호주
    elif ticker_upper.endswith('.NZ'):
        return 'NZ'  # 뉴질랜드

    # 기타
    elif ticker_upper.endswith('.SA'):
        return 'BR'  # 브라질
    elif ticker_upper.endswith('.ME'):
        return 'RU'  # 러시아 (MOEX)
    elif ticker_upper.endswith('.JO'):
        return 'ZA'  # 남아프리카공화국

    # 기본값: 미국 (suffix 없는 경우)
    else:
        return 'US'

def calculate_unlevered_beta(levered_beta, debt, equity, tax_rate):
    """
    하마다 모형으로 Unlevered Beta 계산
    Unlevered Beta = Levered Beta / (1 + (1 - Tax Rate) * (Debt / Equity))
    """
    if pd.isna(levered_beta) or levered_beta is None:
        return None
    if pd.isna(debt) or pd.isna(equity) or equity == 0:
        return levered_beta

    unlevered = levered_beta / (1 + (1 - tax_rate) * (debt / equity))
    return unlevered

def get_korea_10y_treasury_yield(base_date_str, user_rf_rate=0.033):
    """
    무위험이자율 반환 (사용자 입력값)

    Parameters:
    - user_rf_rate: 사용자가 직접 입력한 무위험이자율 (기본값 3.3%)
    """
    st.info(f"💡 무위험이자율 (사용자 입력): {user_rf_rate*100:.2f}%")
    return user_rf_rate
@st.cache_data(ttl=3600)  # <--- [추가] 1시간 동안 데이터를 저장해서 재사용함
def get_gpcm_data(tickers_list, base_date_str, mrp=0.08, kd_pretax=0.035, size_premium=0.0402, target_tax_rate=0.264, user_rf_rate=None, beta_type="5Y"):
    """
    GPCM 데이터 수집 및 엑셀 생성을 위한 데이터 구조 반환

    Parameters:
    - mrp: Market Risk Premium (기본값 8%)
    - kd_pretax: 세전 타인자본비용 (기본값 5%)
    - size_premium: Size Premium (기본값 4.02%, 한국공인회계사회 Micro 기준)
    - target_tax_rate: Target 기업 법인세율 (기본값 26.4%, 한국 대기업 기준)
    - user_rf_rate: 사용자 입력 무위험이자율 (None이면 자동 조회)
    - beta_type: WACC 계산에 사용할 베타 유형 ("5Y" 또는 "2Y", 기본값 "5Y")

    Note: 목표 부채비율은 피어들의 평균 자본구조로 자동 계산됨
          개별 peer의 WACC이 아닌 Target 기업의 WACC을 계산함
    """
    base_dt = pd.to_datetime(base_date_str)

    # 10년 국채수익률 조회 (무위험수익률)
    rf_rate = get_korea_10y_treasury_yield(base_date_str, user_rf_rate)
    
    # ---------------------------------------------------------
    # [설정] 계정 맵핑 (v17: NOA Option, 투자부동산 등)
    # ---------------------------------------------------------
    BS_HIGHLIGHT_MAP = {
        'Cash And Cash Equivalents':           'Cash',
        'Other Short Term Investments':        'Cash',
        'Current Debt And Capital Lease Obligation': 'IBD',
        'Long Term Debt And Capital Lease Obligation': 'IBD',
        'Minority Interest':                   'NCI',
        'Stockholders Equity':                 'Equity',
        'Long Term Equity Investment':                         'NOA(Option)',
        'Investments In Other Ventures Under Equity Method':   'NOA(Option)',
        'Investment In Financial Assets':                      'NOA(Option)',
        'Investmentin Financial Assets':                       'NOA(Option)',
        'Investment Properties':                               'NOA(Option)',
        'Non Current Note Receivables':                        'NOA(Option)',
        'Other Investments':                                   'NOA(Option)',
    }
    BS_SUBTOTAL_EXCLUDE = {
        'Cash Cash Equivalents And Short Term Investments', 'Cash And Short Term Investments',
        'Total Debt', 'Total Capitalization', 'Total Equity Gross Minority Interest',
    }
    PL_HIGHLIGHT_MAP = {'Total Revenue': 'Revenue', 'Operating Income': 'EBIT', 'EBIT': 'EBIT', 'EBITDA': 'EBIT', 'Normalized EBITDA': 'EBIT', 'Net Income Common Stockholders': 'NI_Parent', 'Net Income': 'Net Income'}
    PL_CALC_KEY = {'Total Revenue': 'Revenue', 'Operating Income': 'OpIncome', 'EBIT': 'EBIT_yf', 'EBITDA': 'EBITDA_yf', 'Normalized EBITDA': 'NormEBITDA_yf', 'Net Income Common Stockholders': 'NI_Parent'}
    PL_SORT = {'Total Revenue': 10, 'Operating Revenue': 11, 'Cost Of Revenue': 20, 'Gross Profit': 30, 'Operating Expense': 35, 'Selling General And Administration': 36, 'Research And Development': 37, 'Operating Income': 50, 'EBIT': 55, 'EBITDA': 56, 'Normalized EBITDA': 57, 'Interest Expense': 60, 'Pretax Income': 70, 'Tax Provision': 75, 'Net Income': 90, 'Net Income Common Stockholders': 91, 'Basic EPS': 95, 'Diluted EPS': 96}

    # Data Containers
    gpcm_data = {}
    raw_bs_rows = []
    raw_pl_rows = []
    market_rows = []
    price_abs_dfs = []
    price_rel_dfs = []
    ticker_to_name = {}

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_tickers = len(tickers_list)

    for idx, ticker in enumerate(tickers_list):
        time.sleep(1)
        status_text.text(f"Processing: {ticker}...")
        progress_bar.progress((idx + 1) / total_tickers)

        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            company_name = info.get('longName') or info.get('shortName') or ticker
            ticker_to_name[ticker] = company_name
            currency = info.get('currency', 'USD')

            gpcm = {
                'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                'Base_Date': base_date_str,
                'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA(Option)': 0, 'Equity': 0,
                'Revenue': 0, 'EBIT': 0, 'EBITDA': 0, 'NI_Parent': 0,
                'Close': 0, 'Shares': 0, 'Market_Cap_M': 0, 'PL_Source': '',
                'Exchange': '', 'Market_Index': '',
                'Beta_5Y_Monthly_Raw': None, 'Beta_5Y_Monthly_Adj': None,
                'Beta_2Y_Weekly_Raw': None, 'Beta_2Y_Weekly_Adj': None,
                'Pretax_Income': 0, 'Tax_Rate': 0.22,
                'Debt_Ratio': 0, 'Unlevered_Beta_5Y': None, 'Unlevered_Beta_2Y': None,
            }

            # [0] Price History (10Y)
            try:
                hist_10y_raw = stock.history(start=(base_dt - timedelta(days=365*10+20)).strftime('%Y-%m-%d'), end=base_dt.strftime('%Y-%m-%d'), auto_adjust=False)
                hist_10y = hist_10y_raw['Close'] if 'Close' in hist_10y_raw.columns else hist_10y_raw.iloc[:,0]
                if not hist_10y.empty:
                    abs_s = hist_10y.copy(); abs_s.name = ticker; price_abs_dfs.append(abs_s)
                    rel_s = (hist_10y / hist_10y.iloc[0]) * 100; rel_s.name = ticker; price_rel_dfs.append(rel_s)
            except: pass

            # [1] BS
            q_bs = stock.quarterly_balance_sheet
            bs_shares = None
            if q_bs is not None and not q_bs.empty:
                valid = sorted([d for d in q_bs.columns if d <= base_dt + timedelta(days=7)], reverse=True)
                if valid:
                    latest = valid[0]
                    for acct_name in q_bs.index:
                        val = q_bs.loc[acct_name, latest]
                        if pd.isna(val): continue
                        val_f = float(val)
                        if str(acct_name) == 'Ordinary Shares Number': bs_shares = val_f
                        ev_tag = BS_HIGHLIGHT_MAP.get(str(acct_name), '')
                        if str(acct_name) in BS_SUBTOTAL_EXCLUDE: ev_tag = ''
                        
                        raw_bs_rows.append({
                            'Company': company_name, 'Ticker': ticker, 'Period': latest.strftime('%Y-%m-%d'),
                            'Currency': currency, 'Account': str(acct_name), 'EV_Tag': ev_tag, 'Amount_M': val_f/1e6
                        })
                        if ev_tag: gpcm[ev_tag] += val_f/1e6
            
            gpcm['Shares'] = bs_shares if bs_shares else float(info.get('sharesOutstanding', 0))

            # [2] Market Cap
            try:
                hist = stock.history(start=(base_dt - timedelta(days=10)).strftime('%Y-%m-%d'), end=(base_dt + timedelta(days=1)).strftime('%Y-%m-%d'), auto_adjust=False)
                close = float(hist['Close'].iloc[-1]) if (not hist.empty and 'Close' in hist.columns) else 0.0
                p_date = hist.index[-1].strftime('%Y-%m-%d') if not hist.empty else '-'
            except: close=0.0; p_date='-'
            gpcm['Close'] = close
            gpcm['Market_Cap_M'] = (close * gpcm['Shares'] / 1e6) if gpcm['Shares'] else 0.0
            market_rows.append({
                'Company': company_name, 'Ticker': ticker, 'Base_Date': base_date_str, 'Price_Date': p_date,
                'Currency': currency, 'Close': close, 'Shares': gpcm['Shares'], 'Market_Cap_M': round(gpcm['Market_Cap_M'], 1)
            })

            # [3] PL
            q_is = stock.quarterly_income_stmt
            q_valid = []
            if q_is is not None and not q_is.empty:
                q_valid = sorted([d for d in q_is.columns if d <= base_dt + timedelta(days=7)], reverse=True)[:4]
            
            is_complete = False
            if len(q_valid) == 4:
                if 'Total Revenue' in q_is.index:
                    vals = q_is.loc['Total Revenue', q_valid]
                    if vals.notna().all() and (vals != 0).all(): is_complete = True
            
            pl_source = 'Quarterly (4Q Sum)' if is_complete else 'Annual'
            pl_data = None; pl_dates = []
            if is_complete: pl_data = q_is; pl_dates = q_valid
            else:
                a_is = stock.income_stmt
                if a_is is not None and not a_is.empty:
                    valid_a = sorted([d for d in a_is.columns if d <= base_dt + timedelta(days=7)], reverse=True)
                    if valid_a: pl_dates = [valid_a[0]]; pl_data = a_is
            
            gpcm['PL_Source'] = pl_source
            calc_sums = {'Revenue': 0, 'OpIncome': 0, 'EBIT_yf': 0, 'EBITDA_yf': 0, 'NormEBITDA_yf': 0, 'NI_Parent': 0}

            if pl_data is not None:
                for acct in pl_data.index:
                    acct_str = str(acct)
                    hl_tag = PL_HIGHLIGHT_MAP.get(acct_str, '')
                    calc_key = PL_CALC_KEY.get(acct_str, '')
                    is_eps = 'EPS' in acct_str or 'Per Share' in acct_str
                    is_shares = 'Average Shares' in acct_str
                    for i, d in enumerate(pl_dates):
                        val = pl_data.loc[acct, d]
                        if pd.isna(val): continue
                        val_f = float(val)
                        if is_eps: unit = 'per share'; amt = val_f
                        elif is_shares: unit = 'M shares'; amt = val_f/1e6
                        else: unit = 'M'; amt = val_f/1e6
                        raw_pl_rows.append({
                            'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                            'Account': acct_str, 'GPCM_Tag': hl_tag, 'PL_Source': pl_source,
                            'Q_Label': f"D-{i}Q" if pl_source.startswith('Quarterly') else 'Annual', 
                            'Period': d.strftime('%Y-%m-%d'), 'Amount_M': amt, 'Unit': unit, '_sort': PL_SORT.get(acct_str, 500)
                        })
                        if calc_key and not is_eps and not is_shares: calc_sums[calc_key] += val_f/1e6
                
                gpcm['Revenue'] = calc_sums['Revenue']
                gpcm['EBIT'] = calc_sums['OpIncome']
                ebitda_yf = calc_sums['EBITDA_yf'] if calc_sums['EBITDA_yf'] != 0 else calc_sums['NormEBITDA_yf']
                ebit_yf = calc_sums['EBIT_yf']
                da_amount = (ebitda_yf - ebit_yf) if (ebitda_yf != 0 and ebit_yf != 0) else 0
                gpcm['EBITDA'] = calc_sums['OpIncome'] + da_amount
                gpcm['NI_Parent'] = calc_sums['NI_Parent']

            # [4] Beta Calculation with Retry Logic
            exchange, market_idx = get_market_index(ticker)
            gpcm['Exchange'] = exchange
            gpcm['Market_Index'] = market_idx

            try:
                # [4-1] 5년 주가 데이터 수집 (월말 종가)
                start_5y = (base_dt - timedelta(days=365*5+20)).strftime('%Y-%m-%d')
                end_date = base_dt.strftime('%Y-%m-%d')

                # ⭐ 주가 데이터는 stock.history() 재사용 (이미 yfinance 객체 생성됨)
                stock_data_5y = None
                market_data_5y = None

                try:
                    # 주가 데이터: 이미 생성된 stock 객체 사용
                    stock_hist_5y = stock.history(start=start_5y, end=end_date, auto_adjust=False)
                    if not stock_hist_5y.empty:
                        stock_data_5y = stock_hist_5y

                    # 시장 지수 데이터: yfinance 사용
                    market_hist_5y = yf.download(market_idx, start=start_5y, end=end_date, progress=False)
                    if not market_hist_5y.empty:
                        market_data_5y = market_hist_5y

                except Exception as e:
                    st.warning(f"{ticker} 5Y 데이터 수집 실패: {e}")
                    pass

                # 월말 종가 저장 (엑셀에서 베타 계산에 사용)
                if stock_data_5y is not None and market_data_5y is not None:
                    if not stock_data_5y.empty and not market_data_5y.empty:
                        # Close 컬럼 추출
                        if isinstance(stock_data_5y, pd.DataFrame):
                            stock_prices_5y = stock_data_5y['Close'] if 'Close' in stock_data_5y.columns else stock_data_5y.iloc[:, 0]
                        else:
                            stock_prices_5y = stock_data_5y

                        if isinstance(market_data_5y, pd.DataFrame):
                            market_prices_5y = market_data_5y['Close'] if 'Close' in market_data_5y.columns else market_data_5y.iloc[:, 0]
                        else:
                            market_prices_5y = market_data_5y

                        # 인덱스 확인 및 변환 (timezone-naive DatetimeIndex)
                        if not isinstance(stock_prices_5y.index, pd.DatetimeIndex):
                            stock_prices_5y.index = pd.to_datetime(stock_prices_5y.index)
                        if stock_prices_5y.index.tz is not None:
                            stock_prices_5y.index = stock_prices_5y.index.tz_localize(None)

                        if not isinstance(market_prices_5y.index, pd.DatetimeIndex):
                            market_prices_5y.index = pd.to_datetime(market_prices_5y.index)
                        if market_prices_5y.index.tz is not None:
                            market_prices_5y.index = market_prices_5y.index.tz_localize(None)

                        # 월말 종가 저장
                        stock_monthly_prices = stock_prices_5y.resample('ME').last().dropna()
                        market_monthly_prices = market_prices_5y.resample('ME').last().dropna()

                        # 데이터 저장 (엑셀에서 사용)
                        if len(stock_monthly_prices) >= 12 and len(market_monthly_prices) >= 12:  # 최소 12개월 데이터 필요
                            gpcm['Stock_Monthly_Prices_5Y'] = stock_monthly_prices
                            gpcm['Market_Monthly_Prices_5Y'] = market_monthly_prices
                        else:
                            st.warning(f"{ticker}: 월별 데이터가 부족합니다 (Stock: {len(stock_monthly_prices)}, Market: {len(market_monthly_prices)})")
                            gpcm['Stock_Monthly_Prices_5Y'] = None
                            gpcm['Market_Monthly_Prices_5Y'] = None

                        # 베타는 엑셀에서 계산하므로 None으로 설정
                        gpcm['Beta_5Y_Monthly_Raw'] = None
                        gpcm['Beta_5Y_Monthly_Adj'] = None
                    else:
                        gpcm['Stock_Monthly_Prices_5Y'] = None
                        gpcm['Market_Monthly_Prices_5Y'] = None
                        gpcm['Beta_5Y_Monthly_Raw'] = None
                        gpcm['Beta_5Y_Monthly_Adj'] = None
                else:
                    gpcm['Stock_Monthly_Prices_5Y'] = None
                    gpcm['Market_Monthly_Prices_5Y'] = None
                    gpcm['Beta_5Y_Monthly_Raw'] = None
                    gpcm['Beta_5Y_Monthly_Adj'] = None

                # [4-2] 2년 주간 베타 데이터 수집
                start_2y = (base_dt - timedelta(days=365*2+20)).strftime('%Y-%m-%d')
                stock_data_2y = None
                market_data_2y = None

                try:
                    # 주가 데이터: 이미 생성된 stock 객체 사용
                    stock_hist_2y = stock.history(start=start_2y, end=end_date, auto_adjust=False)
                    if not stock_hist_2y.empty:
                        stock_data_2y = stock_hist_2y

                    # 시장 지수 데이터: yfinance 사용
                    market_hist_2y = yf.download(market_idx, start=start_2y, end=end_date, progress=False)
                    if not market_hist_2y.empty:
                        market_data_2y = market_hist_2y

                except Exception as e:
                    st.warning(f"{ticker} 2Y 데이터 수집 실패: {e}")
                    pass

                # 주간 종가 저장
                if stock_data_2y is not None and market_data_2y is not None:
                    if not stock_data_2y.empty and not market_data_2y.empty:
                        # Close 컬럼 추출
                        if isinstance(stock_data_2y, pd.DataFrame):
                            stock_prices_2y = stock_data_2y['Close'] if 'Close' in stock_data_2y.columns else stock_data_2y.iloc[:, 0]
                        else:
                            stock_prices_2y = stock_data_2y

                        if isinstance(market_data_2y, pd.DataFrame):
                            market_prices_2y = market_data_2y['Close'] if 'Close' in market_data_2y.columns else market_data_2y.iloc[:, 0]
                        else:
                            market_prices_2y = market_data_2y

                        # 인덱스 확인 및 변환
                        if not isinstance(stock_prices_2y.index, pd.DatetimeIndex):
                            stock_prices_2y.index = pd.to_datetime(stock_prices_2y.index)
                        if stock_prices_2y.index.tz is not None:
                            stock_prices_2y.index = stock_prices_2y.index.tz_localize(None)

                        if not isinstance(market_prices_2y.index, pd.DatetimeIndex):
                            market_prices_2y.index = pd.to_datetime(market_prices_2y.index)
                        if market_prices_2y.index.tz is not None:
                            market_prices_2y.index = market_prices_2y.index.tz_localize(None)

                        # 주말 종가 저장 (W-FRI: 금요일 기준)
                        stock_weekly_prices = stock_prices_2y.resample('W-FRI').last().dropna()
                        market_weekly_prices = market_prices_2y.resample('W-FRI').last().dropna()

                        if len(stock_weekly_prices) >= 50 and len(market_weekly_prices) >= 50:  # 최소 50주 데이터 필요
                            gpcm['Stock_Weekly_Prices_2Y'] = stock_weekly_prices
                            gpcm['Market_Weekly_Prices_2Y'] = market_weekly_prices
                        else:
                            gpcm['Stock_Weekly_Prices_2Y'] = None
                            gpcm['Market_Weekly_Prices_2Y'] = None
                    else:
                        gpcm['Stock_Weekly_Prices_2Y'] = None
                        gpcm['Market_Weekly_Prices_2Y'] = None
                else:
                    gpcm['Stock_Weekly_Prices_2Y'] = None
                    gpcm['Market_Weekly_Prices_2Y'] = None

                gpcm['Beta_2Y_Weekly_Raw'] = None
                gpcm['Beta_2Y_Weekly_Adj'] = None

            except Exception as e:
                st.warning(f"Price data collection failed for {ticker}: {e}")
                gpcm['Stock_Monthly_Prices_5Y'] = None
                gpcm['Market_Monthly_Prices_5Y'] = None
                gpcm['Beta_5Y_Monthly_Raw'] = None
                gpcm['Beta_5Y_Monthly_Adj'] = None
                gpcm['Beta_2Y_Weekly_Raw'] = None
                gpcm['Beta_2Y_Weekly_Adj'] = None

            # [5] Tax Rate Calculation
            # Wikipedia에서 법인세율 가져오기
            tax_rates_wiki = get_corporate_tax_rates_from_wikipedia()
            country_code = get_country_from_ticker(ticker)

            if pl_data is not None and 'Pretax Income' in pl_data.index:
                pretax_vals = []
                for d in pl_dates:
                    val = pl_data.loc['Pretax Income', d]
                    if pd.notna(val):
                        pretax_vals.append(float(val) / 1e6)
                if pretax_vals:
                    gpcm['Pretax_Income'] = sum(pretax_vals)

            # 세율 결정
            if country_code == 'KR':
                # 한국: 세전순이익 기반 한계세율 (지방세 포함)
                gpcm['Tax_Rate'] = get_korean_marginal_tax_rate(gpcm['Pretax_Income'])
            else:
                # 기타 국가: Wikipedia에서 크롤링한 세율
                gpcm['Tax_Rate'] = tax_rates_wiki.get(country_code, 0.25)

            # [6] 부채비율 계산 (개별 peer용)
            # 실제 부채비율 계산: IBD / (IBD + Market Cap + NCI)
            total_value = gpcm['IBD'] + gpcm['Market_Cap_M'] + gpcm['NCI']
            if total_value > 0:
                gpcm['Debt_Ratio'] = gpcm['IBD'] / total_value
            else:
                gpcm['Debt_Ratio'] = 0

            gpcm_data[ticker] = gpcm

        except Exception as e:
            st.error(f"Error fetching {ticker}: {e}")
            continue

    # ========================================
    # [7] Target WACC 계산 (별도 시트용)
    # ========================================
    status_text.text("Target WACC 계산 중...")

    # 7-1. 평균 부채비율 계산
    debt_ratios = [gpcm['Debt_Ratio'] for gpcm in gpcm_data.values() if gpcm['Debt_Ratio'] > 0]
    if debt_ratios:
        avg_debt_ratio = np.mean(debt_ratios)
        st.info(f"📊 피어 평균 부채비율 (D/V): {avg_debt_ratio*100:.1f}%")
    else:
        avg_debt_ratio = 0.30
        st.warning(f"⚠️ 부채비율 계산 불가. 기본값 {avg_debt_ratio*100:.0f}% 사용")

    # 7-2. Unlevered Beta 평균 계산 (선택된 beta_type에 따라)
    unlevered_betas = []
    beta_field = 'Unlevered_Beta_5Y' if beta_type == '5Y' else 'Unlevered_Beta_2Y'
    beta_label = "5Y Monthly" if beta_type == '5Y' else "2Y Weekly"

    for gpcm in gpcm_data.values():
        if gpcm[beta_field] is not None and gpcm[beta_field] > 0:
            unlevered_betas.append(gpcm[beta_field])

    if unlevered_betas:
        avg_unlevered_beta = np.mean(unlevered_betas)
        st.info(f"📊 피어 평균 Unlevered Beta ({beta_label}): {avg_unlevered_beta:.4f}")
    else:
        avg_unlevered_beta = 1.0
        st.warning(f"⚠️ Unlevered Beta ({beta_label}) 계산 불가. 기본값 {avg_unlevered_beta:.2f} 사용")

    # 7-3. Target의 Relevered Beta 계산
    # Relevered Beta = Unlevered Beta × (1 + (1 - Tax Rate) × (D/E))
    # D/E = D/V ÷ (1 - D/V)
    if avg_debt_ratio < 1.0:
        target_de_ratio = avg_debt_ratio / (1 - avg_debt_ratio)
        target_relevered_beta = avg_unlevered_beta * (1 + (1 - target_tax_rate) * target_de_ratio)
    else:
        target_relevered_beta = avg_unlevered_beta

    # 7-4. Target의 자기자본비용 (Ke) 계산
    # Ke = Rf + MRP × Relevered Beta + Size Premium
    target_ke = rf_rate + mrp * target_relevered_beta + size_premium

    # 7-5. Target의 세후 타인자본비용 (Kd_aftertax) 계산
    # Kd_aftertax = Kd_pretax × (1 - Tax Rate)
    target_kd_aftertax = kd_pretax * (1 - target_tax_rate)

    # 7-6. Target의 WACC 계산
    # WACC = (E/V) × Ke + (D/V) × Kd_aftertax
    equity_weight = 1 - avg_debt_ratio
    debt_weight = avg_debt_ratio
    target_wacc = equity_weight * target_ke + debt_weight * target_kd_aftertax

    # WACC 계산 결과 저장
    target_wacc_data = {
        'Rf': rf_rate,
        'MRP': mrp,
        'Size_Premium': size_premium,
        'Avg_Unlevered_Beta': avg_unlevered_beta,
        'Target_Tax_Rate': target_tax_rate,
        'Avg_Debt_Ratio': avg_debt_ratio,
        'Target_DE_Ratio': target_de_ratio if avg_debt_ratio < 1.0 else 0,
        'Target_Relevered_Beta': target_relevered_beta,
        'Target_Ke': target_ke,
        'Kd_Pretax': kd_pretax,
        'Target_Kd_Aftertax': target_kd_aftertax,
        'Equity_Weight': equity_weight,
        'Debt_Weight': debt_weight,
        'Target_WACC': target_wacc
    }

    st.success(f"✅ Target WACC: {target_wacc*100:.2f}%")

    status_text.text("Data collection complete!")
    return gpcm_data, raw_bs_rows, raw_pl_rows, market_rows, price_abs_dfs, price_rel_dfs, ticker_to_name, avg_debt_ratio, target_wacc_data


def create_excel(gpcm_data, raw_bs_rows, raw_pl_rows, market_rows, price_abs_dfs, price_rel_dfs, base_date_str, ticker_to_name, target_wacc_data, beta_type="5Y"):
    output = io.BytesIO()
    wb = Workbook(); wb.remove(wb.active)

    # Styles & Colors (v17)
    C_BL='00338D'; C_DB='1E2A5E'; C_MB='005EB8'; C_LB='C3D7EE'; C_PB='E8EFF8'
    C_DG='333333'; C_MG='666666'; C_LG='F5F5F5'; C_BG='B0B0B0'; C_W='FFFFFF'
    C_GR='E2EFDA'; C_YL='FFF8E1'; C_NOA='FCE4EC'

    S1=Side(style='thin',color=C_BG); BD=Border(left=S1,right=S1,top=S1,bottom=S1)
    fT=Font(name='Arial',bold=True,size=14,color=C_BL); fS=Font(name='Arial',size=9,color=C_MG,italic=True)
    fH=Font(name='Arial',bold=True,size=9,color=C_W); fA=Font(name='Arial',size=9,color=C_DG)
    fHL=Font(name='Arial',bold=True,size=9,color=C_DB); fSEC=Font(name='Arial',bold=True,size=10,color=C_W)
    fMUL=Font(name='Arial',bold=True,size=10,color=C_BL); fNOTE=Font(name='Arial',size=8,color=C_MG,italic=True)
    fSTAT=Font(name='Arial',bold=True,size=9,color=C_DB)
    fFRM=Font(name='Arial',size=9,color='000000'); fFRM_B=Font(name='Arial',bold=True,size=9,color='000000')
    fLINK=Font(name='Arial',size=9,color='008000'); fLINK_B=Font(name='Arial',bold=True,size=9,color='008000')

    pH=PatternFill('solid',fgColor=C_BL); pW=PatternFill('solid',fgColor=C_W)
    pST=PatternFill('solid',fgColor=C_LG); pSEC=PatternFill('solid',fgColor=C_DB)
    pSTAT=PatternFill('solid',fgColor=C_LB)

    ev_fills = {'Cash':PatternFill('solid',fgColor=C_GR), 'IBD':PatternFill('solid',fgColor=C_YL),
                'NCI':PatternFill('solid',fgColor=C_PB), 'NOA(Option)':PatternFill('solid',fgColor=C_NOA),
                'Equity':PatternFill('solid',fgColor=C_LB), 'PL_HL':PatternFill('solid',fgColor=C_YL),
                'NI_Parent':PatternFill('solid',fgColor=C_YL)}
    
    def sc(c,fo=None,fi=None,al=None,bd=None,nf=None):
        if fo:c.font=fo
        if fi:c.fill=fi
        if al:c.alignment=al
        if bd:c.border=bd
        if nf:c.number_format=nf
    
    aC=Alignment(horizontal='center',vertical='center',wrap_text=True)
    aL=Alignment(horizontal='left',vertical='center',indent=1)
    aR=Alignment(horizontal='right',vertical='center')
    NF_M='#,##0;(#,##0);"-"'; NF_M1='#,##0.0;(#,##0.0);"-"'; NF_PRC='#,##0.00;(#,##0.00);"-"'
    NF_INT='#,##0;(#,##0);"-"'; NF_EPS='#,##0.00;(#,##0.00);"-"'; NF_X='0.0"x";(0.0"x");"-"'

    # [Sheet 1] BS_Full
    ws_bs = wb.create_sheet('BS_Full')
    if raw_bs_rows:
        df_bs = pd.DataFrame(raw_bs_rows)
        cols = [('Company','Company',18),('Ticker','Ticker',10), ('Period','Period',12),('Currency','Curr',6), ('Account','Account',42),('EV_Tag','EV Tag',14), ('Amount_M','Amount (M)',18)]
        ws_bs.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_bs.cell(1,1,'Balance Sheet Full'),fo=fT)
        r=3
        for i,(lb,key) in enumerate([('Cash','Cash'),('IBD','IBD'),('NCI','NCI'),('NOA(Option)','NOA(Option)'),('Equity','Equity')]):
            sc(ws_bs.cell(r,i+1,lb), fo=Font(name='Arial',size=8,bold=True),fi=ev_fills[key],al=aC,bd=BD)
        r=5
        for i,(col,disp,w) in enumerate(cols): ws_bs.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_bs.cell(r,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
        hdr=r; r+=1
        for _,rd in df_bs.iterrows():
            ev_tag=rd['EV_Tag']; is_hl=bool(ev_tag)
            row_fi=ev_fills.get(ev_tag, pST if r%2==0 else pW) if is_hl else (pST if r%2==0 else pW)
            row_font=fHL if is_hl else fA
            for i,(col,_,_) in enumerate(cols):
                c=ws_bs.cell(r,i+1); v=rd[col]
                if isinstance(v,(float,np.floating)): c.value=round(v,1) if pd.notna(v) else None
                else: c.value=v
                sc(c,fo=row_font,fi=row_fi,al=aR if col=='Amount_M' else aL,bd=BD,nf=NF_M if col=='Amount_M' else None)
            r+=1
        ws_bs.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r-1}"; ws_bs.freeze_panes=f'A{hdr+1}'

    # [Sheet 2] PL_Data
    ws_pl = wb.create_sheet('PL_Data')
    if raw_pl_rows:
        df_pl = pd.DataFrame(raw_pl_rows).sort_values(['Company','_sort','Q_Label'])
        cols = [('Company','Company',18),('Ticker','Ticker',10), ('Currency','Curr',6),('Account','Account',42), ('GPCM_Tag','GPCM Tag',14),('Unit','Unit',10), ('PL_Source','Source',14),('Q_Label','Q Label',9), ('Period','Period',12),('Amount_M','Amount',18)]
        ws_pl.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_pl.cell(1,1,'Income Statement'),fo=fT)
        r=5
        for i,(col,disp,w) in enumerate(cols): ws_pl.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_pl.cell(r,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
        hdr=r; r+=1
        for _,rd in df_pl.iterrows():
            is_hl=bool(rd['GPCM_Tag']); row_fi=ev_fills['PL_HL'] if is_hl else (pST if r%2==0 else pW)
            row_font=fHL if is_hl else fA
            for i,(col,_,_) in enumerate(cols):
                if col=='_sort': continue
                c=ws_pl.cell(r,i+1); v=rd[col]
                if isinstance(v,(float,np.floating)): c.value=round(v,1) if pd.notna(v) else None
                else: c.value=v
                is_eps=rd.get('Unit','')=='per share'; nf=NF_EPS if is_eps else NF_M
                sc(c,fo=row_font,fi=row_fi,al=aR if col=='Amount_M' else aL,bd=BD,nf=nf if col=='Amount_M' else None)
            r+=1
        ws_pl.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r-1}"; ws_pl.freeze_panes=f'A{hdr+1}'

    # [Sheet 3] Market_Cap
    ws_mc = wb.create_sheet('Market_Cap')
    if market_rows:
        df_mkt = pd.DataFrame(market_rows)
        cols = [('Company','Company',18),('Ticker','Ticker',10), ('Base_Date','Base Date',12),('Price_Date','Price Date',12), ('Currency','Curr',6),('Close','Close Price',14), ('Shares','Shares (Ord.)',18),('Market_Cap_M','Mkt Cap (M)',20)]
        ws_mc.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_mc.cell(1,1,'Market Capitalization'),fo=fT)
        ws_mc.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(cols)); sc(ws_mc.cell(2,1,'Mkt Cap = Ordinary Shares Number (자기주식 차감) × Close Price (auto_adjust=False)'), fo=fS)
        r=4
        for i,(col,disp,w) in enumerate(cols): ws_mc.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_mc.cell(r,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
        mc_hdr=r; r+=1
        MC_DATA_START=r
        for _,rd in df_mkt.iterrows():
            ev=(r%2==0)
            for i,(col,_,_) in enumerate(cols):
                c=ws_mc.cell(r,i+1); v=rd[col]
                if isinstance(v,(float,np.floating)): c.value=round(v,2) if pd.notna(v) else None
                else: c.value=v
                nf=NF_PRC if col=='Close' else (NF_INT if col=='Shares' else (NF_M1 if col=='Market_Cap_M' else None))
                sc(c,fo=fA,fi=pST if ev else pW,al=aR if nf else aL,bd=BD,nf=nf)
            r+=1
        ws_mc.auto_filter.ref=f"A{mc_hdr}:{get_column_letter(len(cols))}{r-1}"; ws_mc.freeze_panes=f'A{mc_hdr+1}'

    # [Sheet 3.5] Beta_Calculation (엑셀 수식으로 베타 계산)
    ws_beta = wb.create_sheet('Beta_Calculation')
    wb.move_sheet('Beta_Calculation', offset=-3)  # 맨 앞으로 (Sheet 위치: 0)

    # 제목
    ws_beta.merge_cells('A1:F1')
    sc(ws_beta['A1'], fo=Font(name='Arial', bold=True, size=14, color=C_BL))
    ws_beta['A1'] = 'Beta Calculation (Excel Formulas)'

    ws_beta.merge_cells('A2:F2')
    sc(ws_beta['A2'], fo=Font(name='Arial', size=9, color=C_MG, italic=True))
    ws_beta['A2'] = f'5-Year Monthly & 2-Year Weekly Returns | Base: {base_date_str}'

    r_beta = 4

    # 각 ticker별로 베타 계산 섹션 생성
    beta_result_rows = {}  # ticker: (raw_5y, adj_5y, raw_2y, adj_2y) 매핑

    for idx, (ticker, gpcm) in enumerate(gpcm_data.items()):
        # 회사 정보
        company_name = gpcm['Company']
        market_idx = gpcm['Market_Index']

        # ========== 5Y Monthly Beta Section ==========
        ws_beta.merge_cells(f'A{r_beta}:F{r_beta}')
        sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=10, color=C_W),
           fi=PatternFill('solid', fgColor='607D8B'), al=Alignment(horizontal='center'))
        ws_beta.cell(r_beta, 1, f'{company_name} ({ticker}) vs {market_idx} - 5Y Monthly')
        r_beta += 1

        # 5Y 데이터 확인
        stock_prices_5y = gpcm.get('Stock_Monthly_Prices_5Y')
        market_prices_5y = gpcm.get('Market_Monthly_Prices_5Y')
        raw_5y_row = None
        adj_5y_row = None

        if stock_prices_5y is not None and market_prices_5y is not None and not stock_prices_5y.empty and not market_prices_5y.empty:
            # 헤더
            ws_beta.cell(r_beta, 1, 'Date')
            ws_beta.cell(r_beta, 2, f'{ticker} Price')
            ws_beta.cell(r_beta, 3, f'{market_idx} Price')
            ws_beta.cell(r_beta, 4, f'{ticker} Return')
            ws_beta.cell(r_beta, 5, f'{market_idx} Return')
            for col in range(1, 6):
                sc(ws_beta.cell(r_beta, col), fo=Font(name='Arial', bold=True, size=9, color=C_W),
                   fi=PatternFill('solid', fgColor=C_BL), al=Alignment(horizontal='center'), bd=BD)
            r_beta += 1

            data_start_row = r_beta

            # 공통 날짜 인덱스
            common_dates = stock_prices_5y.index.intersection(market_prices_5y.index)

            # 데이터 행 작성
            for date in common_dates:
                ws_beta.cell(r_beta, 1, date.strftime('%Y-%m'))
                ws_beta.cell(r_beta, 2, float(stock_prices_5y.loc[date]))
                ws_beta.cell(r_beta, 3, float(market_prices_5y.loc[date]))

                # 수익률 계산 (엑셀 수식)
                if r_beta > data_start_row:
                    ws_beta.cell(r_beta, 4).value = f'=(B{r_beta}-B{r_beta-1})/B{r_beta-1}'
                    ws_beta.cell(r_beta, 5).value = f'=(C{r_beta}-C{r_beta-1})/C{r_beta-1}'
                else:
                    ws_beta.cell(r_beta, 4, None)
                    ws_beta.cell(r_beta, 5, None)

                # 스타일
                sc(ws_beta.cell(r_beta, 1), fo=fA, al=aC, bd=BD)
                sc(ws_beta.cell(r_beta, 2), fo=fA, al=aR, bd=BD, nf='#,##0.00')
                sc(ws_beta.cell(r_beta, 3), fo=fA, al=aR, bd=BD, nf='#,##0.00')
                sc(ws_beta.cell(r_beta, 4), fo=fA, al=aR, bd=BD, nf='0.00%')
                sc(ws_beta.cell(r_beta, 5), fo=fA, al=aR, bd=BD, nf='0.00%')

                r_beta += 1

            data_end_row = r_beta - 1

            # 베타 계산 (SLOPE 함수)
            r_beta += 1
            ws_beta.cell(r_beta, 1, 'Raw Beta (5Y Monthly)')
            ws_beta.cell(r_beta, 2).value = f'=SLOPE(D{data_start_row+1}:D{data_end_row},E{data_start_row+1}:E{data_end_row})'
            sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=9), bd=BD)
            sc(ws_beta.cell(r_beta, 2), fo=Font(name='Arial', bold=True, size=9), fi=PatternFill('solid', fgColor='E8F5E9'),
               bd=BD, al=aR, nf='0.0000')
            raw_5y_row = r_beta
            r_beta += 1

            # Adjusted Beta
            ws_beta.cell(r_beta, 1, 'Adjusted Beta (5Y)')
            ws_beta.cell(r_beta, 2).value = f'=2/3*B{r_beta-1}+1/3*1'
            sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=9), bd=BD)
            sc(ws_beta.cell(r_beta, 2), fo=Font(name='Arial', bold=True, size=9), fi=PatternFill('solid', fgColor='E8F5E9'),
               bd=BD, al=aR, nf='0.0000')
            adj_5y_row = r_beta

        else:
            ws_beta.cell(r_beta, 1, 'No 5Y price data available')
            sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', size=9, color='FF0000'))

        r_beta += 2  # 간격

        # ========== 2Y Weekly Beta Section ==========
        ws_beta.merge_cells(f'A{r_beta}:F{r_beta}')
        sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=10, color=C_W),
           fi=PatternFill('solid', fgColor='455A64'), al=Alignment(horizontal='center'))
        ws_beta.cell(r_beta, 1, f'{company_name} ({ticker}) vs {market_idx} - 2Y Weekly')
        r_beta += 1

        # 2Y 데이터 확인
        stock_prices_2y = gpcm.get('Stock_Weekly_Prices_2Y')
        market_prices_2y = gpcm.get('Market_Weekly_Prices_2Y')
        raw_2y_row = None
        adj_2y_row = None

        if stock_prices_2y is not None and market_prices_2y is not None and not stock_prices_2y.empty and not market_prices_2y.empty:
            # 헤더
            ws_beta.cell(r_beta, 1, 'Date')
            ws_beta.cell(r_beta, 2, f'{ticker} Price')
            ws_beta.cell(r_beta, 3, f'{market_idx} Price')
            ws_beta.cell(r_beta, 4, f'{ticker} Return')
            ws_beta.cell(r_beta, 5, f'{market_idx} Return')
            for col in range(1, 6):
                sc(ws_beta.cell(r_beta, col), fo=Font(name='Arial', bold=True, size=9, color=C_W),
                   fi=PatternFill('solid', fgColor=C_BL), al=Alignment(horizontal='center'), bd=BD)
            r_beta += 1

            data_start_row = r_beta

            # 공통 날짜 인덱스
            common_dates = stock_prices_2y.index.intersection(market_prices_2y.index)

            # 데이터 행 작성
            for date in common_dates:
                ws_beta.cell(r_beta, 1, date.strftime('%Y-%m-%d'))
                ws_beta.cell(r_beta, 2, float(stock_prices_2y.loc[date]))
                ws_beta.cell(r_beta, 3, float(market_prices_2y.loc[date]))

                # 수익률 계산 (엑셀 수식)
                if r_beta > data_start_row:
                    ws_beta.cell(r_beta, 4).value = f'=(B{r_beta}-B{r_beta-1})/B{r_beta-1}'
                    ws_beta.cell(r_beta, 5).value = f'=(C{r_beta}-C{r_beta-1})/C{r_beta-1}'
                else:
                    ws_beta.cell(r_beta, 4, None)
                    ws_beta.cell(r_beta, 5, None)

                # 스타일
                sc(ws_beta.cell(r_beta, 1), fo=fA, al=aC, bd=BD)
                sc(ws_beta.cell(r_beta, 2), fo=fA, al=aR, bd=BD, nf='#,##0.00')
                sc(ws_beta.cell(r_beta, 3), fo=fA, al=aR, bd=BD, nf='#,##0.00')
                sc(ws_beta.cell(r_beta, 4), fo=fA, al=aR, bd=BD, nf='0.00%')
                sc(ws_beta.cell(r_beta, 5), fo=fA, al=aR, bd=BD, nf='0.00%')

                r_beta += 1

            data_end_row = r_beta - 1

            # 베타 계산 (SLOPE 함수)
            r_beta += 1
            ws_beta.cell(r_beta, 1, 'Raw Beta (2Y Weekly)')
            ws_beta.cell(r_beta, 2).value = f'=SLOPE(D{data_start_row+1}:D{data_end_row},E{data_start_row+1}:E{data_end_row})'
            sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=9), bd=BD)
            sc(ws_beta.cell(r_beta, 2), fo=Font(name='Arial', bold=True, size=9), fi=PatternFill('solid', fgColor='FFF9C4'),
               bd=BD, al=aR, nf='0.0000')
            raw_2y_row = r_beta
            r_beta += 1

            # Adjusted Beta
            ws_beta.cell(r_beta, 1, 'Adjusted Beta (2Y)')
            ws_beta.cell(r_beta, 2).value = f'=2/3*B{r_beta-1}+1/3*1'
            sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=9), bd=BD)
            sc(ws_beta.cell(r_beta, 2), fo=Font(name='Arial', bold=True, size=9), fi=PatternFill('solid', fgColor='FFF9C4'),
               bd=BD, al=aR, nf='0.0000')
            adj_2y_row = r_beta

        else:
            ws_beta.cell(r_beta, 1, 'No 2Y price data available')
            sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', size=9, color='FF0000'))

        # 결과 저장 (4개 값)
        beta_result_rows[ticker] = (raw_5y_row, adj_5y_row, raw_2y_row, adj_2y_row)

        r_beta += 2  # 다음 회사와 간격

    ws_beta.column_dimensions['A'].width = 15
    ws_beta.column_dimensions['B'].width = 15
    ws_beta.column_dimensions['C'].width = 15
    ws_beta.column_dimensions['D'].width = 15
    ws_beta.column_dimensions['E'].width = 15

    ws_beta.freeze_panes = 'A4'

    # [Sheet 4] GPCM
    ws = wb.create_sheet('GPCM')
    wb.move_sheet('GPCM', offset=-5)  # 맨 앞으로 (Sheet 위치: 0)
    TOTAL_COLS = 34  # WACC 컬럼 제거 (별도 시트로 분리)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS); sc(ws.cell(1,1,'GPCM Valuation Summary with Beta Analysis'), fo=fT)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=TOTAL_COLS); sc(ws.cell(2,1,f'Base: {base_date_str} | Unit: Millions (local currency) | EV = MCap + IBD − Cash + NCI | Target WACC: See WACC_Calculation Sheet'), fo=fS)

    r=4
    sections = [(1,3,'Company Info'),(4,4,'Other Information'),(8,6,'BS → EV Components'),(14,4,'PL (LTM / Annual)'),(18,3,'Market Data'),(21,5,'Valuation Multiples'),(26,9,'Beta & Risk Analysis')]
    for start,span,txt in sections:
        ws.merge_cells(start_row=r, start_column=start, end_row=r, end_column=start+span-1)
        sc(ws.cell(r,start,txt), fo=fSEC, fi=pSEC, al=aC, bd=BD)
        for c_idx in range(start,start+span): sc(ws.cell(r,c_idx), bd=BD)

    r=5
    headers = ['Company','Ticker','Base Date','Curr','PL Source','Exchange','Mkt Index',
               'Cash','IBD','Net Debt','NCI','Equity','EV',
               'Revenue','EBIT','EBITDA','NI (Parent)',
               'Price','Shares','Mkt Cap',
               'EV/EBITDA','EV/EBIT','PER','PBR','PSR',
               'β 5Y Raw','β 5Y Adj','β 2Y Raw','β 2Y Adj','Pretax Inc','Tax Rate','Debt Ratio','Unlevered β 5Y','Unlevered β 2Y']
    widths = [18,10,11,6,16,10,10,
              14,14,14,12,14,16,
              14,14,14,14,
              12,16,16,
              12,12,10,10,10,
              10,10,10,10,14,9,10,12,12]
    for i,(h,w) in enumerate(zip(headers,widths)): ws.column_dimensions[get_column_letter(i+1)].width=w; sc(ws.cell(r,i+1,h), fo=fH, fi=pH, al=aC, bd=BD)
    
    DATA_START=6; n_companies=len(gpcm_data); DATA_END=DATA_START+n_companies-1
    NF_BETA='0.00;(0.00);"-"'; NF_PCT='0.0%;(0.0%);"-"'; NF_RATIO='0.00;(0.00);"-"'
    pBETA=PatternFill('solid',fgColor='E8F5E9')
    for idx,(ticker, gpcm) in enumerate(gpcm_data.items()):
        r=DATA_START+idx; mc_row=MC_DATA_START+idx; ev_row=(r%2==0); base_fi=pST if ev_row else pW
        # A-G: Company Info + Other Info
        vals=[gpcm['Company'],ticker,gpcm['Base_Date'],gpcm['Currency'],gpcm['PL_Source'],gpcm['Exchange'],gpcm['Market_Index']]
        for ci,v in enumerate(vals,1): ws.cell(r,ci,v); sc(ws.cell(r,ci), fo=fA, fi=base_fi, al=aL, bd=BD)

        # H-M: BS → EV Components (Formulas)
        ws.cell(r,8).value=f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"Cash")'; sc(ws.cell(r,8), fo=fLINK_B, fi=ev_fills['Cash'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,9).value=f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"IBD")'; sc(ws.cell(r,9), fo=fLINK_B, fi=ev_fills['IBD'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,10).value=f'=I{r}-H{r}'; sc(ws.cell(r,10), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_M)
        ws.cell(r,11).value=f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"NCI")'; sc(ws.cell(r,11), fo=fLINK_B, fi=ev_fills['NCI'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,12).value=f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"Equity")'; sc(ws.cell(r,12), fo=fLINK_B, fi=ev_fills['Equity'], al=aR, bd=BD, nf=NF_M)
        # M (EV)
        ws.cell(r,13).value=f'=T{r}+I{r}-H{r}+K{r}'; sc(ws.cell(r,13), fo=fFRM_B, fi=PatternFill('solid',fgColor=C_PB), al=aR, bd=BD, nf=NF_M)

        # N-Q: PL (LTM/Annual)
        ws.cell(r,14).value=f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$E:$E,"Revenue")'; sc(ws.cell(r,14), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,15).value=f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"Operating Income")'; sc(ws.cell(r,15), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,16).value=f'=O{r}+SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"EBITDA")-SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"EBIT")'; sc(ws.cell(r,16), fo=fFRM_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,17).value=f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$E:$E,"NI_Parent")'; sc(ws.cell(r,17), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)

        # R-T: Market Data
        ws.cell(r,18).value=f'=Market_Cap!F{mc_row}'; sc(ws.cell(r,18), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_PRC)
        ws.cell(r,19).value=f'=Market_Cap!G{mc_row}'; sc(ws.cell(r,19), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_INT)
        ws.cell(r,20).value=f'=Market_Cap!H{mc_row}'; sc(ws.cell(r,20), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_M1)

        # U-Y: Valuation Multiples
        pMULT=PatternFill('solid',fgColor=C_PB)
        ws.cell(r,21).value=f'=IF(P{r}>0,M{r}/P{r},"N/M")'; sc(ws.cell(r,21), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,22).value=f'=IF(O{r}>0,M{r}/O{r},"N/M")'; sc(ws.cell(r,22), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,23).value=f'=IF(Q{r}>0,T{r}/Q{r},"N/M")'; sc(ws.cell(r,23), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,24).value=f'=IF(L{r}>0,T{r}/L{r},"N/M")'; sc(ws.cell(r,24), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,25).value=f'=IF(N{r}>0,T{r}/N{r},"N/M")'; sc(ws.cell(r,25), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

        # Z-AI: Beta & Risk Analysis
        # Beta 값은 Beta_Calculation 시트에서 엑셀 수식으로 참조
        raw_5y_row, adj_5y_row, raw_2y_row, adj_2y_row = beta_result_rows.get(ticker, (None, None, None, None))

        # Beta 5Y Raw - Beta_Calculation 시트 참조
        if raw_5y_row is not None:
            ws.cell(r,26).value = f'=Beta_Calculation!$B${raw_5y_row}'
        else:
            ws.cell(r,26, None)
        sc(ws.cell(r,26), fo=fLINK, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

        # Beta 5Y Adj - Beta_Calculation 시트 참조
        if adj_5y_row is not None:
            ws.cell(r,27).value = f'=Beta_Calculation!$B${adj_5y_row}'
        else:
            ws.cell(r,27, None)
        sc(ws.cell(r,27), fo=fLINK, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

        # Beta 2Y Raw - Beta_Calculation 시트 참조
        if raw_2y_row is not None:
            ws.cell(r,28).value = f'=Beta_Calculation!$B${raw_2y_row}'
        else:
            ws.cell(r,28, None)
        sc(ws.cell(r,28), fo=fLINK, fi=PatternFill('solid',fgColor='FFF9C4'), al=aR, bd=BD, nf=NF_BETA)

        # Beta 2Y Adj - Beta_Calculation 시트 참조
        if adj_2y_row is not None:
            ws.cell(r,29).value = f'=Beta_Calculation!$B${adj_2y_row}'
        else:
            ws.cell(r,29, None)
        sc(ws.cell(r,29), fo=fLINK, fi=PatternFill('solid',fgColor='FFF9C4'), al=aR, bd=BD, nf=NF_BETA)

        # Pretax Income (Formula)
        ws.cell(r,30).value=f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"Pretax Income")'; sc(ws.cell(r,30), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_M)

        # Tax Rate
        ws.cell(r,31,gpcm['Tax_Rate']); sc(ws.cell(r,31), fo=fA, fi=base_fi, al=aR, bd=BD, nf=NF_PCT)

        # Debt Ratio = IBD / (Market Cap + NCI)
        ws.cell(r,32).value=f'=IF((T{r}+K{r})>0,I{r}/(T{r}+K{r}),0)'; sc(ws.cell(r,32), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_RATIO)

        # Unlevered Beta 5Y = Beta 5Y Adj / (1 + (1 - Tax Rate) * Debt Ratio)
        ws.cell(r,33).value=f'=IF(AA{r}>0,AA{r}/(1+(1-AE{r})*AF{r}),AA{r})'; sc(ws.cell(r,33), fo=fFRM_B, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

        # Unlevered Beta 2Y = Beta 2Y Adj / (1 + (1 - Tax Rate) * Debt Ratio)
        ws.cell(r,34).value=f'=IF(AC{r}>0,AC{r}/(1+(1-AE{r})*AF{r}),AC{r})'; sc(ws.cell(r,34), fo=fFRM_B, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

    # Stats
    r=DATA_END+2
    stat_labels=['Mean','Median','Max','Min']; func_map={'Mean':'AVERAGE','Median':'MEDIAN','Max':'MAX','Min':'MIN'}
    # Multiples: 21-25 (EV/EBITDA, EV/EBIT, PER, PBR, PSR)
    # Betas: 26-29, 33-34 (Beta 5Y Raw, Beta 5Y Adj, Beta 2Y Raw, Beta 2Y Adj, Unlevered Beta 5Y, Unlevered Beta 2Y)
    # Ratios: 32 (Debt Ratio)
    mult_cols=[21,22,23,24,25]
    beta_cols=[26,27,28,29,33,34]
    ratio_cols=[32]  # Debt Ratio만

    for sn in stat_labels:
        sc(ws.cell(r,20,sn), fo=fSTAT, fi=pSTAT, al=aC, bd=BD)
        fn=func_map[sn]
        # Multiples
        for ci in mult_cols:
            col=get_column_letter(ci)
            ws.cell(r,ci).value=f'=IFERROR({fn}({col}{DATA_START}:{col}{DATA_END}),"N/M")'
            sc(ws.cell(r,ci), fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_X)
        # Betas
        for ci in beta_cols:
            col=get_column_letter(ci)
            ws.cell(r,ci).value=f'=IFERROR({fn}({col}{DATA_START}:{col}{DATA_END}),"N/M")'
            sc(ws.cell(r,ci), fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_BETA)
        # Ratios
        for ci in ratio_cols:
            col=get_column_letter(ci)
            ws.cell(r,ci).value=f'=IFERROR({fn}({col}{DATA_START}:{col}{DATA_END}),"N/M")'
            sc(ws.cell(r,ci), fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_RATIO)
        r+=1
    
    # Notes
    r+=2
    notes = [
        '[ Valuation Methodology Notes ]',
        f'• Base Date: {base_date_str} | Unit: Millions (local currency)',
        '• EV = Market Cap + IBD − Cash + NCI',
        '• Cash includes: Cash & Cash Equivalents + Other Short-Term Investments',
        '• NOA(Option) in BS_Full: Long-Term Equity Investment, Investment In Financial Assets, Investment Properties etc.',
        '    → Changes in BS_Full EV Tag will automatically update GPCM sheet',
        '• Net Debt = IBD − Cash',
        '• EBIT = Operating Income only',
        '• EBITDA = Operating Income + D&A (D&A = yf_EBITDA - yf_EBIT)',
        '• PER = Market Cap ÷ Net Income Common Stockholders (NI Parent)',
        '• PBR = Market Cap ÷ Stockholders Equity',
        '• PSR = Market Cap ÷ Revenue',
        '• Market Cap = Ordinary Shares Number × Close Price',
        '• PL Source: LTM prioritized',
        '',
        '[ Beta & Risk Analysis ]',
        '• Data Source:',
        '  - 한국 주식 (.KS, .KQ): Yahoo Finance (yfinance)',
        '  - 해외 주식: Yahoo Finance (yfinance)',
        '• Beta 계산 방법:',
        '  - 5Y Monthly Beta: 5년간 월말 종가 기준 월간 수익률 계산 → 시장지수 대비 선형회귀',
        '  - 2Y Weekly Beta: 2년간 주말 종가 기준 주간 수익률 계산 → 시장지수 대비 선형회귀',
        '  - Raw Beta = Slope of linear regression (Market vs Stock returns)',
        '  - Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1.0 (Bloomberg 방법론)',
        '• Market Index: KOSPI (KS11), KOSDAQ (KQ11), Nikkei 225 (^N225), S&P/TSX (^GSPTSE), etc.',
        '• 값 검증: NaN, inf, 극단값(-10 ~ 10 범위 벗어남) 필터링 → None 처리',
        '• Tax Rate: Wikipedia 기반 법인세율; 한국은 한계세율 적용 (지방세 포함, 2025)',
        '   - Korea: ≤ 200M: 9.9% | 200M-20,000M: 20.9% | 20,000M-300,000M: 23.1% | > 300,000M: 26.4%',
        '• Debt Ratio = IBD ÷ (IBD + Market Cap + NCI)',
        '• Unlevered Beta = Levered Beta ÷ (1 + (1 - Tax Rate) × Debt Ratio) [Hamada Model]',
        '• 베타 값은 Python에서 계산되어 엑셀에 저장됩니다 (실시간 데이터 기반)',
        '',
        '[ Target WACC Calculation ]',
        '• Target WACC은 "WACC_Calculation" 시트에서 별도 계산됩니다.',
        '• Ke = Rf + MRP × Relevered Beta + Size Premium',
        '  - Relevered Beta = Avg Unlevered Beta × (1 + (1 - Tax) × Target D/E)',
        '  - Size Premium: 한국공인회계사회 기준 (Micro: 4.02%, Small: 2.56%, Medium: 1.24%, Large: 0%)',
        '• Kd (Aftertax) = Kd (Pretax) × (1 - Target Tax Rate)',
        '• Target D/V = 피어 평균 부채비율 (자동 계산)',
        '• WACC = (E/V) × Ke + (D/V) × Kd (Aftertax)',
        '',
        '• N/M = Not Meaningful (negative or zero)',
        '• All values in GPCM are calculated via Excel Formulas linking to BS_Full and PL_Data sheets.',
        '', '⚠ Data from Yahoo Finance. Verify with official sources.'
    ]
    for note in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_COLS)
        sc(ws.cell(r,1,note), fo=fNOTE); r+=1

    # 틀고정: BS → EV Components 이전 (H6 = Company/Ticker/등 고정, BS부터 스크롤)
    ws.freeze_panes='H6'

    # [Sheet 4.5] WACC_Calculation (Target 기업의 WACC 계산)
    ws_wacc = wb.create_sheet('WACC_Calculation')
    wb.move_sheet('WACC_Calculation', offset=-5)  # GPCM 다음 위치 (Sheet 위치: 1)

    # WACC 시트 제목
    ws_wacc.merge_cells('A1:D1')
    sc(ws_wacc['A1'], fo=Font(name='Arial', bold=True, size=14, color=C_BL))
    ws_wacc['A1'] = 'Target WACC Calculation'

    ws_wacc.merge_cells('A2:D2')
    sc(ws_wacc['A2'], fo=Font(name='Arial', size=9, color=C_MG, italic=True))
    ws_wacc['A2'] = f'Base: {base_date_str} | Peer Average Method'

    # 스타일 정의
    pWACC_PARAM = PatternFill('solid', fgColor='E3F2FD')  # 연한 파란색 (입력값)
    pWACC_CALC = PatternFill('solid', fgColor='FFF9C4')   # 연한 노란색 (계산값)
    pWACC_RESULT = PatternFill('solid', fgColor='FFE082') # 진한 노란색 (최종 WACC)

    r_wacc = 4

    # Section 1: Input Parameters
    ws_wacc.merge_cells(f'A{r_wacc}:D{r_wacc}')
    sc(ws_wacc.cell(r_wacc, 1), fo=Font(name='Arial', bold=True, size=10, color=C_W),
       fi=PatternFill('solid', fgColor=C_MB), al=Alignment(horizontal='center'))
    ws_wacc.cell(r_wacc, 1, '[ 1 ] Input Parameters')
    r_wacc += 1

    # 헤더
    ws_wacc['A' + str(r_wacc)] = 'Parameter'
    ws_wacc['B' + str(r_wacc)] = 'Value'
    ws_wacc['C' + str(r_wacc)] = 'Format'
    ws_wacc['D' + str(r_wacc)] = 'Note'
    for col in ['A', 'B', 'C', 'D']:
        sc(ws_wacc[col + str(r_wacc)], fo=Font(name='Arial', bold=True, size=9, color=C_W),
           fi=PatternFill('solid', fgColor=C_BL), al=Alignment(horizontal='center'), bd=BD)
    r_wacc += 1

    # Calculate GPCM stats row position for formulas
    # DATA_START = 6, DATA_END = 6 + n_companies - 1
    # Mean row = DATA_END + 2
    mean_row = DATA_END + 2

    # 데이터 행 - Input Parameters (외부 조회/사용자 입력이므로 Python 값 유지)
    wacc_params = [
        ('Risk-Free Rate (Rf)', target_wacc_data['Rf'], f"{target_wacc_data['Rf']*100:.2f}%", '10-year Korea Treasury Yield'),
        ('Market Risk Premium (MRP)', target_wacc_data['MRP'], f"{target_wacc_data['MRP']*100:.1f}%", '한국공인회계사회 기준'),
        ('Size Premium', target_wacc_data['Size_Premium'], f"{target_wacc_data['Size_Premium']*100:.2f}%", '한국공인회계사회 (시가총액 기준)'),
        ('Kd (Pretax)', target_wacc_data['Kd_Pretax'], f"{target_wacc_data['Kd_Pretax']*100:.1f}%", '세전 타인자본비용 (사용자 입력)'),
        ('Target Tax Rate', target_wacc_data['Target_Tax_Rate'], f"{target_wacc_data['Target_Tax_Rate']*100:.1f}%", '한국 대기업 기준 (지방세 포함)'),
    ]

    for param, value, formatted, note in wacc_params:
        ws_wacc.cell(r_wacc, 1, param)
        ws_wacc.cell(r_wacc, 2, value)
        ws_wacc.cell(r_wacc, 3, formatted)
        ws_wacc.cell(r_wacc, 4, note)
        sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD, al=Alignment(horizontal='left'))
        sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_PARAM, bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
        sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
        sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
        r_wacc += 1

    r_wacc += 1

    # Section 2: Peer Analysis (GPCM 시트에서 엑셀 수식으로 참조)
    ws_wacc.merge_cells(f'A{r_wacc}:D{r_wacc}')
    sc(ws_wacc.cell(r_wacc, 1), fo=Font(name='Arial', bold=True, size=10, color=C_W),
       fi=PatternFill('solid', fgColor=C_MB), al=Alignment(horizontal='center'))
    ws_wacc.cell(r_wacc, 1, '[ 2 ] Peer Analysis')
    r_wacc += 1

    # 헤더
    ws_wacc['A' + str(r_wacc)] = 'Metric'
    ws_wacc['B' + str(r_wacc)] = 'Value'
    ws_wacc['C' + str(r_wacc)] = 'Format'
    ws_wacc['D' + str(r_wacc)] = 'Note'
    for col in ['A', 'B', 'C', 'D']:
        sc(ws_wacc[col + str(r_wacc)], fo=Font(name='Arial', bold=True, size=9, color=C_W),
           fi=PatternFill('solid', fgColor=C_BL), al=Alignment(horizontal='center'), bd=BD)
    r_wacc += 1

    # Avg Unlevered Beta - 엑셀 수식으로 GPCM 시트 참조 (선택된 beta_type에 따라)
    row_unlevered_beta = r_wacc
    beta_label = "5Y Monthly" if beta_type == "5Y" else "2Y Weekly"
    beta_col = 'AG' if beta_type == "5Y" else 'AH'  # AG = 컬럼 33 (Unlevered Beta 5Y), AH = 컬럼 34 (Unlevered Beta 2Y)
    ws_wacc.cell(r_wacc, 1, f'Avg Unlevered Beta ({beta_label})')
    ws_wacc.cell(r_wacc, 2).value = f'=GPCM!{beta_col}{mean_row}'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Avg_Unlevered_Beta']:.4f}")
    ws_wacc.cell(r_wacc, 4, '피어 평균 (GPCM Mean)')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.0000')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # Avg Debt Ratio - 엑셀 수식으로 GPCM 시트 참조
    row_debt_ratio = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Avg Debt Ratio (D/V)')
    ws_wacc.cell(r_wacc, 2).value = f'=GPCM!AF{mean_row}'  # 컬럼 32 (AF) = Debt Ratio
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Avg_Debt_Ratio']*100:.1f}%")
    ws_wacc.cell(r_wacc, 4, '피어 평균 자본구조 (GPCM Mean)')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # Target D/E Ratio - 엑셀 수식으로 계산
    row_de_ratio = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Target D/E Ratio')
    ws_wacc.cell(r_wacc, 2).value = f'=B{row_debt_ratio}/(1-B{row_debt_ratio})'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Target_DE_Ratio']:.4f}")
    ws_wacc.cell(r_wacc, 4, '= D/V ÷ (1 - D/V)')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.0000')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    r_wacc += 1

    # Section 3: Target WACC Calculation (엑셀 수식으로 계산)
    ws_wacc.merge_cells(f'A{r_wacc}:D{r_wacc}')
    sc(ws_wacc.cell(r_wacc, 1), fo=Font(name='Arial', bold=True, size=10, color=C_W),
       fi=PatternFill('solid', fgColor=C_MB), al=Alignment(horizontal='center'))
    ws_wacc.cell(r_wacc, 1, '[ 3 ] Target WACC Calculation')
    r_wacc += 1

    # 헤더
    ws_wacc['A' + str(r_wacc)] = 'Component'
    ws_wacc['B' + str(r_wacc)] = 'Value'
    ws_wacc['C' + str(r_wacc)] = 'Format'
    ws_wacc['D' + str(r_wacc)] = 'Formula'
    for col in ['A', 'B', 'C', 'D']:
        sc(ws_wacc[col + str(r_wacc)], fo=Font(name='Arial', bold=True, size=9, color=C_W),
           fi=PatternFill('solid', fgColor=C_BL), al=Alignment(horizontal='center'), bd=BD)
    r_wacc += 1

    # Row references for formulas
    row_rf = 6
    row_mrp = 7
    row_size_premium = 8
    row_kd_pretax = 9
    row_tax = 10

    # Relevered Beta
    row_relevered_beta = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Relevered Beta')
    ws_wacc.cell(r_wacc, 2).value = f'=B{row_unlevered_beta}*(1+(1-B{row_tax})*B{row_de_ratio})'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Target_Relevered_Beta']:.4f}")
    ws_wacc.cell(r_wacc, 4, 'Unlevered β × (1 + (1 - Tax) × D/E)')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.0000')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # Ke (Cost of Equity)
    row_ke = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Ke (Cost of Equity)')
    ws_wacc.cell(r_wacc, 2).value = f'=B{row_rf}+B{row_mrp}*B{row_relevered_beta}+B{row_size_premium}'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Target_Ke']*100:.2f}%")
    ws_wacc.cell(r_wacc, 4, 'Rf + MRP × Relevered β + Size Premium')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # Kd (Aftertax)
    row_kd_aftertax = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Kd (Aftertax)')
    ws_wacc.cell(r_wacc, 2).value = f'=B{row_kd_pretax}*(1-B{row_tax})'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Target_Kd_Aftertax']*100:.2f}%")
    ws_wacc.cell(r_wacc, 4, 'Kd (Pretax) × (1 - Tax Rate)')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # Equity Weight (E/V)
    row_equity_weight = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Equity Weight (E/V)')
    ws_wacc.cell(r_wacc, 2).value = f'=1-B{row_debt_ratio}'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Equity_Weight']*100:.1f}%")
    ws_wacc.cell(r_wacc, 4, '1 - Debt Ratio')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # Debt Weight (D/V)
    row_debt_weight = r_wacc
    ws_wacc.cell(r_wacc, 1, 'Debt Weight (D/V)')
    ws_wacc.cell(r_wacc, 2).value = f'=B{row_debt_ratio}'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Debt_Weight']*100:.1f}%")
    ws_wacc.cell(r_wacc, 4, 'Debt Ratio')
    sc(ws_wacc.cell(r_wacc, 1), fo=fA, bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=fA, fi=pWACC_CALC, bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
    sc(ws_wacc.cell(r_wacc, 3), fo=fA, bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG), bd=BD)
    r_wacc += 1

    # 구분선
    ws_wacc.cell(r_wacc, 1, '━━━━━━━━━━━━')
    ws_wacc.cell(r_wacc, 2, None)
    ws_wacc.cell(r_wacc, 3, '━━━━━━━━━━━━')
    ws_wacc.cell(r_wacc, 4, '━━━━━━━━━━━━━━━━━━━━━━━━━━━')
    for col_idx in range(1, 5):
        sc(ws_wacc.cell(r_wacc, col_idx), bd=BD)
    r_wacc += 1

    # WACC (최종 결과)
    row_wacc_final = r_wacc
    ws_wacc.cell(r_wacc, 1, 'WACC')
    ws_wacc.cell(r_wacc, 2).value = f'=B{row_equity_weight}*B{row_ke}+B{row_debt_weight}*B{row_kd_aftertax}'
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Target_WACC']*100:.2f}%")
    ws_wacc.cell(r_wacc, 4, '(E/V) × Ke + (D/V) × Kd (Aftertax)')
    sc(ws_wacc.cell(r_wacc, 1), fo=Font(name='Arial', bold=True, size=10), bd=BD)
    sc(ws_wacc.cell(r_wacc, 2), fo=Font(name='Arial', bold=True, size=10), fi=pWACC_RESULT,
       bd=BD, al=Alignment(horizontal='right'), nf='0.00%')
    sc(ws_wacc.cell(r_wacc, 3), fo=Font(name='Arial', bold=True, size=10), bd=BD, al=Alignment(horizontal='center'))
    sc(ws_wacc.cell(r_wacc, 4), fo=Font(name='Arial', size=8, color=C_MG, italic=True), bd=BD)
    r_wacc += 1

    # 열 너비 조정
    ws_wacc.column_dimensions['A'].width = 25
    ws_wacc.column_dimensions['B'].width = 12
    ws_wacc.column_dimensions['C'].width = 15
    ws_wacc.column_dimensions['D'].width = 40

    ws_wacc.freeze_panes = 'A4'

    # Named Range 정의 (다른 시트에서 참조 가능)
    from openpyxl.workbook.defined_name import DefinedName

    # WACC_Calculation 시트의 주요 값들에 Named Range 할당
    # 셀 주소 계산: r_wacc는 계속 증가하므로, 고정된 위치 사용
    # Input Parameters는 6행부터 시작 (r=5 헤더, r=6~10 데이터)
    # Peer Analysis는 약 14행부터
    # Target WACC는 마지막 행

    wb.defined_names['Target_WACC'] = DefinedName('Target_WACC', attr_text=f"'WACC_Calculation'!$B${r_wacc-1}")
    wb.defined_names['Target_Rf'] = DefinedName('Target_Rf', attr_text="'WACC_Calculation'!$B$6")
    wb.defined_names['Target_MRP'] = DefinedName('Target_MRP', attr_text="'WACC_Calculation'!$B$7")
    wb.defined_names['Target_Size_Premium'] = DefinedName('Target_Size_Premium', attr_text="'WACC_Calculation'!$B$8")
    wb.defined_names['Target_Kd_Pretax'] = DefinedName('Target_Kd_Pretax', attr_text="'WACC_Calculation'!$B$9")
    wb.defined_names['Target_Tax_Rate'] = DefinedName('Target_Tax_Rate', attr_text="'WACC_Calculation'!$B$10")

    # 참고용 셀 주소 표시
    ws_wacc['A' + str(r_wacc + 2)] = '[ Named Ranges for Reference ]'
    sc(ws_wacc.cell(r_wacc + 2, 1), fo=Font(name='Arial', bold=True, size=9, color=C_MG, italic=True))
    ws_wacc['A' + str(r_wacc + 3)] = '다른 시트에서 참조: =Target_WACC, =Target_Rf 등'
    sc(ws_wacc.cell(r_wacc + 3, 1), fo=Font(name='Arial', size=8, color=C_MG))

    # [Sheet 5] Price_History
    if price_abs_dfs:
        ws_ph = wb.create_sheet('Price_History')
        df_abs = pd.concat(price_abs_dfs, axis=1).sort_index().ffill()
        df_rel = pd.concat(price_rel_dfs, axis=1).sort_index().ffill()
        common_index = df_abs.index
        sc(ws_ph.cell(1,1,'Stock Price History (10 Years)'), fo=fT)
        ws_ph.merge_cells(start_row=1,start_column=1,end_row=1,end_column=10)
        r=3
        ws_ph.cell(r,1,'Date'); sc(ws_ph.cell(r,1), fo=fH, fi=pH, al=aC, bd=BD); ws_ph.column_dimensions['A'].width=12
        c_idx=2
        for col in df_abs.columns:
            sc(ws_ph.cell(r,c_idx,f"{ticker_to_name.get(col,col)} (Abs)"), fo=fH, fi=PatternFill('solid',fgColor='607D8B'), al=aC, bd=BD)
            ws_ph.column_dimensions[get_column_letter(c_idx)].width=16; c_idx+=1
        sc(ws_ph.cell(r,c_idx,""), fi=pW); ws_ph.column_dimensions[get_column_letter(c_idx)].width=2; c_idx+=1
        ws_ph.cell(r,c_idx,'Date'); sc(ws_ph.cell(r,c_idx), fo=fH, fi=pH, al=aC, bd=BD); rel_start=c_idx; c_idx+=1
        for col in df_rel.columns:
            sc(ws_ph.cell(r,c_idx,f"{ticker_to_name.get(col,col)} (Rel)"), fo=fH, fi=pH, al=aC, bd=BD)
            ws_ph.column_dimensions[get_column_letter(c_idx)].width=16; c_idx+=1
        r=4
        for date in common_index:
            dv=date.date(); ws_ph.cell(r,1,dv).number_format='yyyy-mm-dd'; sc(ws_ph.cell(r,1), fo=fA, al=aC, bd=BD); cc=2
            for v in df_abs.loc[date]: ws_ph.cell(r,cc,v).number_format='#,##0.00'; sc(ws_ph.cell(r,cc), fo=fA, al=aR, bd=BD); cc+=1
            cc+=1; ws_ph.cell(r,cc,dv).number_format='yyyy-mm-dd'; sc(ws_ph.cell(r,cc), fo=fA, al=aC, bd=BD); cc+=1
            for v in df_rel.loc[date]: ws_ph.cell(r,cc,v).number_format='#,##0'; sc(ws_ph.cell(r,cc), fo=fA, al=aR, bd=BD); cc+=1
            r+=1
        
        # Monthly Chart Data
        chart_start=r+2; df_m=df_rel.resample('ME').last().dropna(how='all')
        cr=chart_start; sc(ws_ph.cell(cr,1,'[ Chart Data - Monthly Sampled ]'), fo=fNOTE); cr+=1
        hdr_row=cr; sc(ws_ph.cell(cr,1,'Year-Month'), fo=Font(name='Arial',size=8,bold=True,color=C_MG), al=aC)
        for i,cn in enumerate(df_m.columns): sc(ws_ph.cell(cr,i+2,ticker_to_name.get(cn,cn)), fo=Font(name='Arial',size=8,bold=True,color=C_MG), al=aC)
        cr+=1; data_first=cr
        for date in df_m.index:
            ws_ph.cell(cr,1,date.strftime('%Y-%m')); sc(ws_ph.cell(cr,1), fo=Font(name='Arial',size=7,color=C_MG))
            for i,cn in enumerate(df_m.columns):
                v=df_m.loc[date,cn]; 
                if pd.notna(v): ws_ph.cell(cr,i+2,round(v,1))
                sc(ws_ph.cell(cr,i+2), fo=Font(name='Arial',size=7,color=C_MG))
            cr+=1
        data_last=cr-1
        
        # Chart
        chart=LineChart(); chart.title="10-Year Relative Performance (Base=100)"; chart.style=13; chart.height=18; chart.width=34
        chart.y_axis.title="Relative Index"; chart.y_axis.scaling.min=0; chart.y_axis.majorGridlines=ChartLines()
        chart.x_axis.title="Date"; chart.x_axis.tickLblSkip=max(1,(data_last-data_first)//10)
        cats=Reference(ws_ph, min_col=1, min_row=data_first, max_row=data_last); chart.set_categories(cats)
        colors=['1F77B4','FF7F0E','2CA02C','D62728','9467BD','8C564B','E377C2','7F7F7F','BCBD22','17BECF']
        for i in range(len(df_m.columns)):
            s=Series(Reference(ws_ph, min_col=i+2, min_row=hdr_row, max_row=data_last), title_from_data=True)
            s.graphicalProperties.line.solidFill=colors[i%len(colors)]; s.graphicalProperties.line.width=20000; s.smooth=False
            chart.series.append(s)
        ws_ph.add_chart(chart, f"{get_column_letter(rel_start+len(df_m.columns)+1)}3")

    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 5. Streamlit App Layout
# ==========================================
st.set_page_config(page_title="Global GPCM Generator", layout="wide", page_icon="📊")

# ---------------------------------------------------------
# [User Access Log] 접속자 로그 기록 (Console 출력)
# ---------------------------------------------------------
try:
    # Streamlit Cloud (Private) 환경에서 이메일 가져오기
    user_email = st.experimental_user.email if st.experimental_user.email else "Anonymous"
except:
    user_email = "Local_Dev"

# 현재 시간 (이미 import datetime이 되어 있으므로 바로 사용)
now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# 로그 출력 (Manage app > Logs 터미널에서 확인 가능)
print(f"👉 [접속 알림] {now_str} / 사용자: {user_email}")

st.title("📊 GPCM Calculator with yfinance")
st.write("yfinance 라이브러리를 통해 기준일 시점 선정된 Peer들의 재무제표, 주가, 시가총액 등을 크롤링하여 기준일 시점 Peer Group GPCM Multiple을 자동계산하는 어플리케이션입니다(Made by SGJ, 2026-02-10)")

# [Notes Section]
st.markdown("---")
st.subheader("📝 Valuation Methodology Notes")
notes = [
    '• Base Date: User Input | Unit: Millions (local currency)',
    '• EV (Enterprise Value) = Market Cap + Interest-Bearing Debt − Cash + Non-Controlling Interest',
    '• Cash includes: Cash & Cash Equivalents + Other Short-Term Investments',
    '• NOA(Option): Not deducted in EV. (User can change EV Tag in Excel BS_Full sheet)',
    '• EBIT = Operating Income (from PL)',
    '• EBITDA = Operating Income + D&A (Implied from yfinance)',
    '• PER = Market Cap ÷ Net Income Common Stockholders',
    '• PBR = Market Cap ÷ Stockholders Equity',
    '• PSR = Market Cap ÷ Revenue',
    '• Market Cap = Shares × Close Price (auto_adjust=False)',
    '• PL Source: LTM prioritized (Current + Prior Annual - Prior Same Q)',
]
for note in notes:
    st.text(note)

st.subheader("📊 Beta & Risk Analysis")
beta_notes = [
    '📌 Data Source:',
    '  - 한국 주식 (.KS, .KQ): FinanceDataReader 우선 → yfinance 백업',
    '  - 해외 주식: Yahoo Finance (yfinance)',
    '',
    '• Beta 계산 방법:',
    '  - 5Y Monthly Beta: 5년간 월말 종가 → 월간 수익률 → 시장지수 대비 선형회귀',
    '  - 2Y Weekly Beta: 2년간 주말 종가 → 주간 수익률 → 시장지수 대비 선형회귀',
    '  - Raw Beta = Slope of linear regression (Market vs Stock returns)',
    '  - Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1.0 (Bloomberg 방법론)',
    '',
    '• Market Index: KOSPI (KS11), KOSDAQ (KQ11), Nikkei 225, S&P/TSX, DAX, etc.',
    '',
    '• 값 검증: NaN, inf, 극단값(-10 ~ 10 범위 벗어남) 필터링',
    '',
    '• Tax Rate: Wikipedia-sourced corporate tax rates; Korean rates include local tax (2025)',
    '• Debt Ratio = IBD ÷ (IBD + Market Cap + NCI)',
    '• Unlevered Beta = Levered Beta ÷ (1 + (1 - Tax Rate) × Debt Ratio) [Hamada Model]',
]
for note in beta_notes:
    st.text(note)

st.subheader("💰 Target WACC (Weighted Average Cost of Capital)")
wacc_notes = [
    '📌 Target WACC은 엑셀 "WACC_Calculation" 시트에서 별도 계산됩니다.',
    '',
    '• Ke (자기자본비용) = Rf + MRP × Relevered Beta + Size Premium',
    '  - Relevered Beta = Avg Unlevered Beta × (1 + (1 - Tax) × Target D/E)',
    '  - Size Premium: 한국공인회계사회 기준 (Micro/Small/Medium/Large)',
    '',
    '• Kd (Aftertax) = Kd (Pretax) × (1 - Target Tax Rate)',
    '',
    '• Target D/V = 피어 평균 부채비율 (자동 계산)',
    '',
    '• WACC = (E/V) × Ke + (D/V) × Kd (Aftertax)',
]
for note in wacc_notes:
    st.text(note)
st.markdown("---")

# [Sidebar]
with st.sidebar:
    st.header("Settings")
    
    # 1. Date Input (Year/Quarter -> End Date)
    st.subheader("1. 기준일 설정")
    sel_year = st.number_input("Year", min_value=2020, max_value=2030, value=2025)
    sel_qtr = st.selectbox("Quarter", ["1Q", "2Q", "3Q", "4Q"], index=2) # Default 3Q
    
    q_end_map = {"1Q": "03-31", "2Q": "06-30", "3Q": "09-30", "4Q": "12-31"}
    base_date_str = f"{sel_year}-{q_end_map[sel_qtr]}"
    st.info(f"선택된 기준일: {base_date_str}")

    # 2. Ticker Input
    st.subheader("2. Ticker List")
    st.markdown("야후파이낸스 웹사이트 기준 Ticker를 한줄씩 입력하세요")
    default_tickers = """3116.T
7273.T
4246.T
7282.T
MG.TO
EZM.F
200880.KS
038110.KQ
012330.KS
PYT.VI"""
    txt_input = st.text_area("Tickers", value=default_tickers, height=250)

    # 3. WACC 파라미터 설정 (Target 기업용)
    st.subheader("3. Target WACC Parameters")

    st.markdown("**무위험이자율 (Rf)**")
    rf_input = st.number_input("Rf - 무위험이자율 (%)", min_value=0.0, max_value=10.0, value=3.3, step=0.1, format="%.2f",
                                help="한국 10년 국채수익률 (한국은행 경제통계시스템 참고)") / 100

    st.markdown("**자기자본비용 (Ke) 파라미터**")
    mrp_input = st.slider("MRP (시장위험프리미엄)", min_value=7.0, max_value=9.0, value=8.0, step=0.1, format="%.1f%%",
                         help="한국공인회계사회 권장: 7~9%") / 100

    st.markdown("**Size Premium (한국공인회계사회 기준, 2023)**")

    # Size Premium 표 보여주기
    with st.expander("📊 시가총액별 Size Premium 참고표"):
        st.markdown("**3분위수 기준**")
        st.markdown("""
        | 구분 | 시가총액 범위 (억원) | Size Premium |
        |------|---------------------|--------------|
        | **Micro** | < 2,000 | **4.02%** |
        | **Low** | 2,000 ~ 20,000 | 1.37% |
        | **Mid** | > 20,000 | -0.36% |
        """)

        st.markdown("**5분위수 기준**")
        st.markdown("""
        | 구분 | 시가총액 범위 (억원) | Size Premium |
        |------|---------------------|--------------|
        | **5분위 (최소)** | < 2,000 | **4.66%** |
        | **4분위** | 2,000 ~ 5,000 | 3.02% |
        | **3분위** | 5,000 ~ 20,000 | 1.21% |
        | **2분위** | 20,000 ~ 50,000 | 0.06% |
        | **1분위 (최대)** | > 50,000 | -0.58% |
        """)

        st.info("💡 Target 기업의 시가총액에 맞는 Size Premium을 선택하세요.")

    size_premium_options = {
        "3분위 - Micro (4.02%): < 2,000억": 0.0402,
        "3분위 - Low (1.37%): 2,000~20,000억": 0.0137,
        "3분위 - Mid (-0.36%): > 20,000억": -0.0036,
        "5분위 - 5분위/최소 (4.66%): < 2,000억": 0.0466,
        "5분위 - 4분위 (3.02%): 2,000~5,000억": 0.0302,
        "5분위 - 3분위 (1.21%): 5,000~20,000억": 0.0121,
        "5분위 - 2분위 (0.06%): 20,000~50,000억": 0.0006,
        "5분위 - 1분위/최대 (-0.58%): > 50,000억": -0.0058,
        "Size Premium 없음 (0%)": 0.0
    }
    size_premium_choice = st.selectbox("기업 규모 선택", list(size_premium_options.keys()), index=0,
                                       help="Target 기업의 시가총액에 맞는 Size Premium 선택")
    size_premium_input = size_premium_options[size_premium_choice]

    st.markdown("**Beta 계산 기준 선택**")
    beta_type_options = {
        "5년 월간 베타 (5Y Monthly)": "5Y",
        "2년 주간 베타 (2Y Weekly)": "2Y"
    }
    beta_type_choice = st.selectbox("WACC 계산에 사용할 Beta", list(beta_type_options.keys()), index=0,
                                    help="Target WACC 계산 시 사용할 베타 유형을 선택하세요. 두 베타 모두 엑셀에 표시됩니다.")
    beta_type_input = beta_type_options[beta_type_choice]

    st.markdown("**타인자본비용 (Kd) 파라미터**")
    kd_pretax_input = st.number_input("Kd (Pretax) - 세전 이자율 (%)", min_value=0.0, max_value=15.0, value=3.5, step=0.1, format="%.1f") / 100

    st.markdown("**Target 기업 법인세율**")
    target_tax_rate_input = st.number_input("Target 법인세율 (%)", min_value=0.0, max_value=50.0, value=26.4, step=0.1, format="%.1f",
                                            help="한국: 26.4% (대기업 기준, 지방세 포함) | 미국: 21% | 일본: 30.6%") / 100

    st.info(f"💡 목표 부채비율은 피어들의 평균 자본구조로 자동 계산됩니다.")

    # 4. Run Button
    btn_run = st.button("Go, Go, Go 🚀", type="primary")

# [Main Execution]
if btn_run:
    target_tickers = [t.strip() for t in txt_input.split('\n') if t.strip()]

    with st.spinner("데이터 추출 및 분석 중... 잠시만 기다려주세요..."):
        # Run Data Logic with WACC parameters (목표 부채비율은 자동 계산)
        gpcm_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, t_map, avg_debt_ratio, target_wacc_data = get_gpcm_data(
            target_tickers,
            base_date_str,
            mrp=mrp_input,
            kd_pretax=kd_pretax_input,
            size_premium=size_premium_input,
            target_tax_rate=target_tax_rate_input,
            user_rf_rate=rf_input,
            beta_type=beta_type_input
        )
        
        # 1. Summary Table
        st.subheader("📋 GPCM Multiples Summary")
        summary_list = []
        for t, g in gpcm_data.items():
            ev = g['Market_Cap_M'] + g['IBD'] - g['Cash'] + g['NCI'] # NOA Option 미반영
            ev_ebitda = ev/g['EBITDA'] if g['EBITDA']>0 else None
            ev_ebit = ev/g['EBIT'] if g['EBIT']>0 else None
            per = g['Market_Cap_M']/g['NI_Parent'] if g['NI_Parent']>0 else None
            pbr = g['Market_Cap_M']/g['Equity'] if g['Equity']>0 else None
            psr = g['Market_Cap_M']/g['Revenue'] if g['Revenue']>0 else None
            
            summary_list.append({
                'Ticker': t, 'Company': g['Company'],
                'EV/EBITDA': ev_ebitda, 'EV/EBIT': ev_ebit, 'PER': per, 'PBR': pbr, 'PSR': psr
            })
        df_sum = pd.DataFrame(summary_list)
        st.dataframe(df_sum.style.format({
            'EV/EBITDA': '{:.1f}x', 'EV/EBIT': '{:.1f}x', 'PER': '{:.1f}x', 'PBR': '{:.1f}x', 'PSR': '{:.1f}x'
        }, na_rep='N/M'))

        # 평균 자본구조 표시
        st.success(f"✅ **피어 평균 부채비율 (목표 자본구조)**: {avg_debt_ratio*100:.1f}%")

        # 2. Statistics Table
        st.subheader("📊 Multiples Statistics")
       
        # [안전장치] 데이터가 비어있지 않은 경우에만 통계 계산 수행
        if not df_sum.empty:
            stats = []
            for col in ['EV/EBITDA', 'EV/EBIT', 'PER', 'PBR', 'PSR']:
                # 해당 컬럼이 실제로 존재하는지 확인 (KeyError 방지)
                if col in df_sum.columns:
                    vals = [x for x in df_sum[col] if pd.notnull(x)]
                    if vals:
                        stats.append({'Metric': col, 'Mean': np.mean(vals), 'Median': np.median(vals), 'Max': np.max(vals), 'Min': np.min(vals)})
                    else:
                        stats.append({'Metric': col, 'Mean': None, 'Median': None, 'Max': None, 'Min': None})
            
            if stats:
                st.dataframe(pd.DataFrame(stats).set_index('Metric').style.format('{:.1f}x', na_rep='N/M'))
            else:
                st.warning("통계를 산출할 유효한 데이터가 없습니다.")
        else:
            st.error("데이터를 불러오지 못했습니다. 잠시 후 다시 시도해주세요 (Yahoo Rate Limit).")

        # 3. Excel Download
        excel_data = create_excel(gpcm_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, base_date_str, t_map, target_wacc_data, beta_type=beta_type_input)
        
        # 다운로드 버튼 (누르고 있어도 화면 유지됨)
        st.download_button(
            label="📥 Report Download (Excel)",
            data=excel_data,
            file_name=f"Global_GPCM_{base_date_str.replace('-','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
