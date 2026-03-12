# 최신 수정: 2026-02-16 16:00 KST
# 주요 변경사항:
# - D/E Ratio 컬럼 추가 (IBD / (시총 + NCI))
# - Unlevered Beta 계산에 D/E Ratio 사용 (정확한 Hamada 공식)
# - 국내 베타 계산 수정: 시장 지수는 yfinance 사용 (^KS11, ^KQ11)
# - Debt Ratio 수식 수정: IBD/(시총+IBD+NCI) [총부채/총자산]
# - 노트 업데이트: 주가 데이터 소스 명시 (yfinance 통합)

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
# import FinanceDataReader as fdr (Removed as per user request to use yfinance only)
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
        return 'KOSPI', '^KS11'
    elif ticker_upper.endswith('.KQ'):
        return 'KOSDAQ', '^KQ11'
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

@st.cache_data(ttl=86400)
def get_kpmg_tax_rates():
    """
    KPMG 사이트(OECD 데이터 기반) 법인세율 자동 크롤링 및 주요 국가 하드코딩 Fallback
    """
    # [Fallback] 주요 국가 최신 실효 법인세율 (Combined Rate, 2025 KPMG/OECD 기준)
    tax_rates = {
        'Korea': 0.264, 'United States': 0.26, 'Japan': 0.2974, 
        'France': 0.3613, 'Germany': 0.3006, 'Austria': 0.23, 
        'Canada': 0.265, 'United Kingdom': 0.25, 'China': 0.25, 
        'India': 0.30, 'Australia': 0.30, 'Netherlands': 0.258, 
        'Spain': 0.25, 'Italy': 0.24, 'Switzerland': 0.148, # 스위스 평균
        'Ireland': 0.125, 'Vietnam': 0.20, 'Sweden': 0.206, 'Norway': 0.22,
        'Belgium': 0.25, 'Mexico': 0.30
    }
    try:
        url = "https://kpmg.com/dk/en/services/tax/corporate-tax/corporate-tax-rates-table.html"
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=5)
        soup = BeautifulSoup(response.content, 'html.parser')
        table = soup.find('table')
        if table:
            rows = table.find_all('tr')
            for row in rows[1:]:
                cols = row.find_all('td')
                if len(cols) >= 2:
                    country = cols[0].get_text(strip=True)
                    rate_text = cols[1].get_text(strip=True).replace('%','')
                    try:
                        rate = float(rate_text) / 100
                        tax_rates[country] = rate
                    except: continue
    except: pass
    return tax_rates

def get_corporate_tax_rates_from_wikipedia():
    # [KPMG 사이트 우선 사용하도록 랩핑]
    kpmg_rates = get_kpmg_tax_rates()
    # 기존 위키백과 로직은 백업용으로 유지하거나 KPMG 결과와 병합
    return kpmg_rates

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

def get_country_from_ticker(ticker, info=None):
    """
    티커 및 info로부터 국가 코드 추출
    """
    if info and 'country' in info:
        c_val = info.get('country')
        # yfinance info의 country는 'France', 'Austria', 'Germany' 등 풀네임인 경우가 많음
        c_map_inv = {
            'France':'FR', 'Austria':'AT', 'Germany':'DE', 'Japan':'JP', 'South Korea':'KR', 'Korea':'KR',
            'United States':'US', 'United Kingdom':'GB', 'Canada':'CA', 'Netherlands':'NL', 'Spain':'ES', 'Italy':'IT'
        }
        if c_val in c_map_inv: return c_map_inv[c_val]
        if len(str(c_val)) == 2: return c_val.upper()

    ticker_upper = ticker.upper()
    # [수송] 수동 매핑 추가 (사용자 특정 사례)
    if 'EZM.F' in ticker_upper: return 'FR' # OPmobility SE는 프랑스 본사지만 독일 상장
    if 'PYT.VI' in ticker_upper: return 'AT' # Polytec Holding AG는 오스트리아

    # 아시아
    if ticker_upper.endswith('.KS') or ticker_upper.endswith('.KQ'): return 'KR'
    elif ticker_upper.endswith('.T'): return 'JP'
    elif ticker_upper.endswith('.SS') or ticker_upper.endswith('.SZ'): return 'CN'
    elif ticker_upper.endswith('.HK'): return 'HK'
    elif ticker_upper.endswith('.SI'): return 'SG'
    elif ticker_upper.endswith('.TW'): return 'TW'
    elif ticker_upper.endswith('.BO') or ticker_upper.endswith('.NS'): return 'IN'
    # 북미
    elif ticker_upper.endswith('.TO') or ticker_upper.endswith('.V'): return 'CA'
    elif ticker_upper.endswith('.MX'): return 'MX'
    # 유럽 (Suffix 기반 Fallback)
    elif ticker_upper.endswith('.F') or ticker_upper.endswith('.DE') or ticker_upper.endswith('.BE'): return 'DE'
    elif ticker_upper.endswith('.PA'): return 'FR'
    elif ticker_upper.endswith('.L'): return 'GB'
    elif ticker_upper.endswith('.MI'): return 'IT'
    elif ticker_upper.endswith('.MC'): return 'ES'
    elif ticker_upper.endswith('.AS'): return 'NL'
    elif ticker_upper.endswith('.SW') or ticker_upper.endswith('.VX'): return 'CH'
    elif ticker_upper.endswith('.VI'): return 'AT'
    elif ticker_upper.endswith('.BR'): return 'BE'
    elif ticker_upper.endswith('.ST'): return 'SE'
    elif ticker_upper.endswith('.OL'): return 'NO'
    elif ticker_upper.endswith('.CO'): return 'DK'
    elif ticker_upper.endswith('.HE'): return 'FI'
    elif ticker_upper.endswith('.IR'): return 'IE'
    # 오세아니아
    elif ticker_upper.endswith('.AX'): return 'AU'
    elif ticker_upper.endswith('.NZ'): return 'NZ'
    
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
def get_gpcm_data(tickers_list, target_periods, mrp=0.08, kd_pretax=0.035, size_premium=0.0402, target_tax_rate=0.264, user_rf_rate=None, beta_type="5Y", force_annual=False, include_recent=False):
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
    base_period_str = target_periods[-1]
    base_dt = pd.to_datetime(base_period_str)

    # 10년 국채수익률 조회 (무위험수익률)
    rf_rate = get_korea_10y_treasury_yield(base_period_str, user_rf_rate)
    
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
        'NOA':                                                 'NOA',
    }
    BS_SUBTOTAL_EXCLUDE = {
        'Cash Cash Equivalents And Short Term Investments', 'Cash And Short Term Investments',
        'Total Debt', 'Total Capitalization', 'Total Equity Gross Minority Interest',
    }
    PL_HIGHLIGHT_MAP = {'Total Revenue': 'Revenue', 'Operating Income': 'EBIT', 'EBIT': 'EBIT', 'EBITDA': 'EBIT', 'Normalized EBITDA': 'EBIT', 'Net Income Common Stockholders': 'NI_Parent', 'Net Income': 'Net Income'}
    PL_CALC_KEY = {'Total Revenue': 'Revenue', 'Operating Income': 'OpIncome', 'EBIT': 'EBIT_yf', 'EBITDA': 'EBITDA_yf', 'Normalized EBITDA': 'NormEBITDA_yf', 'Net Income Common Stockholders': 'NI_Parent'}
    PL_SORT = {'Total Revenue': 10, 'Operating Revenue': 11, 'Cost Of Revenue': 20, 'Gross Profit': 30, 'Operating Expense': 35, 'Selling General And Administration': 36, 'Research And Development': 37, 'Operating Income': 50, 'EBIT': 55, 'EBITDA': 56, 'Normalized EBITDA': 57, 'Interest Expense': 60, 'Pretax Income': 70, 'Tax Provision': 75, 'Net Income': 90, 'Net Income Common Stockholders': 91, 'Basic EPS': 95, 'Diluted EPS': 96}

    # Labels for indexing (Y, Y-1, Y-2...)
    lookback = len(target_periods)
    rel_labels = []
    for i in range(lookback):
        label = "Y" if i == 0 else f"Y-{i}"
        rel_labels.append(label)
    
    # We will return data keyed by rel_labels
    all_period_data = {lbl: {} for lbl in rel_labels}
    if include_recent:
        all_period_data['Recent'] = {}
    base_label = "Y"
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
            exchange, market_idx = get_market_index(ticker)

            # [핵심] 재무 데이터 가져오기 및 타임존 표준화 (tz-naive)
            a_is = stock.income_stmt
            q_is = stock.quarterly_income_stmt
            a_bs = stock.balance_sheet
            q_bs = stock.quarterly_balance_sheet

            def norm_df(df):
                if df is not None and not df.empty:
                    df.columns = [d.replace(tzinfo=None) if hasattr(d, 'tzinfo') else d for d in df.columns]
                return df

            a_is = norm_df(a_is)
            q_is = norm_df(q_is)
            a_bs = norm_df(a_bs)
            q_bs = norm_df(q_bs)

            if a_is is None or a_is.empty:
                st.warning(f"⚠️ {ticker}: 연간 재무제표를 찾을 수 없습니다. 건너뜜.")
                continue
            
            all_fiscal_dates = sorted([d for d in a_is.columns], reverse=True)

            # [수정] 라벨링 체계 정교화: Y(기준일/LTM), Y-1(직전연간), Y-2...
            labels_to_fetch = [('Y', base_dt)]
            hist_annuals = [d for d in all_fiscal_dates if d < base_dt - timedelta(days=30)]
            
            for i in range(1, lookback):
                if i-1 < len(hist_annuals):
                    labels_to_fetch.append((f'Y-{i}', hist_annuals[i-1]))

            for label, f_dt_obj in labels_to_fetch:
                if hasattr(f_dt_obj, 'tzinfo'): f_dt_obj = f_dt_obj.replace(tzinfo=None)
                f_dt_str = f_dt_obj.strftime('%Y-%m-%d')
                
                gpcm = {
                    'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                    'Base_Date': f_dt_str, # 실제 결산일 표시
                    'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA(Option)': 0, 'Equity': 0,
                    'Revenue': 0, 'EBIT': 0, 'EBITDA': 0, 'NI_Parent': 0,
                    'Close': 0, 'Shares': 0, 'Market_Cap_M': 0, 'PL_Source': 'Annual',
                    'Exchange': exchange, 'Market_Index': market_idx,
                    'Beta_5Y_Monthly_Raw': None, 'Beta_5Y_Monthly_Adj': None,
                    'Beta_2Y_Weekly_Raw': None, 'Beta_2Y_Weekly_Adj': None,
                    'Pretax_Income': 0, 'Tax_Rate': 0.22,
                    'Debt_Ratio': 0, 'Unlevered_Beta_5Y': None, 'Unlevered_Beta_2Y': None,
                }

                # [수정] BS 최신 시점 엄밀 선택 (Strict Max <= Base Date)
                # 연간/분기를 합쳐서 기준일(base_dt) 이전의 가장 최신 날짜 1개를 찾음
                a_bs = stock.balance_sheet
                q_bs = stock.quarterly_balance_sheet
                all_possible_bs_dates = []
                if a_bs is not None and not a_bs.empty:
                    all_possible_bs_dates.extend([d for d in a_bs.columns if d <= base_dt + timedelta(days=7)])
                if q_bs is not None and not q_bs.empty:
                    all_possible_bs_dates.extend([d for d in q_bs.columns if d <= base_dt + timedelta(days=7)])
                
                # 중복 제거 및 내림차순 정렬
                all_possible_bs_dates = sorted(list(set(all_possible_bs_dates)), reverse=True)
                
                target_bs = None
                actual_bs_date = None
                
                if all_possible_bs_dates:
                    # 가장 최신 날짜 선택
                    actual_bs_date = all_possible_bs_dates[0]
                    # 연간에 있으면 연간에서, 없으면 분기에서 데이터 가져옴
                    if a_bs is not None and actual_bs_date in a_bs.columns:
                        target_bs = a_bs[actual_bs_date]
                    elif q_bs is not None and actual_bs_date in q_bs.columns:
                        target_bs = q_bs[actual_bs_date]

                if target_bs is not None:
                    for acct_name in target_bs.index:
                        val = target_bs.loc[acct_name]
                        if pd.isna(val): continue
                        val_f = float(val)
                        if str(acct_name) == 'Ordinary Shares Number': bs_shares = val_f
                        ev_tag = BS_HIGHLIGHT_MAP.get(str(acct_name), '')
                        if str(acct_name) in BS_SUBTOTAL_EXCLUDE: ev_tag = ''
                        
                        raw_bs_rows.append({
                            'Company': company_name, 'Ticker': ticker, 'Period': actual_bs_date.strftime('%Y-%m-%d'), 'Label': label,
                            'Currency': currency, 'Account': str(acct_name), 'EV_Tag': ev_tag, 'Amount_M': val_f/1e6
                        })
                        if ev_tag: gpcm[ev_tag] += val_f/1e6
                
                gpcm['Shares'] = bs_shares if bs_shares else float(info.get('sharesOutstanding', 0))

                # [2] Market Cap (실제 결산일 시점의 주가 사용)
                try:
                    hist = stock.history(start=(f_dt_obj - timedelta(days=10)).strftime('%Y-%m-%d'), end=(f_dt_obj + timedelta(days=1)).strftime('%Y-%m-%d'), auto_adjust=False)
                    close = float(hist['Close'].iloc[-1]) if (not hist.empty and 'Close' in hist.columns) else 0.0
                    p_date = hist.index[-1].strftime('%Y-%m-%d') if not hist.empty else '-'
                except: close=0.0; p_date='-'
                gpcm['Close'] = close
                gpcm['Market_Cap_M'] = (close * gpcm['Shares'] / 1e6) if gpcm['Shares'] else 0.0
                market_rows.append({
                    'Company': company_name, 'Ticker': ticker, 'Base_Date': f_dt_str, 'Price_Date': p_date, 'Label': label,
                    'Currency': currency, 'Close': close, 'Shares': gpcm['Shares'], 'Market_Cap_M': round(gpcm['Market_Cap_M'], 1)
                })

                # [3] PL (해당 결산일의 연간 데이터 사용)
                calc_sums = {'Revenue': 0, 'OpIncome': 0, 'EBIT_yf': 0, 'EBITDA_yf': 0, 'NormEBITDA_yf': 0, 'NI_Parent': 0}
                
                # [수정] 기준일('Y') 시점에 해당 날짜 연간 데이터가 없으면 최신 연간 데이터를 Fallback으로 사용
                f_pl_lookup = f_dt_obj
                if label == 'Y' and f_dt_obj not in a_is.columns:
                    f_pl_lookup = all_fiscal_dates[0] if all_fiscal_dates else None

                if f_pl_lookup and f_pl_lookup in a_is.columns:
                    for acct in a_is.index:
                        acct_str = str(acct)
                        hl_tag = PL_HIGHLIGHT_MAP.get(acct_str, '')
                        calc_key = PL_CALC_KEY.get(acct_str, '')
                        is_eps = 'EPS' in acct_str or 'Per Share' in acct_str
                        is_shares = 'Average Shares' in acct_str
                        
                        val = a_is.loc[acct, f_pl_lookup]
                        if pd.isna(val): continue
                        val_f = float(val)
                        if is_eps: unit = 'per share'; amt = val_f
                        elif is_shares: unit = 'M shares'; amt = val_f/1e6
                        else: unit = 'M'; amt = val_f/1e6
                        
                        raw_pl_rows.append({
                            'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                            'Account': acct_str, 'GPCM_Tag': hl_tag, 'PL_Source': 'Annual',
                            'Q_Label': 'Annual', 'Period': f_pl_lookup.strftime('%Y-%m-%d'), 'Label': label, 
                            'Amount_M': amt, 'Unit': unit, '_sort': PL_SORT.get(acct_str, 500)
                        })
                        if calc_key and not is_eps and not is_shares: calc_sums[calc_key] += val_f/1e6
                else:
                    # 데이터가 전혀 없을 경우에도 GPCM 딕셔너리 구조는 유지 (에러 방지)
                    pass
                
                gpcm['Revenue'] = calc_sums['Revenue']
                gpcm['EBIT'] = calc_sums['OpIncome']
                ebitda_yf = calc_sums['EBITDA_yf'] if calc_sums['EBITDA_yf'] != 0 else calc_sums['NormEBITDA_yf']
                ebit_yf = calc_sums['EBIT_yf']
                da_amount = (ebitda_yf - ebit_yf) if (ebitda_yf != 0 and ebit_yf != 0) else 0
                gpcm['EBITDA'] = calc_sums['OpIncome'] + da_amount
                gpcm['NI_Parent'] = calc_sums['NI_Parent']

                # [4] Tax Rate
                k_rates = get_corporate_tax_rates_from_wikipedia()
                country_code = get_country_from_ticker(ticker, info)
                
                if country_code == 'KR':
                    # Pretax Income 확보
                    if 'Pretax Income' in a_is.index and f_pl_lookup in a_is.columns:
                        pt_val = a_is.loc['Pretax Income', f_pl_lookup]
                        if pd.notna(pt_val): gpcm['Pretax_Income'] = float(pt_val) / 1e6
                    gpcm['Tax_Rate'] = get_korean_marginal_tax_rate(gpcm['Pretax_Income'])
                else: 
                    # KPMG 국가명 맵핑 (KPMG 테이블의 Key와 매칭)
                    c_name_map = {
                        'US':'United States', 'JP':'Japan', 'CN':'China', 'IN':'India', 'CA':'Canada', 
                        'GB':'United Kingdom', 'DE':'Germany', 'FR':'France', 'AU':'Australia',
                        'AT':'Austria', 'NL':'Netherlands', 'BE':'Belgium', 'IT':'Italy', 'ES':'Spain',
                        'SE':'Sweden', 'NO':'Norway', 'DK':'Denmark', 'FI':'Finland', 'IE':'Ireland', 'CH':'Switzerland'
                    }
                    c_full_name = c_name_map.get(country_code, '')
                    gpcm['Tax_Rate'] = k_rates.get(c_full_name, 0.25)

                # [6] LTM Adjustment for Base Period (Y)
                if label == base_label and not force_annual:
                    try:
                        q_is = stock.quarterly_income_stmt
                        if q_is is not None and not q_is.empty:
                            # 1. 기준일(base_dt) 이전의 가용 분기 결산일 찾기
                            valid_q_dates = sorted([d for d in q_is.columns if d <= base_dt + timedelta(days=15)], reverse=True)
                            if valid_q_dates:
                                recent_q_dt = valid_q_dates[0]
                                # 최근 4개 분기 데이터 확보 시도
                                ltm_q_dates = [d for d in q_is.columns if d <= recent_q_dt][:4]
                                
                                if len(ltm_q_dates) == 4:
                                    # 기존 Annual 기반 PL 데이터 제거 (이 티커와 이 레이블에 대해)
                                    # global raw_pl_rows를 수정해야 하므로 주의 (여기서는 지역 변수처럼 사용되고 있다면 문제 없음)
                                    # raw_pl_rows = [r for r in raw_pl_rows if not (r['Ticker'] == ticker and r['Label'] == label)]
                                    # 위 방식은 리스트를 새로 만드므로, 원본 리스트를 유지하기 위해 index를 찾아 지우거나 필터링 다시 수행
                                    filtered_pl = [r for r in raw_pl_rows if not (r['Ticker'] == ticker and r['Label'] == label)]
                                    raw_pl_rows.clear()
                                    raw_pl_rows.extend(filtered_pl)

                                    # 4개 분기 각각에 대해 PL 데이터 추가 (transparency: D-0Q ~ D-3Q)
                                    for i, q_dt in enumerate(ltm_q_dates):
                                        q_col = q_is[q_dt]
                                        q_label_calc = f"D-{i}Q"

                                        for acct in q_is.index:
                                            acct_str = str(acct)
                                            val_q = float(q_col.loc[acct])
                                            hl_tag = PL_HIGHLIGHT_MAP.get(acct_str, '')
                                            is_eps = 'EPS' in acct_str or 'Per Share' in acct_str
                                            amt_q = val_q if is_eps else val_q/1e6
                                            
                                            raw_pl_rows.append({
                                                'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                                                'Account': acct_str, 'GPCM_Tag': hl_tag, 'PL_Source': 'Quarterly (4Q Sum)',
                                                'Q_Label': q_label_calc, 'Period': q_dt.strftime('%Y-%m-%d'), 
                                                'Label': label, 'Amount_M': amt_q, 'Unit': 'per share' if is_eps else 'M', 
                                                '_sort': PL_SORT.get(acct_str, 500)
                                            })

                                    # 4개 분기 합산 PL (GPCM 연산용 내부 변수만 업데이트)
                                    ltm_sum_vals = q_is[ltm_q_dates].sum(axis=1)
                                    calc_sums_ltm = {'Revenue': 0, 'OpIncome': 0, 'EBIT_yf': 0, 'EBITDA_yf': 0, 'NormEBITDA_yf': 0, 'NI_Parent': 0}
                                    
                                    for acct in ltm_sum_vals.index:
                                        acct_str = str(acct)
                                        calc_key = PL_CALC_KEY.get(acct_str, '')
                                        is_eps = 'EPS' in acct_str or 'Per Share' in acct_str
                                        is_shares = 'Average Shares' in acct_str
                                        val_ltm = float(ltm_sum_vals.loc[acct])
                                        if calc_key and not is_eps and not is_shares: calc_sums_ltm[calc_key] += val_ltm/1e6
                                    
                                    # GPCM 데이터 최종 업데이트
                                    gpcm['Revenue'] = calc_sums_ltm['Revenue']
                                    gpcm['EBIT'] = calc_sums_ltm['OpIncome']
                                    ebitda_yf_l = calc_sums_ltm['EBITDA_yf'] if calc_sums_ltm['EBITDA_yf'] != 0 else calc_sums_ltm['NormEBITDA_yf']
                                    ebit_yf_l = calc_sums_ltm['EBIT_yf']
                                    da_amount_l = (ebitda_yf_l - ebit_yf_l) if (ebitda_yf_l != 0 and ebit_yf_l != 0) else 0
                                    gpcm['EBITDA'] = calc_sums_ltm['OpIncome'] + da_amount_l
                                    gpcm['NI_Parent'] = calc_sums_ltm['NI_Parent']
                                    
                                    gpcm['PL_Source'] = 'LTM (4Q Sum)'
                                    gpcm['Base_Date'] = recent_q_dt.strftime('%Y-%m-%d')
                                    
                                    # Tax Rate 재산출 (LTM Pretax Income 기준)
                                    if 'Pretax Income' in ltm_sum_vals.index:
                                        gpcm['Pretax_Income'] = float(ltm_sum_vals.loc['Pretax Income']) / 1e6
                                    
                                    if country_code == 'KR':
                                        gpcm['Tax_Rate'] = get_korean_marginal_tax_rate(gpcm['Pretax_Income'])
                                    else:
                                        c_name_map = {
                                            'US':'United States', 'JP':'Japan', 'CN':'China', 'IN':'India', 'CA':'Canada', 
                                            'GB':'United Kingdom', 'DE':'Germany', 'FR':'France', 'AU':'Australia',
                                            'AT':'Austria', 'NL':'Netherlands', 'BE':'Belgium', 'IT':'Italy', 'ES':'Spain',
                                            'SE':'Sweden', 'NO':'Norway', 'DK':'Denmark', 'FI':'Finland', 'IE':'Ireland', 'CH':'Switzerland'
                                        }
                                        c_full_name = c_name_map.get(country_code, '')
                                        gpcm['Tax_Rate'] = k_rates.get(c_full_name, 0.25)
                    except: pass

                # Save to all_period_data
                all_period_data[label][ticker] = gpcm

            # [핵심] 최신 분기 데이터(Recent) 별도 수집 (재무제표 요약 모드용)
            if include_recent:
                try:
                    q_is_recent = stock.quarterly_income_stmt
                    q_bs_recent = stock.quarterly_balance_sheet
                    
                    if q_is_recent is not None and not q_is_recent.empty and q_bs_recent is not None and not q_bs_recent.empty:
                        recent_f_dt = q_is_recent.columns[0]
                        recent_f_str = recent_f_dt.strftime('%Y-%m-%d')
                        
                        gpcm_recent = {
                            'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                            'Base_Date': recent_f_str,
                            'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA(Option)': 0, 'Equity': 0,
                            'Revenue': 0, 'EBIT': 0, 'EBITDA': 0, 'NI_Parent': 0,
                            'Close': 0, 'Shares': 0, 'Market_Cap_M': 0, 'PL_Source': 'Quarterly (Recent)',
                            'Exchange': exchange, 'Market_Index': market_idx,
                            'Pretax_Income': 0, 'Tax_Rate': 0.22,
                            'Debt_Ratio': 0,
                        }
                        
                        # Recent BS
                        recent_bs_data = q_bs_recent.iloc[:, 0]
                        recent_bs_date_str = q_bs_recent.columns[0].strftime('%Y-%m-%d')
                        bs_shares_r = None
                        for acct_name in recent_bs_data.index:
                            val = recent_bs_data.loc[acct_name]
                            if pd.isna(val): continue
                            val_f = float(val)
                            if str(acct_name) == 'Ordinary Shares Number': bs_shares_r = val_f
                            ev_tag = BS_HIGHLIGHT_MAP.get(str(acct_name), '')
                            if str(acct_name) in BS_SUBTOTAL_EXCLUDE: ev_tag = ''
                            
                            raw_bs_rows.append({
                                'Company': company_name, 'Ticker': ticker, 'Period': recent_bs_date_str, 'Label': 'Recent',
                                'Currency': currency, 'Account': str(acct_name), 'EV_Tag': ev_tag, 'Amount_M': val_f/1e6
                            })
                            if ev_tag: gpcm_recent[ev_tag] += val_f/1e6
                        
                        gpcm_recent['Shares'] = bs_shares_r if bs_shares_r else float(info.get('sharesOutstanding', 0))
                        
                        # Recent Market Cap
                        try:
                            hist_r = stock.history(period='1d', auto_adjust=False)
                            close_r = float(hist_r['Close'].iloc[-1]) if not hist_r.empty else info.get('previousClose', 0)
                            p_date_r = hist_r.index[-1].strftime('%Y-%m-%d') if not hist_r.empty else '-'
                        except: close_r=0.0; p_date_r='-'
                        gpcm_recent['Close'] = close_r
                        gpcm_recent['Market_Cap_M'] = (close_r * gpcm_recent['Shares'] / 1e6) if gpcm_recent['Shares'] else 0.0
                        market_rows.append({
                            'Company': company_name, 'Ticker': ticker, 'Base_Date': recent_f_str, 'Price_Date': p_date_r, 'Label': 'Recent',
                            'Currency': currency, 'Close': close_r, 'Shares': gpcm_recent['Shares'], 'Market_Cap_M': round(gpcm_recent['Market_Cap_M'], 1)
                        })
                        
                        # Recent PL
                        recent_pl_data = q_is_recent.iloc[:, 0]
                        calc_sums_r = {'Revenue': 0, 'OpIncome': 0, 'EBIT_yf': 0, 'EBITDA_yf': 0, 'NormEBITDA_yf': 0, 'NI_Parent': 0}
                        for acct in recent_pl_data.index:
                            acct_str = str(acct)
                            hl_tag = PL_HIGHLIGHT_MAP.get(acct_str, '')
                            calc_key = PL_CALC_KEY.get(acct_str, '')
                            is_eps = 'EPS' in acct_str or 'Per Share' in acct_str
                            is_shares = 'Average Shares' in acct_str
                            
                            val = recent_pl_data.loc[acct]
                            if pd.isna(val): continue
                            val_f = float(val)
                            amt = val_f if is_eps else val_f/1e6
                            unit = 'per share' if is_eps else ('M shares' if is_shares else 'M')
                            
                            raw_pl_rows.append({
                                'Company': company_name, 'Ticker': ticker, 'Currency': currency,
                                'Account': acct_str, 'GPCM_Tag': hl_tag, 'PL_Source': 'Quarterly',
                                'Q_Label': 'Recent', 'Period': recent_f_str, 'Label': 'Recent', 
                                'Amount_M': amt, 'Unit': unit, '_sort': PL_SORT.get(acct_str, 500)
                            })
                            if calc_key and not is_eps and not is_shares: calc_sums_r[calc_key] += val_f/1e6
                            
                        gpcm_recent['Revenue'] = calc_sums_r['Revenue']
                        gpcm_recent['EBIT'] = calc_sums_r['OpIncome']
                        gpcm_recent['EBITDA'] = calc_sums_r['OpIncome'] # Simplified for Recent
                        gpcm_recent['NI_Parent'] = calc_sums_r['NI_Parent']
                        
                        all_period_data['Recent'][ticker] = gpcm_recent
                except: pass

            # [Beta Calculation] Only for the Base Period (Label 'Y')
            base_gpcm = all_period_data.get('Y', {}).get(ticker)
            if base_gpcm:
                # Price History
                try:
                    hist_10y_raw = stock.history(start=(base_dt - timedelta(days=365*10+20)).strftime('%Y-%m-%d'), end=base_dt.strftime('%Y-%m-%d'), auto_adjust=False)
                    hist_10y = hist_10y_raw['Close'] if 'Close' in hist_10y_raw.columns else hist_10y_raw.iloc[:,0]
                    if not hist_10y.empty:
                        abs_s = hist_10y.copy(); abs_s.name = ticker; price_abs_dfs.append(abs_s)
                        rel_s = (hist_10y / hist_10y.iloc[0]) * 100; rel_s.name = ticker; price_rel_dfs.append(rel_s)
                except: pass

                try:
                    start_5y = (base_dt - timedelta(days=365*5+20)).strftime('%Y-%m-%d')
                    end_date = base_dt.strftime('%Y-%m-%d')

                    stock_data_5y = None; market_data_5y = None
                    try:
                        stock_hist_5y = stock.history(start=start_5y, end=end_date, auto_adjust=False)
                        if not stock_hist_5y.empty: stock_data_5y = stock_hist_5y

                        # Always use yfinance for market data as per user request
                        market_data_5y = yf.download(market_idx, start=start_5y, end=end_date, progress=False)
                        
                        if market_data_5y is not None and not market_data_5y.empty:
                            if not isinstance(market_data_5y.index, pd.DatetimeIndex): market_data_5y.index = pd.to_datetime(market_data_5y.index)
                            if market_data_5y.index.tz is not None: market_data_5y.index = market_data_5y.index.tz_localize(None)
                    except Exception as e: st.warning(f"{ticker} 5Y 데이터 수집 실패: {e}")

                    if stock_data_5y is not None and market_data_5y is not None:
                        if not stock_data_5y.empty and not market_data_5y.empty:
                            stock_prices_5y = stock_data_5y['Close'] if isinstance(stock_data_5y, pd.DataFrame) and 'Close' in stock_data_5y.columns else (stock_data_5y.iloc[:, 0] if isinstance(stock_data_5y, pd.DataFrame) else stock_data_5y)
                            market_prices_5y = market_data_5y['Close'] if isinstance(market_data_5y, pd.DataFrame) and 'Close' in market_data_5y.columns else (market_data_5y.iloc[:, 0] if isinstance(market_data_5y, pd.DataFrame) else market_data_5y)
                            
                            if not isinstance(stock_prices_5y.index, pd.DatetimeIndex): stock_prices_5y.index = pd.to_datetime(stock_prices_5y.index)
                            if stock_prices_5y.index.tz is not None: stock_prices_5y.index = stock_prices_5y.index.tz_localize(None)
                            if not isinstance(market_prices_5y.index, pd.DatetimeIndex): market_prices_5y.index = pd.to_datetime(market_prices_5y.index)
                            if market_prices_5y.index.tz is not None: market_prices_5y.index = market_prices_5y.index.tz_localize(None)

                            stock_monthly_prices = stock_prices_5y.resample('ME').last().dropna()
                            market_monthly_prices = market_prices_5y.resample('ME').last().dropna()

                            if len(stock_monthly_prices) >= 12 and len(market_monthly_prices) >= 12:
                                base_gpcm['Stock_Monthly_Prices_5Y'] = stock_monthly_prices
                                base_gpcm['Market_Monthly_Prices_5Y'] = market_monthly_prices
                                
                                # Python-side Beta Calculation
                                s_ret = stock_monthly_prices.pct_change().dropna()
                                m_ret = market_monthly_prices.pct_change().dropna()
                                raw_5y, adj_5y = calculate_beta(s_ret, m_ret)
                                base_gpcm['Beta_5Y_Monthly_Raw'] = raw_5y
                                base_gpcm['Beta_5Y_Monthly_Adj'] = adj_5y
                                
                                # Unlevered Beta Calculation
                                if adj_5y is not None:
                                    equity_m = base_gpcm['Market_Cap_M'] + base_gpcm['NCI']
                                    base_gpcm['Unlevered_Beta_5Y'] = calculate_unlevered_beta(adj_5y, base_gpcm['IBD'], equity_m, base_gpcm['Tax_Rate'])
                            else:
                                st.warning(f"{ticker}: 월별 데이터가 부족합니다")
                                base_gpcm['Stock_Monthly_Prices_5Y'] = None; base_gpcm['Market_Monthly_Prices_5Y'] = None
                                base_gpcm['Beta_5Y_Monthly_Raw'] = None; base_gpcm['Beta_5Y_Monthly_Adj'] = None
                        else:
                            base_gpcm['Stock_Monthly_Prices_5Y'] = None; base_gpcm['Market_Monthly_Prices_5Y'] = None
                            base_gpcm['Beta_5Y_Monthly_Raw'] = None; base_gpcm['Beta_5Y_Monthly_Adj'] = None
                    else:
                        base_gpcm['Stock_Monthly_Prices_5Y'] = None; base_gpcm['Market_Monthly_Prices_5Y'] = None
                        base_gpcm['Beta_5Y_Monthly_Raw'] = None; base_gpcm['Beta_5Y_Monthly_Adj'] = None

                    # 2Y Weekly
                    start_2y = (base_dt - timedelta(days=365*2+20)).strftime('%Y-%m-%d')
                    stock_data_2y = None; market_data_2y = None
                    try:
                        stock_hist_2y = stock.history(start=start_2y, end=end_date, auto_adjust=False)
                        if not stock_hist_2y.empty: stock_data_2y = stock_hist_2y

                        market_data_2y = yf.download(market_idx, start=start_2y, end=end_date, progress=False)
                        if market_data_2y is not None and not market_data_2y.empty:
                            if not isinstance(market_data_2y.index, pd.DatetimeIndex): market_data_2y.index = pd.to_datetime(market_data_2y.index)
                            if market_data_2y.index.tz is not None: market_data_2y.index = market_data_2y.index.tz_localize(None)
                    except Exception as e: st.warning(f"{ticker} 2Y 데이터 수집 실패: {e}")

                    if stock_data_2y is not None and market_data_2y is not None:
                        if not stock_data_2y.empty and not market_data_2y.empty:
                            stock_prices_2y = stock_data_2y['Close'] if isinstance(stock_data_2y, pd.DataFrame) and 'Close' in stock_data_2y.columns else (stock_data_2y.iloc[:, 0] if isinstance(stock_data_2y, pd.DataFrame) else stock_data_2y)
                            market_prices_2y = market_data_2y['Close'] if isinstance(market_data_2y, pd.DataFrame) and 'Close' in market_data_2y.columns else (market_data_2y.iloc[:, 0] if isinstance(market_data_2y, pd.DataFrame) else market_data_2y)

                            if not isinstance(stock_prices_2y.index, pd.DatetimeIndex): stock_prices_2y.index = pd.to_datetime(stock_prices_2y.index)
                            if stock_prices_2y.index.tz is not None: stock_prices_2y.index = stock_prices_2y.index.tz_localize(None)
                            if not isinstance(market_prices_2y.index, pd.DatetimeIndex): market_prices_2y.index = pd.to_datetime(market_prices_2y.index)
                            if market_prices_2y.index.tz is not None: market_prices_2y.index = market_prices_2y.index.tz_localize(None)

                            stock_weekly_prices = stock_prices_2y.resample('W').last().dropna()
                            market_weekly_prices = market_prices_2y.resample('W').last().dropna()

                            if len(stock_weekly_prices) >= 50 and len(market_weekly_prices) >= 50:
                                base_gpcm['Stock_Weekly_Prices_2Y'] = stock_weekly_prices
                                base_gpcm['Market_Weekly_Prices_2Y'] = market_weekly_prices
                                
                                # Python-side Beta Calculation
                                s_ret_w = stock_weekly_prices.pct_change().dropna()
                                m_ret_w = market_weekly_prices.pct_change().dropna()
                                raw_2y, adj_2y = calculate_beta(s_ret_w, m_ret_w)
                                base_gpcm['Beta_2Y_Weekly_Raw'] = raw_2y
                                base_gpcm['Beta_2Y_Weekly_Adj'] = adj_2y
                                
                                # Unlevered Beta Calculation
                                if adj_2y is not None:
                                    equity_m = base_gpcm['Market_Cap_M'] + base_gpcm['NCI']
                                    base_gpcm['Unlevered_Beta_2Y'] = calculate_unlevered_beta(adj_2y, base_gpcm['IBD'], equity_m, base_gpcm['Tax_Rate'])
                            else:
                                base_gpcm['Stock_Weekly_Prices_2Y'] = None; base_gpcm['Market_Weekly_Prices_2Y'] = None
                                base_gpcm['Beta_2Y_Weekly_Raw'] = None; base_gpcm['Beta_2Y_Weekly_Adj'] = None
                        else:
                            base_gpcm['Stock_Weekly_Prices_2Y'] = None; base_gpcm['Market_Weekly_Prices_2Y'] = None
                            base_gpcm['Beta_2Y_Weekly_Raw'] = None; base_gpcm['Beta_2Y_Weekly_Adj'] = None
                    else:
                        base_gpcm['Stock_Weekly_Prices_2Y'] = None; base_gpcm['Market_Weekly_Prices_2Y'] = None
                        base_gpcm['Beta_2Y_Weekly_Raw'] = None; base_gpcm['Beta_2Y_Weekly_Adj'] = None
                except Exception as e: 
                    base_gpcm['Stock_Monthly_Prices_5Y'] = None; base_gpcm['Market_Monthly_Prices_5Y'] = None
                    base_gpcm['Beta_5Y_Monthly_Raw'] = None; base_gpcm['Beta_5Y_Monthly_Adj'] = None
                    base_gpcm['Beta_2Y_Weekly_Raw'] = None; base_gpcm['Beta_2Y_Weekly_Adj'] = None
                    base_gpcm['Stock_Weekly_Prices_2Y'] = None; base_gpcm['Market_Weekly_Prices_2Y'] = None

        except Exception as e:
            st.error(f"Error fetching {ticker}: {e}")
            continue

    # ========================================
    # [7] Target WACC 계산 (별도 시트용 - Base Label 'Y' 기준 데이터로만 수행)
    # ========================================
    # 7-1. 평균 부채비율 계산 (base label 기준)
    base_gpcm_list = list(all_period_data[base_label].values())
    debt_ratios = [g['Debt_Ratio'] for g in base_gpcm_list if g['Debt_Ratio'] > 0]
    if debt_ratios:
        avg_debt_ratio = np.mean(debt_ratios)
        st.info(f"📊 피어 평균 부채비율 (D/V): {avg_debt_ratio*100:.1f}%")
    else:
        avg_debt_ratio = 0.30

    # 7-2. Unlevered Beta 평균 계산 (선택된 beta_type에 따라)
    unlevered_betas = []
    beta_field = 'Unlevered_Beta_5Y' if beta_type == '5Y' else 'Unlevered_Beta_2Y'
    beta_label = "5Y Monthly" if beta_type == '5Y' else "2Y Weekly"

    for gpcm in base_gpcm_list:
        if gpcm[beta_field] is not None and gpcm[beta_field] > 0:
            unlevered_betas.append(gpcm[beta_field])

    if unlevered_betas:
        avg_unlevered_beta = np.mean(unlevered_betas)
        st.info(f"📊 피어 평균 Unlevered Beta ({beta_label}): {avg_unlevered_beta:.4f}")
    else:
        avg_unlevered_beta = 1.0

    # 7-3. Target의 Relevered Beta 계산
    if avg_debt_ratio < 1.0:
        target_de_ratio = avg_debt_ratio / (1 - avg_debt_ratio)
        target_relevered_beta = avg_unlevered_beta * (1 + (1 - target_tax_rate) * target_de_ratio)
    else:
        target_relevered_beta = avg_unlevered_beta

    # 7-4. Target의 자기자본비용 (Ke) 계산
    target_ke = rf_rate + mrp * target_relevered_beta + size_premium

    # 7-5. Target의 세후 타인자본비용 (Kd_aftertax) 계산
    target_kd_aftertax = kd_pretax * (1 - target_tax_rate)

    # 7-6. Target의 WACC 계산
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
    return all_period_data, raw_bs_rows, raw_pl_rows, market_rows, price_abs_dfs, price_rel_dfs, ticker_to_name, avg_debt_ratio, target_wacc_data


def create_excel(all_period_data, raw_bs_rows, raw_pl_rows, market_rows, price_abs_dfs, price_rel_dfs, base_period_str, target_periods, ticker_to_name, target_wacc_data, beta_type="5Y"):
    output = io.BytesIO()
    wb = Workbook(); wb.remove(wb.active)

    base_gpcm_data = all_period_data.get('Y', {})
    # Determine labels in order (Recent, Y, Y-1, Y-2...)
    def label_sort_key(x):
        if x == 'Recent': return -1
        if x == 'Y': return 0
        try: return int(x.split('-')[1])
        except: return 999
    rel_labels = sorted([k for k in all_period_data.keys() if all_period_data[k]], key=label_sort_key)
    
    # Sheet Colors
    COLOR_DARK_BLUE = '00338D'
    COLOR_PURPLE = '6A1B9A'

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
    fHL=Font(name='Arial',bold=True,size=9,color=C_DB); fSEC=Font(name='Arial',bold=True,size=10,color=C_W)
    fMUL=Font(name='Arial',bold=True,size=10,color=C_BL); fNOTE=Font(name='Arial',size=8,color=C_MG,italic=True)
    fSTAT=Font(name='Arial',bold=True,size=9,color=C_DB)
    fFRM=Font(name='Arial',size=9,color='000000'); fFRM_B=Font(name='Arial',bold=True,size=9,color='000000')
    fLINK=Font(name='Arial',size=9,color='008000'); fLINK_B=Font(name='Arial',bold=True,size=9,color='008000')

    pH=PatternFill('solid',fgColor=C_BL); pW=PatternFill('solid',fgColor=C_W)
    pST=PatternFill('solid',fgColor=C_LG); pSEC=PatternFill('solid',fgColor=C_DB)
    pSTAT=PatternFill('solid',fgColor=C_LB); pBETA=PatternFill('solid',fgColor='E8F5E9')

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

    # [Sheet 1] Multiples_Trend (Move to front as per user request)
    ws_trend = wb.create_sheet('Multiples_Trend', 0)
    ws_trend.sheet_properties.tabColor = COLOR_PURPLE
    
    ws_trend.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    sc(ws_trend.cell(1,1,'Multiples Trend Summary (Excel Formulas)'), fo=fT)
    ws_trend.merge_cells(start_row=2, start_column=1, end_row=2, end_column=20)
    sc(ws_trend.cell(2,1,'Base on dynamic BS/PL Sheets per period | Unit: Millions (local currency), Multiple (x)'), fo=fS)

    tr_r = 4
    trend_headers = ['Company', 'Ticker', 'Year Label', 'Period', 'Close', 'Shares', 'Mkt Cap', 'Cash', 'IBD', 'NCI', 'Equity', 'Revenue', 'EBIT', 'EBITDA', 'NI_Parent', 'EV', 'EV/EBITDA', 'PER', 'PBR', 'PSR']
    for i, h in enumerate(trend_headers, 1):
        sc(ws_trend.cell(tr_r, i, h), fo=fH, fi=pH, al=aC, bd=BD)
        
    ws_trend.column_dimensions['A'].width = 18
    ws_trend.column_dimensions['B'].width = 10
    ws_trend.column_dimensions['C'].width = 10 
    ws_trend.column_dimensions['D'].width = 12 
    for col_l in 'EFGHIJKLMNOP': ws_trend.column_dimensions[col_l].width = 14
    for col_l in 'QRST': ws_trend.column_dimensions[col_l].width = 11

    tr_r += 1
    mc_sht = "'Market_Cap'"
    for ticker_idx, (ticker, name) in enumerate(ticker_to_name.items()):
        for label in rel_labels:
            gpcm = all_period_data[label].get(ticker)
            if not gpcm: continue
            
            p_dt = gpcm.get('Base_Date', '-')
            row_fi = pW if ticker_idx % 2 == 0 else pST

            ws_trend.cell(tr_r, 1, name); sc(ws_trend.cell(tr_r, 1), fo=fA, fi=row_fi, al=aL, bd=BD)
            ws_trend.cell(tr_r, 2, ticker); sc(ws_trend.cell(tr_r, 2), fo=fA, fi=row_fi, al=aC, bd=BD)
            ws_trend.cell(tr_r, 3, label); sc(ws_trend.cell(tr_r, 3), fo=fA, fi=row_fi, al=aC, bd=BD)
            ws_trend.cell(tr_r, 4, p_dt); sc(ws_trend.cell(tr_r, 4), fo=fA, fi=row_fi, al=aC, bd=BD)
            
            # E, F, G: Close, Shares, Mkt Cap
            ws_trend.cell(tr_r, 5).value = f'=SUMIFS({mc_sht}!F:F, {mc_sht}!B:B, $B{tr_r}, {mc_sht}!H:H, $C{tr_r})'
            sc(ws_trend.cell(tr_r, 5), fo=fLINK, fi=row_fi, al=aR, bd=BD, nf=NF_PRC)
            ws_trend.cell(tr_r, 6).value = f'=SUMIFS({mc_sht}!G:G, {mc_sht}!B:B, $B{tr_r}, {mc_sht}!H:H, $C{tr_r})'
            sc(ws_trend.cell(tr_r, 6), fo=fLINK, fi=row_fi, al=aR, bd=BD, nf=NF_INT)
            ws_trend.cell(tr_r, 7).value = f'=SUMIFS({mc_sht}!I:I, {mc_sht}!B:B, $B{tr_r}, {mc_sht}!H:H, $C{tr_r})'
            sc(ws_trend.cell(tr_r, 7), fo=fLINK, fi=row_fi, al=aR, bd=BD, nf=NF_M1)

            # H, I, J, K: Cash, IBD, NCI, Equity
            bs_sn = "BS_FULL (base)" if label == 'Y' else f"BS_Full_{label}"
            pl_sn = "PL_Data (base)" if label == 'Y' else f"PL_Data_{label}"
            ws_trend.cell(tr_r, 8).value = f'=SUMIFS(\'{bs_sn}\'!G:G, \'{bs_sn}\'!B:B, $B{tr_r}, \'{bs_sn}\'!F:F, "Cash")'
            ws_trend.cell(tr_r, 9).value = f'=SUMIFS(\'{bs_sn}\'!G:G, \'{bs_sn}\'!B:B, $B{tr_r}, \'{bs_sn}\'!F:F, "IBD")'
            ws_trend.cell(tr_r, 10).value = f'=SUMIFS(\'{bs_sn}\'!G:G, \'{bs_sn}\'!B:B, $B{tr_r}, \'{bs_sn}\'!F:F, "NCI")'
            ws_trend.cell(tr_r, 11).value = f'=SUMIFS(\'{bs_sn}\'!G:G, \'{bs_sn}\'!B:B, $B{tr_r}, \'{bs_sn}\'!F:F, "Equity")'
            for c_idx in range(8, 12): sc(ws_trend.cell(tr_r, c_idx), fo=fLINK, fi=row_fi, al=aR, bd=BD, nf=NF_M)

            # L, M, N, O: Revenue, EBIT, EBITDA, NI_Parent
            # 모든 시트에서 J컬럼이 Amount임 (이미지 1 구조 통일)
            ws_trend.cell(tr_r, 12).value = f'=SUMIFS(\'{pl_sn}\'!$J:$J, \'{pl_sn}\'!$B:$B, $B{tr_r}, \'{pl_sn}\'!$E:$E, "Revenue")'
            ws_trend.cell(tr_r, 13).value = f'=SUMIFS(\'{pl_sn}\'!$J:$J, \'{pl_sn}\'!$B:$B, $B{tr_r}, \'{pl_sn}\'!$D:$D, "Operating Income")'
            ws_trend.cell(tr_r, 14).value = f'=M{tr_r} + SUMIFS(\'{pl_sn}\'!$J:$J, \'{pl_sn}\'!$B:$B, $B{tr_r}, \'{pl_sn}\'!$D:$D, "EBITDA") - SUMIFS(\'{pl_sn}\'!$J:$J, \'{pl_sn}\'!$B:$B, $B{tr_r}, \'{pl_sn}\'!$D:$D, "EBIT")'
            ws_trend.cell(tr_r, 15).value = f'=SUMIFS(\'{pl_sn}\'!$J:$J, \'{pl_sn}\'!$B:$B, $B{tr_r}, \'{pl_sn}\'!$E:$E, "NI_Parent")'
            for c_idx in [12, 13, 15]: sc(ws_trend.cell(tr_r, c_idx), fo=fLINK, fi=row_fi, al=aR, bd=BD, nf=NF_M)
            sc(ws_trend.cell(tr_r, 14), fo=fFRM_B, fi=row_fi, al=aR, bd=BD, nf=NF_M)

            # P: EV = Mkt Cap(G) + IBD(I) - Cash(H) + NCI(J)
            ws_trend.cell(tr_r, 16).value = f'=G{tr_r} + I{tr_r} - H{tr_r} + J{tr_r}'
            sc(ws_trend.cell(tr_r, 16), fo=fFRM_B, fi=PatternFill('solid',fgColor='E3F2FD'), al=aR, bd=BD, nf=NF_M)

            # Q, R, S, T: EV/EBITDA, PER, PBR, PSR
            ws_trend.cell(tr_r, 17).value = f'=IFERROR(IF(N{tr_r}>0, P{tr_r}/N{tr_r}, "N/M"), "N/M")'
            ws_trend.cell(tr_r, 18).value = f'=IFERROR(IF(O{tr_r}>0, G{tr_r}/O{tr_r}, "N/M"), "N/M")'
            ws_trend.cell(tr_r, 19).value = f'=IFERROR(IF(K{tr_r}>0, G{tr_r}/K{tr_r}, "N/M"), "N/M")'
            ws_trend.cell(tr_r, 20).value = f'=IFERROR(IF(L{tr_r}>0, G{tr_r}/L{tr_r}, "N/M"), "N/M")'
            for c_idx in range(17, 21): sc(ws_trend.cell(tr_r, c_idx), fo=fFRM_B, fi=PatternFill('solid',fgColor='FFF9C4'), al=aR, bd=BD, nf=NF_X)

            tr_r += 1

    ws_trend.auto_filter.ref = f"A4:T{tr_r-1}"
    ws_trend.freeze_panes = 'A5'

    # [Sheet 2] BS_Full
    for label in rel_labels:
        bs_rows_p = [r for r in raw_bs_rows if r.get('Label') == label]
        if bs_rows_p:
            sheet_name = "BS_FULL (base)" if label == 'Y' else f'BS_Full_{label}'
            ws_bs = wb.create_sheet(sheet_name)
            ws_bs.sheet_properties.tabColor = COLOR_DARK_BLUE
            df_bs = pd.DataFrame(bs_rows_p)
            cols = [('Company','Company',18),('Ticker','Ticker',10), ('Period','Period',12),('Currency','Curr',6), ('Account','Account',42),('EV_Tag','EV Tag',14), ('Amount_M','Amount (M)',18)]
            ws_bs.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_bs.cell(1,1,f'Balance Sheet ({label})'),fo=fT)
            r=3
            for i,(lb,key) in enumerate([('Cash','Cash'),('IBD','IBD'),('NCI','NCI'),('NOA(Option)','NOA(Option)'),('Equity','Equity')]):
                sc(ws_bs.cell(r,i+1,lb), fo=Font(name='Arial',size=8,bold=True),fi=ev_fills[key],al=aC,bd=BD)
            r=5
            for i,(col,disp,w) in enumerate(cols): ws_bs.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_bs.cell(r,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
            hdr=r; r+=1
            for _,rd in df_bs.iterrows():
                ev_tag=rd.get('EV_Tag', ''); is_hl=bool(ev_tag)
                row_fi=ev_fills.get(ev_tag, pST if r%2==0 else pW) if is_hl else (pST if r%2==0 else pW)
                row_font=fHL if is_hl else fA
                for i,(col,_,_) in enumerate(cols):
                    c=ws_bs.cell(r,i+1); v=rd.get(col, '')
                    if isinstance(v,(float,np.floating)): c.value=round(v,1) if pd.notna(v) else None
                    else: c.value=v
                    sc(c,fo=row_font,fi=row_fi,al=aR if col=='Amount_M' else aL,bd=BD,nf=NF_M if col=='Amount_M' else None)
                r+=1
            ws_bs.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r-1}"; ws_bs.freeze_panes=f'A{hdr+1}'

    # [Sheet 3] PL_Data
    for label in rel_labels:
        pl_rows_p = [r for r in raw_pl_rows if r.get('Label') == label]
        if pl_rows_p:
            is_base = (label == 'Y')
            sheet_name = "PL_Data (base)" if is_base else f'PL_Data_{label}'
            ws_pl = wb.create_sheet(sheet_name)
            ws_pl.sheet_properties.tabColor = COLOR_DARK_BLUE
            
            # 사용자 캡처 이미지 1의 컬럼 구성 및 순서 준수
            # Company, Ticker, Cur, Account, GPCM Tag, Unit, Source, Q Label, Period, Amount
            cols = [
                ('Company','Company',18),('Ticker','Ticker',10), ('Currency','Cur',6),
                ('Account','Account',42), ('GPCM_Tag','GPCM Tag',14), ('Unit','Unit',10),
                ('PL_Source','Source',16), ('Q_Label','Q Label',10), ('Period','Period',12),
                ('Amount_M','Amount',18)
            ]
            
            ws_pl.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_pl.cell(1,1,f'Income Statement ({label})'),fo=fT)
            r_idx=5
            for i,(col,disp,w) in enumerate(cols): ws_pl.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_pl.cell(r_idx,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
            hdr=r_idx; r_idx+=1
            
            df_pl = pd.DataFrame(pl_rows_p).sort_values(['Company','_sort','Q_Label'])
            for _,rd in df_pl.iterrows():
                is_hl=bool(rd.get('GPCM_Tag', '')); row_fi=ev_fills.get('PL_HL', pW) if is_hl else (pST if r_idx%2==0 else pW)
                row_font=fHL if is_hl else fA
                for i,(col,_,_) in enumerate(cols):
                    c=ws_pl.cell(r_idx,i+1); v=rd.get(col, '')
                    if isinstance(v,(float,np.floating)): c.value=round(v,1) if pd.notna(v) else None
                    else: c.value=v
                    is_eps=rd.get('Unit','')=='per share'; nf=NF_EPS if is_eps else NF_M
                    sc(c,fo=row_font,fi=row_fi,al=aR if col=='Amount_M' else aL,bd=BD,nf=nf if col=='Amount_M' else None)
                r_idx+=1
            ws_pl.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r_idx-1}"; ws_pl.freeze_panes=f'A{hdr+1}'

    # [Sheet 4] Market_Cap
    ws_mc = wb.create_sheet('Market_Cap')
    ws_mc.sheet_properties.tabColor = COLOR_DARK_BLUE
    if market_rows:
        df_mkt = pd.DataFrame(market_rows)
        cols = [('Company','Company',18),('Ticker','Ticker',10), ('Base_Date','Base Date',12),('Price_Date','Price Date',12), ('Currency','Curr',6),('Close','Close Price',14), ('Shares','Shares (Ord.)',18),('Label','Label',10), ('Market_Cap_M','Mkt Cap (M)',20)]
        ws_mc.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_mc.cell(1,1,'Market Capitalization'),fo=fT)
        ws_mc.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(cols)); sc(ws_mc.cell(2,1,'Mkt Cap = Ordinary Shares Number (자기주식 차감) × Close Price (auto_adjust=False)'), fo=fS)
        r=4
        for i,(col,disp,w) in enumerate(cols): ws_mc.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_mc.cell(r,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
        mc_hdr=r; r+=1
        for _,rd in df_mkt.iterrows():
            ev=(r%2==0)
            for i,(col,_,_) in enumerate(cols):
                c=ws_mc.cell(r,i+1); v=rd.get(col, '')
                if isinstance(v,(float,np.floating)): c.value=round(v,2) if pd.notna(v) else None
                else: c.value=v
                nf=NF_PRC if col=='Close' else (NF_INT if col=='Shares' else (NF_M1 if col=='Market_Cap_M' else None))
                sc(c,fo=fA,fi=pST if ev else pW,al=aR if nf else aL,bd=BD,nf=nf)
            r+=1
        ws_mc.auto_filter.ref=f"A{mc_hdr}:{get_column_letter(len(cols))}{r-1}"; ws_mc.freeze_panes=f'A{mc_hdr+1}'

    # [Sheet 5] Beta_Calculation
    ws_beta = wb.create_sheet('Beta_Calculation')
    ws_beta.sheet_properties.tabColor = COLOR_DARK_BLUE

    # 제목
    ws_beta.merge_cells('A1:F1')
    sc(ws_beta['A1'], fo=Font(name='Arial', bold=True, size=14, color=C_BL))
    ws_beta['A1'] = 'Beta Calculation (Excel Formulas)'

    ws_beta.merge_cells('A2:F2')
    sc(ws_beta['A2'], fo=Font(name='Arial', size=9, color=C_MG, italic=True))
    ws_beta['A2'] = f'5-Year Monthly & 2-Year Weekly Returns | Base: {base_period_str}'

    r_beta = 4

    # 각 ticker별로 베타 계산 섹션 생성
    beta_result_rows = {}  # ticker: (raw_5y, adj_5y, raw_2y, adj_2y) 매핑

    for idx, (ticker, gpcm) in enumerate(base_gpcm_data.items()):
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

    # [Sheet 7] GPCM (Valuation Summary)
    ws = wb.create_sheet('GPCM')
    ws.sheet_properties.tabColor = COLOR_PURPLE
    TOTAL_COLS = 36  # NOA 컬럼 추가 (35 → 36)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS); sc(ws.cell(1,1,'GPCM Valuation Summary with Beta Analysis'), fo=fT)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=TOTAL_COLS); sc(ws.cell(2,1,f'Base: {base_period_str} | Unit: Millions (local currency) | EV = MCap + IBD − Cash + NCI - NOA | Target WACC: See WACC_Calculation Sheet'), fo=fS)

    r=3
    # 섹션 위치 조정 (NOA 컬럼 추가 버전)
    sections = [
        (1,7,'Company Info'),         # A-G (Exchange, Mkt Index 포함)
        (8,7,'BS → EV Components'),    # H-N (NOA 추가)
        (15,4,'PL (LTM / Annual)'),   # O-R (Revenue at O)
        (19,3,'Market Data'),         # S-U
        (22,5,'Valuation Multiples'), # V-Z
        (27,10,'Beta & Risk Analysis') # AA-AJ
    ]
    for start,span,txt in sections:
        ws.merge_cells(start_row=r, start_column=start, end_row=r, end_column=start+span-1)
        sc(ws.cell(r,start,txt), fo=fSEC, fi=pSEC, al=aC, bd=BD)
        for c_idx in range(start,start+span): sc(ws.cell(r,c_idx), bd=BD)

    r=4
    # 헤더 구성 (NOA 추가)
    headers = ['Company','Ticker','Base Date','Curr', 'PL Source', 'Exchange', 'Mkt Index',
               'Cash','IBD','Net Debt','NCI','NOA','Equity','EV',
               'Revenue','EBIT','EBITDA','NI (Parent)',
               'Price','Shares','Mkt Cap',
               'EV/EBITDA','EV/EBIT','PER','PBR','PSR',
               'β 5Y Raw','β 5Y Adj','β 2Y Raw','β 2Y Adj','Pretax Inc','Tax Rate','D/E Ratio','Debt Ratio','Unlevered β 5Y','Unlevered β 2Y']
    widths = [18,10,11,6,16, 12, 12,
              14,14,14,12,12,14,16,
              14,14,14,16,
              12,18,20,
              12,12,12,12,12,
              10,10,10,10,14,10,10,10,14,14]
    for i,(h,w) in enumerate(zip(headers,widths)): 
        ws.column_dimensions[get_column_letter(i+1)].width=w
        sc(ws.cell(r,i+1,h),fo=fH,fi=pH,al=aC,bd=BD)
    
    DATA_START=5; n_companies=len(base_gpcm_data); DATA_END=DATA_START+n_companies-1
    MC_DATA_START=5
    NF_BETA='0.00;(0.00);"-"'; NF_PCT='0.0%;(0.0%);"-"'; NF_RATIO='0.00;(0.00);"-"'
    
    bs_sn = "BS_FULL (base)"
    pl_sn = "PL_Data (base)"
    
    for idx,(ticker, gpcm) in enumerate(base_gpcm_data.items()):
        r=DATA_START+idx; mc_row=MC_DATA_START+idx; ev_row=(r%2==0); base_fi=pST if ev_row else pW
        # A-G: Company Info (Exchange, Mkt Index 포함)
        vals=[gpcm['Company'],ticker,gpcm['Base_Date'],gpcm['Currency'],gpcm['PL_Source'], gpcm.get('Exchange',''), gpcm.get('Market_Index','')]
        for ci,v in enumerate(vals,1): ws.cell(r,ci,v); sc(ws.cell(r,ci), fo=fA, fi=base_fi, al=aL, bd=BD)

        # H-N: BS → EV Components (Formulas)
        ws.cell(r,8).value=f'=SUMIFS(\'{bs_sn}\'!$G:$G,\'{bs_sn}\'!$B:$B,$B{r},\'{bs_sn}\'!$F:$F,"Cash")'; sc(ws.cell(r,8), fo=fLINK_B, fi=ev_fills['Cash'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,9).value=f'=SUMIFS(\'{bs_sn}\'!$G:$G,\'{bs_sn}\'!$B:$B,$B{r},\'{bs_sn}\'!$F:$F,"IBD")'; sc(ws.cell(r,9), fo=fLINK_B, fi=ev_fills['IBD'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,10).value=f'=I{r}-H{r}'; sc(ws.cell(r,10), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_M)
        ws.cell(r,11).value=f'=SUMIFS(\'{bs_sn}\'!$G:$G,\'{bs_sn}\'!$B:$B,$B{r},\'{bs_sn}\'!$F:$F,"NCI")'; sc(ws.cell(r,11), fo=fLINK_B, fi=ev_fills['NCI'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,12).value=f'=SUMIFS(\'{bs_sn}\'!$G:$G,\'{bs_sn}\'!$B:$B,$B{r},\'{bs_sn}\'!$F:$F,"NOA")'; sc(ws.cell(r,12), fo=fLINK_B, fi=ev_fills['NOA(Option)'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,13).value=f'=SUMIFS(\'{bs_sn}\'!$G:$G,\'{bs_sn}\'!$B:$B,$B{r},\'{bs_sn}\'!$F:$F,"Equity")'; sc(ws.cell(r,13), fo=fLINK_B, fi=ev_fills['Equity'], al=aR, bd=BD, nf=NF_M)
        # N (EV) = MCap(U) + IBD(I) - Cash(H) + NCI(K) - NOA(L)
        ws.cell(r,14).value=f'=U{r}+I{r}-H{r}+K{r}-L{r}'; sc(ws.cell(r,14), fo=fFRM_B, fi=PatternFill('solid',fgColor=C_PB), al=aR, bd=BD, nf=NF_M)

        # O-R: PL (LTM/Annual)
        # Revenue(O=15), EBIT(P=16), EBITDA(Q=17), NI_Parent(R=18)
        ws.cell(r,15).value=f'=SUMIFS(\'{pl_sn}\'!$J:$J,\'{pl_sn}\'!$B:$B,$B{r},\'{pl_sn}\'!$E:$E,"Revenue")'; sc(ws.cell(r,15), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)
        # EBIT: PL_Data D열(Account)이 "Operating Income"인 행만 합산
        ws.cell(r,16).value=f'=SUMIFS(\'{pl_sn}\'!$J:$J,\'{pl_sn}\'!$B:$B,$B{r},\'{pl_sn}\'!$D:$D,"Operating Income")'; sc(ws.cell(r,16), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,17).value=f'=P{r}+SUMIFS(\'{pl_sn}\'!$J:$J,\'{pl_sn}\'!$B:$B,$B{r},\'{pl_sn}\'!$D:$D,"EBITDA")-SUMIFS(\'{pl_sn}\'!$J:$J,\'{pl_sn}\'!$B:$B,$B{r},\'{pl_sn}\'!$D:$D,"EBIT")'; sc(ws.cell(r,17), fo=fFRM_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)
        ws.cell(r,18).value=f'=SUMIFS(\'{pl_sn}\'!$J:$J,\'{pl_sn}\'!$B:$B,$B{r},\'{pl_sn}\'!$E:$E,"NI_Parent")'; sc(ws.cell(r,18), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)

        # S-U: Market Data (Ticker와 Label 'Y'를 기준으로 SUMIFS 조회)
        # Price(S=19), Shares(T=20), Mkt Cap(U=21)
        ws.cell(r,19).value=f'=SUMIFS(Market_Cap!$F:$F,Market_Cap!$B:$B,$B{r},Market_Cap!$H:$H,"Y")'; sc(ws.cell(r,19), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_PRC)
        ws.cell(r,20).value=f'=SUMIFS(Market_Cap!$G:$G,Market_Cap!$B:$B,$B{r},Market_Cap!$H:$H,"Y")'; sc(ws.cell(r,20), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_INT)
        ws.cell(r,21).value=f'=SUMIFS(Market_Cap!$I:$I,Market_Cap!$B:$B,$B{r},Market_Cap!$H:$H,"Y")'; sc(ws.cell(r,21), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_M1)

        # V-Z: Valuation Multiples
        # EV(N=14), EBITDA(Q=17), EBIT(P=16), NI(R=18), Rev(O=15), BV(M=13)
        pMULT=PatternFill('solid',fgColor=C_PB)
        ws.cell(r,22).value=f'=IF(Q{r}>0,N{r}/Q{r},"N/M")'; sc(ws.cell(r,22), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,23).value=f'=IF(P{r}>0,N{r}/P{r},"N/M")'; sc(ws.cell(r,23), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,24).value=f'=IF(R{r}>0,U{r}/R{r},"N/M")'; sc(ws.cell(r,24), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,25).value=f'=IF(M{r}>0,U{r}/M{r},"N/M")'; sc(ws.cell(r,25), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)
        ws.cell(r,26).value=f'=IF(O{r}>0,U{r}/O{r},"N/M")'; sc(ws.cell(r,26), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

        # AA-AJ: Beta & Risk Analysis
        raw_5y_row, adj_5y_row, raw_2y_row, adj_2y_row = beta_result_rows.get(ticker, (None, None, None, None))
        # AA(27), AB(28), AC(29), AD(30)
        if raw_5y_row is not None: ws.cell(r,27).value = f'=Beta_Calculation!$B${raw_5y_row}'
        sc(ws.cell(r,27), fo=fLINK, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)
        if adj_5y_row is not None: ws.cell(r,28).value = f'=Beta_Calculation!$B${adj_5y_row}'
        sc(ws.cell(r,28), fo=fLINK, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)
        if raw_2y_row is not None: ws.cell(r,29).value = f'=Beta_Calculation!$B${raw_2y_row}'
        sc(ws.cell(r,29), fo=fLINK, fi=PatternFill('solid',fgColor='FFF9C4'), al=aR, bd=BD, nf=NF_BETA)
        if adj_2y_row is not None: ws.cell(r,30).value = f'=Beta_Calculation!$B${adj_2y_row}'
        sc(ws.cell(r,30), fo=fLINK, fi=PatternFill('solid',fgColor='FFF9C4'), al=aR, bd=BD, nf=NF_BETA)

        # Pretax Income (AE=31)
        ws.cell(r,31).value=f'=SUMIFS(\'{pl_sn}\'!$J:$J,\'{pl_sn}\'!$B:$B,$B{r},\'{pl_sn}\'!$D:$D,"Pretax Income")'; sc(ws.cell(r,31), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_M)
        
        # Tax Rate (AF=32)
        ws.cell(r,32,gpcm['Tax_Rate']); sc(ws.cell(r,32), fo=fA, fi=base_fi, al=aR, bd=BD, nf=NF_PCT)
        
        # D/E Ratio (AG=33) = IBD(I=9) / (MCap(U=21) + NCI(K=11))
        ws.cell(r,33).value=f'=IF((U{r}+K{r})>0,I{r}/(U{r}+K{r}),0)'; sc(ws.cell(r,33), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_RATIO)
        
        # Debt Ratio (AH=34) = IBD(I=9) / (MCap(U=21) + IBD(I=9) + NCI(K=11))
        ws.cell(r,34).value=f'=IF((U{r}+I{r}+K{r})>0,I{r}/(U{r}+I{r}+K{r}),0)'; sc(ws.cell(r,34), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_RATIO)
        
        # Unlevered Betas (AI=35, AJ=36)
        # Levered Adj Beta(AB, AD) / (1 + (1-Tax)*D/E)
        ws.cell(r,35).value=f'=IF(AB{r}<>0,AB{r}/(1+(1-AF{r})*AG{r}),"")'; sc(ws.cell(r,35), fo=fFRM_B, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)
        ws.cell(r,36).value=f'=IF(AD{r}<>0,AD{r}/(1+(1-AF{r})*AG{r}),"")'; sc(ws.cell(r,36), fo=fFRM_B, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

    # Stats
    r=DATA_END+2
    stat_labels=['Mean','Median','Max','Min']; func_map={'Mean':'AVERAGE','Median':'MEDIAN','Max':'MAX','Min':'MIN'}
    # Multiples: 22-26 (EV/EBITDA, EV/EBIT, PER, PBR, PSR)
    # Betas: 27-30, 35-36 (Beta 5Y Raw, Beta 5Y Adj, Beta 2Y Raw, Beta 2Y Adj, Unlevered Beta 5Y, Unlevered Beta 2Y)
    # Ratios: 33-34 (D/E Ratio, Debt Ratio)
    mult_cols=[22,23,24,25,26]
    beta_cols=[27,28,29,30,35,36]
    ratio_cols=[33,34]  # D/E Ratio, Debt Ratio

    for sn in stat_labels:
        ws.cell(r,21,sn); sc(ws.cell(r,21), fo=fSTAT, fi=pSTAT, al=aC, bd=BD) # Price(S=19), Shares(T=20), Mkt Cap(U=21)
        fn=func_map[sn]
        # Multiples: 22-26
        for ci in mult_cols:
            col=get_column_letter(ci)
            ws.cell(r,ci).value=f'=IFERROR({fn}({col}{DATA_START}:{col}{DATA_END}),"N/M")'
            sc(ws.cell(r,ci), fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_X)
        # Betas: 27, 28, 29, 30, 35, 36
        for ci in beta_cols:
            col=get_column_letter(ci)
            ws.cell(r,ci).value=f'=IFERROR({fn}({col}{DATA_START}:{col}{DATA_END}),"N/M")'
            sc(ws.cell(r,ci), fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_BETA)
        # Ratios: 33, 34
        for ci in ratio_cols:
            col=get_column_letter(ci)
            ws.cell(r,ci).value=f'=IFERROR({fn}({col}{DATA_START}:{col}{DATA_END}),"N/M")'
            sc(ws.cell(r,ci), fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_RATIO)
        r+=1
    
    # Notes
    r+=2
    notes = [
        '[ Valuation Methodology Notes ]',
        f'• Base Date: {base_period_str} | Unit: Millions (local currency)',
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
        '  - 한국 주식 (.KS, .KQ): 주가 yfinance, 시장지수 FinanceDataReader (KS11, KQ11)',
        '  - 해외 주식: Yahoo Finance (yfinance)',
        '• Beta 계산 방법:',
        '  - 5Y Monthly Beta: 5년간 월말 종가 기준 월간 수익률 계산 → 시장지수 대비 선형회귀',
        '  - 2Y Weekly Beta: 2년간 주말 종가 기준 주간 수익률 계산 → 시장지수 대비 선형회귀',
        '  - Raw Beta = Slope of linear regression (Market vs Stock returns)',
        '  - Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1.0 (Bloomberg 방법론)',
        '• Market Index: KOSPI (KS11), KOSDAQ (KQ11), Nikkei 225 (^N225), S&P/TSX (^GSPTSE), etc.',
        '• 값 검증: NaN, inf, 극단값(-10 ~ 10 범위 벗어남) 필터링 → None 처리',
        '• Tax Rate: KPMG Corporate Tax Rates Table (2025 Combined Rate) 기준',
        '   - Korea: ≤ 200M: 9.9% | 200M-20,000M: 20.9% | 20,000M-300,000M: 23.1% | > 300,000M: 26.4%',
        '• D/E Ratio = IBD ÷ (Market Cap + NCI)',
        '• Debt Ratio (D/V) = IBD ÷ (Market Cap + IBD + NCI) [총부채/총자산]',
        '• Unlevered Beta = Levered Beta ÷ (1 + (1 - Tax Rate) × D/E Ratio) [Hamada Model]',
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
        '', '⚠ 주가 데이터: Yahoo Finance (yfinance) | 한국 시장지수: FinanceDataReader | Verify with official sources.'
    ]
    for note in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_COLS)
        sc(ws.cell(r,1,note), fo=fNOTE); r+=1

    # 틀고정: BS → EV Components 이전 (I5 = Company/Ticker/등 고정, BS부터 스크롤)
    ws.freeze_panes='I5'

    # [Sheet 6] WACC_Calculation (Target 기업의 WACC 계산)
    ws_wacc = wb.create_sheet('WACC_Calculation')
    ws_wacc.sheet_properties.tabColor = COLOR_DARK_BLUE

    ws_wacc.merge_cells('A2:D2')
    sc(ws_wacc['A2'], fo=Font(name='Arial', size=9, color=C_MG, italic=True))
    ws_wacc['A2'] = f'Base: {base_period_str} | Peer Average Method'

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
    beta_col = 'AI' if beta_type == "5Y" else 'AJ'  # AI = 컬럼 35 (Unlevered Beta 5Y), AJ = 컬럼 36 (Unlevered Beta 2Y)
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
    ws_wacc.cell(r_wacc, 2).value = f'=GPCM!AH{mean_row}'  # 컬럼 34 (AH) = Debt Ratio
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

    # [Sheet 11] Price_History (Set tab color to black)
    if price_abs_dfs:
        ws_ph = wb.create_sheet('Price_History')
        ws_ph.sheet_properties.tabColor = "000000" # Black tab
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

    # Temp 시트 삭제
    if 'Temp' in wb.sheetnames:
        del wb['Temp']

    # [Reorder Sheets] Strictly per user request:
    # 1. GPCM, 2. BS_FULL (base), 3. PL_Data (base), 4. WACC_Calculation, 5. Beta_Calculation, 6. Multiples_Trend, 7. Market_Cap, 8. Price_History, 9. Historicals...
    
    desired_order = ['GPCM', 'BS_FULL (base)', 'PL_Data (base)', 'WACC_Calculation', 'Beta_Calculation', 'Multiples_Trend', 'Market_Cap']
    
    # Get all actual sheet names
    current_sheets = wb.sheetnames
    
    # Sort history sheets (Y-1, Y-2...)
    hist_sheets = [s for s in current_sheets if any(x in s for x in ['BS_Full_Y-', 'PL_Data_Y-', 'BS_최신', 'PL_최신'])]
    def hist_sort(s):
        if '최신' in s: return -1
        try: return int(s.split('-')[1])
        except: return 999
    hist_sheets.sort(key=hist_sort)
    
    final_order = []
    for d in desired_order:
        if d in current_sheets:
            final_order.append(d)
    
    # Add history sheets
    for s in hist_sheets:
        if s not in final_order:
            final_order.append(s)
            
    # Add any remaining sheets (excluding Price_History)
    for s in current_sheets:
        if s not in final_order and s != 'Price_History':
            final_order.append(s)
            
    # [CRITICAL] ALWAYS Price_History at the VERY END
    if 'Price_History' in current_sheets:
        final_order.append('Price_History')
            
    # Apply order
    new_sheets = []
    for name in final_order:
        new_sheets.append(wb[name])
    wb._sheets = new_sheets

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

def export_historical_excel_global_v2(all_period_data, raw_bs_rows, raw_pl_rows, market_rows, target_periods, ticker_to_name):
    import io, pandas as pd, numpy as np
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()
    wb = Workbook(); wb.remove(wb.active)

    # Determine labels in order (Recent, Y, Y-1, Y-2...)
    def label_sort_key(x):
        if x == 'Recent': return -1
        if x == 'Y': return 0
        try: return int(x.split('-')[1])
        except: return 999
    rel_labels = sorted([k for k in all_period_data.keys() if all_period_data[k]], key=label_sort_key)
    
    # Sheet Colors
    COLOR_DARK_BLUE = '00338D'
    COLOR_PURPLE = '6A1B9A'

    # Styles & Colors (NameError sc 방지용 로컬 정의)
    C_BL='00338D'; C_DB='1E2A5E'; C_MB='005EB8'; C_LB='C3D7EE'; C_PB='E8EFF8'
    C_DG='333333'; C_MG='666666'; C_LG='F5F5F5'; C_BG='B0B0B0'; C_W='FFFFFF'
    C_GR='E2EFDA'; C_YL='FFF8E1'; C_NOA='FCE4EC'

    S1=Side(style='thin',color=C_BG); BD=Border(left=S1,right=S1,top=S1,bottom=S1)
    fT=Font(name='Arial',bold=True,size=14,color=C_BL); fS=Font(name='Arial',size=9,color=C_MG,italic=True)
    fH=Font(name='Arial',bold=True,size=9,color=C_W); fA=Font(name='Arial',size=9,color=C_DG)
    fHL=Font(name='Arial',bold=True,size=9,color=C_DB)
    fFRM=Font(name='Arial',size=9,color='000000'); fFRM_B=Font(name='Arial',bold=True,size=9,color='000000')
    fLINK=Font(name='Arial',size=9,color='008000')

    pH=PatternFill('solid',fgColor=C_BL); pW=PatternFill('solid',fgColor=C_W)
    pST=PatternFill('solid',fgColor=C_LG); pSEC2=PatternFill('solid',fgColor=C_DB)
    pSEC1=PatternFill('solid',fgColor='1565C0') # 진한 파랑

    ev_fills = {'Cash':PatternFill('solid',fgColor=C_GR), 'IBD':PatternFill('solid',fgColor=C_YL),
                'NCI':PatternFill('solid',fgColor=C_PB), 'NOA(Option)':PatternFill('solid',fgColor=C_NOA),
                'Equity':PatternFill('solid',fgColor=C_LB), 'PL_HL':PatternFill('solid',fgColor=C_YL),
                'NI_Parent':PatternFill('solid',fgColor=C_YL)}
    
    aC=Alignment(horizontal='center',vertical='center',wrap_text=True)
    aL=Alignment(horizontal='left',vertical='center',indent=1)
    aR=Alignment(horizontal='right',vertical='center')
    NF_M='#,##0;(#,##0);"-"'; NF_M1='#,##0.0;(#,##0.0);"-"'; NF_PRC='#,##0.00;(#,##0.00);"-"'
    NF_INT='#,##0;(#,##0);"-"'; NF_EPS='#,##0.00;(#,##0.00);"-"'; NF_X='0.0"x";(0.0"x");"-"'; NF_PCT='0.0%;(0.0%);"-"'

    def sc(c,fo=None,fi=None,al=None,bd=None,nf=None):
        if fo:c.font=fo
        if fi:c.fill=fi
        if al:c.alignment=al
        if bd:c.border=bd
        if nf:c.number_format=nf

    # [1] Summary Sheet
    ws_summ = wb.create_sheet('Historical_Summary')
    ws_summ.sheet_properties.tabColor = COLOR_PURPLE
    
    ws_summ.merge_cells('A1:AA1')
    ws_summ['A1'] = f"Historical Financial Summary - yfinance Base (Label: Y, Y-1...)"
    sc(ws_summ['A1'], fo=fT)
    ws_summ.merge_cells('A2:AA2')
    ws_summ['A2'] = "Base on dynamic BS/PL Sheets per period | Unit: Millions (local currency)"
    sc(ws_summ['A2'], fo=fS)
    
    # ... (metrics definition remains same)
    
    mc_map = {} 
    
    ws_summ.cell(row=3, column=1, value="Company"); sc(ws_summ.cell(row=3, column=1), fo=fH, fi=pSEC1, al=aC, bd=BD)
    ws_summ.cell(row=3, column=2, value="Ticker"); sc(ws_summ.cell(row=3, column=2), fo=fH, fi=pSEC1, al=aC, bd=BD)
    ws_summ.merge_cells('A3:A4')
    ws_summ.merge_cells('B3:B4')

    col_idx = 3
    for m_key, m_name, _, _ in [
        ('Financial Date', 'Financial Date', 'Formula_Date', None),
        ('Revenue', 'Revenue(매출)', 'PL_Tag', NF_M), 
        ('Gross Profit', 'Gross Profit', 'PL_Acc', NF_M), 
        ('Operating Income', 'EBIT(영업이익)', 'PL_Acc', NF_M), 
        ('EBITDA_Calc', 'EBITDA', 'Formula', NF_M), 
        ('NI_Parent', 'Net Income(순이익)', 'PL_Tag', NF_M),
        ('Total Assets', 'Total Assets(자산)', 'BS_Acc', NF_M), 
        ('Total Liabilities Net Minority Interest', 'Total Liab(부채)', 'BS_Acc', NF_M), 
        ('Equity', 'Equity(자본)', 'BS_Tag', NF_M),
        ('Cash', 'Cash', 'BS_Tag', NF_M), 
        ('IBD', 'IBD(차입금)', 'BS_Tag', NF_M),
        ('NCI', 'NCI(비지배지분)', 'BS_Tag', NF_M),
        ('NetDebt', 'Net Debt(순부채)', 'Formula', NF_M),
        ('OPM', 'OPM(영업이익률)', 'Formula', NF_PCT), 
        ('DebtRatio', 'Debt Ratio(부채비율)', 'Formula', NF_PCT),
        ('Mkt_Cap', 'Market Cap(시가총액)', 'Mkt', NF_M1),
        ('EV', 'EV(기업가치)', 'Formula', NF_M),
    ]:
        start_col = col_idx
        end_col = col_idx + len(rel_labels) - 1
        ws_summ.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
        ws_summ.cell(row=3, column=start_col, value=m_name)
        sc(ws_summ.cell(row=3, column=start_col), fo=fH, fi=pSEC1, al=aC, bd=BD)
        
        for label in rel_labels:
            ws_summ.cell(row=4, column=col_idx, value=label)
            sc(ws_summ.cell(row=4, column=col_idx), fo=fH, fi=pSEC2, al=aC, bd=BD)
            mc_map[(m_key, label)] = get_column_letter(col_idx)
            col_idx += 1

    r = 5
    for ticker, comp_name in ticker_to_name.items():
        ws_summ.cell(row=r, column=1, value=comp_name); sc(ws_summ.cell(row=r, column=1), fo=fA, bd=BD)
        ws_summ.cell(row=r, column=2, value=ticker); sc(ws_summ.cell(row=r, column=2), fo=fA, al=aC, bd=BD)
        
        c = 3
        for m_key, m_name, m_type, fmt in [
            ('Financial Date', 'Financial Date', 'Formula_Date', None),
            ('Revenue', 'Revenue(매출)', 'PL_Tag', NF_M), 
            ('Gross Profit', 'Gross Profit', 'PL_Acc', NF_M), 
            ('Operating Income', 'EBIT(영업이익)', 'PL_Acc', NF_M), 
            ('EBITDA_Calc', 'EBITDA', 'Formula', NF_M), 
            ('NI_Parent', 'Net Income(순익)', 'PL_Tag', NF_M),
            ('Total Assets', 'Total Assets(자산)', 'BS_Acc', NF_M), 
            ('Total Liabilities Net Minority Interest', 'Total Liab(부채)', 'BS_Acc', NF_M), 
            ('Equity', 'Equity(자본)', 'BS_Tag', NF_M),
            ('Cash', 'Cash', 'BS_Tag', NF_M), 
            ('IBD', 'IBD(차입금)', 'BS_Tag', NF_M),
            ('NCI', 'NCI(비지배지분)', 'BS_Tag', NF_M),
            ('NetDebt', 'Net Debt(순부채)', 'Formula', NF_M),
            ('OPM', 'OPM(영업이익률)', 'Formula', NF_PCT), 
            ('DebtRatio', 'Debt Ratio(부채비율)', 'Formula', NF_PCT),
            ('Mkt_Cap', 'Market Cap(시가총액)', 'Mkt', NF_M1),
            ('EV', 'EV(기업가치)', 'Formula', NF_M),
        ]:
            for label in rel_labels:
                is_recent = (label == 'Recent')
                bs_sn = 'BS_최신' if is_recent else f"BS_Full_{label}"
                pl_sn = 'PL_최신' if is_recent else f"PL_Data_{label}"
                
                v = ""
                if m_type == 'Formula_Date':
                    # Get actual fiscal date for this peer and label
                    gpcm_p = all_period_data[label].get(ticker)
                    v = gpcm_p.get('Base_Date', '-') if gpcm_p else "-"
                elif m_type == 'BS_Tag':
                    v = f'=SUMIFS(\'{bs_sn}\'!G:G, \'{bs_sn}\'!B:B, $B{r}, \'{bs_sn}\'!F:F, "{m_key}")'
                elif m_type == 'BS_Acc':
                    v = f'=SUMIFS(\'{bs_sn}\'!G:G, \'{bs_sn}\'!B:B, $B{r}, \'{bs_sn}\'!E:E, "{m_key}")'
                elif m_type == 'PL_Tag':
                    v = f'=SUMIFS(\'{pl_sn}\'!J:J, \'{pl_sn}\'!B:B, $B{r}, \'{pl_sn}\'!E:E, "{m_key}")'
                elif m_type == 'PL_Acc':
                    v = f'=SUMIFS(\'{pl_sn}\'!J:J, \'{pl_sn}\'!B:B, $B{r}, \'{pl_sn}\'!D:D, "{m_key}")'
                elif m_type == 'Mkt':
                    v = f'=SUMIFS(Market_Cap!I:I, Market_Cap!B:B, $B{r}, Market_Cap!H:H, "{label}")'
                elif m_type == 'Formula':
                    if m_key == 'EBITDA_Calc':
                        ebit_addr = f"{mc_map[('Operating Income', label)]}{r}"
                        v = f'={ebit_addr} + SUMIFS(\'{pl_sn}\'!J:J, \'{pl_sn}\'!B:B, $B{r}, \'{pl_sn}\'!D:D, "EBITDA") - SUMIFS(\'{pl_sn}\'!J:J, \'{pl_sn}\'!B:B, $B{r}, \'{pl_sn}\'!D:D, "EBIT")'
                    elif m_key == 'NetDebt':
                        v = f"={mc_map[('IBD', label)]}{r} - {mc_map[('Cash', label)]}{r}"
                    elif m_key == 'OPM':
                        v = f'=IFERROR({mc_map[('Operating Income', label)]}{r}/{mc_map[('Revenue', label)]}{r}, "")'
                    elif m_key == 'DebtRatio':
                        v = f'=IFERROR({mc_map[('Total Liabilities Net Minority Interest', label)]}{r}/{mc_map[('Equity', label)]}{r}, "")'
                    elif m_key == 'EV':
                        v = f"={mc_map[('Mkt_Cap', label)]}{r} + {mc_map[('IBD', label)]}{r} - {mc_map[('Cash', label)]}{r} + {mc_map[('NCI', label)]}{r}"

                ws_summ.cell(row=r, column=c, value=v)
                font_style = fFRM_B if m_type == 'Formula' else (fA if m_type == 'Formula_Date' else fLINK)
                sc(ws_summ.cell(row=r, column=c), fo=font_style, nf=fmt, bd=BD)
                c += 1
        r += 1

    ws_summ.column_dimensions['A'].width = 18
    ws_summ.column_dimensions['B'].width = 10
    for i in range(3, c):
        ws_summ.column_dimensions[get_column_letter(i)].width = 14
    ws_summ.freeze_panes = "C5"

    # [2] BS_Full, PL_Data 시트들 생성
    # Recent 시트들을 먼저 생성하여 Summary 바로 뒤에 오도록 함 (index 1, 2)
    # bs_idx = 1 # Historical_Summary가 0번
    for label in rel_labels:
        bs_rows_p = [row for row in raw_bs_rows if row.get('Label') == label]
        if bs_rows_p:
            is_recent = (label == 'Recent')
            sheet_title = 'BS_최신' if is_recent else f'BS_Full_{label}'
            # Recent면 인덱스 1에 삽입, 아니면 맨 뒤에 추가
            ws_bs = wb.create_sheet(sheet_title, 1 if is_recent else None)
            ws_bs.sheet_properties.tabColor = COLOR_DARK_BLUE
            
            df_bs = pd.DataFrame(bs_rows_p)
            cols = [('Company','Company',18),('Ticker','Ticker',10), ('Period','Period',12),('Currency','Curr',6), ('Account','Account',42),('EV_Tag','EV Tag',14), ('Amount_M','Amount (M)',18)]
            ws_bs.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_bs.cell(1,1,f'Balance Sheet ({label})'),fo=fT)
            r_idx=3
            for i,(lb,key) in enumerate([('Cash','Cash'),('IBD','IBD'),('NCI','NCI'),('NOA(Option)','NOA(Option)'),('Equity','Equity')]):
                sc(ws_bs.cell(r_idx,i+1,lb), fo=Font(name='Arial',size=8,bold=True),fi=ev_fills[key],al=aC,bd=BD)
            r_idx=5
            for i,(col,disp,w) in enumerate(cols): ws_bs.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_bs.cell(r_idx,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
            hdr=r_idx; r_idx+=1
            for _,rd in df_bs.iterrows():
                ev_tag=rd.get('EV_Tag', ''); is_hl=bool(ev_tag)
                row_fi=ev_fills.get(ev_tag, pST if r_idx%2==0 else pW) if is_hl else (pST if r_idx%2==0 else pW)
                row_font=fHL if is_hl else fA
                for i,(col,_,_) in enumerate(cols):
                    c_cell=ws_bs.cell(r_idx,i+1); v=rd.get(col, '')
                    if isinstance(v,(float,np.floating)): c_cell.value=round(v,1) if pd.notna(v) else None
                    else: c_cell.value=v
                    sc(c_cell,fo=row_font,fi=row_fi,al=aR if col=='Amount_M' else aL,bd=BD,nf=NF_M if col=='Amount_M' else None)
                r_idx+=1
            ws_bs.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r_idx-1}"; ws_bs.freeze_panes=f'A{hdr+1}'

    for label in rel_labels:
        pl_rows_p = [row for row in raw_pl_rows if row.get('Label') == label]
        if pl_rows_p:
            is_recent = (label == 'Recent')
            sheet_title = 'PL_최신' if is_recent else f'PL_Data_{label}'
            # Recent면 인덱스 2에 삽입 (BS_최신이 인덱스 1이므로 그 뒤), 아니면 맨 뒤에 추가
            ws_pl = wb.create_sheet(sheet_title, 2 if is_recent else None)
            ws_pl.sheet_properties.tabColor = COLOR_DARK_BLUE
            
            df_pl = pd.DataFrame(pl_rows_p).sort_values(['Company','_sort','Q_Label'])
            cols = [('Company','Company',18),('Ticker','Ticker',10), ('Currency','Curr',6),('Account','Account',42), ('GPCM_Tag','GPCM Tag',14),('Unit','Unit',10), ('PL_Source','Source',14),('Q_Label','Q Label',9), ('Period','Period',12),('Amount_M','Amount',18)]
            ws_pl.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_pl.cell(1,1,f'Income Statement ({label})'),fo=fT)
            r_idx=5
            for i,(col,disp,w) in enumerate(cols): ws_pl.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_pl.cell(r_idx,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
            hdr=r_idx; r_idx+=1
            for _,rd in df_pl.iterrows():
                is_hl=bool(rd.get('GPCM_Tag', '')); row_fi=ev_fills.get('PL_HL', pW) if is_hl else (pST if r_idx%2==0 else pW)
                row_font=fHL if is_hl else fA
                for i,(col,_,_) in enumerate(cols):
                    if col=='_sort': continue
                    c_cell=ws_pl.cell(r_idx,i+1); v=rd.get(col, '')
                    if isinstance(v,(float,np.floating)): c_cell.value=round(v,1) if pd.notna(v) else None
                    else: c_cell.value=v
                    is_eps=rd.get('Unit','')=='per share'; nf=NF_EPS if is_eps else NF_M
                    sc(c_cell,fo=row_font,fi=row_fi,al=aR if col=='Amount_M' else aL,bd=BD,nf=nf if col=='Amount_M' else None)
                r_idx+=1
            ws_pl.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r_idx-1}"; ws_pl.freeze_panes=f'A{hdr+1}'

    # [3] Market_Cap Sheet
    ws_mc = wb.create_sheet('Market_Cap')
    ws_mc.sheet_properties.tabColor = COLOR_DARK_BLUE
    if market_rows:
        df_mkt = pd.DataFrame(market_rows)
        cols = [('Company','Company',18),('Ticker','Ticker',10), ('Base_Date','Base Date',12),('Price_Date','Price Date',12), ('Currency','Curr',6),('Close','Close Price',14), ('Shares','Shares (Ord.)',18),('Label','Label',10),('Market_Cap_M','Mkt Cap (M)',20)]
        ws_mc.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols)); sc(ws_mc.cell(1,1,'Market Capitalization'),fo=fT)
        r_idx=4
        for i,(col,disp,w) in enumerate(cols): ws_mc.column_dimensions[get_column_letter(i+1)].width=w; sc(ws_mc.cell(r_idx,i+1,disp),fo=fH,fi=pH,al=aC,bd=BD)
        hdr=r_idx; r_idx+=1
        for _,rd in df_mkt.iterrows():
            ev=(r_idx%2==0)
            for i,(col,_,_) in enumerate(cols):
                c_cell=ws_mc.cell(r_idx,i+1); v=rd.get(col, '')
                if isinstance(v,(float,np.floating)): c_cell.value=round(v,2) if pd.notna(v) else None
                else: c_cell.value=v
                nf=NF_PRC if col=='Close' else (NF_INT if col=='Shares' else (NF_M1 if col=='Market_Cap_M' else None))
                sc(c_cell,fo=fA,fi=pST if ev else pW,al=aR if nf else aL,bd=BD,nf=nf)
            r_idx+=1
        ws_mc.auto_filter.ref=f"A{hdr}:{get_column_letter(len(cols))}{r_idx-1}"; ws_mc.freeze_panes=f'A{hdr+1}'

    # [4] Price_History Sheet (Always at the end)
    ws_ph = wb.create_sheet('Price_History')
    ws_ph.sheet_properties.tabColor = '000000'
    # (Assuming market_rows or similar logic would populate this if requested, 
    # but for now ensuring empty sheet at least exists at the end as requested)

    wb.save(output)
    output.seek(0)
    return output

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
    '• Tax Rate: KPMG Corporate Tax Rates Table (2025 Combined Rate) 기준',
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
    
    app_mode = st.selectbox("Analysis Mode", ["GPCM Valuation (Multi-Period)", "Historical FS Summary (과거재무정보 요약)"])
    
    st.subheader("1. 분석 기간 설정")
    lookback_years = st.slider("분석 년수 (N개년)", min_value=1, max_value=10, value=3 if app_mode == "Historical FS Summary (과거재무정보 요약)" else 1)
    
    col1, col2 = st.columns(2)
    with col1:
        base_year = st.number_input("Base Year (Y)", min_value=2010, max_value=2030, value=2024)
    with col2:
        base_quarter = st.selectbox("Base Quarter (Q)", ["1Q", "2Q", "3Q", "4Q"], index=3)
            
    if app_mode == "GPCM Valuation (Multi-Period)":
        cycle_choice = "Quarterly" # Always for LTM logic
    else:
        cycle_choice = "Annual" # Force Annual for summary mode as per user request
            
    # Annual: User might want actual fiscal date, but for selection we use Y-12-31 as proxy
    q_map = {"1Q": "-03-31", "2Q": "-06-30", "3Q": "-09-30", "4Q": "-12-31"}
    
    # Base period date proxy
    if cycle_choice == "Quarterly":
        base_period_str = f"{base_year}{q_map[base_quarter]}"
    else:
        base_period_str = f"{base_year}-12-31"

    target_periods = []
    # Base period is Label 'Y'
    target_periods.append(base_period_str)
    
    # 1) If base is quarter, we must include the latest annual (base_year-12-31 or base_year-1-12-31)
    # The logic: always add N annual periods starting from the current year OR previous year if current is not yet annual-closed.
    # User said: if base is 2025-09-30, 2024-12-31 must be in history.
    
    current_annual_proxy = f"{base_year}-12-31"
    if current_annual_proxy != base_period_str:
        # If base is a mid-year quarter, the latest "finished" annual is base_year - 1
        for i in range(1, lookback_years + 1):
            hist_year = base_year - i
            target_periods.append(f"{hist_year}-12-31")
    else:
        # If base is 12-31, then hist starts from base_year - 1
        for i in range(1, lookback_years):
            hist_year = base_year - i
            target_periods.append(f"{hist_year}-12-31")
    
    target_periods = sorted(list(set(target_periods))) 
    
    if app_mode == "Historical FS Summary (과거재무정보 요약)":
        st.info(f"분석 기준: Y={base_year}부터 과거 {lookback_years}개년 실제 연간재무제표")
    else:
        st.info(f"기준일: {base_period_str} | 과거 추이: {lookback_years-1}개년 연간 데이터")

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
                                            help="KPMG 2025 Combined Rate 기준 | 한국: 26.4% | 미국: 26.0% | 일본: 29.7% | 프랑스: 36.1% | 독일: 30.1%") / 100

    st.info(f"💡 목표 부채비율은 피어들의 평균 자본구조로 자동 계산됩니다.")

    # 4. Run Button
    btn_run = st.button("Go, Go, Go 🚀", type="primary")

# [Main Execution]
if btn_run:
    target_tickers = [t.strip() for t in txt_input.split('\n') if t.strip()]

    if not target_tickers:
        st.warning("분석할 Ticker를 입력하세요.")
        st.stop()

    if app_mode == "Target WACC Calculation Only":
        st.info("지원 예정인 기능입니다 (별도의 WACC 계산기 통합 시). 현재는 GPCM Valuation 모드에서 함께 제공됩니다.")
        st.stop()

    with st.spinner(f"데이터 추출 및 분석 중 ({app_mode})... 잠시만 기다려주세요..."):
        if app_mode == "GPCM Valuation (Multi-Period)":
            # Run Data Logic with WACC parameters
            all_period_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, t_map, avg_debt_ratio, target_wacc_data = get_gpcm_data(
                target_tickers,
                target_periods,
                mrp=mrp_input,
                kd_pretax=kd_pretax_input,
                size_premium=size_premium_input,
                target_tax_rate=target_tax_rate_input,
                user_rf_rate=rf_input,
                beta_type=beta_type_input
            )
            
            # 1. Summary Table (최신 Base Date 기준)
            st.subheader(f"📋 GPCM Multiples Summary (Base: {base_period_str})")
            summary_list = []
            base_gpcm_data = all_period_data.get(base_period_str, {})
            
            for t, g in base_gpcm_data.items():
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
            st.success(f"✅ **피어 평균 부채비율 (목표 자본구조, 최신 기준)**: {avg_debt_ratio*100:.1f}%")

            # 2. Statistics Table
            st.subheader("📊 Multiples Statistics")
           
            if not df_sum.empty:
                stats = []
                for col in ['EV/EBITDA', 'EV/EBIT', 'PER', 'PBR', 'PSR']:
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
            excel_data = create_excel(all_period_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, base_period_str, target_periods, t_map, target_wacc_data, beta_type=beta_type_input)
            
            st.download_button(
                label="📥 GPCM Report Download (Excel)",
                data=excel_data,
                file_name=f"Global_GPCM_{base_period_str.replace('-','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        elif app_mode == "Historical FS Summary (과거재무정보 요약)":
            all_period_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, t_map, _, _ = get_gpcm_data(
                target_tickers,
                target_periods,
                mrp=mrp_input,
                kd_pretax=kd_pretax_input,
                size_premium=size_premium_input,
                target_tax_rate=target_tax_rate_input,
                user_rf_rate=rf_input,
                beta_type="5Y",
                force_annual=True,
                include_recent=True
            )
            
            excel_data = export_historical_excel_global_v2(all_period_data, raw_bs, raw_pl, mkt_rows, target_periods, t_map)
            
            if excel_data:
                st.success("✅ 기간별(년/분기별) 재무정보 및 요약 엑셀 생성이 완료되었습니다.")
                st.download_button(
                    label="📥 Historical FS Summary Download (Excel)",
                    data=excel_data,
                    file_name=f"Historical_FS_Summary_{target_tickers[0]}_etc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("추출된 재무/주가 데이터가 없습니다.")
