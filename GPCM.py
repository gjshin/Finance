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

# ==========================================
# 1. Helper Functions (v17 Logic + Beta Calculation)
# ==========================================

def get_market_index(ticker):
    """티커 기반으로 거래소 및 시장지수 코드 반환"""
    if ticker.endswith('.KS'):
        return 'KOSPI', 'KS11'
    elif ticker.endswith('.KQ'):
        return 'KOSDAQ', 'KQ11'
    elif ticker.endswith('.T'):
        return 'TSE', '^N225'
    elif ticker.endswith('.TO'):
        return 'TSX', '^GSPTSE'
    elif ticker.endswith('.F') or ticker.endswith('.DE'):
        return 'XETRA', '^GDAXI'
    elif ticker.endswith('.VI'):
        return 'VSE', '^ATX'
    else:
        return 'US', '^GSPC'  # S&P 500

def calculate_beta(stock_returns, market_returns, min_periods=20):
    """
    주식 수익률과 시장 수익률로부터 베타 계산
    Returns: raw_beta, adjusted_beta
    """
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

    return raw_beta, adjusted_beta

def get_korean_marginal_tax_rate(pretax_income_millions):
    """
    한국 법인세 한계세율 산출
    과세표준 기준 (단위: 백만원)
    - 2억 이하: 10%
    - 2억 초과 ~ 200억 이하: 20%
    - 200억 초과 ~ 3,000억 이하: 22%
    - 3,000억 초과: 25%
    """
    if pd.isna(pretax_income_millions) or pretax_income_millions == 0:
        return 0.22  # 기본값

    # 백만원 단위로 들어온 값
    if pretax_income_millions <= 200:
        return 0.10
    elif pretax_income_millions <= 20000:
        return 0.20
    elif pretax_income_millions <= 300000:
        return 0.22
    else:
        return 0.25

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
@st.cache_data(ttl=3600)  # <--- [추가] 1시간 동안 데이터를 저장해서 재사용함
def get_gpcm_data(tickers_list, base_date_str):
    """
    GPCM 데이터 수집 및 엑셀 생성을 위한 데이터 구조 반환
    """
    base_dt = pd.to_datetime(base_date_str)
    
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

            # [4] Beta Calculation
            exchange, market_idx = get_market_index(ticker)
            gpcm['Exchange'] = exchange
            gpcm['Market_Index'] = market_idx

            try:
                # 5년 월간 베타 계산
                start_5y = (base_dt - timedelta(days=365*5+20)).strftime('%Y-%m-%d')
                end_date = base_dt.strftime('%Y-%m-%d')

                # FinanceDataReader로 데이터 수집
                if ticker.endswith('.KS') or ticker.endswith('.KQ'):
                    # 한국 주식은 FinanceDataReader 사용
                    stock_data_5y = fdr.DataReader(ticker, start_5y, end_date)
                    market_data_5y = fdr.DataReader(market_idx, start_5y, end_date)
                else:
                    # 해외 주식은 yfinance 사용
                    stock_data_5y = yf.download(ticker, start=start_5y, end=end_date, progress=False)['Close']
                    market_data_5y = yf.download(market_idx, start=start_5y, end=end_date, progress=False)['Close']

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

                    # 월간 수익률 계산
                    stock_monthly = stock_prices_5y.resample('ME').last().pct_change().dropna()
                    market_monthly = market_prices_5y.resample('ME').last().pct_change().dropna()

                    # 베타 계산
                    raw_beta_5y, adj_beta_5y = calculate_beta(stock_monthly, market_monthly)
                    gpcm['Beta_5Y_Monthly_Raw'] = raw_beta_5y
                    gpcm['Beta_5Y_Monthly_Adj'] = adj_beta_5y

                # 2년 주간 베타 계산
                start_2y = (base_dt - timedelta(days=365*2+20)).strftime('%Y-%m-%d')

                if ticker.endswith('.KS') or ticker.endswith('.KQ'):
                    stock_data_2y = fdr.DataReader(ticker, start_2y, end_date)
                    market_data_2y = fdr.DataReader(market_idx, start_2y, end_date)
                else:
                    stock_data_2y = yf.download(ticker, start=start_2y, end=end_date, progress=False)['Close']
                    market_data_2y = yf.download(market_idx, start=start_2y, end=end_date, progress=False)['Close']

                if not stock_data_2y.empty and not market_data_2y.empty:
                    if isinstance(stock_data_2y, pd.DataFrame):
                        stock_prices_2y = stock_data_2y['Close'] if 'Close' in stock_data_2y.columns else stock_data_2y.iloc[:, 0]
                    else:
                        stock_prices_2y = stock_data_2y

                    if isinstance(market_data_2y, pd.DataFrame):
                        market_prices_2y = market_data_2y['Close'] if 'Close' in market_data_2y.columns else market_data_2y.iloc[:, 0]
                    else:
                        market_prices_2y = market_data_2y

                    # 주간 수익률 계산
                    stock_weekly = stock_prices_2y.resample('W').last().pct_change().dropna()
                    market_weekly = market_prices_2y.resample('W').last().pct_change().dropna()

                    # 베타 계산
                    raw_beta_2y, adj_beta_2y = calculate_beta(stock_weekly, market_weekly)
                    gpcm['Beta_2Y_Weekly_Raw'] = raw_beta_2y
                    gpcm['Beta_2Y_Weekly_Adj'] = adj_beta_2y

            except Exception as e:
                st.warning(f"Beta calculation failed for {ticker}: {e}")

            # [5] Pretax Income for Tax Rate Calculation
            if pl_data is not None and 'Pretax Income' in pl_data.index:
                pretax_vals = []
                for d in pl_dates:
                    val = pl_data.loc['Pretax Income', d]
                    if pd.notna(val):
                        pretax_vals.append(float(val) / 1e6)
                if pretax_vals:
                    gpcm['Pretax_Income'] = sum(pretax_vals)

                    # 한국 기업인 경우 한계세율 계산
                    if ticker.endswith('.KS') or ticker.endswith('.KQ'):
                        gpcm['Tax_Rate'] = get_korean_marginal_tax_rate(gpcm['Pretax_Income'])

            gpcm_data[ticker] = gpcm

        except Exception as e:
            st.error(f"Error fetching {ticker}: {e}")
            continue
    
    status_text.text("Data collection complete!")
    return gpcm_data, raw_bs_rows, raw_pl_rows, market_rows, price_abs_dfs, price_rel_dfs, ticker_to_name


def create_excel(gpcm_data, raw_bs_rows, raw_pl_rows, market_rows, price_abs_dfs, price_rel_dfs, base_date_str, ticker_to_name):
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

    # [Sheet 4] GPCM
    ws = wb.create_sheet('GPCM')
    wb.move_sheet('GPCM', offset=-3)
    TOTAL_COLS = 35
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS); sc(ws.cell(1,1,'GPCM Valuation Summary with Beta Analysis'), fo=fT)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=TOTAL_COLS); sc(ws.cell(2,1,f'Base: {base_date_str} | Unit: Millions (local currency) | EV = MCap + IBD − Cash + NCI'), fo=fS)

    r=4
    sections = [(1,3,'Company Info'),(4,4,'Other Information'),(8,6,'BS → EV Components'),(14,4,'PL (LTM / Annual)'),(18,3,'Market Data'),(21,5,'Valuation Multiples'),(26,10,'Beta & Risk Analysis')]
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
               'β 5Y Raw','β 5Y Adj','β 2Y Raw','β 2Y Adj','Pretax Inc','Tax Rate','D/E Ratio','Unlevered β 5Y','Unlevered β 2Y','Debt Ratio']
    widths = [18,10,11,6,16,10,10,
              14,14,14,12,14,16,
              14,14,14,14,
              12,16,16,
              12,12,10,10,10,
              10,10,10,10,14,9,10,12,12,10]
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
        # Beta 5Y Raw, Beta 5Y Adj, Beta 2Y Raw, Beta 2Y Adj
        ws.cell(r,26,gpcm['Beta_5Y_Monthly_Raw']); sc(ws.cell(r,26), fo=fA, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)
        ws.cell(r,27,gpcm['Beta_5Y_Monthly_Adj']); sc(ws.cell(r,27), fo=fA, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)
        ws.cell(r,28,gpcm['Beta_2Y_Weekly_Raw']); sc(ws.cell(r,28), fo=fA, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)
        ws.cell(r,29,gpcm['Beta_2Y_Weekly_Adj']); sc(ws.cell(r,29), fo=fA, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

        # Pretax Income (Formula)
        ws.cell(r,30).value=f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"Pretax Income")'; sc(ws.cell(r,30), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_M)

        # Tax Rate
        ws.cell(r,31,gpcm['Tax_Rate']); sc(ws.cell(r,31), fo=fA, fi=base_fi, al=aR, bd=BD, nf=NF_PCT)

        # D/E Ratio = IBD / Equity
        ws.cell(r,32).value=f'=IF(L{r}>0,I{r}/L{r},0)'; sc(ws.cell(r,32), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_RATIO)

        # Unlevered Beta 5Y = Beta 5Y Adj / (1 + (1 - Tax Rate) * (D/E))
        ws.cell(r,33).value=f'=IF(AA{r}>0,AA{r}/(1+(1-AE{r})*AF{r}),AA{r})'; sc(ws.cell(r,33), fo=fFRM_B, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

        # Unlevered Beta 2Y = Beta 2Y Adj / (1 + (1 - Tax Rate) * (D/E))
        ws.cell(r,34).value=f'=IF(AC{r}>0,AC{r}/(1+(1-AE{r})*AF{r}),AC{r})'; sc(ws.cell(r,34), fo=fFRM_B, fi=pBETA, al=aR, bd=BD, nf=NF_BETA)

        # Debt Ratio = IBD / Market Cap (추가)
        ws.cell(r,35).value=f'=IF(T{r}>0,I{r}/T{r},0)'; sc(ws.cell(r,35), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_RATIO)

    # Stats
    r=DATA_END+2
    stat_labels=['Mean','Median','Max','Min']; func_map={'Mean':'AVERAGE','Median':'MEDIAN','Max':'MAX','Min':'MIN'}
    # Multiples: 21-25 (EV/EBITDA, EV/EBIT, PER, PBR, PSR)
    # Betas: 26-29, 33-35 (Beta 5Y Raw, Beta 5Y Adj, Beta 2Y Raw, Beta 2Y Adj, Unlevered Beta 5Y, Unlevered Beta 2Y, Debt Ratio)
    mult_cols=[21,22,23,24,25]
    beta_cols=[26,27,28,29,33,34]
    ratio_cols=[32,35]

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
        '• Beta 5Y: Calculated using 5-year monthly returns vs market index',
        '• Beta 2Y: Calculated using 2-year weekly returns vs market index',
        '• Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1.0 (Bloomberg methodology)',
        '• Market Index: KOSPI (KS11), KOSDAQ (KQ11), Nikkei 225 (^N225), S&P/TSX (^GSPTSE), etc.',
        '• Tax Rate: Korean marginal corporate tax rate based on Pretax Income',
        '   - ≤ 200M: 10% | 200M-20,000M: 20% | 20,000M-300,000M: 22% | > 300,000M: 25%',
        '• D/E Ratio = IBD ÷ Equity',
        '• Unlevered Beta = Levered Beta ÷ (1 + (1 - Tax Rate) × D/E Ratio) [Hamada Model]',
        '• Debt Ratio = IBD ÷ Market Cap',
        '',
        '• N/M = Not Meaningful (negative or zero)',
        '• All values in GPCM are calculated via Excel Formulas linking to BS_Full and PL_Data sheets.',
        '', '⚠ Data from Yahoo Finance & FinanceDataReader. Verify with official sources.'
    ]
    for note in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_COLS)
        sc(ws.cell(r,1,note), fo=fNOTE); r+=1
    ws.freeze_panes='A6'

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
    '• Beta 5Y: Calculated using 5-year monthly returns vs market index',
    '• Beta 2Y: Calculated using 2-year weekly returns vs market index',
    '• Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1.0 (Bloomberg methodology)',
    '• Market Index: KOSPI (KS11), KOSDAQ (KQ11), Nikkei 225, S&P/TSX, DAX, etc.',
    '• Tax Rate: Korean marginal corporate tax rate based on Pretax Income',
    '• Unlevered Beta = Levered Beta ÷ (1 + (1 - Tax Rate) × D/E Ratio) [Hamada Model]',
    '• D/E Ratio = IBD ÷ Equity | Debt Ratio = IBD ÷ Market Cap',
]
for note in beta_notes:
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
    
    # 3. Run Button
    btn_run = st.button("Go, Go, Go 🚀", type="primary")

# [Main Execution]
if btn_run:
    target_tickers = [t.strip() for t in txt_input.split('\n') if t.strip()]
    
    with st.spinner("데이터 추출 및 분석 중... 잠시만 기다려주세요..."):
        # Run Data Logic
        gpcm_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, t_map = get_gpcm_data(target_tickers, base_date_str)
        
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
        excel_data = create_excel(gpcm_data, raw_bs, raw_pl, mkt_rows, p_abs, p_rel, base_date_str, t_map)
        
        # 다운로드 버튼 (누르고 있어도 화면 유지됨)
        st.download_button(
            label="📥 Report Download (Excel)",
            data=excel_data,
            file_name=f"Global_GPCM_{base_date_str.replace('-','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
