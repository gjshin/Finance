import streamlit as st
import pandas as pd
import OpenDartReader
import FinanceDataReader as fdr
from datetime import datetime, timedelta
import warnings
import numpy as np
import re
import requests
import time
import io # 엑셀 메모리 저장을 위해 추가
import yfinance as yf # 지수 정보 조회를 위해 추가
from bs4 import BeautifulSoup # 주식수 크롤링을 위해 추가

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 최신 수정: 2026-02-17 15:00 KST
# 주요 변경사항:
# - Beta 계산 기능 추가 (5Y Monthly, 2Y Weekly) - FinanceDataReader 사용
# - Beta_Calculation 시트 추가
# - WACC_Calculation 시트 완전 구현 (GPCM.py와 동일)
# - GPCM 시트에 Beta & Risk Analysis 컬럼 추가 (총 35개 컬럼)
# - D/E Ratio 컬럼 추가 (컬럼 32): IBD/(시총+NCI)
# - Debt Ratio 컬럼 이동 (컬럼 33): IBD/(시총+IBD+NCI)
# - Unlevered Beta 수식 수정: D/E 사용 (하마다 모형)
# - 한국 법인세 한계세율 적용 (지방세 포함)
# - Streamlit 사용자 입력 추가: Rf, MRP, Size Premium, Beta Type, Kd, Target Tax Rate
# - 모든 데이터 소스: FinanceDataReader (yfinance 미사용)

# Streamlit 페이지 설정 (가장 먼저 와야 함)
st.set_page_config(page_title="GPCM Calculator", layout="wide")

warnings.filterwarnings('ignore')

# ==========================================
# 0. 전역 설정 및 상수
# ==========================================
RCODE_MAP = {'1Q': '11013', '2Q': '11012', '3Q': '11014', '4Q': '11011'}
QUARTER_INFO = {'1Q': '03-31', '2Q': '06-30', '3Q': '09-30', '4Q': '12-31'}
DEBUG_PL = False  # 로그 출력 줄임

# ==========================================
# 1. Helper Functions
# ==========================================

def get_market_index(ticker):
    """
    티커 기반으로 거래소 및 시장지수 코드 반환 (한국 종목만 지원)
    Returns: (exchange_name, index_symbol)
    """
    # 한국 종목 - FinanceDataReader 기준
    # KS11 (KOSPI)가 fdr에서 실패하는 경우가 많아 yfinance 심볼(^KS11)로 대체
    return 'KRX', '^KS11'  # 기본값: KOSPI

def get_korean_marginal_tax_rate(pretax_income_100m):
    """
    한국 법인세 한계세율 산출 (2025년 기준, 지방세 포함)
    과세표준 기준 (단위: 억원)
    - 2억 이하: 9% (국세) + 0.9% (지방세 10%) = 9.9%
    - 2억 ~ 200억: 19% + 1.9% = 20.9%
    - 200억 ~ 3,000억: 21% + 2.1% = 23.1%
    - 3,000억 초과: 24% + 2.4% = 26.4%
    """
    if pd.isna(pretax_income_100m) or pretax_income_100m == 0:
        return 0.231  # 기본값 (중간 구간)

    # 억원 단위로 들어온 값
    if pretax_income_100m <= 2:
        return 0.099
    elif pretax_income_100m <= 200:
        return 0.209
    elif pretax_income_100m <= 3000:
        return 0.231
    else:
        return 0.264

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

def parse_period(p: str):
    parts = p.strip().split('.')
    return int(parts[0]), parts[1]

def get_base_date_str(year: int, qtr: str):
    return f"{year}-{QUARTER_INFO[qtr]}"

def get_ltm_required_periods(year: int, qtr: str):
    if qtr == '4Q':
        return [(year, '4Q', 'annual')]
    return [
        (year, qtr, 'current_cum'),
        (year - 1, '4Q', 'prior_annual'),
        (year - 1, qtr, 'prior_same_q'),
    ]

@st.cache_resource
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
    except:
        corp_code = None

    if not corp_code and krx_name:
        try:
            corp_code = dart_instance.find_corp_code(krx_name)
        except:
            corp_code = None

    return corp_code, krx_name

def get_stock_price(ticker: str, date_str: str):
    try:
        td = pd.to_datetime(date_str)
        if td > datetime.now():
            return None, None
        df = fdr.DataReader(ticker, td - timedelta(days=10), td)
        if df is not None and not df.empty:
            return float(df.iloc[-1]['Close']), df.index[-1].strftime('%Y-%m-%d')
        return None, None
    except:
        return None, None

def _to_int(x):
    try:
        if x is None:
            return None
        s = str(x).strip().replace(',', '')
        if s == '' or s.lower() == 'nan':
            return None
        return int(float(s))
    except:
        return None

# --- DART 유통주식수 ---
DART_STOCKTOTQY_URL = "https://opendart.fss.or.kr/api/stockTotqySttus.json"

def fetch_dart_distb_shares(api_key, corp_code: str, bsns_year: int, reprt_code: str):
    meta = {'shares': None, 'rcept_no': None, 'stlm_dt': None, 'se': None, 'status': None, 'message': None}
    try:
        params = {
            'crtfc_key': api_key,
            'corp_code': corp_code,
            'bsns_year': str(bsns_year),
            'reprt_code': str(reprt_code),
        }
        resp = requests.get(DART_STOCKTOTQY_URL, params=params, timeout=10)
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

def get_outstanding_shares(api_key, corp_code: str, ticker: str, bsns_year: int, reprt_code: str, df_krx: pd.DataFrame):
    # 1. DART API 조회 (요청한 기준년도/분기)
    shares, meta = fetch_dart_distb_shares(api_key, corp_code, bsns_year, reprt_code)
    if shares is not None and shares > 0:
        return shares, f"DART({reprt_code})", meta

    # 2. 직전 보고서들에서 주식수 조회 시도 (요청 분기에 주식수 누락 시)
    # 순서: 11011 (1Q), 11012 (2Q), 11014 (3Q), 11013 (4Q)
    order = ['11011', '11012', '11014', '11013']
    try:
        current_idx = order.index(reprt_code)
    except:
        current_idx = 3 # 기본 사업보고서 매핑

    cy = bsns_year
    ci = current_idx - 1
    
    # 최근 8개 분기(약 2년치)를 역순으로 훑어 가장 최근 공시된 주식총수를 찾음
    for _ in range(8):
        if ci < 0:
            cy -= 1
            ci = 3
        
        fb_code = order[ci]
        fb_shares, fb_meta = fetch_dart_distb_shares(api_key, corp_code, cy, fb_code)
        
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
    except:
        pass

    return None, 'N/A', meta

# --- BS Matching Logic ---
IBD_AID_ALWAYS = {
    'ifrs-full_CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings',
    'ifrs-full_LongtermBorrowings',
    'ifrs-full_CurrentLeaseLiabilities',
    'ifrs-full_CurrentPortionOfLongtermBorrowings',
    'ifrs-full_ShorttermBorrowings',
    'ifrs-full_NoncurrentLeaseLiabilities',
    'dart_CurrentPortionOfBonds',
    'ifrs-full_BondsIssued',
    'ifrs-full_Borrowings',
}
IBD_AID_PATTERN = re.compile(r'(Borrowings|Bonds|LeaseLiabilit)', re.IGNORECASE)
MEZZ_KW_KR = ['전환사채', '교환사채', '신주인수권부사채', 'BW', 'CB', 'EB', '전환', '상환', '신주인수', '교환']
MEZZ_KW_EN = ['convertible', 'exchangeable', 'bond with warrant', 'bonds with warrants', 'warrant']
IBD_KW_NAME = ['차입금', '사채', '리스부채', 'Borrowings', 'Bond', 'Bonds', 'LeaseLiabilit', 'Lease Liability']
IBD_EXCLUDE = [
    '매입채무', '미지급', '충당', '선수', '예수', '보증금',
    '자산', '대여금', '미수', '매출채권', '미수금', '미수수익',
    '선급', '선급금', '선급비용', '예치금', '보증금',
    '리스채권', '대여', '대출금(자산)',
]

def _norm(s):
    s = "" if s is None else str(s)
    return re.sub(r"\s+", "", s).strip()

def match_bs_ev_component(account_nm, account_id):
    acct = "" if account_nm is None else str(account_nm).strip()
    aid = "" if account_id is None else str(account_id).strip()
    acct_n = _norm(acct)
    acct_u = acct_n.upper()
    acct_l = acct_n.lower()

    if aid in ['ifrs-full_CashAndCashEquivalents', 'ifrs-full_ShorttermDepositsNotClassifiedAsCashEquivalents']:
        return 'Cash', '현금및단기예금'
    if aid == 'ifrs-full_Equity':
        return 'Equity_Total', '자본총계'
    if aid == 'ifrs-full_EquityAttributableToOwnersOfParent':
        return 'Equity_P', '지배기업지분'
    if aid == 'dart_ElementsOfOtherStockholdersEquity':
        return None, None

    if '우선주' not in acct_n:
        mezz_hit = False
        for kw in MEZZ_KW_KR:
            if kw.replace(" ", "") in acct_n: mezz_hit = True; break
        if (not mezz_hit) and any(kw in acct_l for kw in MEZZ_KW_EN): mezz_hit = True
        if (not mezz_hit) and re.search(r'(\bCB\b|\bEB\b|\bBW\b)', acct_u): mezz_hit = True
        if mezz_hit: return 'IBD(Option)', acct

    if not any(ex.replace(" ", "") in acct_n for ex in IBD_EXCLUDE):
        if aid in IBD_AID_ALWAYS: return 'IBD', acct
        if aid and IBD_AID_PATTERN.search(aid): return 'IBD', acct

    if any(k.replace(" ", "") in acct_n for k in IBD_KW_NAME):
        if not any(ex.replace(" ", "") in acct_n for ex in IBD_EXCLUDE):
            return 'IBD', acct

    if ('비지배지분' in acct or '소수주주지분' in acct) and ('귀속' not in acct):
        return 'NCI', '비지배지분'

    noa_keywords = ['관계기업', '지분법', '공동기업', '종속기업', '금융자산', '금융상품']
    noa_exclude = ['단기', '현금', '매출', '보증금', '미수', '대여금', '예치금', '부채', '충당', '손실', '리스채권']
    if any(kw in acct for kw in noa_keywords) and not any(ex in acct for ex in noa_exclude):
        if aid not in ['ifrs-full_CashAndCashEquivalents', 'ifrs-full_ShorttermDepositsNotClassifiedAsCashEquivalents']:
            return 'NOA(Option)', acct
    return None, None

# --- PL Logic ---
PL_REVENUE = {
    '매출액', '수익(매출액)', '수익(매출)', '영업수익',
    '수익', '매출', '총매출액', '총수익', '영업수익',
    '매출액합계', '수익합계', '총영업수익'
}
PL_EBIT    = {'영업이익', '영업이익(손실)', '영업손실', '영업손익'}
PL_NI      = {
    '당기순이익', '당기순이익(손실)', '당기순손실', '당기순손익',
    '분기순이익', '분기순이익(손실)', '분기순손실', '분기순손익',
    '반기순이익', '반기순이익(손실)', '반기순손실', '반기순손익',
    '연결당기순이익', '연결당기순이익(손실)', '연결당기순손실', '연결당기순손익',
    'ProfitLoss', 'ifrs-full_ProfitLoss'
}
PL_PRETAX_INCOME = {
    '법인세비용차감전순이익', '법인세비용차감전순이익(손실)', '법인세차감전순이익',
    '법인세비용차감전계속사업이익', '법인세비용차감전이익', '세전순이익',
    '법인세비용차감전순손실', '세전이익', '법인세차감전이익'
}

def _norm_pl(s):
    s = "" if s is None else str(s).strip()
    return re.sub(r"\s+", "", s)

def match_pl_core_only(account_nm, aid=None):
    if aid == 'ifrs-full_ProfitLoss': return 'NI'
    a = _norm_pl(account_nm)
    if '지배' in a: return None # Exclude subset (지배기업, 비지배기업)
    if '포괄' in a: return None # Exclude Comprehensive Income
    if a in PL_REVENUE: return 'Revenue'
    if a in PL_EBIT:    return 'EBIT'
    if a in PL_NI:      return 'NI'
    if a in PL_PRETAX_INCOME: return 'Pretax_Income'
    return None

def _parse_amount(x):
    v = pd.to_numeric(str(x).replace(',', ''), errors='coerce')
    if pd.isna(v) or v == 0: return None
    return float(v)

def pick_pl_value(row: pd.Series, qtr: str):
    if qtr == '4Q':
        for col in ['thstrm_amount', 'thstrm_add_amount']:
            v = _parse_amount(row.get(col, ''))
            if v is not None: return v
    else:
        for col in ['thstrm_add_amount', 'thstrm_amount']:
            v = _parse_amount(row.get(col, ''))
            if v is not None: return v
    return None

# --- DART PL Fetch Functions (Need dart instance) ---
def safe_finstate(dart_instance, corp_code, year, reprt_code, fs_div=None, max_retry=2):
    last_err = None
    for _ in range(max_retry + 1):
        try:
            if fs_div is None:
                df = dart_instance.finstate(corp_code, year, reprt_code=reprt_code)
            else:
                df = dart_instance.finstate(corp_code, year, reprt_code=reprt_code, fs_div=fs_div)
            if df is not None and not df.empty: return df, None
            return df, None
        except Exception as e:
            last_err = e
            time.sleep(0.4)
    return None, last_err

def safe_finstate_all(dart_instance, corp_code, year, reprt_code, fs_div=None, max_retry=2):
    last_err = None
    for _ in range(max_retry + 1):
        try:
            if fs_div is None:
                df = dart_instance.finstate_all(corp_code, year, reprt_code=reprt_code)
            else:
                df = dart_instance.finstate_all(corp_code, year, reprt_code=reprt_code, fs_div=fs_div)
            if df is not None and not df.empty: return df, None
            return df, None
        except Exception as e:
            last_err = e
            time.sleep(0.4)
    return None, last_err

def fetch_pl_df(dart_instance, corp_code, year, reprt_code):
    for fs in ['CFS', 'OFS']:
        df, err = safe_finstate(dart_instance, corp_code, year, reprt_code, fs_div=fs)
        if df is not None and not df.empty: return df, f'finstate|{fs}', None
    
    df, err = safe_finstate(dart_instance, corp_code, year, reprt_code, fs_div=None)
    if df is not None and not df.empty: return df, 'finstate|no_fs_div', None
    
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

# ==========================================
# 4. 스타일 및 엑셀 유틸
# ==========================================
C_BL='00338D'; C_DB='1E2A5E'; C_LB='C3D7EE'; C_PB='E8EFF8'
C_DG='333333'; C_MG='666666'; C_LG='F5F5F5'; C_BG='B0B0B0'; C_W='FFFFFF'
C_GR='E2EFDA'; C_YL='FFF8E1'; C_NOA='FCE4EC'

S1=Side(style='thin',color=C_BG); BD=Border(left=S1,right=S1,top=S1,bottom=S1)
fT=Font(name='KoPub돋움체 Medium',bold=True,size=14,color=C_BL)
fS=Font(name='KoPub돋움체 Medium',size=9,color=C_MG,italic=True)
fH=Font(name='KoPub돋움체 Medium',bold=True,size=9,color=C_W)
fA=Font(name='KoPub돋움체 Medium',size=9,color=C_DG)
fHL=Font(name='KoPub돋움체 Medium',bold=True,size=9,color=C_DB)
fMUL=Font(name='KoPub돋움체 Medium',bold=True,size=10,color=C_BL)
fNOTE=Font(name='KoPub돋움체 Medium',size=8,color=C_MG,italic=True)
fSTAT=Font(name='KoPub돋움체 Medium',bold=True,size=9,color=C_DB)
fFRM=Font(name='KoPub돋움체 Medium',size=9,color='000000')  # 검은색 (수식/계산값)
fHARD=Font(name='KoPub돋움체 Medium',size=9,color='FF0000') # 빨간색 (하드코딩/외부데이터)
fASSUM=Font(name='KoPub돋움체 Medium',size=9,color='FFC000') # 노란색 (주요 가정사항)
fLINK=Font(name='KoPub돋움체 Medium',size=9,color='008000') # 초록색 (내부시트 링크)
fSEC = Font(name='KoPub돋움체 Medium', bold=True, size=10, color=C_W)

pH=PatternFill('solid',fgColor=C_BL); pW=PatternFill('solid',fgColor=C_W)
pST=PatternFill('solid',fgColor=C_LG); pSTAT=PatternFill('solid',fgColor=C_LB)
pSEC1 = PatternFill('solid', fgColor=C_DB); pSEC2 = PatternFill('solid', fgColor=C_BL)

# ==========================================
# 4.5. [신규] 다기간 재무제표 요약 로직 추가
# ==========================================
def fetch_historical_financials(api_key, target_code_list, periods_to_fetch, dart, status_container, progress_bar, df_krx):
    total = len(target_code_list) * len(periods_to_fetch)
    cnt = 0
    hist_summary = []
    hist_details = []
    
    QTR_TO_CODE = {'1Q': '11011', '2Q': '11012', '3Q': '11014', '4Q': '11013'}

    for ticker in target_code_list:
        corp_code, _ = resolve_company_info(dart, ticker)
        if not corp_code:
            cnt += len(periods_to_fetch); progress_bar.progress(cnt/total)
            continue
        
        comp_name = dart.company(corp_code).get('corp_name', ticker)

        for p in periods_to_fetch:
            year = p['year']
            hist_qtr = p['qtr']
            plabel = p['label']
            
            if not hist_qtr: 
                req_periods = [(year, '4Q', 'annual')]
            else:
                req_periods = [(year, hist_qtr, 'current_cum')]
            
            # BS & EV Components
            ast, liab, eq = np.nan, np.nan, np.nan
            cash, ibd, noa, nci = 0.0, 0.0, 0.0, 0.0
            
            # PL & CF LTM Aggregator
            pl_agg = {'Revenue': 0.0, 'GrossProfit': 0.0, 'EBIT': 0.0, 'NI': 0.0, 'CFO': 0.0, 'CFI': 0.0, 'CFF': 0.0}
            valid_pl_flags = {'Revenue': False, 'GrossProfit': False, 'EBIT': False, 'NI': False, 'CFO': False, 'CFI': False, 'CFF': False}
            
            used_code_current = 'N/A'
            df_fs_current = None
            
            for req_year, req_qtr, role in req_periods:
                primary = QTR_TO_CODE.get(req_qtr, '11013')
                fallbacks = [c for c in ['11013', '11014', '11012', '11011'] if c != primary]
                target_qtrs = [primary] + fallbacks
                
                df_fs = None
                used_code = None
                for rcode in target_qtrs:
                    df_fs, _ = safe_finstate_all(dart, corp_code, req_year, rcode, fs_div='CFS')
                    if df_fs is None or df_fs.empty:
                        df_fs, _ = safe_finstate_all(dart, corp_code, req_year, rcode, fs_div='OFS')
                    if df_fs is not None and not df_fs.empty:
                        used_code = rcode
                        break
                
                if df_fs is None or df_fs.empty:
                    continue # Skip if data is missing
                    
                if role in ('current_cum', 'annual'):
                    used_code_current = used_code
                    df_fs_current = df_fs
                    
                temp_pl = {'Revenue': np.nan, 'GrossProfit': np.nan, 'EBIT': np.nan, 'NI': np.nan, 'CFO': np.nan, 'CFI': np.nan, 'CFF': np.nan}
                
                for row_idx, row in df_fs.iterrows():
                    sj = str(row.get('sj_nm', ''))
                    acc = str(row.get('account_nm', '')).strip()
                    aid = str(row.get('account_id', '')).strip()
                    _raw = _parse_amount(str(row.get('thstrm_amount', '')))
                    val_1m = (_raw / 1000000) if _raw is not None else np.nan
                    
                    if pd.isna(val_1m): continue
                    
                    m_key = ""
                    if '상태' in sj and role in ('current_cum', 'annual'):
                        if acc == '자산총계': m_key = 'Assets'
                        elif acc == '부채총계': m_key = 'Liabilities'
                        elif acc == '자본총계': m_key = 'Equity_Total'
                        ev_comp, _ = match_bs_ev_component(acc, aid)
                        if ev_comp:
                            m_key = ev_comp # 'Cash', 'Cash(Option)', 'IBD', 'IBD(Option)', 'NOA', 'NOA(Option)', 'NCI'
                            
                    elif '손익' in sj and role in ('current_cum', 'annual'):
                        n_acc = _norm_pl(acc)
                        if '지배' not in n_acc and '포괄' not in n_acc:
                            if n_acc in PL_REVENUE: m_key = 'Revenue'
                            elif '매출총이익' in acc: m_key = 'GrossProfit'
                            elif '영업이익' in acc: m_key = 'EBIT'
                            elif '당기순이익' in acc or '분기순이익' in acc or '반기순이익' in acc or aid == 'ifrs-full_ProfitLoss': m_key = 'NI'
                            
                    elif '현금' in sj and role in ('current_cum', 'annual'):
                        if '영업활동' in acc and '흐름' in acc: m_key = 'CFO'
                        elif '투자활동' in acc and '흐름' in acc: m_key = 'CFI'
                        elif '재무활동' in acc and '흐름' in acc: m_key = 'CFF'

                    # Store Raw Data for Details Sheet (Only for current period)
                    if role in ('current_cum', 'annual') and val_1m != 0:
                        hist_details.append({
                            'Company': comp_name, 'Ticker': ticker, 'Period': plabel, 'Report': used_code_current,
                            'M_Key': m_key, 'Type': sj, 'Account_ID': aid, 'Account_NM': acc, 
                            'Amount': val_1m, 'Row_Idx': row_idx
                        })
                    
                    if '상태' in sj and role in ('current_cum', 'annual'):
                        if acc == '자산총계': ast = val_1m
                        elif acc == '부채총계': liab = val_1m
                        elif acc == '자본총계': eq = val_1m
                        
                        ev_comp, _ = match_bs_ev_component(acc, aid)
                        if ev_comp:
                            if ev_comp == 'Cash': cash += val_1m
                            elif ev_comp == 'IBD': ibd += val_1m
                            elif ev_comp == 'NCI': nci += val_1m
                            elif ev_comp == 'NOA': noa += val_1m
                            
                    elif '손익' in sj:
                        n_acc = _norm_pl(acc)
                        _raw_pl = pick_pl_value(row, req_qtr)
                        val_pl = (_raw_pl / 1000000) if _raw_pl is not None else np.nan
                        if not pd.isna(val_pl) and '지배' not in n_acc and '포괄' not in n_acc:
                            if pd.isna(temp_pl['Revenue']) and n_acc in PL_REVENUE: temp_pl['Revenue'] = val_pl
                            if pd.isna(temp_pl['GrossProfit']) and '매출총이익' in acc: temp_pl['GrossProfit'] = val_pl
                            if pd.isna(temp_pl['EBIT']) and '영업이익' in acc: temp_pl['EBIT'] = val_pl
                            if pd.isna(temp_pl['NI']) and '당기순이익' in acc: temp_pl['NI'] = val_pl
                            
                    elif '현금' in sj:
                        if pd.isna(temp_pl['CFO']) and '영업활동' in acc and '흐름' in acc: temp_pl['CFO'] = val_1m
                        if pd.isna(temp_pl['CFI']) and '투자활동' in acc and '흐름' in acc: temp_pl['CFI'] = val_1m
                        if pd.isna(temp_pl['CFF']) and '재무활동' in acc and '흐름' in acc: temp_pl['CFF'] = val_1m
                
                # Apply to aggregator
                for k in temp_pl:
                    v = temp_pl[k]
                    if pd.notna(v):
                        pl_agg[k] += v
                        valid_pl_flags[k] = True

            if df_fs_current is None or df_fs_current.empty:
                hist_summary.append({
                    'Company': comp_name, 'Ticker': ticker, 'Period': plabel, 'Report': 'N/A',
                    'Revenue': np.nan, 'GrossProfit': np.nan, 'EBIT': np.nan, 'NI': np.nan,
                    'Assets': np.nan, 'Liabilities': np.nan, 'Equity': np.nan,
                    'CFO': np.nan, 'CFI': np.nan, 'CFF': np.nan,
                    'Cash': np.nan, 'IBD': np.nan, 'NOA': np.nan, 'NCI': np.nan,
                    'Shares': np.nan, 'Price': np.nan, 'MarketCap': np.nan
                })
                cnt += 1; progress_bar.progress(cnt/total)
                continue

            hist_summary.append({
                'Company': comp_name, 'Ticker': ticker, 'Period': plabel, 'Report': used_code_current,
                'Revenue': pl_agg['Revenue'] if valid_pl_flags['Revenue'] else np.nan, 
                'GrossProfit': pl_agg['GrossProfit'] if valid_pl_flags['GrossProfit'] else np.nan, 
                'EBIT': pl_agg['EBIT'] if valid_pl_flags['EBIT'] else np.nan, 
                'NI': pl_agg['NI'] if valid_pl_flags['NI'] else np.nan,
                'Assets': ast, 'Liabilities': liab, 'Equity_Total': eq,
                'CFO': pl_agg['CFO'] if valid_pl_flags['CFO'] else np.nan, 
                'CFI': pl_agg['CFI'] if valid_pl_flags['CFI'] else np.nan, 
                'CFF': pl_agg['CFF'] if valid_pl_flags['CFF'] else np.nan,
                'Cash': cash, 'IBD': ibd, 'NOA': noa, 'NCI': nci
            })
            
            status_container.update(label=f"다기간 재무데이터 수집 중... {comp_name} ({plabel})")
            cnt += 1; progress_bar.progress(cnt/total)
            
    return pd.DataFrame(hist_summary), pd.DataFrame(hist_details)

def calculate_historical_metrics(df_summ):
    if df_summ.empty: return df_summ
    
    for col in ['OPM', 'GPM', 'ROE', 'DebtRatio', 'NetDebt']:
        df_summ[col] = np.nan
        
    for i, row in df_summ.iterrows():
        rev = row.get('Revenue'); ebit = row.get('EBIT'); gp = row.get('GrossProfit'); ni = row.get('NI')
        eq = row.get('Equity_Total'); liab = row.get('Liabilities')
        cash = row.get('Cash', 0.0); ibd = row.get('IBD', 0.0)
        noa = row.get('NOA', 0.0); nci = row.get('NCI', 0.0)
        
        if rev and rev > 0:
            df_summ.at[i, 'OPM'] = ebit / rev if pd.notna(ebit) else np.nan
            df_summ.at[i, 'GPM'] = gp / rev if pd.notna(gp) else np.nan
        if eq and eq > 0:
            df_summ.at[i, 'ROE'] = ni / eq if pd.notna(ni) else np.nan
            if pd.notna(liab): df_summ.at[i, 'DebtRatio'] = liab / eq
                
        nd = (ibd if pd.notna(ibd) else 0) - (cash if pd.notna(cash) else 0)
        df_summ.at[i, 'NetDebt'] = nd

    return df_summ

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def export_historical_excel(df_summ, df_details, periods_to_fetch):
    import io
    output = io.BytesIO()
    wb = Workbook()
    
    # ---------------------------------------------------------
    # 1. Summary 시트 생성 (Layout A안: 회사 세로, 연도/지표 가로)
    # ---------------------------------------------------------
    ws_summ = wb.active
    ws_summ.title = "Summary"
    
    ws_summ.merge_cells('A1:Z1')
    p_start = periods_to_fetch[0]['label'] if periods_to_fetch else ""
    p_end = periods_to_fetch[-1]['label'] if periods_to_fetch else ""
    ws_summ['A1'] = f"Historical Financial Summary ({p_start} ~ {p_end})"
    sc(ws_summ['A1'], fo=fT)
    
    if df_summ.empty:
        ws_summ['A3'] = "No data available."
    else:
        metrics = [
            ('Revenue', '매출액', NB), ('GrossProfit', '매출총이익', NB), 
            ('EBIT', '영업이익', NB), ('NI', '당기순이익', NB),
            ('Assets', '자산총계', NB), ('Liabilities', '부채총계', NB), ('Equity_Total', '자본총계', NB),
            ('Cash', 'Cash', NB), ('IBD', 'IBD', NB),
            ('NOA', 'NOA', NB), ('NCI', 'NCI', NB),
            ('NetDebt', '순부채(Net Debt)', NB),
            ('CFO', '영업활동현금흐름', NB), ('CFI', '투자활동현금흐름', NB), ('CFF', '재무활동현금흐름', NB),
            ('OPM', '영업이익률', '0.0%'), ('GPM', '매출총이익률', '0.0%'), 
            ('ROE', 'ROE', '0.0%'), ('DebtRatio', '부채비율', '0.0%')
        ]
        
        labels = [p['label'] for p in periods_to_fetch]
        
        # 헤더 그리기 (Row 3: 지표명, Row 4: 기간)
        ws_summ.cell(row=3, column=1, value="Company"); sc(ws_summ.cell(row=3, column=1), fo=fH, fi=pH, al=aC, bd=BD)
        ws_summ.cell(row=3, column=2, value="Ticker"); sc(ws_summ.cell(row=3, column=2), fo=fH, fi=pH, al=aC, bd=BD)
        ws_summ.merge_cells('A3:A4')
        ws_summ.merge_cells('B3:B4')
        
        # 메트릭별 컬럼 알파벳 매핑 저장용
        mc_map = {} # (m_key, plabel) -> 'C'
        plabel_col_idx = {plabel: 5 + i for i, plabel in enumerate(labels)} # Detail 시트의 데이터 컬럼 E부터 시작
        
        col_idx = 3
        for m_key, m_name, _ in metrics:
            start_col = col_idx
            end_col = col_idx + len(labels) - 1
            ws_summ.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
            ws_summ.cell(row=3, column=start_col, value=m_name)
            sc(ws_summ.cell(row=3, column=start_col), fo=fH, fi=pSEC1, al=aC, bd=BD)
            
            for plabel in labels:
                ws_summ.cell(row=4, column=col_idx, value=plabel)
                sc(ws_summ.cell(row=4, column=col_idx), fo=fH, fi=pSEC2, al=aC, bd=BD)
                
                mc_map[(m_key, plabel)] = get_column_letter(col_idx)
                col_idx += 1
                
        # 데이터 쓰기 (Row 5부터 ~ 회사별)
        r = 5
        companies = df_summ['Company'].unique()
        for comp in companies:
            df_comp = df_summ[df_summ['Company'] == comp]
            ticker = df_comp['Ticker'].iloc[0] if not df_comp.empty else ""
            comp_sht = comp[:31] # 엑셀 시트 참조용 이름
            
            ws_summ.cell(row=r, column=1, value=comp); sc(ws_summ.cell(row=r, column=1), fo=fA, bd=BD)
            ws_summ.cell(row=r, column=2, value=ticker); sc(ws_summ.cell(row=r, column=2), fo=fA, al=aC, bd=BD)
            
            # 수식 적용 그룹 (비율/멀티플)
            ratio_keys = ['OPM', 'GPM', 'ROE', 'DebtRatio']
            # Raw Data SUMIFS 그룹
            sumifs_keys = ['Revenue', 'GrossProfit', 'EBIT', 'NI', 'Assets', 'Liabilities', 'Equity_Total', 'Cash', 'IBD', 'NOA', 'NCI', 'CFO', 'CFI', 'CFF']
            
            c = 3
            for m_key, m_name, fmt in metrics:
                for plabel in labels:
                    v = ""
                    dtl_col = get_column_letter(plabel_col_idx[plabel]) # Detail 시트의 타겟 Period 열
                    
                    if m_key in sumifs_keys:
                        # 엑셀 SUMIFS 수식 주입 (매핑 키 A열, 금액 Dtl_Col열)
                        v = f"=SUMIFS('{comp_sht}'!{dtl_col}:{dtl_col}, '{comp_sht}'!$A:$A, \"{m_key}\")"
                    elif m_key == 'NetDebt':
                        v = f"={mc_map[('IBD', plabel)]}{r} - {mc_map[('Cash', plabel)]}{r}"
                    elif m_key in ratio_keys:
                        if m_key == 'OPM': v = f"=IFERROR({mc_map[('EBIT', plabel)]}{r}/{mc_map[('Revenue', plabel)]}{r}, \"\")"
                        elif m_key == 'GPM': v = f"=IFERROR({mc_map[('GrossProfit', plabel)]}{r}/{mc_map[('Revenue', plabel)]}{r}, \"\")"
                        elif m_key == 'ROE': v = f"=IFERROR({mc_map[('NI', plabel)]}{r}/{mc_map[('Equity_Total', plabel)]}{r}, \"\")"
                        elif m_key == 'DebtRatio': v = f"=IFERROR({mc_map[('Liabilities', plabel)]}{r}/{mc_map[('Equity_Total', plabel)]}{r}, \"\")"
                        
                    ws_summ.cell(row=r, column=c, value=v)
                    font_style = fA if m_key in sumifs_keys else fFRM
                    sc(ws_summ.cell(row=r, column=c), fo=font_style, nf=fmt, bd=BD)
                    
                    c += 1
            r += 1
        
        ws_summ.column_dimensions['A'].width = 18
        ws_summ.column_dimensions['B'].width = 10
        for i in range(3, c):
            ws_summ.column_dimensions[get_column_letter(i)].width = 14
        
        ws_summ.freeze_panes = "C5"

    # ---------------------------------------------------------
    # 2. 개별 회사 상세 시트 생성 (세로: 계정, 가로: 연도 피벗 형태)
    # ---------------------------------------------------------
    if not df_details.empty:
        companies = df_details['Company'].unique()
        for comp in companies:
            ws_dtl = wb.create_sheet(title=comp[:31]) # 시트명 제한 31자
            df_c = df_details[df_details['Company'] == comp].copy()
            
            # 헤더 2줄 그리미 (Simplified)
            ws_dtl.merge_cells('A1:H1'); ws_dtl['A1'] = f"{comp} - 상세 재무제표 (Report 기반)"; sc(ws_dtl['A1'], fo=fT)
            ws_dtl.merge_cells('A2:H2'); ws_dtl['A2'] = "DART finstate_all 원본 계정 정보 (최다 추출)"; sc(ws_dtl['A2'], fo=fS)
            
            if df_c.empty:
                ws_dtl['A4'] = "No detailed data available."
                continue
                
            pivot_df = df_c.pivot_table(
                index=['M_Key', 'Type', 'Account_ID', 'Account_NM'], 
                columns='Period', 
                values='Amount', 
                aggfunc='sum'
            ).reset_index()
            
            order_df = df_c.groupby(['M_Key', 'Type', 'Account_ID', 'Account_NM'])['Row_Idx'].min().reset_index()
            pivot_df = pd.merge(pivot_df, order_df, on=['M_Key', 'Type', 'Account_ID', 'Account_NM'], how='left')
            
            # Type 내림차순 정렬 유도 및 DART 한계 극복용 계층형 정렬 맵핑 (KPMG Style)
            sort_map = {'재무상태표': 1, '손익계산서': 2, '포괄손익계산서': 3, '현금흐름표': 4}
            def get_heuristic_rank(row):
                t = str(row['Type']).split()[0].replace('연결', '')
                acc = str(row['Account_NM'])
                idx = row['Row_Idx']
                t_rank = sort_map.get(t, 99) * 1000000
                
                if t in ['손익계산서', '포괄손익계산서']:
                    if '매출액' in acc or '영업수익' in acc: return t_rank + 10000
                    if '원가' in acc: return t_rank + 20000
                    if '매출총이익' in acc: return t_rank + 30000
                    if '판매비' in acc or '관리비' in acc: return t_rank + 40000
                    if '영업이익' in acc or '영업손실' in acc: return t_rank + 50000
                    if '법인세비용' in acc: return t_rank + 80000
                    if '당기순이익' in acc or '당기순손실' in acc: return t_rank + 90000
                    return t_rank + 60000 + idx
                elif t == '현금흐름표':
                    if acc == '영업활동현금흐름' or ('영업활동' in acc and '흐름' in acc): return t_rank + 10000
                    if acc == '투자활동현금흐름' or ('투자활동' in acc and '흐름' in acc): return t_rank + 40000
                    if acc == '재무활동현금흐름' or ('재무활동' in acc and '흐름' in acc): return t_rank + 70000
                    return t_rank + 10000 + idx
                return t_rank + idx
                
            pivot_df['SortKey'] = pivot_df.apply(get_heuristic_rank, axis=1)
            pivot_df = pivot_df.sort_values('SortKey').drop(columns=['SortKey', 'Row_Idx'])
            
            # 헤더 그리기
            labels = [p['label'] for p in periods_to_fetch]
            ws_dtl.cell(row=3, column=1, value="M_Key"); sc(ws_dtl.cell(row=3, column=1), fo=fH, fi=pH, al=aC, bd=BD)
            ws_dtl.cell(row=3, column=2, value="Type"); sc(ws_dtl.cell(row=3, column=2), fo=fH, fi=pH, al=aC, bd=BD)
            ws_dtl.cell(row=3, column=3, value="Account ID"); sc(ws_dtl.cell(row=3, column=3), fo=fH, fi=pH, al=aC, bd=BD)
            ws_dtl.cell(row=3, column=4, value="Account Name"); sc(ws_dtl.cell(row=3, column=4), fo=fH, fi=pH, al=aC, bd=BD)
            
            col_idx = 5
            for plabel in labels:
                ws_dtl.cell(row=3, column=col_idx, value=plabel); sc(ws_dtl.cell(row=3, column=col_idx), fo=fH, fi=pSEC2, al=aC, bd=BD)
                col_idx += 1
                
            r = 4
            for _, row in pivot_df.iterrows():
                ws_dtl.cell(row=r, column=1, value=row.get('M_Key', '')); sc(ws_dtl.cell(row=r, column=1), fo=fA, al=aL, bd=BD)
                ws_dtl.cell(row=r, column=2, value=row.get('Type', '')); sc(ws_dtl.cell(row=r, column=2), fo=fA, al=aL, bd=BD)
                ws_dtl.cell(row=r, column=3, value=row.get('Account_ID', '')); sc(ws_dtl.cell(row=r, column=3), fo=fA, al=aL, bd=BD)
                ws_dtl.cell(row=r, column=4, value=row.get('Account_NM', '')); sc(ws_dtl.cell(row=r, column=4), fo=fA, al=aL, bd=BD)
                
                c = 5
                for plabel in labels:
                    val = row.get(plabel)
                    v = val if pd.notna(val) else ""
                    ws_dtl.cell(row=r, column=c, value=v); sc(ws_dtl.cell(row=r, column=c), fo=fHARD, nf=NB, bd=BD)
                    c += 1
                r += 1
                
            ws_dtl.column_dimensions['A'].width = 15
            ws_dtl.column_dimensions['B'].width = 15
            ws_dtl.column_dimensions['C'].width = 25
            ws_dtl.column_dimensions['D'].width = 35
            for i in range(5, c):
                ws_dtl.column_dimensions[get_column_letter(i)].width = 15
                
            ws_dtl.freeze_panes = "E5"


    wb.save(output)
    output.seek(0)
    return output

pSEC3 = PatternFill('solid', fgColor='2E7D32'); pSEC4 = PatternFill('solid', fgColor='6A1B9A')
pSEC5 = PatternFill('solid', fgColor='C62828'); pSEC6 = PatternFill('solid', fgColor='455A64')

ev_fills = {
    'Cash': PatternFill('solid',fgColor=C_GR), 'IBD': PatternFill('solid',fgColor=C_YL),
    'IBD(Option)': PatternFill('solid',fgColor=C_YL), 'NOA(Option)': PatternFill('solid',fgColor=C_NOA),
    'NOA': PatternFill('solid',fgColor=C_NOA), 'NCI': PatternFill('solid',fgColor=C_PB),
    'Equity': PatternFill('solid',fgColor=C_LB), 'PL_HL': PatternFill('solid',fgColor=C_YL),
}

aC=Alignment(horizontal='center',vertical='center',wrap_text=True)
aL=Alignment(horizontal='left',vertical='center',indent=1)
aR=Alignment(horizontal='right',vertical='center')

NB='#,##0;(#,##0);"-"'; NB1='#,##0.0;(#,##0.0);"-"'; NI_FMT='#,##0;(#,##0);"-"'
NP='₩#,##0;(₩#,##0);"-"'; NF_X='#,##0.0x;(#,##0.0x);"-"'

def sc(c,fo=None,fi=None,al=None,bd=None,nf=None):
    if fo: c.font=fo
    if fi: c.fill=fi
    if al: c.alignment=al
    if bd: c.border=bd
    if nf: c.number_format=nf

def style_range(ws, r1, c1, r2, c2, fo=None, fi=None, al=None, bd=None, nf=None):
    for rr in range(r1, r2+1):
        for cc in range(c1, c2+1):
            sc(ws.cell(rr, cc), fo=fo, fi=fi, al=al, bd=bd, nf=nf)



def add_gpcm_section_row(ws):
    sec_row = 4
    sections = [
        (1, 2,  "Company Info",       pSEC1), (3, 5,  "Other Info",         pSEC2),
        (6, 12, "BS & EV Components", pSEC3), (13,17, "PL(Annual & LTM)",   pSEC4),
        (18,20, "Market Data",        pSEC5), (21,25, "Valuation Multiples", pSEC6),
        (26,35, "Beta & Risk Analysis", PatternFill('solid', fgColor='6A1B9A')),
    ]
    for c1, c2, label, fill in sections:
        ws.merge_cells(start_row=sec_row, start_column=c1, end_row=sec_row, end_column=c2)
        ws.cell(sec_row, c1).value = label
        style_range(ws, sec_row, c1, sec_row, c2, fo=fSEC, fi=fill, al=aC, bd=BD)


# ==========================================

def fetch_financial_data(api_key_input, target_code_list, target_periods, dart, status_container, progress_bar):
    # DART & KRX 로드 패스
    from datetime import datetime, timedelta
    import pandas as pd
    import numpy as np
    df_krx = get_krx_listing()
    
    # 변수 초기화
    base_period_str = target_periods[-1]
    base_year, base_qtr = parse_period(base_period_str)
    base_date_str = get_base_date_str(base_year, base_qtr)

    raw_bs_rows = []
    raw_pl_rows = []
    all_mkt = []
    ticker_to_name = {}

    screen_summary_data = []
    all_multiples = []

    total_tickers = len(target_code_list)

    for idx, ticker in enumerate(target_code_list):
        status_container.write(f"Processing [{ticker}] ({idx+1}/{total_tickers})...")
        progress_bar.progress((idx) / total_tickers)

        corp_code, krx_name = resolve_company_info(dart, ticker)
        if not corp_code:
            status_container.write(f"❌ [{ticker}] DART 고유번호 조회 실패")
            continue

        display_name = krx_name if krx_name else f"Company_{ticker}"
        ticker_to_name[ticker] = display_name

        # 임시 저장소 (화면 출력용) - 최신 기준일 데이터
        temp_metrics = {
            'Company': display_name, 'Ticker': ticker,
            'Market_Cap': 0, 'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA': 0, 'Equity': 0,
            'Revenue': 0, 'EBIT': 0, 'NI': 0, 'Pretax_Income': 0,
            'Stock_Monthly_Prices_5Y': None, 'Market_Monthly_Prices_5Y': None,
            'Stock_Weekly_Prices_2Y': None, 'Market_Weekly_Prices_2Y': None,
            'Exchange': 'KRX', 'Market_Index': 'KS11',
        }

        # DART API Call 최소화를 위한 로컬 캐시
        dart_fs_cache = {}

        for tp in target_periods:
            tyear, tqtr = parse_period(tp)
            required_periods = get_ltm_required_periods(tyear, tqtr)
            
            period_metrics = {
                'Market_Cap': 0, 'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA': 0, 'Equity': 0,
                'Revenue': 0, 'EBIT': 0, 'NI': 0, 'Pretax_Income': 0
            }

            for year, qtr, role in required_periods:
                r_code = RCODE_MAP[qtr]
                bds = get_base_date_str(year, qtr)

                # 1) Market Cap (기준시점만)
                if role in ('current_cum', 'annual'):
                    price, price_date = get_stock_price(ticker, bds)
                    shares, shares_src, sh_meta = get_outstanding_shares(api_key_input, corp_code, ticker, year, r_code, df_krx)

                    mkt_100m = 0
                    if price is not None and shares is not None and shares > 0:
                        mkt_100m = round((price * shares) / 1e8, 1)

                    period_metrics['Market_Cap'] = mkt_100m

                    all_mkt.append({
                        'Company': display_name, 'Ticker': ticker, 'Period': tp,
                        'Price_Date': price_date or bds, 'Close': price,
                        'Outstanding_Shares': shares, 'Market_Cap_100M': mkt_100m,
                        'Shares_Source': shares_src, 'Shares_RceptNo': sh_meta.get('rcept_no'),
                        'Shares_StlmDt': sh_meta.get('stlm_dt'), 'Shares_Se': sh_meta.get('se'),
                        'DART_Status': sh_meta.get('status'), 'DART_Message': sh_meta.get('message'),
                    })

                # 2) BS Fetch (finstate_all: 상세) - CFS 우선 → OFS
                if role in ('current_cum', 'annual'):
                    df_all = None
                    cache_key = f"all_{year}_{r_code}"
                    if cache_key in dart_fs_cache:
                        df_all = dart_fs_cache[cache_key]
                    else:
                        for fs in ['CFS', 'OFS']:
                            try:
                                _df = dart.finstate_all(corp_code, year, reprt_code=r_code, fs_div=fs)
                                if _df is not None and not _df.empty:
                                    df_all = _df
                                    dart_fs_cache[cache_key] = _df
                                    break
                            except:
                                continue

                    if df_all is not None and not df_all.empty:
                        df_bs = df_all[df_all['sj_nm'].astype(str).str.contains('상태표|재정상태', na=False)].copy()
                        for _, row in df_bs.iterrows():
                            amt = pd.to_numeric(str(row.get('thstrm_amount', '')).replace(',', ''), errors='coerce')
                            if pd.isna(amt) or amt == 0: continue

                            acct = str(row.get('account_nm', '')).strip()
                            aid = str(row.get('account_id', '')).strip()
                            ev_comp, _ = match_bs_ev_component(acct, aid)

                            if ev_comp:
                                # 화면 출력용 집계
                                val_100m = amt / 1e8
                                if ev_comp == 'Cash': period_metrics['Cash'] += val_100m
                                elif ev_comp == 'IBD': period_metrics['IBD'] += val_100m
                                elif ev_comp == 'NCI': period_metrics['NCI'] += val_100m
                                elif ev_comp == 'NOA': period_metrics['NOA'] += val_100m
                                elif ev_comp in ['Equity_Total', 'Equity_P']: period_metrics['Equity'] += val_100m

                            raw_bs_rows.append({
                                'Company': display_name, 'Ticker': ticker, 'Period': tp,
                                'sj_nm': row.get('sj_nm', ''), 'account_nm': acct, 'account_id': aid,
                                'EV_Component': ev_comp or '', 'Amount_100M': amt / 1e8,
                            })

                # 3) PL Fetch
                df_is = None
                cache_key_pl = f"pl_{year}_{r_code}"
                if cache_key_pl in dart_fs_cache:
                    df_is, pl_src = dart_fs_cache[cache_key_pl]
                else:
                    df_pl_raw, pl_src, pl_flag = fetch_pl_df(dart, corp_code, year, r_code)
                    if df_pl_raw is not None and not df_pl_raw.empty:
                        df_is = filter_income_statement(df_pl_raw)
                        dart_fs_cache[cache_key_pl] = (df_is, pl_src)
                    
                if df_is is None or df_is.empty: continue

                wanted = {'Revenue', 'EBIT', 'NI', 'Pretax_Income'}
                picked = set()

                for _, row in df_is.iterrows():
                    acct = str(row.get('account_nm', '')).strip()
                    aid = str(row.get('account_id', '')).strip()
                    calc_key = match_pl_core_only(acct, aid)
                    if not calc_key or calc_key not in wanted: continue
                    if calc_key in picked: continue

                    val = pick_pl_value(row, qtr)
                    if val is None: continue

                    amt_100m = val / 1e8
                    raw_pl_rows.append({
                        'Company': display_name, 'Ticker': ticker, 'Period': tp,
                        'Role': role, 'PL_Source': pl_src, 'account_nm': acct,
                        'calc_key': calc_key, 'Amount_100M': amt_100m,
                    })

                    if role in ('current_cum', 'annual'):
                        period_metrics[calc_key] += amt_100m
                    elif role == 'prior_annual':
                        period_metrics[calc_key] += amt_100m
                    elif role == 'prior_same_q':
                        period_metrics[calc_key] -= amt_100m

                    picked.add(calc_key)
                    if picked == wanted: break

            # Period loop ends, append to all_multiples
            all_multiples.append({
                'Company': display_name, 'Ticker': ticker, 'Period': tp,
                **period_metrics
            })
            
            # If this is the main base period, update temp_metrics
            if tp == base_period_str:
                temp_metrics.update(period_metrics)

        # 4) Beta Calculation (5Y Monthly, 2Y Weekly)
        exchange, market_idx = get_market_index(ticker)
        temp_metrics['Exchange'] = exchange
        temp_metrics['Market_Index'] = market_idx

        try:
            # 5년 월간 베타 데이터
            start_5y = (pd.to_datetime(base_date_str) - timedelta(days=365*5+20)).strftime('%Y-%m-%d')
            end_date = base_date_str

            # 주가 데이터 수집 (시장지수는 yf 사용 가능성 대응)
            stock_data_5y = fdr.DataReader(ticker, start_5y, end_date)
            if market_idx.startswith('^'):
                market_data_5y = yf.download(market_idx, start=start_5y, end=end_date, progress=False)
                if isinstance(market_data_5y.columns, pd.MultiIndex):
                    market_data_5y.columns = market_data_5y.columns.droplevel(1) # MultiIndex 제거
            else:
                market_data_5y = fdr.DataReader(market_idx, start_5y, end_date)

            if stock_data_5y is not None and not stock_data_5y.empty and market_data_5y is not None and not market_data_5y.empty:
                # Close 컬럼 추출
                stock_prices_5y = stock_data_5y['Close'] if 'Close' in stock_data_5y.columns else stock_data_5y.iloc[:, 0]
                market_prices_5y = market_data_5y['Close'] if 'Close' in market_data_5y.columns else market_data_5y.iloc[:, 0]

                # 인덱스를 timezone-naive DatetimeIndex로 변환
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

                if len(stock_monthly_prices) >= 12 and len(market_monthly_prices) >= 12:
                    temp_metrics['Stock_Monthly_Prices_5Y'] = stock_monthly_prices
                    temp_metrics['Market_Monthly_Prices_5Y'] = market_monthly_prices

            start_2y = (pd.to_datetime(base_date_str) - timedelta(days=365*2+20)).strftime('%Y-%m-%d')

            stock_data_2y = fdr.DataReader(ticker, start_2y, end_date)
            if market_idx.startswith('^'):
                market_data_2y = yf.download(market_idx, start=start_2y, end=end_date, progress=False)
                if isinstance(market_data_2y.columns, pd.MultiIndex):
                    market_data_2y.columns = market_data_2y.columns.droplevel(1)
            else:
                market_data_2y = fdr.DataReader(market_idx, start_2y, end_date)

            if stock_data_2y is not None and not stock_data_2y.empty and market_data_2y is not None and not market_data_2y.empty:
                # Close 컬럼 추출
                stock_prices_2y = stock_data_2y['Close'] if 'Close' in stock_data_2y.columns else stock_data_2y.iloc[:, 0]
                market_prices_2y = market_data_2y['Close'] if 'Close' in market_data_2y.columns else market_data_2y.iloc[:, 0]

                # 인덱스를 timezone-naive DatetimeIndex로 변환
                if not isinstance(stock_prices_2y.index, pd.DatetimeIndex):
                    stock_prices_2y.index = pd.to_datetime(stock_prices_2y.index)
                if stock_prices_2y.index.tz is not None:
                    stock_prices_2y.index = stock_prices_2y.index.tz_localize(None)

                if not isinstance(market_prices_2y.index, pd.DatetimeIndex):
                    market_prices_2y.index = pd.to_datetime(market_prices_2y.index)
                if market_prices_2y.index.tz is not None:
                    market_prices_2y.index = market_prices_2y.index.tz_localize(None)

                # 주간 종가 저장 (금요일 기준)
                stock_weekly_prices = stock_prices_2y.resample('W-FRI').last().dropna()
                market_weekly_prices = market_prices_2y.resample('W-FRI').last().dropna()

                if len(stock_weekly_prices) >= 50 and len(market_weekly_prices) >= 50:
                    temp_metrics['Stock_Weekly_Prices_2Y'] = stock_weekly_prices
                    temp_metrics['Market_Weekly_Prices_2Y'] = market_weekly_prices

        except Exception as e:
            pass  # Beta 데이터 수집 실패 시 계속 진행

        screen_summary_data.append(temp_metrics)
        time.sleep(0.5) # API 호출 간격 조절

    progress_bar.progress(1.0)
    status_container.update(label="분석 완료!", state="complete", expanded=False)

    # --- 결과 처리 및 엑셀 생성 ---

    return raw_bs_rows, raw_pl_rows, all_mkt, ticker_to_name, screen_summary_data, base_year, base_qtr, base_date_str, all_multiples

def calculate_wacc_and_beta(target_code_list, screen_summary_data, target_tax_rate_input, rf_input, mrp_input, size_premium_input, kd_pretax_input, beta_type_input):
    # 1.5. WACC Calculation (Target 기업용)
    # Beta 시트에서 계산될 Unlevered Beta를 엑셀에서 참조할 것이므로,
    # 여기서는 대략적인 값만 계산 (정확한 값은 엑셀 수식 기반)

    # 피어들의 평균 계산을 위한 준비
    avg_debt_ratios = []
    avg_unlevered_betas_5y = []
    avg_unlevered_betas_2y = []

    for ticker in target_code_list:
        comp_data = next((item for item in screen_summary_data if item["Ticker"] == ticker), None)
        if not comp_data:
            continue

        mkt_cap = comp_data.get('Market_Cap', 0)
        ibd = comp_data.get('IBD', 0)
        nci = comp_data.get('NCI', 0)
        equity = comp_data.get('Equity', 0)
        pretax_income = comp_data.get('Pretax_Income', 0)

        # Debt Ratio (D/V) = IBD / (Mkt Cap + IBD + NCI)
        total_value = mkt_cap + ibd + nci
        if total_value > 0:
            debt_ratio = ibd / total_value
            avg_debt_ratios.append(debt_ratio)

        # D/E Ratio = IBD / (Mkt Cap + NCI)
        equity_value = mkt_cap + nci
        de_ratio = ibd / equity_value if equity_value > 0 else 0

        # 한계세율 계산 (2025년 한국 법인세, 지방세 포함)
        tax_rate = get_korean_marginal_tax_rate(pretax_income)
        comp_data['Tax_Rate'] = tax_rate  # 저장 (나중에 Excel 출력용)

        # Beta 계산 (간단히 수익률 기반)
        stock_monthly_5y = comp_data.get('Stock_Monthly_Prices_5Y')
        market_monthly_5y = comp_data.get('Market_Monthly_Prices_5Y')
        stock_weekly_2y = comp_data.get('Stock_Weekly_Prices_2Y')
        market_weekly_2y = comp_data.get('Market_Weekly_Prices_2Y')

        # 5Y Monthly Beta
        if stock_monthly_5y is not None and market_monthly_5y is not None and not stock_monthly_5y.empty and not market_monthly_5y.empty:
            try:
                common_dates = stock_monthly_5y.index.intersection(market_monthly_5y.index)
                if len(common_dates) > 12:
                    stock_ret = stock_monthly_5y.loc[common_dates].pct_change().dropna()
                    market_ret = market_monthly_5y.loc[common_dates].pct_change().dropna()
                    common_idx = stock_ret.index.intersection(market_ret.index)
                    if len(common_idx) > 10:
                        stock_ret_aligned = stock_ret.loc[common_idx]
                        market_ret_aligned = market_ret.loc[common_idx]
                        cov_matrix = np.cov(stock_ret_aligned, market_ret_aligned)
                        beta_raw = cov_matrix[0, 1] / cov_matrix[1, 1] if cov_matrix[1, 1] != 0 else np.nan
                        beta_adj = (2/3) * beta_raw + (1/3) * 1

                        # Unlevered Beta = Adj Beta / (1 + (1 - Tax Rate) × D/E)
                        # tax_rate는 이미 위에서 get_korean_marginal_tax_rate()로 계산됨
                        if not np.isnan(beta_adj) and equity > 0:
                            unlevered_beta_5y = beta_adj / (1 + (1 - tax_rate) * de_ratio)
                            avg_unlevered_betas_5y.append(unlevered_beta_5y)
            except:
                pass

        # 2Y Weekly Beta
        if stock_weekly_2y is not None and market_weekly_2y is not None and not stock_weekly_2y.empty and not market_weekly_2y.empty:
            try:
                common_dates = stock_weekly_2y.index.intersection(market_weekly_2y.index)
                if len(common_dates) > 50:
                    stock_ret = stock_weekly_2y.loc[common_dates].pct_change().dropna()
                    market_ret = market_weekly_2y.loc[common_dates].pct_change().dropna()
                    common_idx = stock_ret.index.intersection(market_ret.index)
                    if len(common_idx) > 20:
                        stock_ret_aligned = stock_ret.loc[common_idx]
                        market_ret_aligned = market_ret.loc[common_idx]
                        cov_matrix = np.cov(stock_ret_aligned, market_ret_aligned)
                        beta_raw = cov_matrix[0, 1] / cov_matrix[1, 1] if cov_matrix[1, 1] != 0 else np.nan
                        beta_adj = (2/3) * beta_raw + (1/3) * 1

                        # Unlevered Beta
                        # tax_rate는 이미 위에서 get_korean_marginal_tax_rate()로 계산됨
                        if not np.isnan(beta_adj) and equity > 0:
                            unlevered_beta_2y = beta_adj / (1 + (1 - tax_rate) * de_ratio)
                            avg_unlevered_betas_2y.append(unlevered_beta_2y)
            except:
                pass

    # 평균값 계산
    avg_debt_ratio = np.mean(avg_debt_ratios) if avg_debt_ratios else 0.3

    # Beta Type에 따라 선택
    if beta_type_input == "5Y":
        avg_unlevered_beta = np.mean(avg_unlevered_betas_5y) if avg_unlevered_betas_5y else 0.8
    else:
        avg_unlevered_beta = np.mean(avg_unlevered_betas_2y) if avg_unlevered_betas_2y else 0.8

    # Target D/E Ratio 계산
    target_de_ratio = avg_debt_ratio / (1 - avg_debt_ratio) if avg_debt_ratio < 1 else 0

    # Relevered Beta 계산
    target_relevered_beta = avg_unlevered_beta * (1 + (1 - target_tax_rate_input) * target_de_ratio)

    # Ke (자기자본비용) 계산
    target_ke = rf_input + mrp_input * target_relevered_beta + size_premium_input

    # Kd (타인자본비용, 세후)
    kd_aftertax = kd_pretax_input * (1 - target_tax_rate_input)

    # E/V, D/V
    equity_weight = 1 - avg_debt_ratio
    debt_weight = avg_debt_ratio

    # Target WACC
    target_wacc = equity_weight * target_ke + debt_weight * kd_aftertax

    # WACC 데이터 저장
    target_wacc_data = {
        'Rf': rf_input,
        'MRP': mrp_input,
        'Size_Premium': size_premium_input,
        'Avg_Unlevered_Beta': avg_unlevered_beta,
        'Target_Tax_Rate': target_tax_rate_input,
        'Avg_Debt_Ratio': avg_debt_ratio,
        'Target_DE_Ratio': target_de_ratio,
        'Target_Relevered_Beta': target_relevered_beta,
        'Target_Ke': target_ke,
        'Kd_Pretax': kd_pretax_input,
        'Kd_Aftertax': kd_aftertax,
        'Equity_Weight': equity_weight,
        'Debt_Weight': debt_weight,
        'Target_WACC': target_wacc
    }
    return target_wacc_data, avg_debt_ratio

def export_gpcm_excel(base_period_str, base_qtr, target_code_list, screen_summary_data, raw_bs_rows, raw_pl_rows, all_mkt, ticker_to_name, target_wacc_data, beta_type_input, notes_list, avg_debt_ratio, base_date_str, df_screen, target_periods):
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.workbook.defined_name import DefinedName
    # 2. 엑셀 생성 (메모리)
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    # (기존 엑셀 생성 로직 그대로 활용 - 함수화 하지 않고 바로 실행)
    # Sheet 1: BS_Full
    ws_bs = wb.create_sheet('BS_Full')
    ws_bs.merge_cells('A1:H1'); ws_bs['A1'] = "BS_Full (Balance Sheet Detail)"; sc(ws_bs['A1'], fo=fT)
    ws_bs.merge_cells('A2:H2'); ws_bs['A2'] = "Logic: finstate_all(CFS→OFS) 재무상태표 라인아이템 수집 후 EV_Component 태깅 | Unit: 억원"; sc(ws_bs['A2'], fo=fS)
    cols = [('Company',15), ('Ticker',10), ('Period',10), ('sj_nm',15),('account_nm',35), ('account_id',40), ('EV_Component',12), ('Amount_100M',15)]
    header_row = 4
    ws_bs.append([]); ws_bs.append([c[0] for c in cols])
    for i, (_, w) in enumerate(cols): ws_bs.column_dimensions[get_column_letter(i+1)].width = w; sc(ws_bs.cell(header_row, i+1), fo=fH, fi=pH, al=aC, bd=BD)
    r = header_row + 1
    if raw_bs_rows:
        for rd in raw_bs_rows:
            ev_comp = rd['EV_Component']; is_hl = bool(ev_comp)
            fill_key = 'Equity' if ev_comp in ['Equity_P', 'Equity_Total'] else ev_comp
            row_fi = ev_fills.get(fill_key, pW) if is_hl else pW; row_fo = fHL if is_hl else fA
            vals = [rd['Company'], rd['Ticker'], rd['Period'], rd['sj_nm'],rd['account_nm'], rd['account_id'], rd['EV_Component'], rd['Amount_100M']]
            for i, v in enumerate(vals): sc(ws_bs.cell(r, i+1), fo=row_fo, fi=row_fi, al=aR if i==7 else aL, nf=NB if i==7 else None, bd=BD); ws_bs.cell(r, i+1).value = v
            r += 1
    ws_bs.auto_filter.ref = f"A{header_row}:H{r-1}"; ws_bs.freeze_panes = f"A{header_row+1}"

    # Sheet 2: PL_Data
    ws_pl = wb.create_sheet('PL_Data')
    ws_pl.merge_cells('A1:H1'); ws_pl['A1'] = "PL_Data (Income Statement Core Only)"; sc(ws_pl['A1'], fo=fT)
    ws_pl.merge_cells('A2:H2'); ws_pl['A2'] = "Logic: IS 추출 후 매출/영업이익/순이익 3개 계정만 엄격 추출 | Unit: 억원"; sc(ws_pl['A2'], fo=fS)
    cols = [('Company',15), ('Ticker',10), ('Period',10), ('Role',15),('PL_Source',16), ('account_nm',35), ('calc_key',12), ('Amount_100M',15)]
    header_row = 4
    ws_pl.append([]); ws_pl.append([c[0] for c in cols])
    for i, (_, w) in enumerate(cols): ws_pl.column_dimensions[get_column_letter(i+1)].width = w; sc(ws_pl.cell(header_row, i+1), fo=fH, fi=pH, al=aC, bd=BD)
    r = header_row + 1
    if raw_pl_rows:
        for rd in raw_pl_rows:
            vals = [rd['Company'], rd['Ticker'], rd['Period'], rd['Role'],rd['PL_Source'], rd['account_nm'], rd['calc_key'], rd['Amount_100M']]
            for i, v in enumerate(vals): sc(ws_pl.cell(r, i+1), fo=fHL, fi=ev_fills['PL_HL'], al=aR if i==7 else aL, nf=NB if i==7 else None, bd=BD); ws_pl.cell(r, i+1).value = v
            r += 1
    ws_pl.auto_filter.ref = f"A{header_row}:H{r-1}"; ws_pl.freeze_panes = f"A{header_row+1}"

    # Sheet 3: Market_Cap
    ws_mc = wb.create_sheet('Market_Cap')
    ws_mc.merge_cells('A1:M1'); ws_mc['A1'] = "Market_Cap (Price & Shares)"; sc(ws_mc['A1'], fo=fT)
    ws_mc.merge_cells('A2:M2'); ws_mc['A2'] = "Logic: 종가(FDR) × 유통주식수(DART) | Unit: 억원"; sc(ws_mc['A2'], fo=fS)
    cols = [('Company',15), ('Ticker',10), ('Period',10), ('Price_Date',12), ('Close',12),('Shares',16), ('Market_Cap_100M',18),('Shares_Source',12), ('Shares_RcpNo',16), ('Shares_StlmDt',12), ('Shares_Se',10),('DART_Status',10), ('DART_Message',40)]
    header_row = 4
    ws_mc.append([]); ws_mc.append([c[0] for c in cols])
    for i, (_, w) in enumerate(cols): ws_mc.column_dimensions[get_column_letter(i+1)].width = w; sc(ws_mc.cell(header_row, i+1), fo=fH, fi=pH, al=aC, bd=BD)
    r = header_row + 1
    if all_mkt:
        for rd in all_mkt:
            vals = [rd.get('Company'), rd.get('Ticker'), rd.get('Period'), rd.get('Price_Date'), rd.get('Close'),rd.get('Outstanding_Shares'), rd.get('Market_Cap_100M'),rd.get('Shares_Source'), rd.get('Shares_RceptNo'), rd.get('Shares_StlmDt'), rd.get('Shares_Se'),rd.get('DART_Status'), rd.get('DART_Message')]
            for i, v in enumerate(vals):
                c = ws_mc.cell(r, i+1); c.value = v
                nf = NP if i==4 else (NI_FMT if i==5 else (NB1 if i==6 else None)); al = aR if i in [4,5,6] else aL
                sc(c, fo=fA, fi=pW, al=al, nf=nf, bd=BD)
            r += 1
    ws_mc.auto_filter.ref = f"A{header_row}:M{r-1}"; ws_mc.freeze_panes = f"A{header_row+1}"

    # Sheet 4: LTM_Calc
    ws_ltm = wb.create_sheet('LTM_Calc')
    ws_ltm.merge_cells('A1:I1'); ws_ltm['A1'] = "LTM_Calc (Revenue/EBIT/NI/Pretax Inc)"; sc(ws_ltm['A1'], fo=fT)
    ws_ltm.merge_cells('A2:I2'); ws_ltm['A2'] = "모든 선택 기간별 LTM 계산 내역 | Unit: 억원"; sc(ws_ltm['A2'], fo=fS)
    cols = [('Company',15), ('Ticker',10), ('Period',10), ('Calc_Key',12),('Current_Cum(A)',15), ('Prior_Annual(B)',15), ('Prior_SameQ(C)',15), ('LTM_Value',15), ('Note',10)]
    header_row = 4
    ws_ltm.append([]); ws_ltm.append([c[0] for c in cols])
    for i, (_, w) in enumerate(cols): ws_ltm.column_dimensions[get_column_letter(i+1)].width = w; sc(ws_ltm.cell(header_row, i+1), fo=fH, fi=pH, al=aC, bd=BD)
    r = header_row + 1
    ltm_keys = ['Revenue', 'EBIT', 'NI', 'Pretax_Income']
    for ticker in target_code_list:
        comp_name = ticker_to_name.get(ticker, ticker)
        for tp in target_periods:
            qtr_suffix = tp.split('.')[-1] if '.' in tp else '4Q'
            for k in ltm_keys:
                ws_ltm.cell(r, 1, comp_name); sc(ws_ltm.cell(r, 1), fo=fA, fi=pW, al=aL, bd=BD)
                ws_ltm.cell(r, 2, ticker);    sc(ws_ltm.cell(r, 2), fo=fA, fi=pW, al=aL, bd=BD)
                ws_ltm.cell(r, 3, tp);        sc(ws_ltm.cell(r, 3), fo=fA, fi=pW, al=aL, bd=BD)
                ws_ltm.cell(r, 4, k);         sc(ws_ltm.cell(r, 4), fo=fA, fi=pW, al=aL, bd=BD)
                # Formula: SUMIFS sum_range, r1, criteria1, r2, criteria2...
                ws_ltm.cell(r, 5).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!C:C, C{r}, PL_Data!G:G, D{r}, PL_Data!D:D, "current_cum")'; sc(ws_ltm.cell(r,5), fo=fLINK, fi=pW, nf=NB, bd=BD)
                ws_ltm.cell(r, 6).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!C:C, C{r}, PL_Data!G:G, D{r}, PL_Data!D:D, "prior_annual")'; sc(ws_ltm.cell(r,6), fo=fLINK, fi=pW, nf=NB, bd=BD)
                ws_ltm.cell(r, 7).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!C:C, C{r}, PL_Data!G:G, D{r}, PL_Data!D:D, "prior_same_q")'; sc(ws_ltm.cell(r,7), fo=fLINK, fi=pW, nf=NB, bd=BD)
                if qtr_suffix == '4Q':
                    ws_ltm.cell(r, 8).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!C:C, C{r}, PL_Data!G:G, D{r}, PL_Data!D:D, "annual")'; note = 'Annual'
                else:
                    ws_ltm.cell(r, 8).value = f'=E{r}+F{r}-G{r}'; note = 'A+B-C'
                sc(ws_ltm.cell(r,8), fo=fFRM, fi=pW, nf=NB, bd=BD); ws_ltm.cell(r, 9).value = note; sc(ws_ltm.cell(r,9), fo=fA, fi=pW, al=aC, bd=BD)
                r += 1
    ws_ltm.auto_filter.ref = f"A{header_row}:I{r-1}"; ws_ltm.freeze_panes = f"A{header_row+1}"

    # Sheet 3.5: Beta_Calculation
    ws_beta = wb.create_sheet('Beta_Calculation')
    ws_beta.merge_cells('A1:F1')
    sc(ws_beta['A1'], fo=Font(name='Arial', bold=True, size=14, color=C_BL))
    ws_beta['A1'] = 'Beta Calculation (Excel Formulas)'

    ws_beta.merge_cells('A2:F2')
    sc(ws_beta['A2'], fo=Font(name='Arial', size=9, color=C_MG, italic=True))
    ws_beta['A2'] = f'5-Year Monthly & 2-Year Weekly Returns | Base: {base_period_str}'

    r_beta = 4
    beta_result_rows = {}  # ticker: (raw_5y, adj_5y, raw_2y, adj_2y) 매핑

    for idx, ticker in enumerate(target_code_list):
        comp_data = next((item for item in screen_summary_data if item["Ticker"] == ticker), None)
        if not comp_data:
            continue

        company_name = comp_data['Company']
        market_idx = comp_data['Market_Index']

        # ========== 5Y Monthly Beta Section ==========
        ws_beta.merge_cells(f'A{r_beta}:F{r_beta}')
        sc(ws_beta.cell(r_beta, 1), fo=Font(name='Arial', bold=True, size=10, color=C_W),
           fi=PatternFill('solid', fgColor='607D8B'), al=Alignment(horizontal='center'))
        ws_beta.cell(r_beta, 1, f'{company_name} ({ticker}) vs {market_idx} - 5Y Monthly')
        r_beta += 1

        stock_prices_5y = comp_data.get('Stock_Monthly_Prices_5Y')
        market_prices_5y = comp_data.get('Market_Monthly_Prices_5Y')
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

        stock_prices_2y = comp_data.get('Stock_Weekly_Prices_2Y')
        market_prices_2y = comp_data.get('Market_Weekly_Prices_2Y')
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

        # 결과 저장
        beta_result_rows[ticker] = (raw_5y_row, adj_5y_row, raw_2y_row, adj_2y_row)

        r_beta += 2  # 다음 회사와 간격

    ws_beta.column_dimensions['A'].width = 15
    ws_beta.column_dimensions['B'].width = 15
    ws_beta.column_dimensions['C'].width = 15
    ws_beta.column_dimensions['D'].width = 15
    ws_beta.column_dimensions['E'].width = 15

    ws_beta.freeze_panes = 'A4'

    # Sheet 4: WACC_Calculation (완전 구현 - GPCM.py와 동일)
    ws_wacc = wb.create_sheet('WACC_Calculation')
    ws_wacc.merge_cells('A1:D1')
    sc(ws_wacc['A1'], fo=Font(name='Arial', bold=True, size=14, color=C_BL))
    ws_wacc['A1'] = 'Target WACC Calculation'

    ws_wacc.merge_cells('A2:D2')
    sc(ws_wacc['A2'], fo=Font(name='Arial', size=9, color=C_MG, italic=True))
    ws_wacc['A2'] = f'Base: {base_period_str} | Peer Average Method'

    # 스타일 정의
    C_MB = '005EB8'
    pWACC_PARAM = PatternFill('solid', fgColor='E3F2FD')
    pWACC_CALC = PatternFill('solid', fgColor='FFF9C4')
    pWACC_RESULT = PatternFill('solid', fgColor='FFE082')

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
    # DATA_START = 6 (header_row + 1), DATA_END depends on number of companies
    # Mean row = DATA_END + 2
    n_companies = len(target_code_list)
    DATA_START = 6
    DATA_END = 6 + n_companies - 1
    mean_row = DATA_END + 2

    # 데이터 행 - Input Parameters
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

    # Section 2: Peer Analysis
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

    # Avg Unlevered Beta - 엑셀 수식으로 GPCM 시트 참조
    row_unlevered_beta = r_wacc
    beta_label = "5Y Monthly" if beta_type_input == "5Y" else "2Y Weekly"
    beta_col = 'AH' if beta_type_input == "5Y" else 'AI'  # AH = 컬럼 34 (Unlevered Beta 5Y), AI = 컬럼 35 (Unlevered Beta 2Y)
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
    ws_wacc.cell(r_wacc, 2).value = f'=GPCM!AG{mean_row}'  # 컬럼 33 (AG) = Debt Ratio (D/V)
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

    # Section 3: Target WACC Calculation
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
    ws_wacc.cell(r_wacc, 3, f"{target_wacc_data['Kd_Aftertax']*100:.2f}%")
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

    wb.defined_names['Target_WACC'] = DefinedName('Target_WACC', attr_text=f"'WACC_Calculation'!$B${row_wacc_final}")
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

    # Sheet 1: GPCM (맨 앞)
    ws = wb.create_sheet('GPCM')
    wb.move_sheet('GPCM', offset=-6)  # 맨 앞으로 이동 (index 0)
    # 시트 순서: GPCM, WACC_Calculation, Beta_Calculation, BS_Full, PL_Data, Market_Cap, LTM_Calc
    wb.move_sheet('WACC_Calculation', offset=-4)  # GPCM 다음 (index 1)
    wb.move_sheet('Beta_Calculation', offset=-3)  # WACC 다음 (index 2)
    TOTAL_COLS = 35
    ws.merge_cells(f'A1:{get_column_letter(TOTAL_COLS)}1'); ws['A1'] = "GPCM Valuation Summary with Beta Analysis"; sc(ws['A1'], fo=fT)
    ws.merge_cells(f'A2:{get_column_letter(TOTAL_COLS)}2'); ws['A2'] = f"Base: {base_period_str} | Unit: 억원 | EV = MCap + IBD − Cash + NCI − NOA"; sc(ws['A2'], fo=fS)
    add_gpcm_section_row(ws)
    headers = ['Company','Ticker','Base Date','Curr','PL Source','Cash','IBD','NOA','Net Debt','NCI','Equity','EV','Revenue','EBIT','D&A','EBITDA','NI','Price','Shares','Mkt Cap','EV/EBITDA','EV/EBIT','PER','PBR','PSR','β 5Y Raw','β 5Y Adj','β 2Y Raw','β 2Y Adj','Pretax Inc','Tax Rate','D/E Ratio','Debt Ratio (D/V)','Unlevered β 5Y','Unlevered β 2Y']
    widths = [18, 10, 11, 6, 13, 13, 13, 13, 13, 12, 13, 15, 13, 13, 10, 13, 13, 12, 15, 15, 12, 12, 10, 10, 10, 10, 10, 10, 10, 13, 9, 10, 10, 12, 12]
    header_row = 5
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        sc(ws.cell(header_row, i, h), fo=fH, fi=pH, al=aC, bd=BD)
    r = header_row + 1
    for ticker in target_code_list:
        comp_name = ticker_to_name.get(ticker, ticker); bg = pST if (r % 2 == 0) else pW
        ws.cell(r,1, comp_name); ws.cell(r,2, ticker); ws.cell(r,3, base_period_str); ws.cell(r,4, 'KRW'); ws.cell(r,5, 'LTM')
        for c in range(1, 6): sc(ws.cell(r,c), fo=fA, fi=bg, al=aL, bd=BD)
        ws.cell(r,6).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!C:C, C{r}, BS_Full!G:G, "Cash")'; sc(ws.cell(r,6), fo=fLINK, fi=ev_fills['Cash'], nf=NB, bd=BD)
        ws.cell(r,7).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!C:C, C{r}, BS_Full!G:G, "IBD")'; sc(ws.cell(r,7), fo=fLINK, fi=ev_fills['IBD'], nf=NB, bd=BD)
        ws.cell(r,8).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!C:C, C{r}, BS_Full!G:G, "NOA")'; sc(ws.cell(r,8), fo=fLINK, fi=ev_fills['NOA'], nf=NB, bd=BD)
        ws.cell(r,9).value = f'=G{r}-F{r}-H{r}'; sc(ws.cell(r,9), fo=fFRM, fi=bg, nf=NB, bd=BD)
        ws.cell(r,10).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!C:C, C{r}, BS_Full!G:G, "NCI")'; sc(ws.cell(r,10), fo=fLINK, fi=ev_fills['NCI'], nf=NB, bd=BD)
        ws.cell(r,11).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!C:C, C{r}, BS_Full!G:G, "Equity_Total")'; sc(ws.cell(r,11), fo=fLINK, fi=ev_fills['Equity'], nf=NB, bd=BD)
        ws.cell(r,12).value = f'=T{r}+G{r}-F{r}+J{r}-H{r}'; sc(ws.cell(r,12), fo=fFRM, fi=bg, nf=NB, bd=BD)
        ws.cell(r,13).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, C{r}, LTM_Calc!D:D, "Revenue")'; sc(ws.cell(r,13), fo=fLINK, fi=ev_fills['PL_HL'], nf=NB, bd=BD)
        ws.cell(r,14).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, C{r}, LTM_Calc!D:D, "EBIT")'; sc(ws.cell(r,14), fo=fLINK, fi=ev_fills['PL_HL'], nf=NB, bd=BD)
        sc(ws.cell(r,15), fi=PatternFill('solid', fgColor='FFFF00'), nf=NB, bd=BD) # D&A 수기
        ws.cell(r,16).value = f'=N{r}+O{r}'; sc(ws.cell(r,16), fo=fFRM, fi=bg, nf=NB, bd=BD)
        ws.cell(r,17).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, C{r}, LTM_Calc!D:D, "NI")'; sc(ws.cell(r,17), fo=fLINK, fi=ev_fills['PL_HL'], nf=NB, bd=BD)
        ws.cell(r,18).value = f'=SUMIFS(Market_Cap!E:E, Market_Cap!B:B, B{r}, Market_Cap!C:C, C{r})'; sc(ws.cell(r,18), fo=fLINK, nf=NP, bd=BD)
        ws.cell(r,19).value = f'=SUMIFS(Market_Cap!F:F, Market_Cap!B:B, B{r}, Market_Cap!C:C, C{r})'; sc(ws.cell(r,19), fo=fLINK, nf=NI_FMT, bd=BD)
        ws.cell(r,20).value = f'=SUMIFS(Market_Cap!G:G, Market_Cap!B:B, B{r}, Market_Cap!C:C, C{r})'; sc(ws.cell(r,20), fo=fLINK, nf=NB1, bd=BD)
        pMULT = PatternFill('solid', fgColor=C_PB)
        ws.cell(r,21).value = f'=IF(P{r}>0, L{r}/P{r}, "N/M")'; sc(ws.cell(r,21), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
        ws.cell(r,22).value = f'=IF(N{r}>0, L{r}/N{r}, "N/M")'; sc(ws.cell(r,22), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
        ws.cell(r,23).value = f'=IF(Q{r}>0, T{r}/Q{r}, "N/M")'; sc(ws.cell(r,23), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
        ws.cell(r,24).value = f'=IF(K{r}>0, T{r}/K{r}, "N/M")'; sc(ws.cell(r,24), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
        ws.cell(r,25).value = f'=IF(M{r}>0, T{r}/M{r}, "N/M")'; sc(ws.cell(r,25), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)

        # Z-AH: Beta & Risk Analysis (26-34번 컬럼)
        pBETA = PatternFill('solid', fgColor='E8F5E9')
        pBETA2 = PatternFill('solid', fgColor='FFF9C4')
        NF_BETA = '0.00;(0.00);"-"'
        NF_PCT = '0.0%;(0.0%);"-"'

        # Beta 값은 Beta_Calculation 시트에서 참조
        beta_rows = beta_result_rows.get(ticker, (None, None, None, None))
        if beta_rows[0]:  # Raw 5Y
            ws.cell(r,26).value = f'=Beta_Calculation!B{beta_rows[0]}'
            sc(ws.cell(r,26), fo=fLINK, fi=pBETA, al=aR, nf=NF_BETA, bd=BD)
        else:
            ws.cell(r,26, ''); sc(ws.cell(r,26), fo=fA, fi=pBETA, al=aR, nf=NF_BETA, bd=BD)

        if beta_rows[1]:  # Adj 5Y
            ws.cell(r,27).value = f'=Beta_Calculation!B{beta_rows[1]}'
            sc(ws.cell(r,27), fo=fLINK, fi=pBETA, al=aR, nf=NF_BETA, bd=BD)
        else:
            ws.cell(r,27, ''); sc(ws.cell(r,27), fo=fA, fi=pBETA, al=aR, nf=NF_BETA, bd=BD)

        if beta_rows[2]:  # Raw 2Y
            ws.cell(r,28).value = f'=Beta_Calculation!B{beta_rows[2]}'
            sc(ws.cell(r,28), fo=fLINK, fi=pBETA2, al=aR, nf=NF_BETA, bd=BD)
        else:
            ws.cell(r,28, ''); sc(ws.cell(r,28), fo=fA, fi=pBETA2, al=aR, nf=NF_BETA, bd=BD)

        if beta_rows[3]:  # Adj 2Y
            ws.cell(r,29).value = f'=Beta_Calculation!B{beta_rows[3]}'
            sc(ws.cell(r,29), fo=fLINK, fi=pBETA2, al=aR, nf=NF_BETA, bd=BD)
        else:
            ws.cell(r,29, ''); sc(ws.cell(r,29), fo=fA, fi=pBETA2, al=aR, nf=NF_BETA, bd=BD)

        # 컬럼 30: Pretax Inc (LTM_Calc에서 참조)
        ws.cell(r,30).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, C{r}, LTM_Calc!D:D, "Pretax_Income")'; sc(ws.cell(r,30), fo=fLINK, fi=bg, al=aR, nf=NB, bd=BD)

        # 컬럼 31: Tax Rate (한국 법인세 한계세율, 2025년 기준, 지방세 포함)
        # 2억 이하: 9.9%, 2~200억: 20.9%, 200~3000억: 23.1%, 3000억 초과: 26.4%
        ws.cell(r,31).value = f'=IF(AD{r}<=2, 0.099, IF(AD{r}<=200, 0.209, IF(AD{r}<=3000, 0.231, 0.264)))'
        sc(ws.cell(r,31), fo=fFRM, fi=bg, al=aR, nf=NF_PCT, bd=BD)

        # 컬럼 32: D/E Ratio = IBD / (Mkt Cap + NCI)
        ws.cell(r,32).value = f'=IF(T{r}+J{r}>0, G{r}/(T{r}+J{r}), 0)'; sc(ws.cell(r,32), fo=fFRM, fi=bg, al=aR, nf=NF_PCT, bd=BD)

        # 컬럼 33: Debt Ratio (D/V) = IBD / (Mkt Cap + IBD + NCI)
        ws.cell(r,33).value = f'=IF(T{r}+G{r}+J{r}>0, G{r}/(T{r}+G{r}+J{r}), 0)'; sc(ws.cell(r,33), fo=fFRM, fi=bg, al=aR, nf=NF_PCT, bd=BD)

        # 컬럼 34: Unlevered Beta 5Y = β 5Y Adj / (1 + (1 - Tax Rate) × D/E Ratio)
        # 컬럼 27 (AA) = β 5Y Adj, 컬럼 31 (AE) = Tax Rate, 컬럼 32 (AF) = D/E Ratio
        ws.cell(r,34).value = f'=IF(AA{r}>0, AA{r}/(1+(1-AE{r})*AF{r}), "")'; sc(ws.cell(r,34), fo=fFRM, fi=pBETA, al=aR, nf=NF_BETA, bd=BD)

        # 컬럼 35: Unlevered Beta 2Y = β 2Y Adj / (1 + (1 - Tax Rate) × D/E Ratio)
        # 컬럼 29 (AC) = β 2Y Adj, 컬럼 31 (AE) = Tax Rate, 컬럼 32 (AF) = D/E Ratio
        ws.cell(r,35).value = f'=IF(AC{r}>0, AC{r}/(1+(1-AE{r})*AF{r}), "")'; sc(ws.cell(r,35), fo=fFRM, fi=pBETA2, al=aR, nf=NF_BETA, bd=BD)
        r += 1
    r_end = r - 1; r += 1
    for stat, fn in [('Mean','AVERAGE'), ('Median','MEDIAN'), ('Max','MAX'), ('Min','MIN')]:
        ws.cell(r, 20, stat); sc(ws.cell(r,20), fo=fSTAT, fi=pSTAT, al=aC, bd=BD)
        # Valuation Multiples (21-25)
        for c in range(21, 26):
            col = get_column_letter(c)
            ws.cell(r, c).value = f'=IFERROR({fn}({col}{header_row+1}:{col}{r_end}), "N/M")'
            sc(ws.cell(r,c), fo=fSTAT, fi=pSTAT, nf=NF_X, bd=BD)
        # Beta & Risk (26-35)
        for c in range(26, 36):
            col = get_column_letter(c)
            if c in [26, 27, 28, 29, 34, 35]:  # Beta 컬럼 (34=Unlevered β 5Y, 35=Unlevered β 2Y)
                ws.cell(r, c).value = f'=IFERROR({fn}({col}{header_row+1}:{col}{r_end}), "")'
                sc(ws.cell(r,c), fo=fSTAT, fi=pSTAT, nf=NF_BETA, bd=BD)
            elif c == 31:  # Tax Rate
                ws.cell(r, c).value = f'=IFERROR({fn}({col}{header_row+1}:{col}{r_end}), "")'
                sc(ws.cell(r,c), fo=fSTAT, fi=pSTAT, nf=NF_PCT, bd=BD)
            elif c in [32, 33]:  # D/E Ratio, Debt Ratio (D/V)
                ws.cell(r, c).value = f'=IFERROR({fn}({col}{header_row+1}:{col}{r_end}), "")'
                sc(ws.cell(r,c), fo=fSTAT, fi=pSTAT, nf=NF_PCT, bd=BD)
            else:  # Pretax Inc (30)
                ws.cell(r, c).value = f'=IFERROR({fn}({col}{header_row+1}:{col}{r_end}), "")'
                sc(ws.cell(r,c), fo=fSTAT, fi=pSTAT, nf=NB, bd=BD)
        r += 1
    r += 2
    for note in notes_list: ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_COLS); sc(ws.cell(r, 1, note), fo=fNOTE); r += 1
    ws.freeze_panes = f"F{header_row+1}"  # Cash 컬럼부터 스크롤

    # === Multiples_Trend Sheet generation ===
    ws_trend = wb.create_sheet('Multiples_Trend')
    ws_trend.merge_cells('A1:M1'); ws_trend['A1'] = "Multiples Trend (PER, PBR, PSR, EV/EBIT)"; sc(ws_trend['A1'], fo=fT)
    ws_trend.merge_cells('A2:M2'); ws_trend['A2'] = "모든 타겟 기간별 Valuation Multiples 흐름 요약 (Formula 기반)"; sc(ws_trend['A2'], fo=fS)
    
    # 0:Comp, 1:Tick, 2:Per, 3:MC, 4:EV, 5:Rev, 6:EBIT, 7:NI, 8:Eq, 9:EV/EB, 10:PER, 11:PSR, 12:PBR
    cols_t = [('Company',15), ('Ticker',10), ('Period',10), ('Market_Cap',15), ('EV',15), ('Revenue(LTM)',15), ('EBIT(LTM)',15), ('NI(LTM)',15), ('Equity', 15), ('EV/EBIT',12), ('PER',10), ('PSR',10), ('PBR', 10)]
    header_row_t = 4
    ws_trend.append([]); ws_trend.append([c[0] for c in cols_t])
    for i, (_, w) in enumerate(cols_t): ws_trend.column_dimensions[get_column_letter(i+1)].width = w; sc(ws_trend.cell(header_row_t, i+1), fo=fH, fi=pH, al=aC, bd=BD)
    
    rt = header_row_t + 1
    if df_screen is not None and not df_screen.empty:
        # Ticker x Period 순회 (df_screen 기반)
        for _, row_data in df_screen.iterrows():
            ticker = row_data.get('Ticker')
            period = row_data.get('Period')
            comp_name = row_data.get('Company')
            
            # Static basic info
            ws_trend.cell(rt, 1, comp_name); sc(ws_trend.cell(rt, 1), fo=fA, fi=pW, al=aL, bd=BD)
            ws_trend.cell(rt, 2, ticker);    sc(ws_trend.cell(rt, 2), fo=fA, fi=pW, al=aL, bd=BD)
            ws_trend.cell(rt, 3, period);    sc(ws_trend.cell(rt, 3), fo=fA, fi=pW, al=aL, bd=BD)
            
            # MC & EV from Market_Cap sheet
            ws_trend.cell(rt, 4).value = f'=SUMIFS(Market_Cap!G:G, Market_Cap!B:B, B{rt}, Market_Cap!C:C, C{rt})'; sc(ws_trend.cell(rt, 4), fo=fLINK, nf=NB, bd=BD)
            ws_trend.cell(rt, 5).value = f'=D{rt}+SUMIFS(BS_Full!H:H, BS_Full!B:B, B{rt}, BS_Full!C:C, C{rt}, BS_Full!G:G, "IBD")-SUMIFS(BS_Full!H:H, BS_Full!B:B, B{rt}, BS_Full!C:C, C{rt}, BS_Full!G:G, "Cash")+SUMIFS(BS_Full!H:H, BS_Full!B:B, B{rt}, BS_Full!C:C, C{rt}, BS_Full!G:G, "NCI")-SUMIFS(BS_Full!H:H, BS_Full!B:B, B{rt}, BS_Full!C:C, C{rt}, BS_Full!G:G, "NOA")'
            sc(ws_trend.cell(rt, 5), fo=fFRM, nf=NB, bd=BD)
            
            # LTM Figures from LTM_Calc (Sumifs calc_key)
            ws_trend.cell(rt, 6).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{rt}, LTM_Calc!C:C, C{rt}, LTM_Calc!D:D, "Revenue")'; sc(ws_trend.cell(rt, 6), fo=fLINK, nf=NB, bd=BD)
            ws_trend.cell(rt, 7).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{rt}, LTM_Calc!C:C, C{rt}, LTM_Calc!D:D, "EBIT")'; sc(ws_trend.cell(rt, 7), fo=fLINK, nf=NB, bd=BD)
            ws_trend.cell(rt, 8).value = f'=SUMIFS(LTM_Calc!H:H, LTM_Calc!B:B, B{rt}, LTM_Calc!C:C, C{rt}, LTM_Calc!D:D, "NI")'; sc(ws_trend.cell(rt, 8), fo=fLINK, nf=NB, bd=BD)
            
            # Equity from BS_Full
            ws_trend.cell(rt, 9).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{rt}, BS_Full!C:C, C{rt}, BS_Full!G:G, "Equity_Total")'; sc(ws_trend.cell(rt, 9), fo=fLINK, nf=NB, bd=BD)
            
            # Multiples by Formula
            pMULT = PatternFill('solid', fgColor=C_PB)
            ws_trend.cell(rt, 10).value = f'=IF(G{rt}>0, E{rt}/G{rt}, "N/M")'; sc(ws_trend.cell(rt, 10), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
            ws_trend.cell(rt, 11).value = f'=IF(H{rt}>0, D{rt}/H{rt}, "N/M")'; sc(ws_trend.cell(rt, 11), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
            ws_trend.cell(rt, 12).value = f'=IF(F{rt}>0, D{rt}/F{rt}, "N/M")'; sc(ws_trend.cell(rt, 12), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
            ws_trend.cell(rt, 13).value = f'=IF(I{rt}>0, D{rt}/I{rt}, "N/M")'; sc(ws_trend.cell(rt, 13), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
            
            rt += 1
            
    ws_trend.auto_filter.ref = f"A{header_row_t}:M{rt-1}"
    ws_trend.freeze_panes = f"D{header_row_t+1}" # Scroll from Market_Cap

    wb.save(output)
    output.seek(0)

    return output
# 5. Streamlit App Layout & Logic
# ==========================================

# 사이드바 UI
with st.sidebar:
    st.header("Settings")
    
    # 좌측 1 : 기능 모드 선택 (신규)
    ui_mode = st.radio(
        "분석 모드 선택",
        ["GPCM Valuation (기존)", "다기간 재무제표 요약 (신규)"],
        index=0,
        help="GPCM 기반 가치평가 모드와 여러 회사의 과거 N년치 재무제표 요약 모드 중 하나를 선택하세요."
    )
    
    st.markdown("---")
    
    # 공통 입력 1: OpenDart API Key
    api_key_input = st.text_input("OpenDart API Key", type="password", help="OpenDart API 키를 입력하세요.")
    
    # 모드별 입력 파라미터 분기
    if ui_mode == "GPCM Valuation (기존)":
        # 다기간 GPCM 파라미터
        current_year = datetime.now().year
        
        st.write("**GPCM 분석 대상 기간**")
        g_cycle = st.radio("분석 주기", ["분기별 (Quarterly)", "연간별 (Annual)"], index=0, horizontal=True, help="연간별 선택 시 각 연도의 4Q(사업보고서) 데이터만 추출하여 트렌드를 구성합니다.")
        
        col1, col2 = st.columns(2)
        with col1:
            st.write("**시작 기간**")
            g_start_year = st.number_input("시작 연도", min_value=2015, max_value=2030, value=current_year - 1, step=1, key="gsy")
            g_start_qtr = "1Q"
            if g_cycle == "분기별 (Quarterly)":
                g_start_qtr = st.selectbox("시작 분기", ["1Q", "2Q", "3Q", "4Q"], index=0, key="gsq")
        with col2:
            st.write("**종료 기간 (기본 Base Date)**")
            g_end_year = st.number_input("종료 연도", min_value=2015, max_value=2030, value=current_year, step=1, key="gey")
            g_end_qtr = "4Q"
            if g_cycle == "분기별 (Quarterly)":
                g_end_qtr = st.selectbox("종료 분기", ["1Q", "2Q", "3Q", "4Q"], index=2, key="geq")
            
        target_periods = []
        qtrs = ["1Q", "2Q", "3Q", "4Q"]
        for y in range(g_start_year, g_end_year + 1):
            if g_cycle == "연간별 (Annual)":
                # 연간 모드에서는 해당 연도 4Q 데이터를 추가
                # 단, 종료 연도(gey)의 경우 사용자가 의도한 최신 데이터가 반영되도록 함.
                target_periods.append(f"{y}.4Q")
            else:
                s_idx = qtrs.index(g_start_qtr) if y == g_start_year else 0
                e_idx = qtrs.index(g_end_qtr) if y == g_end_year else 3
                for q_idx in range(s_idx, e_idx + 1):
                    target_periods.append(f"{y}.{qtrs[q_idx]}")
                
        if not target_periods:
            st.error("종료 기간이 시작 기간보다 빠릅니다.")
            st.stop()
            
        base_period_str = target_periods[-1]
        base_year, base_qtr = parse_period(base_period_str)
        base_date_display = get_base_date_str(base_year, base_qtr)
        
        st.info(f"Target WACC 기준일 (최신기간 적용): {base_date_display} (말일)")
        
        st.subheader("Target Companies")
        tickers_input = st.text_area("대상회사의 종목코드를 한줄씩 입력하세요", value="000250\n039030\n005290", height=150)

        st.subheader("Target WACC Parameters")
        rf_input = st.number_input("Rf - 무위험이자율 (%)", min_value=0.0, max_value=10.0, value=3.3, step=0.1, format="%.2f") / 100
        mrp_input = st.slider("MRP (시장위험프리미엄)", min_value=7.0, max_value=9.0, value=8.0, step=0.1, format="%.1f%%") / 100

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
        size_premium_choice = st.selectbox("기업 규모 선택", list(size_premium_options.keys()), index=0)
        size_premium_input = size_premium_options[size_premium_choice]

        beta_type_options = {"5년 월간 베타 (5Y Monthly)": "5Y", "2년 주간 베타 (2Y Weekly)": "2Y"}
        beta_type_choice = st.selectbox("WACC 계산에 사용할 Beta", list(beta_type_options.keys()), index=0)
        beta_type_input = beta_type_options[beta_type_choice]

        kd_pretax_input = st.number_input("Kd (Pretax) - 세전 이자율 (%)", min_value=0.0, max_value=15.0, value=3.5, step=0.1, format="%.1f") / 100
        target_tax_rate_input = st.number_input("Target 법인세율 (%)", min_value=0.0, max_value=50.0, value=26.4, step=0.1, format="%.1f") / 100

        run_btn = st.button("Go,Go,Go 🚀", type="primary", key="btn_gpcm")

    else:
        # 신규 다기간 재무제표 요약 파라미터
        current_year = datetime.now().year
        
        # 연간 vs 분기 조회 선택
        hist_period_type = st.radio(
            "조회 기준",
            ["연간 (사업보고서)", "분기 선택"],
            index=0,
            help="'연간'은 매년 연간 재무제표를 조회\n'분기 선택'은 특정 분기의 재무제표를 순차적으로 조회"
        )
        
        periods_to_fetch = []
        if hist_period_type == "연간 (사업보고서)":
            col1, col2 = st.columns(2)
            with col1:
                start_year = st.number_input("시작 연도", min_value=2015, max_value=2030, value=current_year - 3, step=1)
            with col2:
                end_year = st.number_input("종료 연도", min_value=2015, max_value=2030, value=current_year - 1, step=1)
            
            for y in range(start_year, end_year + 1):
                periods_to_fetch.append({'year': y, 'qtr': None, 'label': f"{y}년"})
        else:
            col1, col2 = st.columns(2)
            with col1:
                st.write("**시작 기간**")
                start_year = st.number_input("시작 연도", min_value=2015, max_value=2030, value=current_year - 1, step=1, key="sy_qtr")
                start_qtr = st.selectbox("시작 분기", ["1Q", "2Q", "3Q", "4Q"], index=0, key="sq_qtr")
            with col2:
                st.write("**종료 기간**")
                end_year = st.number_input("종료 연도", min_value=2015, max_value=2030, value=current_year, step=1, key="ey_qtr")
                end_qtr = st.selectbox("종료 분기", ["1Q", "2Q", "3Q", "4Q"], index=3, key="eq_qtr")
            
            qtrs = ["1Q", "2Q", "3Q", "4Q"]
            for y in range(start_year, end_year + 1):
                s_idx = qtrs.index(start_qtr) if y == start_year else 0
                e_idx = qtrs.index(end_qtr) if y == end_year else 3
                for q_idx in range(s_idx, e_idx + 1):
                    periods_to_fetch.append({'year': y, 'qtr': qtrs[q_idx], 'label': f"{y}년 {qtrs[q_idx]}"})
        
        st.subheader("Target Companies")
        tickers_input = st.text_area("대상회사의 종목코드를 한줄씩 입력하세요", value="000250\n039030\n005290", height=150)
        
        run_btn = st.button("재무제표 일괄 조회 🚀", type="primary", key="btn_hist")


# 메인 UI
if ui_mode == "GPCM Valuation (기존)":
    st.title("GPCM Calculator with Dart/KRX")
    st.markdown("""
    Opendartreader, Financedatareadr 라이브러리를 활용하여 기준일 시점 선정된 Peer의 재무제표, 주가, 유통주식수 등을 크롤링하여 GPCM Multiple을 계산하는 App 입니다. 
    해당 App 사용을 위해서는 **OpenDart API 인증키**를 개별적으로 발급받으셔야 합니다. 
    **감가상각비는 Dart에서 자동으로 불러올 수 없으니 EBITDA 계산 시 엑셀에서 수기로 입력하셔야 합니다.**
    (Made by SGJ _260211)
    """)

    st.markdown("---")
    st.subheader("📝 Valuation Methodology Notes")
    notes_list = [
        f'• Base Date: {base_period_str} ({base_date_display}) | Unit: 억원 (KRW 100M)',
        '• 공통: 연결재무제표 작성 시 CFS 우선, 미존재 시 OFS 기준으로 수집',
        '• PL: 요약 손익계산서에서 매출액/영업이익/당기순이익 3개 계정만 엄격 추출',
        '• PL Fetch: finstate(CFS/OFS) → finstate(no fs_div) → finstate_all fallback',
        '• Shares: DART(stockTotqySttus) 유통주식수(distb_stock_co) 우선, 미공시 시 DART 과거보고서 fallback',
        '• EV = Market Cap + IBD − Cash + NCI − NOA',
        '• Net Debt = IBD − Cash − NOA',
        '• IBD(Option): CB/EB/BW 등 메자닌은 기본적으로 IBD(Option)으로 태깅되어 EV/NetDebt에서 제외됨',
        '• NOA(Option): 투자자산/관계기업 등은 기본적으로 NOA(Option)으로 태깅되어 EV/NetDebt에서 제외됨',
        '• LTM = Current Cumulative + Prior Annual − Prior Same Quarter Cumulative (단, 4Q는 Annual)',
        '• Beta: 5년 월간 & 2년 주간 수익률 기준 (FinanceDataReader 사용)',
        '• Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1',
        '• D/E Ratio = IBD / (Market Cap + NCI)',
        '• Debt Ratio (D/V) = IBD / (Market Cap + IBD + NCI)',
        '• Unlevered Beta = Levered Beta / (1 + (1 - Tax Rate) × D/E Ratio)',
        '• Tax Rate: 한국 법인세 한계세율 (지방세 포함, 세전순이익 기준)',
    ]
    for note in notes_list:
        st.text(note)
    st.markdown("---")
else:
    st.title("📚 다기간 과거 재무제표 및 지표 요약 조회")
    st.markdown("""
    선택한 회사들의 과거 **연간 및 분기 재무제표(재무상태표, 손익계산서, 현금흐름표)**를 DART API를 통해 일괄 조회하고 엑셀로 추출합니다.
    - **조회 대상 기간**: 지정하신 시작/종료 연도 및 분기에 해당하는 공시 자료를 순차적으로 모두 수집합니다.
    - **결과물**: 여러 회사를 가로축으로 한눈에 비교할 수 있는 Summary 시트 + 각 회사별 상세 과거 재무제표 시트
    - **현금흐름표 처리**: 영업활동/투자활동/재무활동 등 대분류 현금흐름만 Summary에 표시되며 상세정보는 개별 시트에 기록됩니다.
    """)
    st.markdown("---")

# ▼▼▼▼▼ [추가 1] DART 접속 객체를 캐싱하는 함수 (if run_btn 바로 위에 넣으세요) ▼▼▼▼▼
@st.cache_resource
def get_dart_reader(api_key):
    return OpenDartReader(api_key)
# ▲▲▲▲▲ [여기까지 추가] ▲▲▲▲▲


# 실행 로직
if run_btn:
    if not api_key_input:
        st.error("OpenDart API Key를 입력해주세요.")
    else:
        target_code_list = [t.strip() for t in tickers_input.split('\n') if t.strip()]
        if not target_code_list:
            st.error("종목코드를 입력해주세요.")
        else:
            if ui_mode == "GPCM Valuation (기존)":
                # ==========================================
                # [모드 1] 기존 GPCM Valuation 로직
                # ==========================================
                status_container = st.status("GPCM 데이터 분석 중...", expanded=True)
                progress_bar = st.progress(0)

                try:
                    dart = get_dart_reader(api_key_input)
                except Exception as e:
                    st.error(f"DART 서버 접속 실패: {e}")
                    st.stop()

                # 변수 초기화 및 데이터 수집
                raw_bs_rows, raw_pl_rows, all_mkt, ticker_to_name, screen_summary_data, base_year, base_qtr, base_date_str, all_multiples = fetch_financial_data(
                    api_key_input, target_code_list, target_periods, dart, status_container, progress_bar)

                # 1. 화면 출력용 DataFrame 구성
                df_screen = pd.DataFrame(all_multiples)
                if not df_screen.empty:
                    df_screen['EV'] = df_screen['Market_Cap'] + df_screen['IBD'] - df_screen['Cash'] + df_screen['NCI'] - df_screen['NOA']
                    df_screen['EV/EBIT'] = np.where(df_screen['EBIT'] > 0, df_screen['EV'] / df_screen['EBIT'], np.nan)
                    df_screen['PER'] = np.where(df_screen['NI'] > 0, df_screen['Market_Cap'] / df_screen['NI'], np.nan)
                    df_screen['PSR'] = np.where(df_screen['Revenue'] > 0, df_screen['Market_Cap'] / df_screen['Revenue'], np.nan)

                    st.subheader("📊 Multiples Table (Preview)")
                    st.dataframe(df_screen[['Company', 'Period', 'Market_Cap', 'EV', 'Revenue', 'EBIT', 'NI', 'EV/EBIT', 'PER', 'PSR']]
                                 .style.format("{:.1f}", subset=['Market_Cap','EV','Revenue','EBIT','NI'])
                                 .format("{:.1f}x", subset=['EV/EBIT','PER','PSR'], na_rep="N/M"))
                    
                    st.subheader("📈 Statistics (Mean/Median - Latest Period)")
                    latest_df = df_screen[df_screen['Period'] == base_period_str]
                    if not latest_df.empty:
                        stats = latest_df[['EV/EBIT', 'PER', 'PSR']].agg(['mean', 'median', 'max', 'min'])
                        st.dataframe(stats.style.format("{:.1f}x"))

                # 1.5. WACC Calculation (Target 기업용)
                target_wacc_data, avg_debt_ratio = calculate_wacc_and_beta(
                    target_code_list, screen_summary_data, target_tax_rate_input, rf_input, mrp_input, size_premium_input, kd_pretax_input, beta_type_input)

                # 2. 엑셀 생성 (메모리)
                output = export_gpcm_excel(
                    base_period_str, base_qtr, target_code_list, screen_summary_data, raw_bs_rows, raw_pl_rows, all_mkt, ticker_to_name,
                    target_wacc_data, beta_type_input, notes_list, avg_debt_ratio, base_date_str, df_screen, target_periods)
                st.success("분석 완료! 아래 버튼을 눌러 리포트를 다운로드하세요.")
                st.download_button(
                    label="📥 Report Download (Excel)",
                    data=output,
                    file_name=f"KR_GPCM_Fixed_{base_period_str.replace('.','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            elif ui_mode == "다기간 재무제표 요약 (신규)":
                # ==========================================
                # [모드 2] 다기간 재무제표 요약 로직
                # ==========================================
                status_container = st.status(f"다기간 재무제표 데이터 수집 중... (대상: {len(target_code_list)}개 기업 / 기간: {len(periods_to_fetch)}개)", expanded=True)
                progress_bar = st.progress(0)

                try:
                    dart = get_dart_reader(api_key_input)
                except Exception as e:
                    st.error(f"DART 서버 접속 실패: {e}")
                    st.stop()
                
                # 1. 데이터 수집
                df_krx = get_krx_listing()
                df_summ, df_details = fetch_historical_financials(
                    api_key_input, target_code_list, periods_to_fetch,
                    dart, status_container, progress_bar, df_krx
                )
                
                # 2. 지표 계산
                status_container.update(label="지표 계산 및 엑셀 리포트 생성 중...")
                df_summ = calculate_historical_metrics(df_summ)
                
                # 3. 엑셀 생성
                if not df_summ.empty:
                    output = export_historical_excel(df_summ, df_details, periods_to_fetch)
                    
                    status_container.update(label="분석 완료!", state="complete")
                    st.success("데이터 추출이 완료되었습니다. 아래 버튼을 눌러 리포트를 다운로드하세요.")
                    
                    st.subheader("📊 Summary Preview")
                    # 화면 표시용: 일부 핵심 지표만 나열
                    preview_cols = ['Company', 'Period', 'Revenue', 'EBIT', 'NI', 'OPM', 'ROE']
                    avail_cols = [c for c in preview_cols if c in df_summ.columns]
                    
                    # 포맷 적용 (문자열/정수 컬럼 제외)
                    num_cols = [c for c in avail_cols if c in ('Revenue', 'EBIT', 'NI')]
                    pct_cols = [c for c in avail_cols if c in ('OPM', 'ROE')]
                    
                    styler = df_summ[avail_cols].style
                    if num_cols: styler = styler.format("{:,.0f}", subset=num_cols, na_rep="")
                    if pct_cols: styler = styler.format("{:.1%}", subset=pct_cols, na_rep="")
                    st.dataframe(styler)
                    
                    st.download_button(
                        label="📥 Report Download (Excel)",
                        data=output,
                        file_name=f"KR_Historical_Financials_{start_year}_to_{end_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    status_container.update(label="데이터 수집 실패", state="error")
                    st.warning("수집된 데이터가 없습니다. 종목코드나 연도를 다시 한번 확인해주세요.")
