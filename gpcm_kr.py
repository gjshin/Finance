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

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
def parse_period(p: str):
    parts = p.strip().split('.')
    return int(parts[0]), parts[1]

def get_base_date_str(year: int, qtr: str):
    return f"{year}-{QUARTER_INFO[qtr]}"

def get_ltm_required_periods(base_period: str):
    year, qtr = parse_period(base_period)
    if qtr == '4Q':
        return [(year, '4Q', 'annual')]
    return [
        (year, qtr, 'current_cum'),
        (year - 1, '4Q', 'prior_annual'),
        (year - 1, qtr, 'prior_same_q'),
    ]

@st.cache_resource
def get_krx_listing():
    return fdr.StockListing('KRX')

def resolve_company_info(dart_instance, ticker: str):
    df_krx = get_krx_listing()
    rows = df_krx[df_krx['Code'] == ticker]
    krx_name = rows.iloc[0]['Name'] if not rows.empty else None

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
    shares, meta = fetch_dart_distb_shares(api_key, corp_code, bsns_year, reprt_code)
    if shares is not None and shares > 0:
        return shares, 'DART', meta

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
PL_REVENUE = {'매출액', '수익(매출액)', '수익(매출)', '영업수익'}
PL_EBIT    = {'영업이익', '영업이익(손실)', '영업손실', '영업손익'}
PL_NI      = {'당기순이익', '당기순이익(손실)', '당기순손실', '분기순이익', '분기순이익(손실)', '분기순손실', '반기순이익', '반기순이익(손실)', '반기순손실', '연결당기순이익', '연결당기순이익(손실)', '연결당기순손실'}

def _norm_pl(s):
    s = "" if s is None else str(s).strip()
    return re.sub(r"\s+", "", s)

def match_pl_core_only(account_nm):
    a = _norm_pl(account_nm)
    if a in PL_REVENUE: return 'Revenue'
    if a in PL_EBIT:    return 'EBIT'
    if a in PL_NI:      return 'NI'
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
fT=Font(name='Arial',bold=True,size=14,color=C_BL)
fS=Font(name='Arial',size=9,color=C_MG,italic=True)
fH=Font(name='Arial',bold=True,size=9,color=C_W)
fA=Font(name='Arial',size=9,color=C_DG)
fHL=Font(name='Arial',bold=True,size=9,color=C_DB)
fMUL=Font(name='Arial',bold=True,size=10,color=C_BL)
fNOTE=Font(name='Arial',size=8,color=C_MG,italic=True)
fSTAT=Font(name='Arial',bold=True,size=9,color=C_DB)
fFRM=Font(name='Arial',size=9,color='000000'); fLINK=Font(name='Arial',size=9,color='008000')
fSEC = Font(name='Arial', bold=True, size=10, color=C_W)

pH=PatternFill('solid',fgColor=C_BL); pW=PatternFill('solid',fgColor=C_W)
pST=PatternFill('solid',fgColor=C_LG); pSTAT=PatternFill('solid',fgColor=C_LB)
pSEC1 = PatternFill('solid', fgColor=C_DB); pSEC2 = PatternFill('solid', fgColor=C_BL)
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

def add_sheet_title_block(ws, title, subtitle, end_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws.cell(1, 1).value = title
    sc(ws.cell(1,1), fo=fT, al=aL)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws.cell(2, 1).value = subtitle
    sc(ws.cell(2,1), fo=fS, al=aL)

def add_gpcm_section_row(ws):
    sec_row = 4
    sections = [
        (1, 2,  "Company Info",       pSEC1), (3, 5,  "Other Info",         pSEC2),
        (6, 12, "BS & EV Components", pSEC3), (13,17, "PL(Annual & LTM)",   pSEC4),
        (18,20, "Market Data",        pSEC5), (21,25, "Valuation Multiples", pSEC6),
    ]
    for c1, c2, label, fill in sections:
        ws.merge_cells(start_row=sec_row, start_column=c1, end_row=sec_row, end_column=c2)
        ws.cell(sec_row, c1).value = label
        style_range(ws, sec_row, c1, sec_row, c2, fo=fSEC, fi=fill, al=aC, bd=BD)


# ==========================================
# 5. Streamlit App Layout & Logic
# ==========================================

# 사이드바 UI
with st.sidebar:
    st.header("Settings")
    
    # 좌측 1 : OpenDart API Key
    api_key_input = st.text_input("OpenDart API Key", type="password", help="OpenDart API 키를 입력하세요.")
    
    # 좌측 2 : 기준일 설정
    current_year = datetime.now().year
    sel_year = st.number_input("Year", min_value=2020, max_value=2030, value=current_year, step=1)
    sel_qtr = st.selectbox("Quarter", ["1Q", "2Q", "3Q", "4Q"], index=2)
    
    base_period_str = f"{sel_year}.{sel_qtr}"
    base_date_display = get_base_date_str(sel_year, sel_qtr)
    st.info(f"기준일: {base_date_display} (말일)")
    
    # 좌측 3 : 종목코드 입력
    st.subheader("Target Companies")
    tickers_input = st.text_area("대상회사의 종목코드를 한줄씩 입력하세요", value="000250\n039030\n005290", height=150)
    
    # 좌측 4 : 실행 버튼
    run_btn = st.button("Go,Go,Go 🚀", type="primary")

# 메인 UI
st.title("GPCM Calculator with Dart/KRX")

st.markdown("""
Opendartreader, Financedatareadr 라이브러리를 활용하여 기준일 시점 선정된 Peer의 재무제표, 주가, 유통주식수 등을 크롤링하여 GPCM Multiple을 계산하는 App 입니다. 
해당 App 사용을 위해서는 **OpenDart API 인증키**를 개별적으로 발급받으셔야 합니다. 
**감가상각비는 Dart에서 자동으로 불러올 수 없으니 EBITDA 계산 시 엑셀에서 수기로 입력하셔야 합니다.**
(Made by SGJ _260211)
""")

# Valuation Methodology Notes (Always visible)
st.markdown("---")
st.subheader("📝 Valuation Methodology Notes")
notes_list = [
    f'• Base Date: {base_period_str} ({base_date_display}) | Unit: 억원 (KRW 100M)',
    '• 공통: 연결재무제표 작성 시 CFS 우선, 미존재 시 OFS 기준으로 수집',
    '• PL: 요약 손익계산서에서 매출액/영업이익/당기순이익 3개 계정만 엄격 추출',
    '• PL Fetch: finstate(CFS/OFS) → finstate(no fs_div) → finstate_all fallback',
    '• Shares: DART(stockTotqySttus) 유통주식수(distb_stock_co) 우선, 미공시 시 KRX fallback',
    '• EV = Market Cap + IBD − Cash + NCI − NOA',
    '• Net Debt = IBD − Cash − NOA',
    '• IBD(Option): CB/EB/BW 등 메자닌은 기본적으로 IBD(Option)으로 태깅되어 EV/NetDebt에서 제외됨',
    '• NOA(Option): 투자자산/관계기업 등은 기본적으로 NOA(Option)으로 태깅되어 EV/NetDebt에서 제외됨',
    '• LTM = Current Cumulative + Prior Annual − Prior Same Quarter Cumulative (단, 4Q는 Annual)',
]
for note in notes_list:
    st.text(note)
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
        # 데이터 처리 시작
        target_code_list = [t.strip() for t in tickers_input.split('\n') if t.strip()]
        
        if not target_code_list:
            st.error("종목코드를 입력해주세요.")
        else:
            # 상태 표시 컨테이너
            status_container = st.status("데이터 분석 중...", expanded=True)
            progress_bar = st.progress(0)

            # ▼▼▼▼▼ [수정 2] API 초기화 (캐싱 함수 사용 + 에러 처리) ▼▼▼▼▼
            try:
                dart = get_dart_reader(api_key_input)
            except Exception as e:
                st.error(f"DART 서버 접속 실패 (잠시 후 다시 시도하거나 로컬에서 실행하세요): {e}")
                st.stop()
            # ▲▲▲▲▲ [여기까지 수정] ▲▲▲▲▲
            
            # 변수 초기화
            base_year, base_qtr = parse_period(base_period_str)
            base_date_str = get_base_date_str(base_year, base_qtr)
            required_periods = get_ltm_required_periods(base_period_str)
            
            raw_bs_rows = []
            raw_pl_rows = []
            all_mkt = []
            ticker_to_name = {}
            
            # 화면 표시용 요약 데이터를 담을 리스트
            screen_summary_data = []

            df_krx = get_krx_listing()
            
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
                
                # 임시 저장소 (화면 출력용)
                temp_metrics = {
                    'Company': display_name, 'Ticker': ticker,
                    'Market_Cap': 0, 'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA': 0,
                    'Revenue': 0, 'EBIT': 0, 'NI': 0
                }

                for year, qtr, role in required_periods:
                    period = f"{year}.{qtr}"
                    r_code = RCODE_MAP[qtr]
                    bds = f"{year}-{QUARTER_INFO[qtr]}"

                    # 1) Market Cap (기준시점만)
                    if role in ('current_cum', 'annual'):
                        price, price_date = get_stock_price(ticker, bds)
                        shares, shares_src, sh_meta = get_outstanding_shares(api_key_input, corp_code, ticker, year, r_code, df_krx)

                        mkt_100m = 0
                        if price is not None and shares is not None and shares > 0:
                            mkt_100m = round((price * shares) / 1e8, 1)
                        
                        temp_metrics['Market_Cap'] = mkt_100m

                        all_mkt.append({
                            'Company': display_name, 'Ticker': ticker, 'Period': period,
                            'Price_Date': price_date or bds, 'Close': price,
                            'Outstanding_Shares': shares, 'Market_Cap_100M': mkt_100m,
                            'Shares_Source': shares_src, 'Shares_RceptNo': sh_meta.get('rcept_no'),
                            'Shares_StlmDt': sh_meta.get('stlm_dt'), 'Shares_Se': sh_meta.get('se'),
                            'DART_Status': sh_meta.get('status'), 'DART_Message': sh_meta.get('message'),
                        })

                    # 2) BS Fetch (finstate_all: 상세) - CFS 우선 → OFS
                    if role in ('current_cum', 'annual'):
                        df_all = None
                        for fs in ['CFS', 'OFS']:
                            try:
                                _df = dart.finstate_all(corp_code, year, reprt_code=r_code, fs_div=fs)
                                if _df is not None and not _df.empty:
                                    df_all = _df
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
                                    if ev_comp == 'Cash': temp_metrics['Cash'] += val_100m
                                    elif ev_comp == 'IBD': temp_metrics['IBD'] += val_100m
                                    elif ev_comp == 'NCI': temp_metrics['NCI'] += val_100m
                                    elif ev_comp == 'NOA': temp_metrics['NOA'] += val_100m

                                raw_bs_rows.append({
                                    'Company': display_name, 'Ticker': ticker, 'Period': period,
                                    'sj_nm': row.get('sj_nm', ''), 'account_nm': acct, 'account_id': aid,
                                    'EV_Component': ev_comp or '', 'Amount_100M': amt / 1e8,
                                })

                    # 3) PL Fetch
                    df_pl_raw, pl_src, pl_flag = fetch_pl_df(dart, corp_code, year, r_code)
                    if df_pl_raw is None or df_pl_raw.empty: continue

                    df_is = filter_income_statement(df_pl_raw)
                    if df_is is None or df_is.empty: continue

                    wanted = {'Revenue', 'EBIT', 'NI'}
                    picked = set()
                    
                    for _, row in df_is.iterrows():
                        acct = str(row.get('account_nm', '')).strip()
                        calc_key = match_pl_core_only(acct)
                        if not calc_key or calc_key not in wanted: continue
                        if calc_key in picked: continue

                        val = pick_pl_value(row, qtr)
                        if val is None: continue

                        amt_100m = val / 1e8
                        raw_pl_rows.append({
                            'Company': display_name, 'Ticker': ticker, 'Period': period,
                            'Role': role, 'PL_Source': pl_src, 'account_nm': acct,
                            'calc_key': calc_key, 'Amount_100M': amt_100m,
                        })
                        
                        # 화면 출력용 LTM 계산 (단순 합산 로직이 필요하므로 여기서는 각 Role별 값을 저장하고 나중에 합산해야 함)
                        # 여기서는 구조상 복잡해지므로 DataFrame 생성 시 처리
                        picked.add(calc_key)
                        if picked == wanted: break
                
                screen_summary_data.append(temp_metrics)
                time.sleep(0.5) # API 호출 간격 조절

            progress_bar.progress(1.0)
            status_container.update(label="분석 완료!", state="complete", expanded=False)

            # --- 결과 처리 및 엑셀 생성 ---
            
            # 1. 화면 출력용 DataFrame 구성 (LTM 계산 포함)
            if raw_pl_rows:
                df_pl_all = pd.DataFrame(raw_pl_rows)
                # LTM 계산 로직 구현 for Screen
                ltm_res = []
                for ticker in target_code_list:
                    d_sub = df_pl_all[df_pl_all['Ticker'] == ticker]
                    metrics = {'Revenue':0, 'EBIT':0, 'NI':0}
                    for k in metrics.keys():
                        curr_cum = d_sub[(d_sub['calc_key']==k) & (d_sub['Role']=='current_cum')]['Amount_100M'].sum()
                        prior_ann = d_sub[(d_sub['calc_key']==k) & (d_sub['Role']=='prior_annual')]['Amount_100M'].sum()
                        prior_same = d_sub[(d_sub['calc_key']==k) & (d_sub['Role']=='prior_same_q')]['Amount_100M'].sum()
                        annual_only = d_sub[(d_sub['calc_key']==k) & (d_sub['Role']=='annual')]['Amount_100M'].sum()
                        
                        if base_qtr == '4Q':
                            metrics[k] = annual_only
                        else:
                            metrics[k] = curr_cum + prior_ann - prior_same
                    
                    # 기본 정보 병합
                    base_info = next((item for item in screen_summary_data if item["Ticker"] == ticker), None)
                    if base_info:
                        ev = base_info['Market_Cap'] + base_info['IBD'] - base_info['Cash'] + base_info['NCI'] - base_info['NOA']
                        ltm_res.append({
                            'Company': base_info['Company'],
                            'Ticker': ticker,
                            'Market Cap': base_info['Market_Cap'],
                            'EV': ev,
                            'Revenue(LTM)': metrics['Revenue'],
                            'EBIT(LTM)': metrics['EBIT'],
                            'NI(LTM)': metrics['NI'],
                            'EV/EBIT': ev / metrics['EBIT'] if metrics['EBIT'] > 0 else np.nan,
                            'PER': base_info['Market_Cap'] / metrics['NI'] if metrics['NI'] > 0 else np.nan,
                            'PBR': np.nan, # Equity 정보는 BS 상세에서 가져와야 함 (여기선 생략)
                            'PSR': base_info['Market_Cap'] / metrics['Revenue'] if metrics['Revenue'] > 0 else np.nan
                        })
                
                df_screen = pd.DataFrame(ltm_res)
                
                st.subheader("📊 Multiples Table (Preview)")
                st.dataframe(df_screen.style.format("{:.1f}", subset=['Market Cap','EV','Revenue(LTM)','EBIT(LTM)','NI(LTM)'])
                                            .format("{:.1f}x", subset=['EV/EBIT','PER','PSR'], na_rep="N/M"))
                
                st.subheader("📈 Statistics (Mean/Median)")
                if not df_screen.empty:
                    stats = df_screen[['EV/EBIT', 'PER', 'PSR']].agg(['mean', 'median', 'max', 'min'])
                    st.dataframe(stats.style.format("{:.1f}x"))


            # 2. 엑셀 생성 (메모리)
            output = io.BytesIO()
            wb = Workbook()
            wb.remove(wb.active)
            
            # (기존 엑셀 생성 로직 그대로 활용 - 함수화 하지 않고 바로 실행)
            # Sheet 1: BS_Full
            ws_bs = wb.create_sheet('BS_Full')
            add_sheet_title_block(ws_bs,"BS_Full (Balance Sheet Detail)","Logic: finstate_all(CFS→OFS) 재무상태표 라인아이템 수집 후 EV_Component 태깅 | Unit: 억원",8)
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
            add_sheet_title_block(ws_pl,"PL_Data (Income Statement Core Only)","Logic: IS 추출 후 매출/영업이익/순이익 3개 계정만 엄격 추출 | Unit: 억원",8)
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
            add_sheet_title_block(ws_mc,"Market_Cap (Price & Shares)","Logic: 종가(FDR) × 유통주식수(DART) | Unit: 억원",13)
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
            add_sheet_title_block(ws_ltm,"LTM_Calc (Revenue/EBIT/NI)",f"Base: {base_period_str} | Unit: 억원",8)
            cols = [('Company',15), ('Ticker',10), ('Calc_Key',12),('Current_Cum(A)',15), ('Prior_Annual(B)',15), ('Prior_SameQ(C)',15), ('LTM_Value',15), ('Note',10)]
            header_row = 4
            ws_ltm.append([]); ws_ltm.append([c[0] for c in cols])
            for i, (_, w) in enumerate(cols): ws_ltm.column_dimensions[get_column_letter(i+1)].width = w; sc(ws_ltm.cell(header_row, i+1), fo=fH, fi=pH, al=aC, bd=BD)
            r = header_row + 1
            ltm_keys = ['Revenue', 'EBIT', 'NI']
            for ticker in target_code_list:
                comp_name = ticker_to_name.get(ticker, ticker)
                for k in ltm_keys:
                    ws_ltm.cell(r, 1, comp_name); sc(ws_ltm.cell(r, 1), fo=fA, fi=pW, al=aL, bd=BD)
                    ws_ltm.cell(r, 2, ticker);    sc(ws_ltm.cell(r, 2), fo=fA, fi=pW, al=aL, bd=BD)
                    ws_ltm.cell(r, 3, k);         sc(ws_ltm.cell(r, 3), fo=fA, fi=pW, al=aL, bd=BD)
                    ws_ltm.cell(r, 4).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!G:G, C{r}, PL_Data!D:D, "current_cum")'; sc(ws_ltm.cell(r,4), fo=fLINK, fi=pW, nf=NB, bd=BD)
                    ws_ltm.cell(r, 5).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!G:G, C{r}, PL_Data!D:D, "prior_annual")'; sc(ws_ltm.cell(r,5), fo=fLINK, fi=pW, nf=NB, bd=BD)
                    ws_ltm.cell(r, 6).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!G:G, C{r}, PL_Data!D:D, "prior_same_q")'; sc(ws_ltm.cell(r,6), fo=fLINK, fi=pW, nf=NB, bd=BD)
                    if base_qtr == '4Q':
                        ws_ltm.cell(r, 7).value = f'=SUMIFS(PL_Data!H:H, PL_Data!B:B, B{r}, PL_Data!G:G, C{r}, PL_Data!D:D, "annual")'; note = 'Annual'
                    else:
                        ws_ltm.cell(r, 7).value = f'=D{r}+E{r}-F{r}'; note = 'A+B-C'
                    sc(ws_ltm.cell(r,7), fo=fFRM, fi=pW, nf=NB, bd=BD); ws_ltm.cell(r, 8).value = note; sc(ws_ltm.cell(r,8), fo=fA, fi=pW, al=aC, bd=BD)
                    r += 1
            ws_ltm.auto_filter.ref = f"A{header_row}:H{r-1}"; ws_ltm.freeze_panes = f"A{header_row+1}"

            # Sheet 5: GPCM
            ws = wb.create_sheet('GPCM')
            wb.move_sheet('GPCM', offset=-4)
            ws.merge_cells('A1:Y1'); ws['A1'] = "GPCM Valuation Summary"; sc(ws['A1'], fo=fT)
            ws.merge_cells('A2:Y2'); ws['A2'] = f"Base: {base_period_str} | Unit: 억원 | EV = MCap + IBD − Cash + NCI − NOA (NOA optional)"; sc(ws['A2'], fo=fS)
            add_gpcm_section_row(ws)
            headers = ['Company','Ticker','Base Date','Curr','PL Source','Cash','IBD','NOA','Net Debt','NCI','Equity','EV','Revenue','EBIT','D&A','EBITDA','NI','Price','Shares','Mkt Cap','EV/EBITDA','EV/EBIT','PER','PBR','PSR']
            header_row = 5
            for i, h in enumerate(headers, 1):
                ws.column_dimensions[get_column_letter(i)].width = 13 if i > 5 else 10
                if i == 1: ws.column_dimensions['A'].width = 18
                sc(ws.cell(header_row, i, h), fo=fH, fi=pH, al=aC, bd=BD)
            r = header_row + 1
            for ticker in target_code_list:
                comp_name = ticker_to_name.get(ticker, ticker); bg = pST if (r % 2 == 0) else pW
                ws.cell(r,1, comp_name); ws.cell(r,2, ticker); ws.cell(r,3, base_period_str); ws.cell(r,4, 'KRW'); ws.cell(r,5, 'LTM')
                for c in range(1, 6): sc(ws.cell(r,c), fo=fA, fi=bg, al=aL, bd=BD)
                ws.cell(r,6).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!G:G, "Cash")'; sc(ws.cell(r,6), fo=fLINK, fi=ev_fills['Cash'], nf=NB, bd=BD)
                ws.cell(r,7).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!G:G, "IBD")'; sc(ws.cell(r,7), fo=fLINK, fi=ev_fills['IBD'], nf=NB, bd=BD)
                ws.cell(r,8).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!G:G, "NOA")'; sc(ws.cell(r,8), fo=fLINK, fi=ev_fills['NOA'], nf=NB, bd=BD)
                ws.cell(r,9).value = f'=G{r}-F{r}-H{r}'; sc(ws.cell(r,9), fo=fFRM, fi=bg, nf=NB, bd=BD)
                ws.cell(r,10).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!G:G, "NCI")'; sc(ws.cell(r,10), fo=fLINK, fi=ev_fills['NCI'], nf=NB, bd=BD)
                ws.cell(r,11).value = f'=SUMIFS(BS_Full!H:H, BS_Full!B:B, B{r}, BS_Full!G:G, "Equity_Total")'; sc(ws.cell(r,11), fo=fLINK, fi=ev_fills['Equity'], nf=NB, bd=BD)
                ws.cell(r,12).value = f'=T{r}+G{r}-F{r}+J{r}-H{r}'; sc(ws.cell(r,12), fo=fFRM, fi=bg, nf=NB, bd=BD)
                ws.cell(r,13).value = f'=SUMIFS(LTM_Calc!G:G, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, "Revenue")'; sc(ws.cell(r,13), fo=fLINK, fi=ev_fills['PL_HL'], nf=NB, bd=BD)
                ws.cell(r,14).value = f'=SUMIFS(LTM_Calc!G:G, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, "EBIT")'; sc(ws.cell(r,14), fo=fLINK, fi=ev_fills['PL_HL'], nf=NB, bd=BD)
                sc(ws.cell(r,15), fi=PatternFill('solid', fgColor='FFFF00'), nf=NB, bd=BD) # D&A 수기
                ws.cell(r,16).value = f'=N{r}+O{r}'; sc(ws.cell(r,16), fo=fFRM, fi=bg, nf=NB, bd=BD)
                ws.cell(r,17).value = f'=SUMIFS(LTM_Calc!G:G, LTM_Calc!B:B, B{r}, LTM_Calc!C:C, "NI")'; sc(ws.cell(r,17), fo=fLINK, fi=ev_fills['PL_HL'], nf=NB, bd=BD)
                ws.cell(r,18).value = f'=SUMIFS(Market_Cap!E:E, Market_Cap!B:B, B{r})'; sc(ws.cell(r,18), fo=fLINK, nf=NP, bd=BD)
                ws.cell(r,19).value = f'=SUMIFS(Market_Cap!F:F, Market_Cap!B:B, B{r})'; sc(ws.cell(r,19), fo=fLINK, nf=NI_FMT, bd=BD)
                ws.cell(r,20).value = f'=SUMIFS(Market_Cap!G:G, Market_Cap!B:B, B{r})'; sc(ws.cell(r,20), fo=fLINK, nf=NB1, bd=BD)
                pMULT = PatternFill('solid', fgColor=C_PB)
                ws.cell(r,21).value = f'=IF(P{r}>0, L{r}/P{r}, "N/M")'; sc(ws.cell(r,21), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
                ws.cell(r,22).value = f'=IF(N{r}>0, L{r}/N{r}, "N/M")'; sc(ws.cell(r,22), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
                ws.cell(r,23).value = f'=IF(Q{r}>0, T{r}/Q{r}, "N/M")'; sc(ws.cell(r,23), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
                ws.cell(r,24).value = f'=IF(K{r}>0, T{r}/K{r}, "N/M")'; sc(ws.cell(r,24), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
                ws.cell(r,25).value = f'=IF(M{r}>0, T{r}/M{r}, "N/M")'; sc(ws.cell(r,25), fo=fMUL, fi=pMULT, nf=NF_X, bd=BD)
                r += 1
            r_end = r - 1; r += 1
            for stat, fn in [('Mean','AVERAGE'), ('Median','MEDIAN'), ('Max','MAX'), ('Min','MIN')]:
                ws.cell(r, 20, stat); sc(ws.cell(r,20), fo=fSTAT, fi=pSTAT, al=aC, bd=BD)
                for c in range(21, 26): col = get_column_letter(c); ws.cell(r, c).value = f'=IFERROR({fn}({col}{header_row+1}:{col}{r_end}), "N/M")'; sc(ws.cell(r,c), fo=fSTAT, fi=pSTAT, nf=NF_X, bd=BD)
                r += 1
            r += 2
            for note in notes_list: ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=25); sc(ws.cell(r, 1, note), fo=fNOTE); r += 1
            ws.freeze_panes = f"A{header_row+1}"

            wb.save(output)
            output.seek(0)
            
            st.success("분석 완료! 아래 버튼을 눌러 리포트를 다운로드하세요.")
            st.download_button(
                label="📥 Report Download (Excel)",
                data=output,
                file_name=f"KR_GPCM_Fixed_{base_period_str.replace('.','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
