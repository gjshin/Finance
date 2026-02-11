import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import warnings
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.chart.axis import ChartLines

warnings.filterwarnings('ignore')

# ==========================================
# 1. 설정
# ==========================================
target_tickers = [
    '3116.T',
    '7273.T',
    '4246.T',
    '7282.T',
    'MG.TO',
    'EZM.F',
    '200880.KS',
    '038110.KQ',
    '012330.KS',
    'PYT.VI',
]

BASE_DATE = '2025-09-30'

print("📡 회사명 자동 조회 중...")
ticker_to_name = {}
for t in target_tickers:
    try:
        _info = yf.Ticker(t).info
        name = _info.get('longName') or _info.get('shortName') or t
        ticker_to_name[t] = name
        print(f"  ✅ {t} → {name}")
    except:
        ticker_to_name[t] = t
        print(f"  ⚠️ {t} → (이름 조회 실패, 티커 사용)")

# ==========================================
# 2. 계정 맵핑 (★ v17: NOA→NOA(Option), 투자부동산 추가)
# ==========================================
# ★ Cash 항목은 그대로 유지, NOA는 NOA(Option)으로만 태깅 (GPCM EV에 미반영)
BS_HIGHLIGHT_MAP = {
    # --- Cash (유지) ---
    'Cash And Cash Equivalents':           'Cash',
    'Other Short Term Investments':        'Cash',
    # --- IBD ---
    'Current Debt And Capital Lease Obligation':    'IBD',
    'Long Term Debt And Capital Lease Obligation':   'IBD',
    # --- NCI ---
    'Minority Interest':                   'NCI',
    # --- Equity ---
    'Stockholders Equity':                 'Equity',
    # --- NOA(Option): 투자성 자산 전체 (사용자 취사선택용) ---
    'Long Term Equity Investment':                        'NOA(Option)',
    'Investments In Other Ventures Under Equity Method':  'NOA(Option)',
    'Investment In Financial Assets':                     'NOA(Option)',
    'Investmentin Financial Assets':                      'NOA(Option)',
    'Investment Properties':                              'NOA(Option)',
    'Non Current Note Receivables':                       'NOA(Option)',
    'Other Investments':                                  'NOA(Option)',
}

BS_SUBTOTAL_EXCLUDE = {
    'Cash Cash Equivalents And Short Term Investments',
    'Cash And Short Term Investments',
    'Total Debt', 'Total Capitalization',
    'Total Equity Gross Minority Interest',
}

PL_HIGHLIGHT_MAP = {
    'Total Revenue':                   'Revenue',
    'Operating Income':                'EBIT',
    'EBIT':                            'EBIT',
    'EBITDA':                          'EBIT',
    'Normalized EBITDA':               'EBIT',
    'Net Income Common Stockholders':  'NI_Parent',
    'Net Income':                      'Net Income',
}

PL_SORT = {
    'Total Revenue': 10, 'Operating Revenue': 11, 'Cost Of Revenue': 20, 'Gross Profit': 30,
    'Operating Expense': 35, 'Selling General And Administration': 36, 'Research And Development': 37,
    'Operating Income': 50, 'EBIT': 55, 'EBITDA': 56, 'Normalized EBITDA': 57,
    'Interest Expense': 60, 'Pretax Income': 70, 'Tax Provision': 75,
    'Net Income': 90, 'Net Income Common Stockholders': 91, 'Basic EPS': 95, 'Diluted EPS': 96,
}

# ==========================================
# 3. KPMG Style
# ==========================================
C_BL='00338D'; C_DB='1E2A5E'; C_MB='005EB8'; C_LB='C3D7EE'; C_PB='E8EFF8'
C_DG='333333'; C_MG='666666'; C_LG='F5F5F5'; C_BG='B0B0B0'; C_W='FFFFFF'
C_GR='E2EFDA'; C_YL='FFF8E1'; C_NOA='FCE4EC'

S1=Side(style='thin',color=C_BG); BD=Border(left=S1,right=S1,top=S1,bottom=S1)
fT=Font(name='Arial',bold=True,size=14,color=C_BL)
fS=Font(name='Arial',size=9,color=C_MG,italic=True)
fH=Font(name='Arial',bold=True,size=9,color=C_W)
fA=Font(name='Arial',size=9,color=C_DG)
fHL=Font(name='Arial',bold=True,size=9,color=C_DB)
fSEC=Font(name='Arial',bold=True,size=10,color=C_W)
fMUL=Font(name='Arial',bold=True,size=10,color=C_BL)
fNOTE=Font(name='Arial',size=8,color=C_MG,italic=True)
fSTAT=Font(name='Arial',bold=True,size=9,color=C_DB)
# ★ 수식셀 = 검정폰트 (금융모델 관례: 수식은 검정)
fFRM=Font(name='Arial',size=9,color='000000')
fFRM_B=Font(name='Arial',bold=True,size=9,color='000000')
# ★ 시트간 참조 = 녹색폰트
fLINK=Font(name='Arial',size=9,color='008000')
fLINK_B=Font(name='Arial',bold=True,size=9,color='008000')

pH=PatternFill('solid',fgColor=C_BL); pW=PatternFill('solid',fgColor=C_W)
pST=PatternFill('solid',fgColor=C_LG); pSEC=PatternFill('solid',fgColor=C_DB)
pSTAT=PatternFill('solid',fgColor=C_LB)

ev_fills = {
    'Cash': PatternFill('solid',fgColor=C_GR),
    'IBD': PatternFill('solid',fgColor=C_YL),
    'NCI': PatternFill('solid',fgColor=C_PB),
    'NOA(Option)': PatternFill('solid',fgColor=C_NOA),
    'Equity': PatternFill('solid',fgColor=C_LB),
    'PL_HL': PatternFill('solid',fgColor=C_YL),
    'NI_Parent': PatternFill('solid',fgColor=C_YL),
}

aC=Alignment(horizontal='center',vertical='center',wrap_text=True)
aL=Alignment(horizontal='left',vertical='center',indent=1)
aR=Alignment(horizontal='right',vertical='center')

NF_M='#,##0;(#,##0);"-"'; NF_M1='#,##0.0;(#,##0.0);"-"'; NF_PRC='#,##0.00;(#,##0.00);"-"'
NF_INT='#,##0;(#,##0);"-"'; NF_EPS='#,##0.00;(#,##0.00);"-"'; NF_X='0.0"x";(0.0"x");"-"'

def sc(c, fo=None, fi=None, al=None, bd=None, nf=None):
    if fo: c.font = fo
    if fi: c.fill = fi
    if al: c.alignment = al
    if bd: c.border = bd
    if nf: c.number_format = nf

# ==========================================
# 4. 데이터 수집
# ==========================================
print(f"\n🚀 Global GPCM 데이터 수집 (v17)")
base_dt = pd.to_datetime(BASE_DATE)

gpcm_data = {}  # 순서 유지 (Python 3.7+ dict)
raw_bs_rows = []
raw_pl_rows = []
market_rows = []
price_abs_dfs = []
price_rel_dfs = []

for ticker in target_tickers:
    company_name = ticker_to_name.get(ticker, ticker)
    print(f"🏢 [{company_name}] ({ticker})")
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        currency = info.get('currency', 'USD')

        gpcm = {
            'Company': company_name, 'Ticker': ticker, 'Currency': currency,
            'Base_Date': BASE_DATE, 'PL_Source': '',
        }

        # 10년 주가
        try:
            hist_10y_raw = stock.history(
                start=(base_dt - timedelta(days=365*10+20)).strftime('%Y-%m-%d'),
                end=base_dt.strftime('%Y-%m-%d'), auto_adjust=False)
            hist_10y = hist_10y_raw['Close'] if 'Close' in hist_10y_raw.columns else hist_10y_raw.iloc[:,0]
            if not hist_10y.empty:
                abs_s = hist_10y.copy(); abs_s.name = ticker
                price_abs_dfs.append(abs_s)
                rel_s = (hist_10y / hist_10y.iloc[0]) * 100; rel_s.name = ticker
                price_rel_dfs.append(rel_s)
        except: pass

        # BS
        q_bs = stock.quarterly_balance_sheet
        bs_shares = None
        if q_bs is not None and not q_bs.empty:
            valid = sorted([d for d in q_bs.columns if d <= base_dt + timedelta(days=7)], reverse=True)
            if valid:
                latest = valid[0]
                for acct_name in q_bs.index:
                    val = q_bs.loc[acct_name, latest]
                    if pd.isna(val): continue
                    if str(acct_name) == 'Ordinary Shares Number':
                        bs_shares = val

                    ev_tag = BS_HIGHLIGHT_MAP.get(str(acct_name), '')
                    if str(acct_name) in BS_SUBTOTAL_EXCLUDE: ev_tag = ''

                    raw_bs_rows.append({
                        'Company': company_name, 'Ticker': ticker,
                        'Period': latest.strftime('%Y-%m-%d'),
                        'Currency': currency, 'Account': str(acct_name),
                        'EV_Tag': ev_tag, 'Amount_M': val/1e6
                    })

        shares = bs_shares if bs_shares else info.get('sharesOutstanding', 0)

        # Market Cap
        try:
            hist = stock.history(
                start=(base_dt - timedelta(days=10)).strftime('%Y-%m-%d'),
                end=(base_dt + timedelta(days=1)).strftime('%Y-%m-%d'), auto_adjust=False)
            close = hist['Close'].iloc[-1] if (not hist.empty and 'Close' in hist.columns) else 0
            p_date = hist.index[-1].strftime('%Y-%m-%d') if not hist.empty else '-'
        except:
            close = 0; p_date = '-'

        mcap_m = (close * shares / 1e6) if shares else 0
        market_rows.append({
            'Company': company_name, 'Ticker': ticker,
            'Base_Date': BASE_DATE, 'Price_Date': p_date,
            'Currency': currency, 'Close': close,
            'Shares': shares, 'Market_Cap_M': round(mcap_m, 1)
        })

        # PL
        q_is = stock.quarterly_income_stmt
        q_valid = []
        if q_is is not None and not q_is.empty:
            q_valid = sorted([d for d in q_is.columns if d <= base_dt + timedelta(days=7)], reverse=True)[:4]

        is_complete = False
        if len(q_valid) == 4:
            if 'Total Revenue' in q_is.index:
                vals = q_is.loc['Total Revenue', q_valid]
                if vals.notna().all() and (vals != 0).all():
                    is_complete = True

        pl_source = 'Quarterly (4Q Sum)' if is_complete else 'Annual'
        pl_data = None; pl_dates = []
        if is_complete:
            pl_data = q_is; pl_dates = q_valid
        else:
            a_is = stock.income_stmt
            if a_is is not None and not a_is.empty:
                valid_a = sorted([d for d in a_is.columns if d <= base_dt + timedelta(days=7)], reverse=True)
                if valid_a:
                    pl_dates = [valid_a[0]]; pl_data = a_is

        gpcm['PL_Source'] = pl_source

        if pl_data is not None:
            for acct in pl_data.index:
                acct_str = str(acct)
                hl_tag = PL_HIGHLIGHT_MAP.get(acct_str, '')
                is_eps = 'EPS' in acct_str or 'Per Share' in acct_str
                is_shares = 'Average Shares' in acct_str

                for i, d in enumerate(pl_dates):
                    val = pl_data.loc[acct, d]
                    if pd.isna(val): continue
                    if is_eps: unit = 'per share'; amt = val
                    elif is_shares: unit = 'M shares'; amt = val/1e6
                    else: unit = 'M'; amt = val/1e6

                    raw_pl_rows.append({
                        'Company': company_name, 'Ticker': ticker,
                        'Currency': currency, 'Account': acct_str,
                        'GPCM_Tag': hl_tag, 'Unit': unit,
                        'PL_Source': pl_source,
                        'Q_Label': f"D-{i}Q" if pl_source.startswith('Quarterly') else 'Annual',
                        'Period': d.strftime('%Y-%m-%d'),
                        'Amount_M': amt, '_sort': PL_SORT.get(acct_str, 500)
                    })

        gpcm_data[ticker] = gpcm
    except Exception as e:
        print(f"  ⚠️ Error: {e}")
        continue


# ==========================================
# 5. 엑셀 생성 (★ v17: 수식 기반 GPCM)
# ==========================================
print(f"\n📄 엑셀 생성 중...")
wb = Workbook(); wb.remove(wb.active)
output_file = f"Global_GPCM_v17_{BASE_DATE.replace('-','')}.xlsx"

# ==========================================================================
# Sheet 1: BS_Full  (★ GPCM보다 먼저 생성 → 수식 참조 대상)
# ==========================================================================
ws_bs = wb.create_sheet('BS_Full')
if raw_bs_rows:
    df_bs = pd.DataFrame(raw_bs_rows)
    bs_cols = [
        ('Company','Company',18), ('Ticker','Ticker',10), ('Period','Period',12),
        ('Currency','Curr',6), ('Account','Account',42),
        ('EV_Tag','EV Tag',14), ('Amount_M','Amount (M)',18)
    ]
    ws_bs.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(bs_cols))
    sc(ws_bs.cell(1,1,'Balance Sheet Full'), fo=fT)
    # 범례 (row 3)
    r = 3
    for i, (lb, key) in enumerate([('Cash','Cash'),('IBD','IBD'),('NCI','NCI'),
                                    ('NOA(Option)','NOA(Option)'),('Equity','Equity')]):
        sc(ws_bs.cell(r, i+1, lb),
           fo=Font(name='Arial',size=8,bold=True), fi=ev_fills[key], al=aC, bd=BD)
    # 헤더 (row 5)
    r = 5
    for i, (col, disp, w) in enumerate(bs_cols):
        ws_bs.column_dimensions[get_column_letter(i+1)].width = w
        sc(ws_bs.cell(r, i+1, disp), fo=fH, fi=pH, al=aC, bd=BD)
    hdr = r; r += 1
    # 데이터
    for _, rd in df_bs.iterrows():
        ev_tag = rd['EV_Tag']
        is_hl = bool(ev_tag)
        row_fi = ev_fills.get(ev_tag, pST if r%2==0 else pW) if is_hl else (pST if r%2==0 else pW)
        row_font = fHL if is_hl else fA
        for i, (col, _, _) in enumerate(bs_cols):
            c = ws_bs.cell(r, i+1)
            v = rd[col]
            if isinstance(v, (float, np.floating)):
                c.value = round(v, 1) if pd.notna(v) else None
            else:
                c.value = v
            sc(c, fo=row_font, fi=row_fi,
               al=aR if col=='Amount_M' else aL, bd=BD,
               nf=NF_M if col=='Amount_M' else None)
        r += 1
    ws_bs.auto_filter.ref = f"A{hdr}:{get_column_letter(len(bs_cols))}{r-1}"
    ws_bs.freeze_panes = f'A{hdr+1}'

# ==========================================================================
# Sheet 2: PL_Data  (★ GPCM보다 먼저 생성)
# ==========================================================================
ws_pl = wb.create_sheet('PL_Data')
if raw_pl_rows:
    df_pl = pd.DataFrame(raw_pl_rows).sort_values(['Company','_sort','Q_Label'])
    pl_cols = [
        ('Company','Company',18), ('Ticker','Ticker',10), ('Currency','Curr',6),
        ('Account','Account',42), ('GPCM_Tag','GPCM Tag',14), ('Unit','Unit',10),
        ('PL_Source','Source',14), ('Q_Label','Q Label',9),
        ('Period','Period',12), ('Amount_M','Amount',18)
    ]
    ws_pl.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(pl_cols))
    sc(ws_pl.cell(1,1,'Income Statement'), fo=fT)
    r = 5
    for i, (col, disp, w) in enumerate(pl_cols):
        ws_pl.column_dimensions[get_column_letter(i+1)].width = w
        sc(ws_pl.cell(r, i+1, disp), fo=fH, fi=pH, al=aC, bd=BD)
    hdr = r; r += 1
    for _, rd in df_pl.iterrows():
        is_hl = bool(rd['GPCM_Tag'])
        row_fi = ev_fills['PL_HL'] if is_hl else (pST if r%2==0 else pW)
        row_font = fHL if is_hl else fA
        for i, (col, _, _) in enumerate(pl_cols):
            if col == '_sort': continue
            c = ws_pl.cell(r, i+1)
            v = rd[col]
            if isinstance(v, (float, np.floating)):
                c.value = round(v, 1) if pd.notna(v) else None
            else:
                c.value = v
            is_eps = rd.get('Unit','') == 'per share'
            nf = NF_EPS if is_eps else NF_M
            sc(c, fo=row_font, fi=row_fi,
               al=aR if col=='Amount_M' else aL, bd=BD,
               nf=nf if col=='Amount_M' else None)
        r += 1
    ws_pl.auto_filter.ref = f"A{hdr}:{get_column_letter(len(pl_cols))}{r-1}"
    ws_pl.freeze_panes = f'A{hdr+1}'

# ==========================================================================
# Sheet 3: Market_Cap  (★ GPCM보다 먼저 생성)
# ==========================================================================
ws_mc = wb.create_sheet('Market_Cap')
if market_rows:
    df_mkt = pd.DataFrame(market_rows)
    mc_cols = [
        ('Company','Company',18), ('Ticker','Ticker',10),
        ('Base_Date','Base Date',12), ('Price_Date','Price Date',12),
        ('Currency','Curr',6), ('Close','Close Price',14),
        ('Shares','Shares (Ord.)',18), ('Market_Cap_M','Mkt Cap (M)',20)
    ]
    ws_mc.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(mc_cols))
    sc(ws_mc.cell(1,1,'Market Capitalization'), fo=fT)
    ws_mc.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(mc_cols))
    sc(ws_mc.cell(2,1,'Mkt Cap = Ordinary Shares Number (자기주식 차감) × Close Price (auto_adjust=False)'), fo=fS)
    r = 4
    for i, (col, disp, w) in enumerate(mc_cols):
        ws_mc.column_dimensions[get_column_letter(i+1)].width = w
        sc(ws_mc.cell(r, i+1, disp), fo=fH, fi=pH, al=aC, bd=BD)
    mc_hdr = r; r += 1
    # ★ Market_Cap 데이터 시작행 = 5
    MC_DATA_START = r
    for _, rd in df_mkt.iterrows():
        ev = (r % 2 == 0)
        for i, (col, _, _) in enumerate(mc_cols):
            c = ws_mc.cell(r, i+1)
            v = rd[col]
            if isinstance(v, (float, np.floating)):
                c.value = round(v, 2) if pd.notna(v) else None
            else:
                c.value = v
            nf = NF_PRC if col=='Close' else (NF_INT if col=='Shares' else (NF_M1 if col=='Market_Cap_M' else None))
            sc(c, fo=fA, fi=pST if ev else pW, al=aR if nf else aL, bd=BD, nf=nf)
        r += 1
    ws_mc.auto_filter.ref = f"A{mc_hdr}:{get_column_letter(len(mc_cols))}{r-1}"
    ws_mc.freeze_panes = f'A{mc_hdr+1}'


# ==========================================================================
# Sheet 4: GPCM  (★ v17: 전체 엑셀 수식 기반)
# ==========================================================================
# ★ GPCM 컬럼 레이아웃 (23열):
#  A(1):  Company         K(11): EV              U(21): PER
#  B(2):  Ticker          L(12): Revenue         V(22): PBR
#  C(3):  Base Date       M(13): EBIT            W(23): PSR
#  D(4):  Curr            N(14): EBITDA
#  E(5):  PL Source       O(15): NI (Parent)
#  F(6):  Cash            P(16): Price
#  G(7):  IBD             Q(17): Shares
#  H(8):  Net Debt        R(18): Mkt Cap
#  I(9):  NCI             S(19): EV/EBITDA
#  J(10): Equity          T(20): EV/EBIT
#
# Formulas:
#  F-G, I-J: SUMIFS from BS_Full
#  H: =G-F (Net Debt)
#  K: =R+G-F+I (EV = MCap + IBD - Cash + NCI)
#  L, O: SUMIFS from PL_Data (by GPCM_Tag)
#  M: SUMIFS from PL_Data (by Account = "Operating Income")
#  N: =M + SUMIFS("EBITDA")- SUMIFS("EBIT")
#  P-R: = references to Market_Cap sheet
#  S-W: IF formulas for multiples
# ==========================================================================

ws = wb.create_sheet('GPCM')
# 시트를 맨 앞으로 이동
wb.move_sheet('GPCM', offset=-3)

TOTAL_COLS = 23
# Title
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS)
sc(ws.cell(1, 1, 'GPCM Valuation Summary'), fo=fT)
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=TOTAL_COLS)
sc(ws.cell(2, 1, f'Base: {BASE_DATE} | Unit: Millions (local currency) | EV = MCap + IBD − Cash + NCI'), fo=fS)

# Row 4: Section headers
r = 4
sections = [
    (1,  3, 'Company Info'),
    (4,  2, 'Other Information'),
    (6,  6, 'BS → EV Components'),       # Cash,IBD,NetDebt,NCI,Equity,EV
    (12, 4, 'PL (LTM / Annual)'),         # Revenue,EBIT,EBITDA,NI(Parent)
    (16, 3, 'Market Data'),               # Price,Shares,MktCap
    (19, 5, 'Valuation Multiples'),       # EV/EBITDA,EV/EBIT,PER,PBR,PSR
]
for start, span, txt in sections:
    ws.merge_cells(start_row=r, start_column=start, end_row=r, end_column=start+span-1)
    sc(ws.cell(r, start, txt), fo=fSEC, fi=pSEC, al=aC, bd=BD)
    for c_idx in range(start, start+span):
        sc(ws.cell(r, c_idx), bd=BD)

# Row 5: Column headers
r = 5
headers = [
    'Company', 'Ticker', 'Base Date', 'Curr', 'PL Source',
    'Cash', 'IBD', 'Net Debt', 'NCI', 'Equity',
    'EV',
    'Revenue', 'EBIT', 'EBITDA', 'NI (Parent)',
    'Price', 'Shares', 'Mkt Cap',
    'EV/EBITDA', 'EV/EBIT', 'PER', 'PBR', 'PSR'
]
widths = [
    18, 10, 11, 6, 16,
    14, 14, 14, 12, 14,
    16,
    14, 14, 14, 14,
    12, 16, 16,
    12, 12, 10, 10, 10
]
for i, (h, w) in enumerate(zip(headers, widths)):
    ws.column_dimensions[get_column_letter(i+1)].width = w
    sc(ws.cell(r, i+1, h), fo=fH, fi=pH, al=aC, bd=BD)

# Row 6+: Company data (★ 엑셀 수식)
DATA_START = 6
n_companies = len(gpcm_data)
DATA_END = DATA_START + n_companies - 1

for idx, (ticker, gpcm) in enumerate(gpcm_data.items()):
    r = DATA_START + idx
    mc_row = MC_DATA_START + idx  # Market_Cap 시트의 대응 행
    ev_row = (r % 2 == 0)
    base_fi = pST if ev_row else pW

    # ─── A-E: Company Info & Other (하드코딩, 블루텍스트) ───
    ws.cell(r, 1, gpcm['Company'])     # A: Company
    ws.cell(r, 2, ticker)               # B: Ticker
    ws.cell(r, 3, gpcm['Base_Date'])    # C: Base Date
    ws.cell(r, 4, gpcm['Currency'])     # D: Currency
    ws.cell(r, 5, gpcm['PL_Source'])    # E: PL Source
    for ci in range(1, 6):
        sc(ws.cell(r, ci), fo=fA, fi=base_fi, al=aL, bd=BD)

    # ─── F-J: BS → EV Components (SUMIFS 수식, 녹색텍스트) ───
    # F(6): Cash
    ws.cell(r, 6).value = f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"Cash")'
    sc(ws.cell(r, 6), fo=fLINK_B, fi=ev_fills['Cash'], al=aR, bd=BD, nf=NF_M)

    # G(7): IBD
    ws.cell(r, 7).value = f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"IBD")'
    sc(ws.cell(r, 7), fo=fLINK_B, fi=ev_fills['IBD'], al=aR, bd=BD, nf=NF_M)

    # H(8): Net Debt = IBD - Cash
    ws.cell(r, 8).value = f'=G{r}-F{r}'
    sc(ws.cell(r, 8), fo=fFRM_B, fi=base_fi, al=aR, bd=BD, nf=NF_M)

    # I(9): NCI
    ws.cell(r, 9).value = f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"NCI")'
    sc(ws.cell(r, 9), fo=fLINK_B, fi=ev_fills['NCI'], al=aR, bd=BD, nf=NF_M)

    # J(10): Equity
    ws.cell(r, 10).value = f'=SUMIFS(BS_Full!$G:$G,BS_Full!$B:$B,$B{r},BS_Full!$F:$F,"Equity")'
    sc(ws.cell(r, 10), fo=fLINK_B, fi=ev_fills['Equity'], al=aR, bd=BD, nf=NF_M)

    # K(11): EV = Mkt Cap + IBD - Cash + NCI
    ws.cell(r, 11).value = f'=R{r}+G{r}-F{r}+I{r}'
    sc(ws.cell(r, 11), fo=fFRM_B, fi=PatternFill('solid',fgColor=C_PB), al=aR, bd=BD, nf=NF_M)

    # ─── L-O: PL (SUMIFS 수식, 녹색텍스트) ───
    # L(12): Revenue (by GPCM_Tag)
    ws.cell(r, 12).value = f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$E:$E,"Revenue")'
    sc(ws.cell(r, 12), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)

    # M(13): EBIT (by Account = "Operating Income" 만)
    ws.cell(r, 13).value = f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"Operating Income")'
    sc(ws.cell(r, 13), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)

    # N(14): EBITDA = Operating Income + (yf_EBITDA + yf_NormEBITDA - yf_EBIT)
    ebitda_formula = (
        f'=M{r}'
        f'+SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"EBITDA")'
        f'-SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$D:$D,"EBIT")'
    )
    ws.cell(r, 14).value = ebitda_formula
    sc(ws.cell(r, 14), fo=fFRM_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)

    # O(15): NI Parent (by GPCM_Tag)
    ws.cell(r, 15).value = f'=SUMIFS(PL_Data!$J:$J,PL_Data!$B:$B,$B{r},PL_Data!$E:$E,"NI_Parent")'
    sc(ws.cell(r, 15), fo=fLINK_B, fi=ev_fills['PL_HL'], al=aR, bd=BD, nf=NF_M)

    # ─── P-R: Market Data (시트간 참조, 녹색텍스트) ───
    # P(16): Price
    ws.cell(r, 16).value = f'=Market_Cap!F{mc_row}'
    sc(ws.cell(r, 16), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_PRC)

    # Q(17): Shares
    ws.cell(r, 17).value = f'=Market_Cap!G{mc_row}'
    sc(ws.cell(r, 17), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_INT)

    # R(18): Mkt Cap
    ws.cell(r, 18).value = f'=Market_Cap!H{mc_row}'
    sc(ws.cell(r, 18), fo=fLINK, fi=base_fi, al=aR, bd=BD, nf=NF_M1)

    # ─── S-W: Valuation Multiples (IF 수식, 검정볼드) ───
    pMULT = PatternFill('solid', fgColor=C_PB)

    # S(19): EV/EBITDA
    ws.cell(r, 19).value = f'=IF(N{r}>0,K{r}/N{r},"N/M")'
    sc(ws.cell(r, 19), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

    # T(20): EV/EBIT
    ws.cell(r, 20).value = f'=IF(M{r}>0,K{r}/M{r},"N/M")'
    sc(ws.cell(r, 20), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

    # U(21): PER = Mkt Cap / NI Parent
    ws.cell(r, 21).value = f'=IF(O{r}>0,R{r}/O{r},"N/M")'
    sc(ws.cell(r, 21), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

    # V(22): PBR = Mkt Cap / Equity
    ws.cell(r, 22).value = f'=IF(J{r}>0,R{r}/J{r},"N/M")'
    sc(ws.cell(r, 22), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

    # W(23): PSR = Mkt Cap / Revenue
    ws.cell(r, 23).value = f'=IF(L{r}>0,R{r}/L{r},"N/M")'
    sc(ws.cell(r, 23), fo=fMUL, fi=pMULT, al=aR, bd=BD, nf=NF_X)

# ─── 통계행 (Mean/Median/Max/Min) ───
r = DATA_END + 2  # 한줄 빈행
stat_labels = ['Mean', 'Median', 'Max', 'Min']
excel_stat_funcs = {
    'Mean': 'AVERAGE', 'Median': 'MEDIAN', 'Max': 'MAX', 'Min': 'MIN'
}
# 멀티플 열: S(19), T(20), U(21), V(22), W(23)
mult_cols = [19, 20, 21, 22, 23]

for stat_name in stat_labels:
    fn = excel_stat_funcs[stat_name]
    # 라벨 in R(18)
    sc(ws.cell(r, 18, stat_name), fo=fSTAT, fi=pSTAT, al=aC, bd=BD)

    for ci in mult_cols:
        col_letter = get_column_letter(ci)
        # AVERAGE/MEDIAN/MAX/MIN은 텍스트("N/M") 무시, IFERROR로 전체 N/M 대비
        formula = f'=IFERROR({fn}({col_letter}{DATA_START}:{col_letter}{DATA_END}),"N/M")'
        c = ws.cell(r, ci)
        c.value = formula
        sc(c, fo=fSTAT, fi=pSTAT, al=aR, bd=BD, nf=NF_X)
    r += 1

# ─── Methodology Notes ───
r += 2
notes = [
    '[ Valuation Methodology Notes ]',
    f'• Base Date: {BASE_DATE}  |  Unit: Millions (local currency)',
    '• EV (Enterprise Value) = Market Cap + Interest-Bearing Debt − Cash + Non-Controlling Interest',
    '• Cash includes: Cash & Cash Equivalents + Other Short-Term Investments (단기금융자산/단기금융상품)',
    '• NOA(Option) in BS_Full: Long-Term Equity Investment, Investment In Financial Assets, Investment Properties 등',
    '    → 사용자가 BS_Full에서 EV Tag를 변경하면 GPCM에 자동 반영됩니다',
    '• Net Debt = IBD − Cash',
    '• EBIT = Operating Income only (yfinance의 EBIT는 비영업손익 포함이므로 사용하지 않음)',
    '• EBITDA = Operating Income + D&A  (D&A = yfinance EBITDA − yfinance EBIT)',
    '• PER = Market Cap ÷ Net Income Common Stockholders (지배기업당기순이익)',
    '• PBR = Market Cap ÷ Stockholders Equity (지배기업순자산, NCI 제외)',
    '• PSR = Market Cap ÷ Revenue  |  EV/EBITDA = EV ÷ EBITDA  |  EV/EBIT = EV ÷ EBIT',
    '• Market Cap = Ordinary Shares Number (유통주식수, 자기주식 차감) × Close Price (단순종가, auto_adjust=False)',
    '• PL Source: 기준일 직전 4개 분기 합산(LTM) 우선, 불가시 직전 연간 데이터 사용',
    '• N/M = Not Meaningful (음수 또는 0인 경우 멀티플 미산출)',
    '• GPCM 시트의 BS/PL 값은 SUMIFS 수식으로 BS_Full, PL_Data 시트를 참조합니다.',
    '    BS_Full/PL_Data의 Tag를 변경하면 GPCM 계산이 자동으로 갱신됩니다.',
    '',
    '⚠ 야후파이낸스의 데이터를 취합하여 계산한 결과이며, Dart/KRX/CIQ/Bloomberg 등 공신력 있는 외부기관 Data가 아님을 주의하세요',
]
for note in notes:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOTAL_COLS)
    sc(ws.cell(r, 1, note), fo=fNOTE)
    r += 1

ws.freeze_panes = 'A6'


# ==========================================================================
# Sheet 5: Price_History
# ==========================================================================
if price_abs_dfs and price_rel_dfs:
    ws_ph = wb.create_sheet(title='Price_History')
    df_abs = pd.concat(price_abs_dfs, axis=1).sort_index().ffill()
    df_rel = pd.concat(price_rel_dfs, axis=1).sort_index().ffill()
    common_index = df_abs.index

    sc(ws_ph.cell(1, 1, 'Stock Price History (10 Years)'), fo=fT)
    ws_ph.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)

    r = 3
    ws_ph.cell(r, 1, 'Date'); sc(ws_ph.cell(r, 1), fo=fH, fi=pH, al=aC, bd=BD)
    ws_ph.column_dimensions[get_column_letter(1)].width = 12
    c_idx = 2
    for col in df_abs.columns:
        disp = ticker_to_name.get(col, col)
        sc(ws_ph.cell(r, c_idx, f"{disp} (Abs)"), fo=fH, fi=PatternFill('solid',fgColor='607D8B'), al=aC, bd=BD)
        ws_ph.column_dimensions[get_column_letter(c_idx)].width = 16; c_idx += 1
    sc(ws_ph.cell(r, c_idx, ""), fi=pW); ws_ph.column_dimensions[get_column_letter(c_idx)].width = 2; c_idx += 1
    ws_ph.cell(r, c_idx, 'Date'); sc(ws_ph.cell(r, c_idx), fo=fH, fi=pH, al=aC, bd=BD)
    rel_date_col = c_idx; c_idx += 1
    rel_start_col = c_idx
    for col in df_rel.columns:
        disp = ticker_to_name.get(col, col)
        sc(ws_ph.cell(r, c_idx, f"{disp} (Rel)"), fo=fH, fi=pH, al=aC, bd=BD)
        ws_ph.column_dimensions[get_column_letter(c_idx)].width = 16; c_idx += 1

    r = 4
    for date in common_index:
        dv = date.date()
        ws_ph.cell(r, 1, dv).number_format = 'yyyy-mm-dd'; sc(ws_ph.cell(r, 1), fo=fA, al=aC, bd=BD)
        cc = 2
        for v in df_abs.loc[date]:
            ws_ph.cell(r, cc, v).number_format = '#,##0.00'; sc(ws_ph.cell(r, cc), fo=fA, al=aR, bd=BD); cc += 1
        cc += 1
        ws_ph.cell(r, cc, dv).number_format = 'yyyy-mm-dd'; sc(ws_ph.cell(r, cc), fo=fA, al=aC, bd=BD); cc += 1
        for v in df_rel.loc[date]:
            ws_ph.cell(r, cc, v).number_format = '#,##0'; sc(ws_ph.cell(r, cc), fo=fA, al=aR, bd=BD); cc += 1
        r += 1

    # 월별 샘플 차트
    chart_data_start_row = r + 2
    df_rel_monthly = df_rel.resample('ME').last().dropna(how='all')
    cr = chart_data_start_row
    sc(ws_ph.cell(cr, 1, '[ Chart Data - Monthly Sampled ]'), fo=fNOTE); cr += 1
    chart_hdr_row = cr
    sc(ws_ph.cell(cr, 1, 'Year-Month'), fo=Font(name='Arial',size=8,bold=True,color=C_MG), al=aC)
    for i, cn in enumerate(df_rel_monthly.columns):
        sc(ws_ph.cell(cr, i+2, ticker_to_name.get(cn, cn)),
           fo=Font(name='Arial',size=8,bold=True,color=C_MG), al=aC)
    cr += 1
    chart_data_first_row = cr
    for date in df_rel_monthly.index:
        ws_ph.cell(cr, 1, date.strftime('%Y-%m'))
        sc(ws_ph.cell(cr, 1), fo=Font(name='Arial',size=7,color=C_MG))
        for i, cn in enumerate(df_rel_monthly.columns):
            val = df_rel_monthly.loc[date, cn]
            if pd.notna(val): ws_ph.cell(cr, i+2, round(val, 1))
            sc(ws_ph.cell(cr, i+2), fo=Font(name='Arial',size=7,color=C_MG))
        cr += 1
    chart_data_last_row = cr - 1
    n_series = len(df_rel_monthly.columns)

    chart = LineChart()
    chart.title = "10-Year Relative Performance (Base=100)"
    chart.style = 13; chart.height = 18; chart.width = 34
    chart.y_axis.title = "Relative Index"; chart.y_axis.scaling.min = 0
    chart.y_axis.tickLblPos = "low"; chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.majorUnit = 50; chart.y_axis.numFmt = '#,##0'
    chart.x_axis.title = "Date"; chart.x_axis.tickLblPos = "low"
    n_months = chart_data_last_row - chart_data_first_row + 1
    skip = max(1, n_months // 10)
    chart.x_axis.tickLblSkip = skip; chart.x_axis.tickMarkSkip = skip
    cats = Reference(ws_ph, min_col=1, min_row=chart_data_first_row, max_row=chart_data_last_row)
    chart.set_categories(cats)
    colors = ['1F77B4','FF7F0E','2CA02C','D62728','9467BD','8C564B','E377C2','7F7F7F','BCBD22','17BECF']
    for i in range(n_series):
        dr = Reference(ws_ph, min_col=i+2, min_row=chart_hdr_row, max_row=chart_data_last_row)
        s = Series(dr, title_from_data=True)
        s.graphicalProperties.line.solidFill = colors[i % len(colors)]
        s.graphicalProperties.line.width = 20000; s.smooth = False
        chart.series.append(s)
    chart_col = rel_start_col + n_series + 1
    ws_ph.add_chart(chart, f"{get_column_letter(chart_col)}3")


# ==========================================================================
# 저장
# ==========================================================================
wb.save(output_file)
print(f"\n🎉 완료! 파일명: {output_file}")
print(f"\n📌 v17 주요 변경사항:")
print(f"   1. NOA → NOA(Option) 태깅 (GPCM EV 계산에 미반영, 사용자 선택용)")
print(f"   2. GPCM 시트 전체 엑셀 수식화 (SUMIFS/IF/AVERAGE/MEDIAN 등)")
print(f"      - BS_Full/PL_Data의 Tag 변경 시 GPCM 자동 갱신")
print(f"      - 녹색: 시트간 참조, 검정: 계산수식 (금융모델 관례)")
print(f"   3. Net Debt 열 추가 (= IBD − Cash)")
print(f"   4. EV를 BS→EV Components 섹션으로 이동")
print(f"   5. 투자부동산(Investment Properties) 등 NOA(Option) 확장")
print(f"   6. 야후파이낸스 데이터 주의 문구 추가")
