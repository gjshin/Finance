"""gpcm_kr.py 회귀 테스트.

다기간 재무데이터 모드가 화면에서 고른 기간의 보고서를 실제로 조회하는지 확인한다.
과거에 분기→보고서코드 매핑이 뒤집혀 있어 '연간'을 고르면 1분기보고서를 가져왔다
(매출이 1/4로 나오는데 오류는 나지 않아 알아채기 어려웠다).
"""
import re, sys, types, pandas as pd, numpy as np
from pathlib import Path
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
import gpcm_kr as M

RCODE_NAME = {'11011':'사업보고서(연간)', '11012':'반기보고서', '11013':'1분기보고서', '11014':'3분기보고서'}
asked = []

def mkfs():
    return pd.DataFrame([
        {'sj_nm':'연결재무상태표','account_nm':'자산총계','account_id':'ifrs-full_Assets','thstrm_amount':'1000','thstrm_add_amount':''},
        {'sj_nm':'연결재무상태표','account_nm':'부채총계','account_id':'ifrs-full_Liabilities','thstrm_amount':'400','thstrm_add_amount':''},
        {'sj_nm':'연결재무상태표','account_nm':'자본총계','account_id':'ifrs-full_Equity','thstrm_amount':'600','thstrm_add_amount':''},
        {'sj_nm':'연결포괄손익계산서','account_nm':'매출액','account_id':'ifrs-full_Revenue','thstrm_amount':'800','thstrm_add_amount':'800'},
    ])

class D:
    corp_codes = pd.DataFrame({'corp_code':['00126380'],'corp_name':['테스트'],'stock_code':['005930']})
    def find_corp_code(self, c): return '00126380'
    def company(self, c): return {'corp_name':'테스트'}
    def finstate_all(self, corp, y, reprt_code='11011', fs_div='CFS'):
        if fs_div == 'CFS': asked.append(str(reprt_code))
        return mkfs() if fs_div == 'CFS' else pd.DataFrame()

M.get_krx_listing = lambda: pd.DataFrame({'Code':['005930'],'Name':['테스트'],'Stocks':[100]})
class C:
    def write(self,*a,**k): pass
    def update(self,*a,**k): pass
class B:
    def progress(self,*a,**k): pass

def first_requested(periods):
    asked.clear()
    M.fetch_historical_financials('K', ['005930'], periods, D(), C(), B(), None)
    return asked[0] if asked else None

cases = [
    ("연간 (사업보고서)", [{'year':2024,'qtr':None,'label':'2024년'}], '11011'),
    ("1분기",            [{'year':2024,'qtr':'1Q','label':'2024.1Q'}], '11013'),
    ("반기(2Q)",         [{'year':2024,'qtr':'2Q','label':'2024.2Q'}], '11012'),
    ("3분기",            [{'year':2024,'qtr':'3Q','label':'2024.3Q'}], '11014'),
    ("4Q(연간)",         [{'year':2024,'qtr':'4Q','label':'2024.4Q'}], '11011'),
]

fails = 0
print(f"{'화면 선택':<20}{'실제 조회한 보고서':<22}{'조회했어야 할 것':<22}")
print("-" * 66)
for label, periods, want in cases:
    got = first_requested(periods)
    ok = got == want
    if not ok: fails += 1
    mark = "OK" if ok else "잘못됨"
    print(f"{label:<20}{RCODE_NAME.get(got,got):<22}{RCODE_NAME[want]:<22}{mark}")

# --- 손익 계정 매칭 (2단계: 완전일치 → 관대) -----------------------------------
print()
print("손익 계정 매칭")
print("-" * 66)

def chk(label, ok, detail=''):
    global fails
    if not ok: fails += 1
    print(f"{label:<52}{'OK' if ok else '잘못됨 — ' + str(detail)}")

# 1단계는 종전 그대로 — 완전 일치만 인정한다
chk('1단계: 매출액 → Revenue', M.match_pl_core_only('매출액') == 'Revenue')
chk('1단계: 번호 붙은 표기는 여전히 못 잡는다',
    M.match_pl_core_only('Ⅰ. 영업수익') is None, M.match_pl_core_only('Ⅰ. 영업수익'))

# 2단계가 표기 흔들림과 표준태그를 줍는다
for name in ('Ⅰ. 영업수익', '1. 매출액', '매출액(주7)', '영업수익', '(1) 수익(매출액)'):
    chk(f'2단계: {name!r} → Revenue', M.match_pl_lenient(name) == 'Revenue',
        M.match_pl_lenient(name))
chk('2단계: 이름이 낯설어도 표준태그로 잡는다',
    M.match_pl_lenient('영업수익(게임)', 'ifrs-full_Revenue') == 'Revenue')
chk('2단계: 영업이익 태그', M.match_pl_lenient('Ⅱ. 영업손익', 'dart_OperatingIncomeLoss') == 'EBIT')
chk('2단계: 세전이익 태그',
    M.match_pl_lenient('법인세차감전순손익', 'ifrs-full_ProfitLossBeforeTax') == 'Pretax_Income')
chk('2단계: 괄호가 의미를 갖는 표기는 그대로', M.match_pl_lenient('당기순이익(손실)') == 'NI')

# 오검출 방지 — 넓혔다고 엉뚱한 걸 잡으면 조용히 틀린다
for bad in ('지배기업소유주지분순이익', '총포괄손익', '상품매출원가', '매출원가',
            '영업외수익', '기타수익', '매출채권'):
    chk(f'오검출 방지: {bad!r} 은 안 잡힌다', M.match_pl_lenient(bad) is None,
        M.match_pl_lenient(bad))
chk('오검출 방지: 품목별 수익 태그는 안 쓴다',
    M.match_pl_lenient('상품매출', 'ifrs-full_RevenueFromSaleOfGoods') is None,
    M.match_pl_lenient('상품매출', 'ifrs-full_RevenueFromSaleOfGoods'))

# --- 엑셀 세율 수식 (사업연도별) -----------------------------------------------
print()
print("엑셀 Tax Rate 수식")
print("-" * 66)
# 결손(0 이하)은 파이썬과 같이 2구간 세율 — 종전 수식은 음수를 첫 구간으로 읽었다
F_2025 = ('=IF(AE6<=0, 0.209, IF(AE6<=2, 0.099, IF(AE6<=200, 0.209, '
          'IF(AE6<=3000, 0.231, 0.264))))')
chk('FY2025 수식 (결손 → 2구간)', M.korean_tax_rate_formula('AE6', 2025) == F_2025,
    M.korean_tax_rate_formula('AE6', 2025))
chk('FY2024 도 동일', M.korean_tax_rate_formula('AE6', 2024) == F_2025)
chk('FY2026 은 개정 세율(11%/22%/24.2%/27.5%)',
    M.korean_tax_rate_formula('AE6', 2026) ==
    '=IF(AE6<=0, 0.22, IF(AE6<=2, 0.11, IF(AE6<=200, 0.22, IF(AE6<=3000, 0.242, 0.275))))',
    M.korean_tax_rate_formula('AE6', 2026))
for fy in (2024, 2025, 2026):
    br = M.get_korean_tax_brackets(fy)
    f = M.korean_tax_rate_formula('X1', fy)
    chk(f'FY{fy} 수식 세율이 파이썬 표와 일치',
        all(f'{r:g}' in f for _, r in br), f)


def _eval_tax_formula(formula, x):
    """중첩 IF(셀<=상한, 세율, …) 수식을 값 x 로 풀어 본다 (엑셀 비교와 같은 순서)."""
    expr = formula.lstrip('=').replace('X1', repr(float(x)))
    while expr.startswith('IF('):
        m = re.match(r'IF\(([-\d.e+]+)<=([-\d.e+]+), ([-\d.e+]+), (.*)\)$', expr)
        a, b, c, rest = m.groups()
        if float(a) <= float(b):
            return float(c)
        expr = rest
    return float(expr)


# 같은 파일 안에서 엑셀 셀과 파이썬 요약이 갈리면 안 된다 (결손 피어에서 한 번 갈렸다)
for fy in (2024, 2025, 2026):
    f = M.korean_tax_rate_formula('X1', fy)
    for x in (-2588.1, -0.5, 0, 1, 2, 2.5, 150, 200, 2522.9, 3000, 3001, 50000):
        xl, py = _eval_tax_formula(f, x), M.get_korean_marginal_tax_rate(x, fy)
        chk(f'FY{fy} 세전이익 {x}억: 엑셀 수식 = 파이썬', abs(xl - py) < 1e-12, (xl, py))

# 설명 문구가 '=' 로 시작하면 엑셀이 수식으로 읽어 오류(#NAME?)를 띄운다
for src in ('gpcm_kr.py', 'GPCM.py'):
    text = (Path(__file__).resolve().parent / src).read_text(encoding='utf-8')
    bad = re.findall(r"['\"]= [^'\"]*['\"]", text)
    chk(f'{src}: "= " 로 시작하는 설명 문구 없음', not bad, bad)

# --- 베타 기준지수 ---------------------------------------------------------------
print()
print("베타 기준지수")
print("-" * 66)
chk('모든 종목이 KOSPI 단일 기준 (코스닥도 동일)',
    len({M.get_market_index(t)[1] for t in ('005930', '247540', '091990', '000660')}) == 1,
    {t: M.get_market_index(t)[1] for t in ('005930', '247540')})
chk('기준지수는 ^KS11', M.get_market_index('247540')[1] == '^KS11', M.get_market_index('247540'))
# 조서에 적히는 방법론이 실제 동작과 어긋나면 안 된다 (한 번 틀리게 적은 적이 있다)

# --- 한국공인회계사회 참고치 (MRP·SRP) -------------------------------------------
print()
print("한공회 MRP·SRP")
print("-" * 66)
chk('MRP 가이던스 범위 7~9%', M.MRP_GUIDANCE == (0.07, 0.09), M.MRP_GUIDANCE)
chk('발표일 2026.06.05', M.CPA_GUIDANCE_DATE == '2026.06.05', M.CPA_GUIDANCE_DATE)
chk('3분위 SRP (4.02 / 1.19 / -0.45)',
    [r[1] for r in M.SRP_TERTILE] == [0.0402, 0.0119, -0.0045], [r[1] for r in M.SRP_TERTILE])
chk('5분위 SRP (4.86 / 2.67 / 0.97 / -0.06 / -0.51)',
    [r[1] for r in M.SRP_QUINTILE] == [0.0486, 0.0267, 0.0097, -0.0006, -0.0051],
    [r[1] for r in M.SRP_QUINTILE])
chk('시총으로 구간을 찾는다',
    (M.srp_for_market_cap(1500), M.srp_for_market_cap(5000), M.srp_for_market_cap(30000))
    == (0.0402, 0.0119, -0.0045))
chk('범위 밖은 지어내지 않는다', M.srp_for_market_cap(10) is None, M.srp_for_market_cap(10))

# 해외판(GPCM.py)은 국내 앱을 임포트할 수 없어 값을 복사해 둔다 — 어긋나면 안 된다
import re as _re, os as _os
_gl = (Path(_os.path.dirname(_os.path.abspath(__file__))) / 'GPCM.py').read_text(encoding='utf-8')
_blk = _gl.split('size_premium_options = {')[1].split('}')[0]
# 라벨 안에도 숫자가 있으므로, 닫는 따옴표 뒤의 값만 읽는다
_gl_vals = [float(v) for v in _re.findall(r'"\s*:\s*(-?\d+\.?\d*)', _blk)]
_kr_vals = [r[1] for r in M.SRP_TERTILE] + [r[1] for r in M.SRP_QUINTILE] + [0.0]
chk('해외판 SRP 값이 국내판과 동일', _gl_vals == _kr_vals, (_gl_vals, _kr_vals))

# --- 베타 기준지수 선택 ------------------------------------------------------------
print()
print("베타 기준지수 선택")
print("-" * 66)
MAP = {'005930': 'KOSPI', '247540': 'KOSDAQ'}
chk('기본값은 코스피 일괄 — 코스닥도 ^KS11 (기존 동작 불변)',
    M.get_market_index('247540', 'KOSPI', MAP)[1] == '^KS11',
    M.get_market_index('247540', 'KOSPI', MAP))
chk('인자를 안 주면 코스피 일괄', M.get_market_index('247540')[1] == '^KS11')
chk('소속시장: 코스닥 종목은 ^KQ11',
    M.get_market_index('247540', 'MARKET', MAP)[1] == '^KQ11')
chk('소속시장: 코스피 종목은 ^KS11',
    M.get_market_index('005930', 'MARKET', MAP)[1] == '^KS11')
chk('시장 판별 실패는 ^KS11 + 판정 None (조용히 넘기지 않음)',
    M.get_market_index('999999', 'MARKET', {})[1:] == ('^KS11', None),
    M.get_market_index('999999', 'MARKET', {}))
chk('Notes 가 선택과 일치 (KOSPI)',
    '단일 기준' in M.beta_basis_note('KOSPI')[0] and 'KQ11' not in ' '.join(M.beta_basis_note('KOSPI')))
chk('Notes 가 선택과 일치 (MARKET)',
    'KQ11' in ' '.join(M.beta_basis_note('MARKET')) and 'MRP' in ' '.join(M.beta_basis_note('MARKET')))

# 해외판도 같은 선택지를 써야 한다 (같은 회사를 다른 기준으로 재면 안 된다)
_gl2 = (Path(_os.path.dirname(_os.path.abspath(__file__))) / 'GPCM.py').read_text(encoding='utf-8')
_gl_basis = dict(_re.findall(r"'(KOSPI|MARKET)':\s*'([^']+)'", _gl2.split('BETA_BASIS = {')[1].split('}')[0]))
chk('해외판 BETA_BASIS 가 국내판과 동일', _gl_basis == M.BETA_BASIS, (_gl_basis, M.BETA_BASIS))
chk('해외판 기본값도 코스피 일괄', "beta_basis='KOSPI'" in _gl2)

chk('일별 주가 구간 목록이 국내·해외 동일',
    _re.findall(r"DAILY_PRICE_SPANS = \{([^}]*)\}", _gl2)[0].strip() ==
    _re.findall(r"DAILY_PRICE_SPANS = \{([^}]*)\}",
                (Path(_os.path.dirname(_os.path.abspath(__file__))) / 'gpcm_kr.py').read_text(encoding='utf-8'))[0].strip(),
    _re.findall(r"DAILY_PRICE_SPANS = \{([^}]*)\}", _gl2))
chk('해외판도 기준일 이후를 안 담는다 (조회 상한이 기준일)', 'base_dt + timedelta(days=1)' in _gl2)

# 해외 Price_History — 영업일 기준 수정주가 (앞값 채움 금지)
_ph = _gl2.split("[Sheet 11] Price_History")[1].split("[Sheet")[0]
_ph_code = '\n'.join(l for l in _ph.splitlines() if not l.strip().startswith('#'))
chk('해외 일별 주가는 앞값으로 채우지 않는다', '.ffill()' not in _ph_code,
    [l.strip() for l in _ph_code.splitlines() if 'ffill' in l])
# Abs 는 dropna 로 실거래일만 남기고, Rel 은 남더라도 NaN 대신 빈칸으로 쓴다
chk('결측은 NaN 이 아니라 빈칸으로 쓴다',
    'dropna()' in _ph and 'pd.isna(rv)' in _ph)
chk('시트에 수정주가·실거래일 기준을 밝힌다',
    '수정주가' in _ph and '실제 거래된 날만' in _ph)
chk('해외도 종목별 독립 표로 쓴다 (한 표에 날짜를 맞추지 않는다)',
    'for b, col in enumerate(df_abs.columns)' in _ph and 'df_abs[col].dropna()' in _ph)
chk('베타·주가 모두 auto_adjust 계열을 쓴다', "bundle['price_series'] = hist_adj" in _gl2)

# --- 0.8.5: 상장사 명단 열 구성 (사본의 열이 바뀌었다) -----------------------------
print()
print("상장사 명단 열 구성")
print("-" * 66)
_new = pd.DataFrame([{'Code': '4370', 'Name': '농심', 'Market': 'KOSPI', 'Sector': None,
                      'Industry': '기타 식품 제조업', 'Products': '라면,스낵', 'SettleMonth': '12월'},
                     {'Code': '0013V0', 'Name': '삼진식품', 'Market': 'KOSDAQ', 'Sector': '중견기업부',
                      'Industry': '수산물 가공 및 저장 처리업', 'Products': '어묵', 'SettleMonth': '12월'}])
_n = M.normalize_krx_desc(_new)
chk('새 구성: 업종은 Industry 열에서', _n.loc[0, 'Sector'] == '기타 식품 제조업', _n.to_dict('records'))
chk('새 구성: 주요제품은 Products 열에서', _n.loc[0, 'Industry'] == '라면,스낵')
chk('새 구성: 소속부는 Dept 로 옮긴다(업종 자리에 소속부가 들어가면 안 된다)',
    _n.loc[1, 'Dept'] == '중견기업부' and _n.loc[1, 'Sector'] == '수산물 가공 및 저장 처리업')
chk('종목코드 6자리 유지(영문 포함 코드 그대로)', list(_n['Code']) == ['004370', '0013V0'], list(_n['Code']))
_old = pd.DataFrame([{'Code': '004370', 'Name': '농심', 'Market': 'KOSPI', 'Sector': '기타 식품 제조업',
                      'Industry': '라면,스낵', 'SettleMonth': '12월'}])
_o = M.normalize_krx_desc(_old)
chk('예전 구성은 그대로 읽는다', _o.loc[0, 'Sector'] == '기타 식품 제조업' and _o.loc[0, 'Industry'] == '라면,스낵')

# --- 0.8.5: 2년 주간 베타 창 — 정확히 104개, 마지막 날짜는 실제 거래일 ---------------
print()
print("2년 주간 베타 창")
print("-" * 66)
_days = pd.bdate_range('2024-01-01', '2026-07-10')
_rng = np.random.default_rng(0)
_st = pd.Series(100 * np.cumprod(1 + _rng.normal(0, 0.01, len(_days))), index=_days)
_mk = pd.Series(100 * np.cumprod(1 + _rng.normal(0, 0.01, len(_days))), index=_days)
_ws, _wm = M.weekly_pair_2y(_st, _mk, '2026-06-30')
chk('주간 종가 105개 → 수익률 104개', len(_ws) == 105 and len(_wm) == 105, len(_ws))
chk('마지막 날짜는 기준일(화요일), 기준일 뒤 금요일이 아니다',
    _ws.index[-1] == pd.Timestamp('2026-06-30'), _ws.index[-1])
chk('기준일 뒤 자료 없음', (_ws.index <= pd.Timestamp('2026-06-30')).all())
chk('종목·지수 날짜가 같다', list(_ws.index) == list(_wm.index))
_gap = _st.drop(_st.loc['2025-03-03':'2025-03-21'].index)   # 3주 거래정지
_gs, _gm = M.weekly_pair_2y(_gap, _mk, '2026-06-30')
chk('정지 구간이 있으면 창을 늘려 메우지 않고 관측치가 준다', len(_gs) == 102, len(_gs))
_hol = _st.drop(pd.Timestamp('2026-06-30'))                  # 기준일 휴장 가정
_hs, _ = M.weekly_pair_2y(_hol, _mk, '2026-06-30')
chk('기준일이 휴장이면 직전 거래일로 찍힌다', _hs.index[-1] == pd.Timestamp('2026-06-29'), _hs.index[-1])
_d2 = pd.bdate_range('2021-01-01', '2024-04-10')
_st2 = pd.Series(100 * np.cumprod(1 + _rng.normal(0, 0.01, len(_d2))), index=_d2)
_mk2 = pd.Series(100 * np.cumprod(1 + _rng.normal(0, 0.01, len(_d2))), index=_d2)
_es, _em = M.weekly_pair_2y(_st2, _mk2, '2024-03-31')               # 일요일 분기말
chk('기준일이 일요일이어도 수익률 104개(한 주가 비지 않는다)', len(_es) == 105 and len(_em) == 105, len(_es))
chk('일요일 기준일 → 마지막은 직전 금요일', _es.index[-1] == pd.Timestamp('2024-03-29'), _es.index[-1])
_ss, _ = M.weekly_pair_2y(_st.drop(_st.loc['2026-06-15':'2026-06-30'].index), _mk, '2026-06-30')
chk('기준일 직전 종목 거래정지면 창을 뒤로 밀지 않고 관측치가 준다(지수 거래일 기준)',
    len(_ss) == 102 and _ss.index[-1] == pd.Timestamp('2026-06-12'), (len(_ss), _ss.index[-1]))
chk('수집부가 이 함수를 쓴다(금요일 리샘플 제거)',
    "weekly_pair_2y(" in open(Path(__file__).resolve().parent / 'gpcm_kr.py', encoding='utf-8').read()
    and "stock_prices_2y.resample('W-FRI')" not in open(Path(__file__).resolve().parent / 'gpcm_kr.py', encoding='utf-8').read())

# --- 0.8.5: 상장 우선주 시가 -------------------------------------------------------
print()
print("상장 우선주 시가")
print("-" * 66)
_krx = pd.DataFrame([
    {'Code': '001680', 'Name': '대상', 'Stocks': 34_000_000},
    {'Code': '001685', 'Name': '대상우', 'Stocks': 1_370_223},
    {'Code': '264900', 'Name': '크라운제과', 'Stocks': 13_000_000},
    {'Code': '26490K', 'Name': '크라운제과우', 'Stocks': 772_939},
    {'Code': '005380', 'Name': '현대차', 'Stocks': 200_000_000},
    {'Code': '005385', 'Name': '현대차우', 'Stocks': 24_000_000},
    {'Code': '005387', 'Name': '현대차2우B', 'Stocks': 36_000_000},
    {'Code': '005389', 'Name': '현대차3우B', 'Stocks': 2_000_000},
    {'Code': '004370', 'Name': '농심', 'Stocks': 6_000_000},
    {'Code': '004379', 'Name': '농심다른회사', 'Stocks': 1},   # 앞 5자리만 같은 가짜
])
chk('대상 → 대상우(001685)', M.find_listed_preferred('001680', _krx) == [('001685', '대상우')])
chk('크라운제과 → 영문 포함 코드 26490K', M.find_listed_preferred('264900', _krx) == [('26490K', '크라운제과우')])
chk('여러 종류(우·2우B·3우B)를 모두 찾는다', len(M.find_listed_preferred('005380', _krx)) == 3)
chk('앞 5자리만 같은 다른 이름은 잡지 않는다', M.find_listed_preferred('004370', _krx) == [])
_prices = {'001685': (14580.0, '2026-06-30'), '26490K': (7240.0, '2026-06-30'),
           '005385': (100.0, '2026-06-30'), '005387': (100.0, '2026-06-30')}
_orig_price = M.get_stock_price
M.get_stock_price = lambda code, d: _prices.get(code, (None, None))
try:
    _v, _i = M.get_preferred_market_value('001680', '2026-06-30', _krx, {'pref_shares': 1_370_223})
    chk('한 종류: DART 우선주 유통주식수 × 기준일 종가', abs(_v - round(14580 * 1_370_223 / 1e8, 1)) < 1e-9, (_v, _i))
    chk('근거에 DART 출처', _i['items'][0]['source'].startswith('DART'), _i)
    _v2, _i2 = M.get_preferred_market_value('005380', '2026-06-30', _krx, {'pref_shares': 62_000_000})
    chk('한 종류라도 종가를 못 구하면 0 → 장부로 되돌아간다(일부만 더하지 않는다)',
        _v2 == 0 and _i2['problems'], (_v2, _i2))
    _v3, _i3 = M.get_preferred_market_value('004370', '2026-06-30', _krx, {})
    chk('상장 우선주가 없으면 (0, None)', _v3 == 0 and _i3 is None)
    # 기준일 명단: 오늘은 상장폐지돼 조회일 명단에 없는 우선주도 찾는다
    _krx_now = _krx[_krx['Code'] != '001685']
    _v4, _i4 = M.get_preferred_market_value('001680', '2026-06-30', _krx_now, {'pref_shares': 1_370_223},
                                            df_asof=_krx, asof_date='2026-06-30')
    chk('기준일 명단으로 찾으면 그 뒤 상장폐지된 우선주도 잡는다', _v4 > 0 and _i4['asof_roster'], (_v4, _i4))
    chk('근거에 기준일 명단 날짜', _i4['roster'] == '기준일 명단(2026-06-30)', _i4.get('roster'))
    _v5, _i5 = M.get_preferred_market_value('001680', '2026-06-30', _krx_now, {'pref_shares': 1_370_223})
    chk('기준일 명단이 없고 조회일 명단에도 없으면 0 + 비상장 우선주 근거(경고용)',
        _v5 == 0 and _i5['unlisted_pref_shares'] == 1_370_223 and not _i5['asof_roster'], _i5)
    _krx_asof_multi = _krx.copy()
    _krx_asof_multi.loc[_krx_asof_multi['Code'] == '005385', 'Stocks'] = 11
    _prices.update({'005389': (100.0, '2026-06-30')})
    _v6, _i6 = M.get_preferred_market_value('005380', '2026-06-30', _krx, {'pref_shares': 62_000_000},
                                            df_asof=_krx_asof_multi, asof_date='2026-06-30')
    chk('여러 종류면 기준일 명단의 종류별 상장주식수를 쓴다',
        [it['shares'] for it in _i6['items'] if it['code'] == '005385'] == [11]
        and all(it['source'] == 'KRX 상장주식수' for it in _i6['items']) and _i6['roster'] == '기준일 명단(2026-06-30)', _i6)
finally:
    M.get_stock_price = _orig_price
_src = open(Path(__file__).resolve().parent / 'gpcm_kr.py', encoding='utf-8').read()
chk('GPCM 시트 우선주 = 시가 있으면 시가, 없으면 장부',
    "_pm = f'SUMIFS(Market_Cap!Q:Q" in _src and "=IF({_pm}>0, {_pm}" in _src)
chk('파이썬 계산도 같은 규칙(시가가 있으면 Preferred 를 시가로 바꾼다)',
    "period_metrics['Preferred'] = pref_mkt_100m" in _src)
chk('수집부는 기준일 종목 명단으로 우선주를 찾는다',
    "get_krx_price_listing_as_of(bds)" in _src and "df_asof=_pr_df" in _src)

# --- 0.8.5: Data_Quality — 결손 피어 경고, 결정계수 정보 ---------------------------
print()
print("Data_Quality 보강")
print("-" * 66)
_q = M.QualityLog()
_wk = pd.date_range('2024-07-05', periods=105, freq='W-FRI')
_mkw = pd.Series(100 * np.cumprod(1 + _rng.normal(0, 0.02, 105)), index=_wk)
_stw = pd.Series(100 * np.cumprod(1 + _rng.normal(0, 0.02, 105)), index=_wk)   # 시장과 무관 → 결정계수 낮음
_comp = [{'Ticker': '001680', 'Company': '결손㈜', 'Market_Cap': 5000, 'IBD': 3000, 'NCI': 0,
          'Preferred': 0, 'Equity': 4000, 'Pretax_Income': -2588.1,
          'Stock_Weekly_Prices_2Y': _stw, 'Market_Weekly_Prices_2Y': _mkw},
         {'Ticker': '004370', 'Company': '흑자㈜', 'Market_Cap': 19000, 'IBD': 700, 'NCI': 0,
          'Preferred': 0, 'Equity': 28000, 'Pretax_Income': 2522.9,
          'Stock_Weekly_Prices_2Y': _mkw * (1 + _rng.normal(0, 0.001, 105)), 'Market_Weekly_Prices_2Y': _mkw}]
M.calculate_wacc_and_beta(['001680', '004370'], _comp, 0.242, 0.03925, 0.08, -0.0045, 0.04375, '2Y',
                          fiscal_year=2026, quality=_q)
_tax_w = [r for r in _q.rows if r['Item'] == 'Tax Rate']
chk('결손 피어는 세율 경고(WARN)', len(_tax_w) == 1 and _tax_w[0]['Level'] == 'WARN'
    and _tax_w[0]['Ticker'] == '001680' and '22.0%' in _tax_w[0]['Message'], _tax_w)
_r2_i = [r for r in _q.rows if r['Level'] == 'INFO' and '결정계수' in r['Message']]
chk('결정계수 0.1 미만은 정보(INFO)로만, 흑자·고설명력 회사는 표시 없음',
    [r['Ticker'] for r in _r2_i] == ['001680'], _r2_i)

print()
print(f"잘못된 항목 {fails}건")
sys.exit(1 if fails else 0)
