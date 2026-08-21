"""gpcm_kr.py 회귀 테스트.

다기간 재무데이터 모드가 화면에서 고른 기간의 보고서를 실제로 조회하는지 확인한다.
과거에 분기→보고서코드 매핑이 뒤집혀 있어 '연간'을 고르면 1분기보고서를 가져왔다
(매출이 1/4로 나오는데 오류는 나지 않아 알아채기 어려웠다).
"""
import sys, types, pandas as pd, numpy as np
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
OLD_2025 = '=IF(AE6<=2, 0.099, IF(AE6<=200, 0.209, IF(AE6<=3000, 0.231, 0.264)))'
chk('FY2025 수식이 종전과 글자까지 같다', M.korean_tax_rate_formula('AE6', 2025) == OLD_2025,
    M.korean_tax_rate_formula('AE6', 2025))
chk('FY2024 도 동일', M.korean_tax_rate_formula('AE6', 2024) == OLD_2025)
chk('FY2026 은 개정 세율(11%/22%/24.2%/27.5%)',
    M.korean_tax_rate_formula('AE6', 2026) ==
    '=IF(AE6<=2, 0.11, IF(AE6<=200, 0.22, IF(AE6<=3000, 0.242, 0.275)))',
    M.korean_tax_rate_formula('AE6', 2026))
for fy in (2024, 2025, 2026):
    br = M.get_korean_tax_brackets(fy)
    f = M.korean_tax_rate_formula('X1', fy)
    chk(f'FY{fy} 수식 세율이 파이썬 표와 일치',
        all(f'{r:g}' in f for _, r in br), f)

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

print()
print(f"잘못된 항목 {fails}건")
sys.exit(1 if fails else 0)
