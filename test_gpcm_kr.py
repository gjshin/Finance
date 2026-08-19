"""gpcm_kr.py 회귀 테스트.

다기간 재무데이터 모드가 화면에서 고른 기간의 보고서를 실제로 조회하는지 확인한다.
과거에 분기→보고서코드 매핑이 뒤집혀 있어 '연간'을 고르면 1분기보고서를 가져왔다
(매출이 1/4로 나오는데 오류는 나지 않아 알아채기 어려웠다).
"""
import sys, types, pandas as pd, numpy as np
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

print()
print(f"잘못된 항목 {fails}건")
sys.exit(1 if fails else 0)
