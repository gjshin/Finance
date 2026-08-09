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

print()
print(f"잘못된 항목 {fails}건")
sys.exit(1 if fails else 0)
