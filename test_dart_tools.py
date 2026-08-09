"""dart_tools.py 회귀 테스트.

네트워크 없이 가짜 DART 응답으로 돌린다. 비율은 손으로 계산한 값과 대조한다.
"""

import sys
from datetime import datetime

import pandas as pd

import dart_tools as T


def _row(sj, account, aid, thstrm, frmtrm=None, add=None):
    return {'sj_nm': sj, 'account_nm': account, 'account_id': aid,
            'thstrm_amount': f'{thstrm:,}' if thstrm is not None else '',
            'frmtrm_amount': f'{frmtrm:,}' if frmtrm is not None else '',
            'thstrm_add_amount': f'{add:,}' if add is not None else ''}


# 손으로 계산하기 쉬운 숫자로 구성한 가상 회사
BS = [
    _row('연결재무상태표', '자산총계', 'ifrs-full_Assets', 2000, 1600),
    _row('연결재무상태표', '유동자산', 'ifrs-full_CurrentAssets', 800),
    _row('연결재무상태표', '매출채권및기타채권', 'ifrs-full_TradeAndOtherCurrentReceivables', 250, 150),
    _row('연결재무상태표', '재고자산', 'ifrs-full_Inventories', 120, 80),
    _row('연결재무상태표', '부채총계', 'ifrs-full_Liabilities', 1200),
    _row('연결재무상태표', '유동부채', 'ifrs-full_CurrentLiabilities', 400),
    _row('연결재무상태표', '자본총계', 'ifrs-full_Equity', 800, 600),
]
IS = [
    _row('연결포괄손익계산서', '매출액', 'ifrs-full_Revenue', 1000),
    _row('연결포괄손익계산서', '매출원가', 'ifrs-full_CostOfSales', 600),
    _row('연결포괄손익계산서', '매출총이익', 'ifrs-full_GrossProfit', 400),
    _row('연결포괄손익계산서', '영업이익', 'dart_OperatingIncomeLoss', 150),
    _row('연결포괄손익계산서', '이자비용', 'ifrs-full_InterestExpense', 50),
    _row('연결포괄손익계산서', '당기순이익', 'ifrs-full_ProfitLoss', 105),
]


class FakeReader:
    """OpenDartReader 대역. 네트워크를 타지 않는다."""

    def __init__(self, rows=None, report_df=None, doc=None):
        self.rows = BS + IS if rows is None else rows
        self.report_df = report_df
        self.doc = doc
        self.calls = 0

    def find_corp_code(self, name):
        return '00126380' if name else None

    def finstate_all(self, corp, year, reprt_code='11011', fs_div='CFS'):
        self.calls += 1
        if fs_div != 'CFS':
            return pd.DataFrame()
        return pd.DataFrame(self.rows)

    def report(self, corp, key_word, bsns_year, reprt_code='11011'):
        self.calls += 1
        return self.report_df if self.report_df is not None else pd.DataFrame()

    def list(self, corp=None, start=None, end=None, kind='', **kw):
        return pd.DataFrame([{'rcept_no': '20240312000736', 'rcept_dt': '20240312',
                              'corp_name': '테스트', 'report_nm': '감사보고서(2023.12)',
                              'flr_nm': '한영회계법인', 'rm': ''}])

    def document(self, rcp_no, cache=True):
        return self.doc or ''

    # --- 확장 기능 대역 ---
    corp_codes = pd.DataFrame({
        'corp_code': ['00126380', '00999999'],
        'corp_name': ['테스트전자', '테스트비상장'],
        'stock_code': ['005930', '']})           # 두 번째는 비상장(종목코드 없음)

    def company(self, corp):
        return {'corp_name': '테스트전자', 'induty_code': '264',
                'est_dt': '19690113', 'ceo_nm': '홍길동',
                'status': '000', 'message': '정상'}

    def finstate(self, corp, bsns_year, reprt_code='11011'):
        self.calls += 1
        return pd.DataFrame([{'sj_nm': '재무상태표', 'account_nm': '자산총계',
                              'thstrm_amount': '2,000', 'frmtrm_amount': '1,600'}])

    def event(self, corp, key_word, start=None, end=None):
        return pd.DataFrame([{'corp_code': 'x', 'rcept_no': '20240501000123',
                              'report_nm': f'{key_word} 관련 주요사항보고서'}])

    def regstate(self, corp, key_word, start=None, end=None):
        return pd.DataFrame([{'corp_code': 'x', 'rcept_no': '20240601000456',
                              'report_nm': f'{key_word} 증권신고서'}])

    def major_shareholders(self, corp):
        return pd.DataFrame([{'corp_code': 'x', 'repror': '홍길동', 'stkrt': '15.2'}])

    def major_shareholders_exec(self, corp):
        return pd.DataFrame([{'corp_code': 'x', 'repror': '임원A', 'sp_stock_lmp_cnt': '1000'}])

    def list_date_ex(self, date=None, cache=True):
        return pd.DataFrame([
            {'rcept_no': '1', 'corp_name': 'A사', 'report_nm': '감사보고서(2023.12)'},
            {'rcept_no': '2', 'corp_name': 'B사', 'report_nm': '분기보고서(2024.03)'}])

    def sub_docs(self, s, match=None):
        return pd.DataFrame([{'title': '감사보고서', 'url': 'http://x'}])

    def attach_files(self, s):
        return pd.DataFrame([{'name': '재무제표.xlsx', 'url': 'http://y'}])

    def xbrl_taxonomy(self, sj_div):
        return pd.DataFrame([{'account_id': 'ifrs-full_Assets', 'account_nm': '자산총계'}])


def client(**kw):
    return T.DartClient(api_key=None, reader=FakeReader(**kw))


PASS, FAIL = [], []


def check(name, got, want, tol=1e-9):
    ok = (got is None and want is None) or (
        got is not None and want is not None and abs(got - want) <= tol)
    (PASS if ok else FAIL).append(name)
    mark = 'OK ' if ok else 'FAIL'
    print(f"  [{mark}] {name}: 계산={got}  기대={want}")


def check_true(name, cond, detail=''):
    (PASS if cond else FAIL).append(name)
    print(f"  [{'OK ' if cond else 'FAIL'}] {name} {detail}")


print("=== 1. 비율 계산 (손계산 대조) ===")
LAST_Y = T.latest_filed_period()[0] - 1
r = T.compute_ratios(client(), ['테스트'], [LAST_Y])['results'][0]
v = r['ratios']

# 매출채권회전율 = 매출 1000 / 평균매출채권 (250+150)/2 = 200  ->  5.0회
check('매출채권회전율 (평균잔액)', v['매출채권회전율'], 5.0)
# 재고자산회전율 = 매출원가 600 / 평균재고 (120+80)/2 = 100  ->  6.0회
check('재고자산회전율 (평균잔액)', v['재고자산회전율'], 6.0)
# 총자산회전율 = 1000 / (2000+1600)/2 = 1800  ->  0.56회
check('총자산회전율', v['총자산회전율'], 0.56)
# 부채비율 = 1200 / 800 = 150%  (기말)
check('부채비율', v['부채비율'], 150.0)
# 유동비율 = 800 / 400 = 200%
check('유동비율', v['유동비율'], 200.0)
# 이자보상배율 = 영업이익 150 / 이자비용 50 = 3.0배
check('이자보상배율', v['이자보상배율'], 3.0)
# ROE = 105 / 평균자본 (800+600)/2 = 700 = 15%
check('ROE (평균자본)', v['ROE'], 15.0)
# ROA = 105 / 1800 = 5.83%
check('ROA (평균자산)', v['ROA'], 5.83)
# 매출총이익률 = 400/1000 = 40%
check('매출총이익률', v['매출총이익률'], 40.0)
# 영업이익률 = 150/1000 = 15%
check('영업이익률', v['영업이익률'], 15.0)

print("\n=== 2. 감사 증적: 어떤 계정을 썼는지 남기는가 ===")
check_true('분모에 평균 사용을 명시', '평균' in r['sources']['매출채권회전율']['분모'],
           f"→ {r['sources']['매출채권회전율']['분모']}")
check_true('기말 사용도 구분해 표시', '기말' in r['sources']['부채비율']['분모'],
           f"→ {r['sources']['부채비율']['분모']}")
check_true('연결/별도 구분 표시', r['basis'] == '연결', f"→ {r['basis']}")

print("\n=== 3. 계정이 없을 때 0이 아니라 null 인가 ===")
no_inv = [x for x in BS + IS if '재고' not in x['account_nm'] and '이자비용' not in x['account_nm']]
r2 = T.compute_ratios(client(rows=no_inv), ['테스트'], [LAST_Y])['results'][0]
check('재고 없음 → null', r2['ratios']['재고자산회전율'], None)
check('이자비용 없음 → null', r2['ratios']['이자보상배율'], None)
check_true('못 구한 이유를 남김', bool(r2['not_available'].get('재고자산회전율')),
           f"→ {r2['not_available'].get('재고자산회전율')}")
check_true('나머지 비율은 정상 계산', r2['ratios']['부채비율'] == 150.0)

print("\n=== 4. 0으로 나누기 방어 ===")
zero_eq = [(_row('연결재무상태표', '자본총계', 'ifrs-full_Equity', 0)
            if x['account_nm'] == '자본총계' else x) for x in BS + IS]
r3 = T.compute_ratios(client(rows=zero_eq), ['테스트'], [LAST_Y])['results'][0]
check('자본 0 → 부채비율 null', r3['ratios']['부채비율'], None)
check_true("자본잠식을 '계정 없음'이 아니라 '0'으로 설명",
           '0이라' in str(r3['not_available'].get('부채비율', '')),
           f"→ {r3['not_available'].get('부채비율', '')}")

# 영업이익이 진짜 0인 회사: 영업이익률은 null이 아니라 0.0 이어야 한다
zero_op = [(_row('연결포괄손익계산서', '영업이익', 'dart_OperatingIncomeLoss', 0)
            if x['account_nm'] == '영업이익' else x) for x in BS + IS]
r3b = T.compute_ratios(client(rows=zero_op), ['테스트'], [LAST_Y],
                       ratios=['영업이익률'])['results'][0]
check('영업이익 0 → 0.0 (null 아님)', r3b['ratios']['영업이익률'], 0.0)

print("\n=== 5. 미래 기간 요청 차단 ===")
future = datetime.now().year + 1
r4 = T.compute_ratios(client(), ['테스트'], [future])['results'][0]
check_true('아직 공시 전이면 에러 반환', 'error' in r4, f"→ {r4.get('error', '')[:45]}")
check_true('조회 가능한 최근 기간을 안내', '조회 가능한' in r4.get('error', ''))

print("\n=== 6. 같은 조회 반복 시 캐시 동작 ===")
c = client()
T.compute_ratios(c, ['테스트'], [LAST_Y])
first = c.dart.calls
T.compute_ratios(c, ['테스트'], [LAST_Y])
check_true('두 번째 호출은 통신 없음', c.dart.calls == first, f"(호출 {first}회 유지)")

print("\n=== 7. 분기보고서는 누적금액 칸을 보는가 ===")
q_rows = [
    _row('연결재무상태표', '자산총계', 'ifrs-full_Assets', 2000, 1600),
    _row('연결재무상태표', '자본총계', 'ifrs-full_Equity', 800, 600),
    # 분기보고서: thstrm_amount=당분기 3개월(300), thstrm_add_amount=누적(900)
    _row('연결포괄손익계산서', '매출액', 'ifrs-full_Revenue', 300, None, 900),
    _row('연결포괄손익계산서', '당기순이익', 'ifrs-full_ProfitLoss', 30, None, 105),
]
r5 = T.compute_ratios(client(rows=q_rows), ['테스트'], [LAST_Y], quarter='3Q',
                      ratios=['총자산회전율', 'ROE'])['results'][0]
# 누적 900 / 평균자산 1800 = 0.5   (3개월치 300을 쓰면 0.17이 되어 틀림)
check('3Q 총자산회전율 = 누적 기준', r5['ratios']['총자산회전율'], 0.5)
check('3Q ROE = 누적 기준', r5['ratios']['ROE'], 15.0)

print("\n=== 8. 재무제표 조회 ===")
fs = T.get_financial_statements(client(), ['테스트'], LAST_Y, statement='IS')
res = fs['results'][0]
check_true('손익계산서만 걸러짐', all('손익' in x['statement'] for x in res['rows']),
           f"({len(res['rows'])}행)")
fs2 = T.get_financial_statements(client(), ['테스트'], LAST_Y, account_filter='매출채권')
check_true('계정 키워드 필터', len(fs2['results'][0]['rows']) == 1)

print("\n=== 9. 정기보고서 주요정보 ===")
bad = T.get_report_item(client(), ['테스트'], '없는항목', [LAST_Y])
check_true('잘못된 항목 → 가능한 목록 안내', 'available' in bad,
           f"({len(bad.get('available', []))}종)")

# 목록이 라이브러리와 어긋나면 조회가 실패하므로 실제 키 집합과 대조한다
import inspect, re as _re
import OpenDartReader.dart_report as _rep
_lib = set(_re.findall(r"'([가-힣]+)'\s*:\s*'[a-zA-Z]+'", inspect.getsource(_rep)))
check_true('지원 항목이 라이브러리와 일치', _lib == set(T.REPORT_ITEMS),
           f"(라이브러리 {len(_lib)}종 / 내 목록 {len(T.REPORT_ITEMS)}종)")
ok = T.get_report_item(
    client(report_df=pd.DataFrame([{'corp_code': 'x', 'rcept_no': 'y',
                                    'inv_prm': '자회사A', 'trmend_blce_qy': '100'}])),
    ['테스트'], '타법인출자', [LAST_Y])
rows = ok['results'][0]['rows']
check_true('타법인출자 조회', len(rows) == 1 and rows[0]['inv_prm'] == '자회사A')
check_true('회사 내부코드는 제거, 접수번호는 보존(근거 추적용)',
           'corp_code' not in rows[0] and 'rcept_no' in rows[0])

print("\n=== 10. 공시 원문 섹션 추출 ===")
doc = ("<html><p>I. 회사의 개요</p>본문 개요입니다."
       "<p>II. 사업의 내용</p>게임 개발 및 퍼블리싱 사업을 영위합니다."
       "<p>III. 재무에 관한 사항</p>재무 본문."
       "<p>주석 8. 매출채권</p>매출채권 손실충당금 내역입니다.</html>")
toc = T.get_filing_document(client(doc=doc), '20240312000736')
check_true('목차 추출', any('사업의 내용' in s for s in toc['sections']),
           f"→ {toc['sections'][:4]}")
sec = T.get_filing_document(client(doc=doc), '20240312000736', section='사업의 내용')
check_true('해당 절만 반환', '게임 개발' in sec['content'] and '재무 본문' not in sec['content'])
note = T.get_filing_document(client(doc=doc), '20240312000736', section='주석 8')
check_true('감사보고서 주석 추출', '손실충당금' in note['content'])
miss = T.get_filing_document(client(doc=doc), '20240312000736', section='없는절')
check_true('없는 절은 안내 메시지', 'error' in miss)

print("\n=== 11. 확장 기능 (OpenDartReader 전 기능) ===")

# 기업개황
prof = T.get_company_profile(client(), ['테스트'])['results'][0]['profile']
check_true('기업개황 조회', prof.get('induty_code') == '264' and prof.get('ceo_nm') == '홍길동')
check_true('응답코드는 제거', 'status' not in prof and 'message' not in prof)

# 주요사항보고서 36종
ev = T.get_major_events(client(), ['테스트'], '소송')
check_true('주요사항보고서 조회', ev['results'][0]['rows'][0]['report_nm'].startswith('소송'))
check_true('접수번호 보존', 'rcept_no' in ev['results'][0]['rows'][0])
bad_ev = T.get_major_events(client(), ['테스트'], '없는사건')
check_true('잘못된 사건 → 36종 목록 안내',
           'available' in bad_ev and len(bad_ev['available']) == 36,
           f"({len(bad_ev.get('available', []))}종)")

# 증권신고서
rs = T.get_registration_statements(client(), ['테스트'], '합병')
check_true('증권신고서 조회', '합병' in rs['results'][0]['rows'][0]['report_nm'])
check_true('잘못된 종류 → 6종 목록 안내',
           len(T.get_registration_statements(client(), ['t'], 'xx')['available']) == 6)

# 지분공시 2종
sh = T.get_shareholding_reports(client(), ['테스트'], '대량보유')
check_true('대량보유(5%) 조회', sh['results'][0]['rows'][0]['stkrt'] == '15.2')
ex = T.get_shareholding_reports(client(), ['테스트'], '임원주요주주')
check_true('임원·주요주주 조회', ex['results'][0]['rows'][0]['repror'] == '임원A')
check_true('잘못된 종류 안내', 'available' in T.get_shareholding_reports(client(), ['t'], 'xx'))

# 날짜별 전체 공시 + 키워드
day = T.list_filings_by_date(client(), '2024-03-12')
check_true('날짜별 공시 조회', day['total'] == 2)
day_f = T.list_filings_by_date(client(), '2024-03-12', keyword='감사보고서')
check_true('보고서명 키워드 필터', day_f['total'] == 1
           and '감사보고서' in day_f['filings'][0]['report_nm'])

# 첨부문서
att = T.get_filing_attachments(client(), '20240312000736')
check_true('하위문서 목록', att['sub_documents'][0]['title'] == '감사보고서')
check_true('첨부파일 목록', att['attached_files'][0]['name'] == '재무제표.xlsx')

# XBRL 표준계정
tax = T.get_xbrl_taxonomy(client(), 'BS1')
check_true('XBRL 표준계정 조회', tax['accounts'][0]['account_id'] == 'ifrs-full_Assets')

# 요약 재무제표
summ = T.get_financial_statements(client(), ['테스트'], LAST_Y, summary=True)
check_true('요약 재무제표(주요계정)', summ['summary'] is True
           and summ['results'][0]['rows'][0]['amount'] == 2000.0)

# 비상장사 검색 — KRX 목록에 없어도 DART 전체 목록에서 찾아야 한다
import types as _types
_c = client()
_c._listing = pd.DataFrame({'Code': ['005930'], 'Name': ['테스트전자'],
                            'Market': ['KOSPI'], 'Sector': ['전자부품']})
found = T.find_companies(_c, query='테스트')
names = {c['name']: c['listed'] for c in found['companies']}
check_true('상장사 검색', names.get('테스트전자') is True, f"→ {list(names)}")
check_true('비상장사도 함께 검색', names.get('테스트비상장') is False)


print("\n=== 11-2. DART 공식 재무지표 (라이브러리 미지원 → 직접 호출) ===")
import dart_tools as _T

_calls = []


def _fake_get(api_key, endpoint, params, timeout=10):
    _calls.append((endpoint, params.get('idx_cl_code')))
    return [{'corp_name': '테스트', 'idx_cl_nm': '수익성지표',
             'idx_nm': '영업이익률', 'idx_val': '15.0'}], None


_real_get = _T._dart_get
_T._dart_get = _fake_get
try:
    c = T.DartClient(api_key='KEY', reader=FakeReader())
    r = T.get_financial_indicators(c, ['테스트'], [LAST_Y], categories=['수익성'])
    row = r['results'][0]['rows'][0]
    check_true('재무지표 조회', row['indicator'] == '영업이익률' and row['value'] == '15.0')
    check_true('분류명은 DART 응답값 사용', row['category'] == '수익성지표',
               '(내 코드 매핑이 틀려도 안전)')
    check_true('단일회사는 fnlttSinglIndx', _calls[0][0] == 'fnlttSinglIndx.json')

    _calls.clear()
    T.get_financial_indicators(c, ['테스트', '테스트2'], [LAST_Y], categories=['수익성'])
    check_true('여러 회사는 fnlttCmpnyIndx', _calls[0][0] == 'fnlttCmpnyIndx.json')

    # 새 클라이언트로 재야 앞 테스트의 캐시에 영향받지 않는다
    _calls.clear()
    c2 = T.DartClient(api_key='KEY', reader=FakeReader())
    r2 = T.get_financial_indicators(c2, ['테스트'], [LAST_Y])
    check_true('분류 생략 시 4종 모두 조회', len(_calls) == 4,
               f"({[x[1] for x in _calls]})")
    check_true('4종 결과가 다 담김', len(r2['results'][0]['rows']) == 4)

    _calls.clear()
    T.get_financial_indicators(c2, ['테스트'], [LAST_Y])
    check_true('같은 조회 반복 시 통신 없음(캐시)', len(_calls) == 0)

    bad = T.get_financial_indicators(c, ['테스트'], [LAST_Y], categories=['없는분류'])
    check_true('모르는 분류 → 가능한 목록 안내', 'available' in bad)

    fut = T.get_financial_indicators(c, ['테스트'], [datetime.now().year + 1])
    check_true('미공시 기간 차단', 'error' in fut['results'][0])

    # 키가 없으면 이 기능만 막히고 나머지는 계속 동작해야 한다
    nokey = T.get_financial_indicators(T.DartClient(api_key=None, reader=FakeReader()),
                                       ['테스트'], [LAST_Y])
    check_true('API 키 없으면 이 기능만 차단', 'error' in nokey)
finally:
    _T._dart_get = _real_get


print("\n=== 12. MCP 서버 계층 (Claude가 실제로 부르는 경로) ===")
import asyncio
import json
import os

os.environ.pop('OPENDART_API_KEY', None)
import dart_mcp as M

# 키가 없으면 조용히 실패하지 않고 발급처를 안내해야 한다
M._client = None
try:
    M.client()
    check_true('API 키 없을 때 안내', False, '(오류 없이 통과됨)')
except RuntimeError as e:
    check_true('API 키 없을 때 발급처 안내',
               'OPENDART_API_KEY' in str(e) and 'opendart.fss.or.kr' in str(e))

# 도구가 6개 다 등록됐는지
tools = asyncio.run(M.server.list_tools())
check_true('도구 14종 등록', len(tools) == 14, f"({len(tools)}개)")

M._client = T.DartClient(api_key=None, reader=FakeReader())


def call_tool(name, args):
    res = asyncio.run(M.server.call_tool(name, args))
    if res.is_error:
        return {'_error': str(res.content[0].text if res.content else '')}
    return json.loads(res.content[0].text)


# MCP를 거쳐 나온 값이 직접 계산과 같아야 한다
direct = T.compute_ratios(client(), ['테스트'], [LAST_Y])['results'][0]['ratios']
viamcp = call_tool('compute_ratios', {'companies': ['테스트'], 'years': [LAST_Y]})['results'][0]['ratios']
check_true('MCP 경유 결과 = 직접 호출 결과', direct == viamcp,
           f"(부채비율 {viamcp['부채비율']}%, 매출채권회전율 {viamcp['매출채권회전율']}회)")

# 공시검색 → 접수번호 → 원문 주석 추출까지 이어지는지
M._client = T.DartClient(api_key=None, reader=FakeReader(
    doc="<p>I. 회사의 개요</p>개요<p>주석 8. 매출채권</p>손실충당금 내역입니다."))
found = call_tool('search_filings', {'company': '테스트', 'kind': 'F'})
rcept = found['filings'][0]['rcept_no']
note = call_tool('get_filing_document', {'rcept_no': rcept, 'section': '주석 8'})
check_true('감사보고서 검색 → 주석 추출 연결', '손실충당금' in note.get('content', ''),
           f"(접수번호 {rcept})")


print("\n" + "=" * 60)
print(f"통과 {len(PASS)}건 / 실패 {len(FAIL)}건")
if FAIL:
    print("실패 항목:")
    for f in FAIL:
        print(f"  - {f}")
    sys.exit(1)
print("전부 통과")
