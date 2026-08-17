"""gpcm_mcp.py 회귀 테스트 — Streamlit 없이 MCP 도구 경로를 검증한다.

이 환경은 DART·KRX가 차단돼 있어 test_gpcm_kr_quality.py 와 같은 스텁 패턴을 쓴다:
gpcm_kr 모듈 전역을 원숭이 패칭해 파이프라인을 통째로 돌리고, 산출 엑셀까지 연다.

따로 확인하는 것 하나 — stdout 위생. MCP 는 stdout 이 프로토콜 통로라,
gpcm_kr 임포트(사이드바 실행)가 한 글자라도 흘리면 클라이언트 파서가 깨진다.
"""
import os, subprocess, sys, tempfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

os.environ["DART_API_KEY"] = "TESTKEY-" + "0" * 32

import gpcm_mcp as W

M = W._load()

# --- 파이프라인 스텁 (test_gpcm_kr_quality.py 와 동일한 대역) -----------------

BS = pd.DataFrame([
    {'sj_nm': '연결재무상태표', 'account_nm': '현금및현금성자산', 'account_id': 'ifrs-full_CashAndCashEquivalents', 'thstrm_amount': '100000000000'},
    {'sj_nm': '연결재무상태표', 'account_nm': '자본총계', 'account_id': 'ifrs-full_Equity', 'thstrm_amount': '600000000000'},
])
PL = pd.DataFrame([
    {'sj_nm': '연결포괄손익계산서', 'account_nm': '매출액', 'account_id': 'ifrs-full_Revenue',
     'thstrm_amount': '800000000000', 'thstrm_add_amount': '800000000000'},
    {'sj_nm': '연결포괄손익계산서', 'account_nm': '영업이익', 'account_id': 'dart_OperatingIncomeLoss',
     'thstrm_amount': '80000000000', 'thstrm_add_amount': '80000000000'},
])


class Dart:
    corp_codes = pd.DataFrame({'corp_code': ['00126380'], 'corp_name': ['테스트'], 'stock_code': ['005930']})
    def find_corp_code(self, c): return '00126380'
    def company(self, c): return {'corp_name': '테스트'}
    def finstate_all(self, corp, y, reprt_code='11011', fs_div='CFS'):
        return BS.copy() if fs_div == 'CFS' else pd.DataFrame()


def stub_pipeline():
    M.get_krx_listing = lambda: pd.DataFrame({'Code': ['005930'], 'Name': ['테스트'], 'Stocks': [100]})
    M.get_stock_price = lambda t, d: (70000, d)
    M.get_outstanding_shares = lambda *a, **k: (1_000_000, 'KRX', {})
    M.fetch_pl_df = lambda d, c, y, r: (PL.copy(), 'CFS', 'OK')
    M.check_dart_reachable = lambda timeout=10: (True, None)
    M.get_dart_reader = lambda key: Dart()
    M.time.sleep = lambda s: None  # 티커당 0.5초 대기 생략


fails = []


def check(label, ok, detail=''):
    print(f"{'OK  ' if ok else '실패'}  {label}{('  — ' + str(detail)) if detail and not ok else ''}")
    if not ok:
        fails.append(label)


def expect_error(label, fn, fragment):
    try:
        fn()
    except Exception as e:
        check(label, fragment in str(e), str(e)[:120])
    else:
        check(label, False, '에러가 나야 하는데 성공했다')


# --- 1. 기간 조립 --------------------------------------------------------------
check('단일 기간', W._build_periods('2025.4Q') == ['2025.4Q'])
check('연도 걸친 범위', W._build_periods('2024.3Q', '2025.2Q') ==
      ['2024.3Q', '2024.4Q', '2025.1Q', '2025.2Q'])
expect_error('역순 입력 거부', lambda: W._build_periods('2025.4Q', '2025.1Q'), '빠릅니다')
expect_error('형식 오류 거부', lambda: W._build_periods('2025-4Q'), '형식')

# --- 2. 시작 전 검증 -----------------------------------------------------------
stub_pipeline()
expect_error('빈 티커 거부', lambda: W.run_gpcm([], '2025.4Q'), '비어')
expect_error('6자리 아닌 티커 거부', lambda: W.run_gpcm(['5930'], '2025.4Q'), '6자리')
expect_error('beta_type 검증', lambda: W.run_gpcm(['005930'], '2025.4Q', beta_type='10Y'), '5Y')

M.check_dart_reachable = lambda timeout=10: (False, 'timeout')
expect_error('DART 불통이면 시작 전에 거부', lambda: W.run_gpcm(['005930'], '2025.4Q'), '연결할 수 없습니다')
M.check_dart_reachable = lambda timeout=10: (True, None)

saved_key = os.environ.pop('DART_API_KEY')
expect_error('인증키 없으면 거부', lambda: W.run_gpcm(['005930'], '2025.4Q'), '인증키가 없습니다')
# 확장(.mcpb)에서 입력칸이 비면 플레이스홀더 리터럴이 그대로 온다 — 없음으로 취급해야 한다
os.environ['DART_API_KEY'] = '${user_config.dart_api_key}'
expect_error('플레이스홀더 리터럴도 없음 취급', lambda: W.run_gpcm(['005930'], '2025.4Q'), '인증키가 없습니다')
os.environ['DART_API_KEY'] = saved_key

# --- 3. 실행 → 완료 → 엑셀 ----------------------------------------------------
W.OUTPUT_DIR = Path(tempfile.mkdtemp()) / 'GPCM'

r = W.run_gpcm(['005930'], '2024.1Q')
check('시작 즉시 running 으로 돌아온다', r['state'] == 'running', r)
check('출력 경로를 미리 알려준다', r['output_file'].endswith('.xlsx'), r['output_file'])

job = W._jobs[r['job_id']]
job['thread'].join(timeout=60)

s = W.gpcm_status(r['job_id'])
check('완료 상태 전이', s['state'] == 'done', s.get('error', s['state']))
check('진행률 100', s['progress_pct'] == 100, s['progress_pct'])
check('Data_Quality 요약이 실린다', 'errors' in s.get('data_quality', {}), s.get('data_quality'))
check('Target WACC 가 실린다', isinstance(s.get('target_wacc_pct'), float), s.get('target_wacc_pct'))
check('Data_Quality 확인 안내문', 'Data_Quality' in s.get('note', ''))

out = Path(s['file'])
check('엑셀 파일이 실제로 쓰였다', out.exists(), s['file'])
wb = load_workbook(out)
check('GPCM·Data_Quality 시트 존재', {'GPCM', 'Data_Quality'} <= set(wb.sheetnames), wb.sheetnames)
check('Notes 의 Base Date 가 요청 기간', any(
    '2024.1Q' in str(c.value) for row in wb['GPCM'].iter_rows() for c in row if c.value))

# --- 3-b. GPCM 시트 열 배치 (우선주를 EV 구성요소로 옮긴 뒤의 등가성) ----------
gp = wb['GPCM']
hdr = [c.value for c in gp[5][:len(M.GPCM_COL_DEFS)]]
check('헤더 순서가 열 정의와 일치', hdr == [h for _, h, _ in M.GPCM_COL_DEFS], hdr[:14])
check('우선주(장부)가 11열(K) — NCI와 Equity 사이', hdr[10] == '우선주(장부)', hdr[8:13])
check('BS & EV Components 섹션이 F4:M4', 'F4:M4' in {str(rng) for rng in gp.merged_cells.ranges}
      and gp.cell(4, 6).value == 'BS & EV Components')
check('Equity Adj 섹션은 사라짐', all(gp.cell(4, c).value != 'Equity Adj' for c in range(1, 40)))
check('EV 수식이 새 열 문자를 참조 (U=MktCap, K=우선주)',
      gp.cell(6, 13).value == '=U6+K6+G6-F6+J6-H6', gp.cell(6, 13).value)
check('우선주 셀은 BS_Full의 Preferred 를 집계', 'Preferred' in (gp.cell(6, 11).value or ''), gp.cell(6, 11).value)
check('D/E 수식이 우선주 새 위치(K) 참조', 'K6' in (gp.cell(6, 33).value or ''), gp.cell(6, 33).value)
wacc_refs = {c.value for row in wb['WACC_Calculation'].iter_rows() for c in row
             if isinstance(c.value, str) and c.value.startswith('=GPCM!')}
check('WACC 시트의 GPCM 참조가 이동된 열(AI·AH)로', 
      any('=GPCM!AI' in v for v in wacc_refs) and any('=GPCM!AH' in v for v in wacc_refs), wacc_refs)
check('Cash 고정틀 유지', gp.freeze_panes == 'F6', gp.freeze_panes)

check('job_id 생략 시 최근 작업', W.gpcm_status()['job_id'] == r['job_id'])
expect_error('없는 job_id', lambda: W.gpcm_status('job-999'), '찾을 수 없습니다')

# --- 4. 실행 중 중복 시작 거부 -------------------------------------------------
import threading
gate = threading.Event()
M.fetch_financial_data_orig = M.fetch_financial_data
def slow_fetch(*a, **k):
    gate.wait(timeout=30)
    return M.fetch_financial_data_orig(*a, **k)
M.fetch_financial_data = slow_fetch
r2 = W.run_gpcm(['005930'], '2024.1Q')
expect_error('실행 중이면 새 작업 거부', lambda: W.run_gpcm(['005930'], '2024.1Q'), '이미 실행 중')
gate.set()
W._jobs[r2['job_id']]['thread'].join(timeout=60)
M.fetch_financial_data = M.fetch_financial_data_orig

# --- 5. 실패 경로 --------------------------------------------------------------
def boom(*a, **k):
    raise RuntimeError('의도한 폭발')
M.fetch_financial_data = boom
r3 = W.run_gpcm(['005930'], '2024.1Q')
W._jobs[r3['job_id']]['thread'].join(timeout=30)
s3 = W.gpcm_status(r3['job_id'])
check('실패 상태와 원인이 남는다', s3['state'] == 'failed' and '의도한 폭발' in s3['error'], s3.get('error', '')[:80])
M.fetch_financial_data = M.fetch_financial_data_orig

# --- 6. stdout 위생 (별도 프로세스에서 임포트만) --------------------------------
proc = subprocess.run(
    [sys.executable, '-c', 'import gpcm_mcp; gpcm_mcp._load()'],
    cwd=os.path.dirname(os.path.abspath(__file__)),
    env={**os.environ, 'DART_API_KEY': 'x' * 40},
    capture_output=True, timeout=300)
check('임포트·로드가 stdout 에 아무것도 안 쓴다', proc.stdout == b'',
      proc.stdout[:200])

# --- 7. 오늘 상장사 명단 (KRX 차단 환경 — 응답을 흉내 내 검증) --------------------
IND = pd.DataFrame([
    {'Code': '005930', 'Name': '삼성전자', 'Market': 'KOSPI', 'Sector': '반도체 제조업',
     'Industry': 'DRAM, NAND Flash', 'SettleMonth': '12월'},
    {'Code': '000660', 'Name': 'SK하이닉스', 'Market': 'KOSPI', 'Sector': '반도체 제조업',
     'Industry': 'DRAM', 'SettleMonth': '12월'},
    {'Code': '111111', 'Name': '삼월결산㈜', 'Market': 'KOSDAQ', 'Sector': '반도체 제조업',
     'Industry': '반도체 장비', 'SettleMonth': '3월'},
    {'Code': '222222', 'Name': '농심', 'Market': 'KOSPI', 'Sector': '음식료품',
     'Industry': '라면', 'SettleMonth': '12월'},
])
M.get_krx_industry_listing = lambda: IND.copy()
W._roster_memo.update(at=0.0, df=None)  # 메모 비우기

r = W.list_krx_companies()
check('빈 query 는 업종 집계만', 'sectors' in r and 'companies' not in r, r)
check('업종별 종목수가 맞는다',
      {s['sector']: s['count'] for s in r['sectors']} == {'반도체 제조업': 3, '음식료품': 1}, r['sectors'])

r = W.list_krx_companies('반도체')
check('업종명 매칭 3개 (장비 주요제품 포함)', r['meta']['count'] == 3, r['meta'])
check('주요제품이 실린다', r['companies'][1]['product'] == 'DRAM, NAND Flash', r['companies'][1])
check('3월 결산 플래그', [c['fiscalMonthNot12'] for c in r['companies']] == [False, False, True])

r = W.list_krx_companies('반도체', december_only=True)
check('12월 결산 필터', r['meta']['count'] == 2, r['meta'])

r = W.list_krx_companies('라면')
check('주요제품으로도 찾는다', r['meta']['count'] == 1 and r['companies'][0]['name'] == '농심', r)

W._roster_memo.update(at=0.0, df=None)
M.get_krx_industry_listing = lambda: pd.DataFrame()
expect_error('KRX 실패는 명확한 에러 (조용한 폴백 없음)', lambda: W.list_krx_companies('반도체'), 'KRX')
W._roster_memo.update(at=0.0, df=None)

# --- 8. 거래정지 이력 점검 (fdr 응답을 흉내 내 검증) -----------------------------
days = pd.bdate_range('2024-01-01', periods=100)
def _px(idx): 
    return pd.DataFrame({'Close': [100.0]*len(idx)}, index=idx)

class FakeFdr:
    def DataReader(self, symbol, start=None, end=None):
        if symbol == 'KS11':                 return _px(days)
        if symbol == '000001':               return _px(days)                          # 정상
        if symbol == '000002':               return _px(days[:40].append(days[50:]))   # 중간 10거래일 결측
        if symbol == '000003':               return _px(days[50:])                     # 늦은 상장
        if symbol == '000004':               return _px(days[:-6])                     # 최근 6거래일 결측
        if symbol == '000005':               return pd.DataFrame()                     # 조회 실패
        raise RuntimeError(symbol)

M.fdr = FakeFdr()
g = W.check_trading_gaps(['000001', '000002', '000003', '000004', '000005'])
by = {x['code']: x for x in g['results']}
check('정상 종목은 무표시', by['000001']['flag'] is False and not by['000001']['suspectedHalts'])
check('중간 10거래일 결측을 정지 의심으로 잡는다',
      by['000002']['flag'] and by['000002']['suspectedHalts'][0]['tradingDays'] == 10, by['000002'])
check('늦은 상장은 정지로 오인하지 않는다', by['000003']['flag'] is False
      and by['000003']['observedFrom'] == days[50].strftime('%Y-%m-%d'), by['000003'])
check('최근 결측은 현재 정지 중으로 표시', by['000004']['currentlySuspended'] is True, by['000004'])
check('조회 실패는 failed 로 드러난다', g.get('failed') == ['000005'], g.get('failed'))
check('자동 배제 금지 안내가 실린다', '자동 배제' in g['meta']['note'])

print()
print(f"잘못된 항목 {len(fails)}건")
sys.exit(1 if fails else 0)
