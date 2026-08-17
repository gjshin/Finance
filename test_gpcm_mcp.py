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

M = W.M

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
    [sys.executable, '-c', 'import gpcm_mcp'],
    cwd=os.path.dirname(os.path.abspath(__file__)),
    env={**os.environ, 'DART_API_KEY': 'x' * 40},
    capture_output=True, timeout=300)
check('임포트가 stdout 에 아무것도 안 쓴다', proc.stdout == b'',
      proc.stdout[:200])

print()
print(f"잘못된 항목 {len(fails)}건")
sys.exit(1 if fails else 0)
