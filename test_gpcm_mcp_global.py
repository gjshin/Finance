"""run_gpcm_global(해외 상장사 — 야후파이낸스) 회귀 테스트. 네트워크를 쓰지 않는다.

GPCM.fetch_ticker_bundle 만 가짜로 바꾸고, 그 뒤(get_gpcm_data 의 WACC 계산 → create_excel →
엑셀 저장 → gpcm_status)는 진짜 코드를 그대로 태운다. 확인하는 것:
  - 입력 검증(빈 목록·형식·중복·기준 기간·베타 종류)
  - 기간 조립이 앱 사이드바와 같은지
  - 백그라운드 실행 → 엑셀 생성 → 상태 보고(WACC·Data_Quality·자료 못 받은 티커)
  - 사용자 입력 티커 출처·금리 출처가 Data_Quality 에 남는지
  - 실행 중 GPCM.st 를 잠시 바꾼 것이 끝나면 원래대로 돌아오는지

    uv run --no-project --with-requirements requirements-mcp.txt python test_gpcm_mcp_global.py
"""
import os
import sys
import time
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("DART_API_KEY", "TESTKEY-" + "0" * 32)

import numpy as np
import pandas as pd
from openpyxl import load_workbook

import gpcm_mcp as W

G = W._load_global()
REAL_ST = G.st

NAMES = {"AAA": "Alpha Corp", "7273.T": "Ikuyo", "MG.TO": "Magna", "200880.KS": "Seoyon E-Hwa"}

fails = []


def check(label, ok, detail=""):
    print(f"{'OK  ' if ok else '실패'}  {label}{('  — ' + str(detail)) if detail and not ok else ''}")
    if not ok:
        fails.append(label)


def expect_error(label, fn, needle):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        check(label, needle in str(e), str(e))
        return
    check(label, False, "예외가 나지 않음")


def fake_bundle(ticker, base_period_str, lookback, force_annual, include_recent, beta_basis="KOSPI"):
    if ticker == "BAD.X":
        return {"ticker": ticker, "ok": False, "error": "no data", "company": ticker, "periods": {},
                "bs_rows": [], "pl_rows": [], "market_rows": [], "price_series": None,
                "flags": [{"Ticker": ticker, "Label": "Y", "Level": G.SEV_ERROR, "Item": "수집",
                           "Message": "야후에 자료 없음"}]}
    i = list(NAMES).index(ticker)
    g = G._empty_gpcm(NAMES[ticker], ticker, "USD", "NMS", "^GSPC", base_period_str, "LTM")
    g.update({
        "BS_Date": base_period_str, "PL_Date": base_period_str, "Cash": 100 + i, "IBD": 300 + i, "NCI": 10,
        "Equity": 1000 + i, "Revenue": 5000 + i, "EBIT": 600 + i, "EBITDA": 800 + i, "DA": 200,
        "NI_Parent": 400 + i, "Close": 50 + i, "Shares": 1e8, "Market_Cap_M": 5000 + i,
        "Equity_Value_M": 5000 + i,
        "Beta_5Y_Monthly_Raw": 1.1, "Beta_5Y_Monthly_Adj": 1.07, "Beta_5Y_R2": .5, "Beta_5Y_StdErr": .1,
        "Beta_5Y_N": 60, "Beta_2Y_Weekly_Raw": 1.2, "Beta_2Y_Weekly_Adj": 1.13, "Beta_2Y_R2": .4,
        "Beta_2Y_StdErr": .1, "Beta_2Y_N": 104, "Pretax_Income": 500, "DE_Ratio": .3, "Debt_Ratio": .2 + i / 100,
        "Unlevered_Beta_5Y": .9 + i / 100, "Unlevered_Beta_2Y": .95,
    })
    idx = pd.bdate_range("2021-07-01", base_period_str)
    s = pd.Series(np.linspace(40 + i, 60 + i, len(idx)), index=idx, name=ticker)
    flags = []
    if ticker == "MG.TO":
        flags.append({"Ticker": ticker, "Label": "Y", "Level": G.SEV_WARN, "Item": "EBITDA",
                      "Message": "EBITDA 항목 없음 — 감가상각비 0 처리"})
    return {"ticker": ticker, "ok": True, "error": None, "company": NAMES[ticker], "currency": "USD",
            "exchange": "NMS", "market_index": "^GSPC", "periods": {"Y": g}, "bs_rows": [], "pl_rows": [],
            "market_rows": [], "price_series": s, "flags": flags}


G.fetch_ticker_bundle = fake_bundle
tmp = Path(tempfile.mkdtemp())
W.OUTPUT_DIR = tmp

print("1. 입력 검증")
expect_error("빈 티커 거부", lambda: W.run_gpcm_global([], "2026.2Q"), "비어")
expect_error("티커 형식 거부(고치지 않음)", lambda: W.run_gpcm_global(["7273 T"], "2026.2Q"), "형식이 아닙니다")
expect_error("같은 티커 두 번 거부", lambda: W.run_gpcm_global(["AAA", "AAA"], "2026.2Q"), "두 번")
expect_error("기준 기간 형식", lambda: W.run_gpcm_global(["AAA"], "2026-06-30"), "2026.2Q")
expect_error("beta_type 검증", lambda: W.run_gpcm_global(["AAA"], "2026.2Q", beta_type="10Y"), "5Y")
expect_error("이상치 방식 검증", lambda: W.run_gpcm_global(["AAA"], "2026.2Q", outlier_method="zz"), "iqr")
expect_error("대상 숫자 키 검증", lambda: W.run_gpcm_global(["AAA"], "2026.2Q", target_financials={"Sales": 1}), "Revenue")

print("2. 기간 조립(앱 사이드바와 같음)")
check("분기 기준 + 직전 연말 1개", W._global_periods("2026.2Q", 1) == ["2025-12-31", "2026-06-30"], W._global_periods("2026.2Q", 1))
check("연말 기준 3년", W._global_periods("2025.4Q", 3) == ["2023-12-31", "2024-12-31", "2025-12-31"], W._global_periods("2025.4Q", 3))


def wait(job_id, sec=120):
    end = time.time() + sec
    while time.time() < end:
        st = W.gpcm_status(job_id)
        if st["state"] != "running":
            return st
        time.sleep(0.2)
    return W.gpcm_status(job_id)


print("3. 실행 → 엑셀 → 상태")
tickers = ["AAA", "7273.T", "MG.TO", "200880.KS", "BAD.X"]
r = W.run_gpcm_global(tickers, "2026.2Q", rf=4.0, mrp=8.0, size_premium=0.97, kd_pretax=5.0, tax_rate=24.2,
                      beta_type="5Y", rate_source="국고채 10년 4.000% (가상 예시)",
                      target_financials={"Revenue": 1000, "Shares": 1e6})
check("바로 돌아옴(백그라운드)", r["state"] == "running" and r["kind"].startswith("global"), r)
st = wait(r["job_id"])
check("완료", st["state"] == "done", st.get("error"))
if st["state"] == "done":
    f = Path(st["file"])
    check("엑셀 저장", f.exists() and f.name.startswith("Global_GPCM_20260630_"), f)
    check("자료 못 받은 티커 보고", st.get("failed_tickers") == ["BAD.X"], st.get("failed_tickers"))
    check("WACC 보고", isinstance(st.get("target_wacc_pct"), float) and 0 < st["target_wacc_pct"] < 30, st.get("target_wacc_pct"))
    check("WACC 구성 보고", {"ke_pct", "avg_unlevered_beta", "peer_avg_debt_ratio_pct"} <= set(st.get("wacc_detail", {})), st.get("wacc_detail"))
    check("EBITDA 빠진 회사 보고", st["data_quality"].get("ebitda_missing") == ["MG.TO"], st["data_quality"])
    check("해외판 안내문", "PL Date" in st["note"], st["note"])
    wb = load_workbook(f)
    check("주요 시트", {"GPCM", "Data_Quality", "WACC_Calculation"} <= set(wb.sheetnames)
          or {"GPCM", "Data_Quality"} <= set(wb.sheetnames), wb.sheetnames)
    dq = wb["Data_Quality"]
    text = " ".join(str(c.value) for row in dq.iter_rows() for c in row if c.value is not None)
    check("Data_Quality: 사용자 입력 티커 출처", "사용자가 사전에 검토해 입력한 티커" in text)
    check("Data_Quality: 금리 출처", "국고채 10년 4.000%" in text)
    check("Implied_Valuation 에 대상 숫자", "Implied_Valuation" in wb.sheetnames, wb.sheetnames)
check("GPCM.st 원래대로", G.st is REAL_ST)

print("4. 동시 실행 막기")
W._jobs["job-x"] = {"id": "job-x", "state": "running"}
expect_error("실행 중이면 거부", lambda: W.run_gpcm_global(["AAA"], "2026.2Q"), "이미 실행 중")
W._jobs.pop("job-x")

print("5. 도구 등록")
check("run_gpcm_global 등록", "run_gpcm_global" in W._registered_tools(), W._registered_tools())

print()
print("잘못된 항목 0건" if not fails else f"잘못된 항목 {len(fails)}건: {fails}")
sys.exit(1 if fails else 0)
