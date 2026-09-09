"""GPCM.py 회귀 테스트 — Price_History 시트가 끝까지 만들어지는지.

일별 주가 시트를 회사별 블록으로 새로 짜면서(v0.8.2) 옛 레이아웃의 열 위치
변수 `rel_start` 가 사라졌는데, 차트를 붙이는 마지막 줄만 그 이름을 계속
참조했다. 배포된 앱이 계산을 다 끝낸 뒤 NameError 로 죽었다.

이 줄은 일별 주가 구간이 켜져 있을 때만 지나가므로, 시트를 실제로 만들어
봐야 걸린다. 여기서는 가짜 주가·재무 데이터로 create_excel 을 끝까지 태운다.
네트워크는 쓰지 않는다.

    python test_gpcm_global.py
"""
import io
import logging
import os
import sys

logging.disable(logging.WARNING)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import openpyxl
import pandas as pd

import GPCM as M

NAMES = {'AAA': 'Alpha Corp', 'BBB': 'Beta Ltd', 'CCC': 'Gamma Inc'}
TARGET_WACC = {
    'Rf': .043, 'MRP': .055, 'Size_Premium': .0182, 'Avg_Unlevered_Beta': .92,
    'Beta_Basis': '5Y Monthly', 'Beta_Label': '5Y Monthly Adj',
    'Target_Tax_Rate': .21, 'Avg_Debt_Ratio': .23,
    'Debt_Ratio_Basis': '피어 3개 평균', 'Target_DE_Ratio': .30,
    'Target_Relevered_Beta': 1.14, 'Target_Ke': .124, 'Kd_Pretax': .05,
    'Target_Kd_Aftertax': .0395, 'Equity_Weight': .77, 'Debt_Weight': .23,
    'Target_WACC': .1052, 'Outlier_Method': 'none',
}


def fake_gpcm(ticker, i):
    g = M._empty_gpcm(NAMES[ticker], ticker, 'USD', 'NMS', '^GSPC',
                      '2024-12-31', 'Annual')
    g.update({
        'BS_Date': '2024-12-31', 'Cash': 100 + i, 'IBD': 300 + i, 'NCI': 10,
        'Equity': 1000 + i, 'Revenue': 5000 + i, 'EBIT': 600 + i,
        'EBITDA': 800 + i, 'DA': 200, 'NI_Parent': 400 + i,
        'Close': 50 + i, 'Shares': 1e8, 'Market_Cap_M': 5000 + i,
        'Equity_Value_M': 5000 + i,
        'Beta_5Y_Monthly_Raw': 1.1, 'Beta_5Y_Monthly_Adj': 1.07,
        'Beta_5Y_R2': .5, 'Beta_5Y_StdErr': .1, 'Beta_5Y_N': 60,
        'Beta_2Y_Weekly_Raw': 1.2, 'Beta_2Y_Weekly_Adj': 1.13,
        'Beta_2Y_R2': .4, 'Beta_2Y_StdErr': .1, 'Beta_2Y_N': 104,
        'Pretax_Income': 500, 'DE_Ratio': .3, 'Debt_Ratio': .23,
        'Unlevered_Beta_5Y': .9, 'Unlevered_Beta_2Y': .95,
    })
    return g


def fake_prices(tickers, empty_for=()):
    """종목별 일별 종가. empty_for 에 든 종목은 데이터가 없는 회사를 흉내낸다."""
    idx = pd.bdate_range('2020-01-01', '2024-12-31')
    abs_series = []
    for i, t in enumerate(tickers):
        if t in empty_for:
            abs_series.append(pd.Series(np.nan, index=idx, name=t))
            continue
        abs_series.append(
            pd.Series(np.linspace(40 + i * 5, 60 + i * 5, len(idx)), index=idx, name=t))
    df_abs = pd.concat(abs_series, axis=1).sort_index()
    common = df_abs.dropna()
    base = common.iloc[0] if not common.empty else df_abs.ffill().bfill().iloc[0]
    df_rel = (df_abs / base) * 100
    rel_series = [df_rel[c].dropna().rename(c) for c in df_rel.columns]
    return abs_series, rel_series


def build_workbook(tickers, span='5Y', empty_for=()):
    p_abs, p_rel = fake_prices(tickers, empty_for)
    data = M.create_excel(
        {'Y': {t: fake_gpcm(t, i) for i, t in enumerate(tickers)}},
        [], [], [], p_abs, p_rel,
        '2024-12-31', [{'label': 'Y', 'year': 2024, 'qtr': '4Q'}],
        {t: NAMES[t] for t in tickers}, TARGET_WACC,
        beta_type='5Y', data_quality_rows=[], target_inputs={},
        daily_price_span=span)
    raw = data.getvalue() if hasattr(data, 'getvalue') else data
    return openpyxl.load_workbook(io.BytesIO(raw))


def check(name, fn):
    try:
        detail = fn()
    except Exception as exc:
        return name, False, f'{type(exc).__name__}: {exc}'
    return name, True, detail


def three_companies():
    wb = build_workbook(['AAA', 'BBB', 'CCC'])
    ws = wb['Price_History']
    assert len(ws._charts) == 1, f'차트 {len(ws._charts)}개'
    col = ws._charts[0].anchor._from.col + 1          # openpyxl 은 0부터 센다
    # 표 3개가 1~12열을 쓴다. 차트는 그보다 오른쪽이어야 겹치지 않는다.
    assert col > 3 * 4, f'차트가 {col}열 — 마지막 표(12열)와 겹친다'
    return f'차트 1개, {col}열'


def one_company():
    wb = build_workbook(['AAA'])
    ws = wb['Price_History']
    col = ws._charts[0].anchor._from.col + 1
    assert col > 4, f'차트가 {col}열 — 표(4열)와 겹친다'
    return f'차트 1개, {col}열'


def company_without_prices():
    """가운데 종목에 주가가 없어도(블록을 건너뛰어도) 열 계산이 맞아야 한다."""
    wb = build_workbook(['AAA', 'BBB', 'CCC'], empty_for=('BBB',))
    ws = wb['Price_History']
    col = ws._charts[0].anchor._from.col + 1
    assert col > 3 * 4, f'차트가 {col}열 — 건너뛴 종목의 자리도 비워 두므로 겹친다'
    return f'차트 1개, {col}열'


def span_off():
    wb = build_workbook(['AAA', 'BBB'], span='off')
    assert 'Price_History' not in wb.sheetnames, '구간을 껐는데 시트가 생겼다'
    return '시트 없음'


CASES = [
    ('3개사 — Price_History 생성', three_companies),
    ('1개사', one_company),
    ('주가 없는 종목 포함', company_without_prices),
    ('일별 주가 구간 off', span_off),
]

if __name__ == '__main__':
    print(f"{'경우':<30}{'결과':<8}{'내용'}")
    print('-' * 70)
    fails = 0
    for name, fn in CASES:
        _, ok, detail = check(name, fn)
        if not ok:
            fails += 1
        print(f'{name:<30}{"OK" if ok else "실패":<8}{detail}')
    print('-' * 70)
    print('전부 통과' if not fails else f'{fails}건 실패')
    sys.exit(1 if fails else 0)
