"""계산 결과를 Claude 가 읽을 수 있는 형태로 줄인다.

두 가지를 지켜야 한다.

1. **워크북에서 숫자를 읽어오면 안 된다.** GPCM·LTM_Calc·Multiples_Trend 의
   밸류에이션 수치는 전부 엑셀 수식이고 openpyxl 은 계산 결과를 저장하지 않는다.
   엑셀로 열기 전에는 값이 없다. 그래서 요약은 파이썬이 들고 있는 값으로 만든다.

2. **JSON 으로 나갈 수 있어야 한다.** screen_summary_data 에는 주가 시계열
   (pandas Series) 이 들어 있어서 그대로 내보내면 전송 단계에서 죽는다.
   여기서 스칼라만 뽑아낸다. NaN 도 null 로 바꾼다 — JSON 에는 NaN 이 없다.
"""

import math

import numpy as np
import pandas as pd

MULTIPLE_KEYS = ('EV/EBIT', 'PER', 'PSR')


def jsonable(value):
    """numpy·pandas 타입과 NaN 을 표준 파이썬 값으로 바꾼다."""
    if value is None:
        return None
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        number = float(value)
        return None if math.isnan(number) or math.isinf(number) else number
    if isinstance(value, (np.bool_, bool)):
        return bool(value)
    if isinstance(value, dict):
        return {str(k): jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [jsonable(v) for v in value]
    if value is pd.NaT or (isinstance(value, float) and pd.isna(value)):
        return None
    return value


def multiples_at(df_screen, base_period_str):
    """기준 기간의 회사별 배수. df_screen 은 orphans.build_screen_frame 이 만든 것."""
    if df_screen is None or df_screen.empty:
        return []
    rows = df_screen[df_screen['Period'] == base_period_str]
    keep = ['Company', 'Ticker', 'Market_Cap', 'EV', 'Revenue', 'EBIT', 'NI',
            'EV/EBIT', 'PER', 'PSR']
    out = []
    for _, row in rows.iterrows():
        out.append({k: jsonable(row[k]) for k in keep if k in rows.columns})
    return out


def statistics(df_screen, base_period_str):
    """기준 기간 배수의 평균·중앙값·최대·최소.

    화면 미리보기(원본 L2681-2688)가 보여주던 것과 같은 것을 응답에 싣는다.
    """
    if df_screen is None or df_screen.empty:
        return {}
    rows = df_screen[df_screen['Period'] == base_period_str]
    if rows.empty:
        return {}
    stats = {}
    for key in MULTIPLE_KEYS:
        if key not in rows.columns:
            continue
        series = pd.to_numeric(rows[key], errors='coerce').dropna()
        if series.empty:
            stats[key] = {'mean': None, 'median': None, 'max': None, 'min': None,
                          'count': 0}
            continue
        stats[key] = {
            'mean': jsonable(series.mean()),
            'median': jsonable(series.median()),
            'max': jsonable(series.max()),
            'min': jsonable(series.min()),
            'count': int(series.count()),
        }
    return stats


def quality_report(quality, limit=20):
    """품질 기록. 건수만 주면 Claude 가 무시하므로 실제 문장을 함께 싣는다.

    ERROR 와 WARN 을 앞에 둔다 — 사용자에게 가장 먼저 말해야 하는 내용이다.
    """
    order = {'ERROR': 0, 'WARN': 1, 'INFO': 2}
    rows = sorted(quality.rows, key=lambda r: order.get(r['Level'], 9))
    counts = {}
    for row in quality.rows:
        counts[row['Level'].lower()] = counts.get(row['Level'].lower(), 0) + 1
    return {
        'counts': counts,
        'total': len(quality.rows),
        'shown': min(limit, len(rows)),
        'rows': [dict(r) for r in rows[:limit]],
    }


def historical_summary(df_summ, limit=200):
    """모드 2 요약. 회사 × 기간 한 줄씩."""
    if df_summ is None or df_summ.empty:
        return []
    keep = ['Company', 'Ticker', 'Period', 'Report', 'Revenue', 'GrossProfit',
            'EBIT', 'NI', 'Assets', 'Liabilities', 'Equity_Total',
            'CFO', 'CFI', 'CFF', 'NetDebt', 'OPM', 'GPM', 'ROE', 'DebtRatio']
    cols = [c for c in keep if c in df_summ.columns]
    return [{c: jsonable(row[c]) for c in cols}
            for _, row in df_summ.head(limit).iterrows()]
