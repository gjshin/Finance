"""기간·기준일 계산. gpcm_kr.py 에서 그대로 옮겼다."""

from datetime import datetime, timedelta

import pandas as pd

from .constants import QUARTER_INFO


def parse_period(p: str):
    parts = p.strip().split('.')
    return int(parts[0]), parts[1]

def get_base_date_str(year: int, qtr: str):
    return f"{year}-{QUARTER_INFO[qtr]}"

# 정기보고서 법정 제출기한 (결산일로부터 경과일수)
FILING_DEADLINE_DAYS = {'1Q': 45, '2Q': 45, '3Q': 45, '4Q': 90}

def is_period_filed(year: int, qtr: str, asof: datetime = None):
    """해당 분기 정기보고서의 제출기한이 지났는지 여부 (= 조회 가능한지)"""
    asof = asof or datetime.now()
    qtr_end = pd.to_datetime(f"{year}-{QUARTER_INFO[qtr]}")
    return (qtr_end + timedelta(days=FILING_DEADLINE_DAYS[qtr])) <= asof

def get_latest_filed_period(asof: datetime = None):
    """오늘 기준으로 실제 공시가 끝난 가장 최근 분기를 (연도, 분기)로 반환"""
    asof = asof or datetime.now()
    y, q_idx = asof.year, 3
    for _ in range(12):
        qtr = ['1Q', '2Q', '3Q', '4Q'][q_idx]
        if is_period_filed(y, qtr, asof):
            return y, qtr
        q_idx -= 1
        if q_idx < 0:
            y -= 1
            q_idx = 3
    return asof.year - 1, '4Q'

def get_ltm_required_periods(year: int, qtr: str):
    if qtr == '4Q':
        return [(year, '4Q', 'annual')]
    return [
        (year, qtr, 'current_cum'),
        (year - 1, '4Q', 'prior_annual'),
        (year - 1, qtr, 'prior_same_q'),
    ]
