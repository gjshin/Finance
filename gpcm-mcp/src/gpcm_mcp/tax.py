"""한국 법인세 한계세율과 하마다 언레버링. gpcm_kr.py 에서 그대로 옮겼다."""

from datetime import datetime

import pandas as pd


# 한국 법인세 한계세율표 (사업연도별, 지방소득세 10% 포함)
# 각 구간: (과세표준 상한(억원), 한계세율)  — 상한 None = 초과 구간
# · FY2018~2022 : 국세 10 / 20 / 22 / 25%
# · FY2023~2025 : 국세  9 / 19 / 21 / 24%  (2022년 세법개정, 1%p 인하)
# · FY2026~     : 국세 10 / 20 / 22 / 25%  (2025년 세법개정, 1%p 인상 환원)
KR_TAX_BRACKETS_PRE2023 = [(2, 0.110), (200, 0.220), (3000, 0.242), (None, 0.275)]
KR_TAX_BRACKETS_2023 = [(2, 0.099), (200, 0.209), (3000, 0.231), (None, 0.264)]
KR_TAX_BRACKETS_2026 = [(2, 0.110), (200, 0.220), (3000, 0.242), (None, 0.275)]


def get_korean_tax_brackets(fiscal_year):
    """사업연도에 적용되는 한국 법인세 한계세율표 반환."""
    try:
        fy = int(fiscal_year)
    except (TypeError, ValueError):
        fy = datetime.now().year
    if fy >= 2026:
        return KR_TAX_BRACKETS_2026
    if fy >= 2023:
        return KR_TAX_BRACKETS_2023
    return KR_TAX_BRACKETS_PRE2023


def get_korean_marginal_tax_rate(pretax_income_100m, fiscal_year=None):
    """
    한국 법인세 한계세율 산출 (지방소득세 포함, 사업연도별 세율표 적용)

    Parameters:
    - pretax_income_100m: 세전이익 (억원). 과세표준의 대용치로 사용.
    - fiscal_year: 해당 재무제표의 결산 연도. 미지정 시 현재 연도 기준.

    Note: 결손(음수) 기업은 한계세율 개념이 성립하지 않으므로
          '2억 초과 ~ 200억' 구간 세율을 적용한다.
    """
    brackets = get_korean_tax_brackets(fiscal_year)

    if pd.isna(pretax_income_100m) or pretax_income_100m <= 0:
        return brackets[1][1]

    for upper, rate in brackets:
        if upper is None or pretax_income_100m <= upper:
            return rate
    return brackets[-1][1]

def calculate_unlevered_beta(levered_beta, debt, equity, tax_rate):
    """
    하마다 모형으로 Unlevered Beta 계산
    Unlevered Beta = Levered Beta / (1 + (1 - Tax Rate) * (Debt / Equity))
    """
    if pd.isna(levered_beta) or levered_beta is None:
        return None
    if pd.isna(debt) or pd.isna(equity) or equity == 0:
        return levered_beta

    unlevered = levered_beta / (1 + (1 - tax_rate) * (debt / equity))
    return unlevered
