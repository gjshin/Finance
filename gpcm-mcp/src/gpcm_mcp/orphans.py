"""화면 코드 안에 섞여 있던 계산.

gpcm_kr.py 의 모드 1 은 함수 세 개로 끝나지 않는다. `if run_btn:` 블록과 사이드바
안에 export_gpcm_excel 로 곧장 들어가는 계산이 네 토막 더 있다. UI 처럼 생겼지만
UI 가 아니다 — 빠뜨리면 워크북이 달라진다.

| 원본 위치 | 무엇 |
|---|---|
| L2377-2388 | target_periods 구성 |
| L2549-2569 | notes_list (GPCM 시트 하단 주석) |
| L2645-2661 | 기준기간 수집 실패 진단 → quality 에 추가 |
| L2643, L2672-2676 | df_screen 과 EV/배수 파생 |

순서 함정 하나: 진단 결과는 fetch_financial_data 가 끝난 **뒤**, export 가 불리기
**전**에 quality 에 붙는다. 이 순서가 Data_Quality 시트의 행 순서를 정한다.
"""

import numpy as np
import pandas as pd

from .constants import SEV_ERROR, SEV_WARN
from .periods import get_base_date_str, parse_period

QUARTERS = ["1Q", "2Q", "3Q", "4Q"]


def build_target_periods(start_year, start_qtr, end_year, end_qtr, cycle='quarterly'):
    """'YYYY.NQ' 목록을 만든다. 원본 L2377-2388 과 같은 규칙.

    연간 모드는 각 연도의 4Q 만 넣는다 (사업보고서 기준). 시작·종료 모두 포함이다.
    마지막 항목이 기준 기간이 된다.
    """
    target_periods = []
    qtrs = QUARTERS
    for y in range(start_year, end_year + 1):
        if cycle == 'annual':
            target_periods.append(f"{y}.4Q")
        else:
            s_idx = qtrs.index(start_qtr) if y == start_year else 0
            e_idx = qtrs.index(end_qtr) if y == end_year else 3
            for q_idx in range(s_idx, e_idx + 1):
                target_periods.append(f"{y}.{qtrs[q_idx]}")
    return target_periods


def build_periods_to_fetch(start_year, end_year, start_qtr=None, end_qtr=None):
    """모드 2 용 {year, qtr, label} 목록. 원본 L2504-2530 과 같은 규칙.

    qtr 을 주지 않으면 연간(사업보고서)만 조회한다.
    label 은 그냥 표시용이 아니다 — df_summ['Period'] 이자 엑셀 열 머리글이고,
    Summary 시트의 SUMIFS 가 이 문자열로 회사별 시트를 찾는다.
    """
    periods = []
    if start_qtr is None and end_qtr is None:
        for y in range(start_year, end_year + 1):
            periods.append({'year': y, 'qtr': None, 'label': f"{y}년"})
        return periods

    start_qtr = start_qtr or '1Q'
    end_qtr = end_qtr or '4Q'
    for y in range(start_year, end_year + 1):
        s_idx = QUARTERS.index(start_qtr) if y == start_year else 0
        e_idx = QUARTERS.index(end_qtr) if y == end_year else 3
        for q_idx in range(s_idx, e_idx + 1):
            periods.append({'year': y, 'qtr': QUARTERS[q_idx],
                            'label': f"{y}년 {QUARTERS[q_idx]}"})
    return periods


def build_notes(base_period_str):
    """GPCM 시트 아래에 붙는 방법론 주석. 원본 L2549-2569 그대로.

    조서에 배수를 옮기는 사람이 이 목록을 읽고 EV 정의와 LTM 규칙을 확인한다.
    문구를 바꾸면 워크북이 달라지므로 원본과 함께만 바꾼다.
    """
    base_year, base_qtr = parse_period(base_period_str)
    base_date_display = get_base_date_str(base_year, base_qtr)
    return [
        f'• Base Date: {base_period_str} ({base_date_display}) | Unit: 억원 (KRW 100M)',
        '• 공통: 연결재무제표 작성 시 CFS 우선, 미존재 시 OFS 기준으로 수집',
        '• PL: 요약 손익계산서에서 매출액/영업이익/당기순이익 3개 계정만 엄격 추출',
        '• PL Fetch: finstate(요약) → finstate_all(CFS/OFS) fallback',
        '• Shares: DART(stockTotqySttus) 유통주식수(distb_stock_co) 우선, 미공시 시 DART 과거보고서 fallback',
        '• EV = Market Cap + 우선주(장부) + IBD − Cash + NCI − NOA',
        '• Net Debt = IBD − Cash − NOA',
        '• IBD(Option): CB/EB/BW 등 메자닌은 기본적으로 IBD(Option)으로 태깅되어 EV/NetDebt에서 제외됨',
        '• NOA(Option): 투자자산/관계기업 등은 기본적으로 NOA(Option)으로 태깅되어 EV/NetDebt에서 제외됨',
        '• LTM = Current Cumulative + Prior Annual − Prior Same Quarter Cumulative (단, 4Q는 Annual)',
        '• Beta: 5년 월간 & 2년 주간 수익률 기준 (FinanceDataReader 사용)',
        '• Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1',
        '• D/E Ratio = IBD / (Market Cap + 우선주 + NCI)',
        '• Debt Ratio (D/V) = IBD / (Market Cap + 우선주 + IBD + NCI)',
        '• 우선주: BS의 우선주자본금(액면) 기준. 시가총액은 보통주만 반영하므로 자기자본가치에 가산',
        '• Unlevered Beta = Levered Beta / (1 + (1 - Tax Rate) × D/E Ratio)',
        '• Tax Rate: 한국 법인세 한계세율 (지방소득세 포함, 세전순이익 기준, 사업연도별 세율표 적용)',
        '   - FY2023~2025: 2억 이하 9.9% | 2~200억 20.9% | 200~3,000억 23.1% | 3,000억 초과 26.4%',
        '   - FY2026~    : 2억 이하 11.0% | 2~200억 22.0% | 200~3,000억 24.2% | 3,000억 초과 27.5% (2025년 세법개정)',
    ]


def diagnose_base_period(all_multiples, base_period_str, quality):
    """기준 기간을 제대로 수집했는지 본다. 원본 L2645-2661 그대로.

    조용히 0 이 남는 것을 막는 장치다. 시가총액·매출·자본 중 하나라도 비면
    배수와 WACC 가 통째로 왜곡되는데, 오류는 나지 않는다.

    quality 를 **여기서** 채우는 순서가 중요하다 — 이 행들이 수집 단계의 행 뒤에
    붙어야 Data_Quality 시트가 원본과 같은 순서로 나온다.
    """
    problems = []
    base_rows = [m for m in all_multiples if m['Period'] == base_period_str]
    if not base_rows:
        problems.append(f"기준 기간({base_period_str}) 데이터를 한 건도 수집하지 못했습니다.")
    for m in base_rows:
        empty_fields = [k for k in ('Market_Cap', 'Revenue', 'Equity') if not m.get(k)]
        if empty_fields:
            problems.append(f"[{m['Company']}] {base_period_str} — {', '.join(empty_fields)} 값이 없습니다.")
    if problems:
        for p in problems:
            quality.add(SEV_ERROR, '', '', f'기준기간 {base_period_str}', p)
    return problems


def build_screen_frame(all_multiples):
    """배수 표. 원본 L2643 + L2672-2676 그대로.

    export_gpcm_excel 은 이 프레임을 (회사 × 기간) 행 목록으로만 쓰지만,
    여기서 만든 EV·EV/EBIT·PER·PSR 은 요약 응답의 숫자가 된다. 워크북 쪽 값은
    엑셀 수식이라 열기 전에는 값이 없다.
    """
    df_screen = pd.DataFrame(all_multiples)
    if not df_screen.empty:
        df_screen['EV'] = df_screen['Market_Cap'] + df_screen['Preferred'] + df_screen['IBD'] - df_screen['Cash'] + df_screen['NCI'] - df_screen['NOA']
        df_screen['EV/EBIT'] = np.where(df_screen['EBIT'] > 0, df_screen['EV'] / df_screen['EBIT'], np.nan)
        df_screen['PER'] = np.where(df_screen['NI'] > 0, df_screen['Market_Cap'] / df_screen['NI'], np.nan)
        df_screen['PSR'] = np.where(df_screen['Revenue'] > 0, df_screen['Market_Cap'] / df_screen['Revenue'], np.nan)
    return df_screen


def count_quality(quality):
    """레벨별 건수. 원본은 화면에 띄웠지만(L2664-2670) 여기서는 응답에 싣는다."""
    return {
        'error': sum(1 for r in quality.rows if r['Level'] == SEV_ERROR),
        'warn': sum(1 for r in quality.rows if r['Level'] == SEV_WARN),
    }
