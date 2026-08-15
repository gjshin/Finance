"""계산·엑셀·요약을 이어 붙인다.

서버(server.py)와 계산 계층 사이의 유일한 이음매다. 원본 gpcm_kr.py 의
`if run_btn:` 블록이 하던 일을 그대로 한다 — 순서까지.

특히 기준기간 진단은 수집이 끝난 **뒤**, 엑셀을 만들기 **전**에 품질 기록에
붙어야 한다. 그 순서가 Data_Quality 시트의 행 순서를 정한다.
"""

import json

from . import orphans, output, summarize
from .dartio import check_dart_reachable, get_dart_reader
from .excel.gpcm_book import export_gpcm_excel
from .excel.historical_book import export_historical_excel
from .gpcm import calculate_wacc_and_beta, fetch_financial_data
from .historical import calculate_historical_metrics, fetch_historical_financials
from .periods import get_base_date_str, parse_period
from .progress import NullProgress

BETA_TYPES = ('5Y', '2Y')


class InputError(ValueError):
    """도구 인자가 잘못됐다. 조회를 시작하기 전에 걸러낸다."""


def _check_tickers(tickers):
    if not tickers:
        raise InputError('종목코드를 하나 이상 주셔야 합니다. '
                         '6자리 숫자입니다 (예: 005930).')
    bad = [t for t in tickers if not (isinstance(t, str) and t.strip().isdigit()
                                      and len(t.strip()) == 6)]
    if bad:
        raise InputError(f'종목코드는 6자리 숫자여야 합니다. 잘못된 값: {bad}')
    return [t.strip() for t in tickers]


def _check_periods(periods, what):
    # 원본은 기간 목록이 비면 len(tickers)*len(periods) 가 0 이 되어
    # ZeroDivisionError 로 죽는다. 조회를 시작하기 전에 막는다.
    if not periods:
        raise InputError(f'{what} 기간이 비어 있습니다. '
                         '종료 시점이 시작 시점보다 앞서지 않는지 확인해주세요.')
    return periods


def _preflight():
    ok, reason = check_dart_reachable()
    if not ok:
        raise InputError(
            'DART 서버에 접속할 수 없습니다. API 키 문제가 아니라 네트워크에서 막힌 상태입니다.\n'
            '- 국내에서 실행해주세요. DART 는 해외 접속을 제한합니다.\n'
            '- 사내망이라면 방화벽에서 opendart.fss.or.kr 을 허용해야 합니다.\n'
            f'(원인: {reason})'
        )


def _assert_jsonable(payload):
    """전송 단계에서 죽지 않도록 여기서 먼저 확인한다.

    주가 시계열이 응답에 섞여 나가는 사고가 가장 흔한데, 그건 MCP 전송에서
    알아보기 어려운 오류로 나타난다.
    """
    json.dumps(payload, ensure_ascii=False)
    return payload


def run_gpcm(tickers, target_periods, *, rf=0.033, mrp=0.08, size_premium=0.0402,
             kd_pretax=0.035, target_tax_rate=0.264, beta_type='5Y',
             peer_selection=None, progress=None, dart=None, preflight=True):
    """모드 1 — GPCM 배수와 WACC."""
    tickers = _check_tickers(tickers)
    target_periods = _check_periods(target_periods, 'GPCM 분석')
    if beta_type not in BETA_TYPES:
        raise InputError(f"beta_type 은 {BETA_TYPES} 중 하나여야 합니다. 받은 값: {beta_type!r}")

    progress = progress or NullProgress()
    if preflight:
        _preflight()
    dart = dart if dart is not None else get_dart_reader()

    base_period_str = target_periods[-1]

    (raw_bs_rows, raw_pl_rows, all_mkt, ticker_to_name, screen_summary_data,
     base_year, base_qtr, base_date_str, all_multiples, quality) = fetch_financial_data(
        None, tickers, target_periods, dart, progress, progress)

    # --- 여기서부터 원본의 UI 블록이 하던 일 (순서 유지) ---
    problems = orphans.diagnose_base_period(all_multiples, base_period_str, quality)
    df_screen = orphans.build_screen_frame(all_multiples)
    notes_list = orphans.build_notes(base_period_str)

    target_wacc_data, avg_debt_ratio = calculate_wacc_and_beta(
        tickers, screen_summary_data, target_tax_rate, rf, mrp, size_premium,
        kd_pretax, beta_type, fiscal_year=base_year)

    warnings = list(problems)
    if target_wacc_data['Target_WACC'] <= rf:
        warnings.append(
            f"계산된 WACC({target_wacc_data['Target_WACC'] * 100:.2f}%)이 "
            f"무위험이자율({rf * 100:.2f}%)보다 낮습니다. 정상적인 결과가 아닙니다 — "
            "시가총액이 0으로 수집되어 자본구조·베타가 붕괴한 경우입니다."
        )

    book = export_gpcm_excel(
        base_period_str, base_qtr, tickers, screen_summary_data, raw_bs_rows,
        raw_pl_rows, all_mkt, ticker_to_name, target_wacc_data, beta_type,
        notes_list, avg_debt_ratio, base_date_str, df_screen, target_periods,
        quality, peer_selection)

    path = output.save(book, output.build_name('KR_GPCM', base_period_str.replace('.', '_')))

    return _assert_jsonable({
        'status': 'done',
        'file': {'path': str(path), 'uri': path.as_uri(),
                 'bytes': path.stat().st_size},
        'base_period': base_period_str,
        'base_date': base_date_str,
        'unit': '억원 (KRW 100M)',
        'periods': list(target_periods),
        'companies': [{'ticker': t, 'name': ticker_to_name.get(t),
                       'collected': t in ticker_to_name} for t in tickers],
        'multiples': summarize.multiples_at(df_screen, base_period_str),
        'statistics': summarize.statistics(df_screen, base_period_str),
        'wacc': summarize.jsonable(dict(target_wacc_data)),
        'quality': summarize.quality_report(quality),
        'warnings': warnings,
        'notes': notes_list,
    })


def run_historical(tickers, periods_to_fetch, *, progress=None, dart=None,
                   preflight=True):
    """모드 2 — 다기간 재무제표 요약."""
    tickers = _check_tickers(tickers)
    periods_to_fetch = _check_periods(periods_to_fetch, '재무제표 조회')

    progress = progress or NullProgress()
    if preflight:
        _preflight()
    dart = dart if dart is not None else get_dart_reader()

    df_summ, df_details = fetch_historical_financials(
        None, tickers, periods_to_fetch, dart, progress, progress, None)
    df_summ = calculate_historical_metrics(df_summ)

    labels = [p['label'] for p in periods_to_fetch]
    if df_summ.empty:
        return _assert_jsonable({
            'status': 'empty',
            'periods': labels,
            'warnings': ['수집된 데이터가 없습니다. 종목코드나 연도를 확인해주세요. '
                         '아직 공시되지 않은 기간일 수도 있습니다.'],
        })

    book = export_historical_excel(df_summ, df_details, periods_to_fetch)
    path = output.save(book, output.build_name(
        'KR_Historical', labels[0], 'to', labels[-1]))

    collected = set(df_summ[df_summ['Revenue'].notna()]['Ticker']) \
        if 'Revenue' in df_summ.columns else set()
    return _assert_jsonable({
        'status': 'done',
        'file': {'path': str(path), 'uri': path.as_uri(),
                 'bytes': path.stat().st_size},
        'unit': '백만원 (KRW 1M)',
        'periods': labels,
        'companies': [{'ticker': t, 'collected': t in collected} for t in tickers],
        'rows': summarize.historical_summary(df_summ),
        'warnings': [] if collected else [
            '어느 회사에서도 손익 항목을 채우지 못했습니다. 기간이 아직 공시되지 '
            '않았거나 종목코드가 잘못됐을 수 있습니다.'],
    })


def period_label_of(period_str):
    year, qtr = parse_period(period_str)
    return get_base_date_str(year, qtr)
