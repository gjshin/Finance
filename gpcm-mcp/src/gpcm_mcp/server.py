"""국내 상장사 GPCM 계산 MCP 서버.

실행:
    set OPENDART_API_KEY=발급받은키
    python -m gpcm_mcp.server              # 로컬(stdio) — Claude 데스크톱 앱, Claude Code
    python -m gpcm_mcp.server --http       # 원격(HTTP)

실제 계산은 gpcm.py / historical.py 에 있고, 이어 붙이는 것은 runner.py 가 한다.
이 파일은 도구 설명과 연결만 담당한다.
"""

import sys

from mcp.server.mcpserver import MCPServer

from . import jobs, runner
from .dartio import api_key_configured, check_dart_reachable
from .listings import get_krx_industry_listing, peer_candidate_rows
from .orphans import build_periods_to_fetch, build_target_periods
from .periods import get_base_date_str, get_latest_filed_period, is_period_filed, parse_period
from .progress import JobProgress

server = MCPServer(
    name="gpcm-kr",
    instructions=(
        "국내 상장사 유사기업 배수(GPCM)와 WACC 를 계산해 엑셀로 만드는 도구입니다. "
        "DART 공시 재무제표와 KRX 주가를 직접 조회해 계산합니다.\n\n"
        "사용 순서: check_dart_access → (필요하면 latest_filed_period 로 기준일 확인) "
        "→ gpcm_valuation 또는 historical_financials → 오래 걸리면 gpcm_job_status.\n\n"
        "**DART 는 해외 접속을 막습니다.** 국내에서 실행해야 하고, 안 되면 "
        "check_dart_access 가 먼저 알려줍니다.\n\n"
        "결과의 quality 항목을 반드시 사용자에게 먼저 알려주세요. ERROR 가 있으면 "
        "그 회사의 배수는 쓰면 안 됩니다. 값이 null 이면 '0' 이 아니라 '수집하지 "
        "못했다' 는 뜻입니다. 빠진 값을 추측해서 채우지 마세요.\n\n"
        "이 도구가 돌려주는 배수·WACC 는 파이썬이 계산한 값입니다. 엑셀 파일 안의 "
        "수치는 수식이라 엑셀로 열어야 값이 보입니다 — 같은 계산이니 둘이 다르지 않습니다.\n\n"
        "감가상각비는 DART 에서 자동으로 받을 수 없습니다. EBITDA 가 필요하면 "
        "엑셀 GPCM 시트의 노란색 D&A 열에 직접 넣어야 한다고 알려주세요."
    ),
)


def _wait(job, wait_seconds):
    """잠깐 기다려 보고, 끝났으면 결과를, 아니면 job_id 를 돌려준다."""
    job.done_event.wait(timeout=max(0, wait_seconds))
    if job.status == 'done':
        return job.result
    if job.status == 'failed':
        return {'status': 'failed', 'job_id': job.id, 'error': job.error}
    snap = job.snapshot()
    snap['hint'] = (f"아직 조회 중입니다. gpcm_job_status(job_id='{job.id}') 로 "
                    "진행 상황을 확인하세요. 회사 수 × 기간 수만큼 DART 를 조회하므로 "
                    "회사가 많으면 몇 분 걸립니다.")
    return snap


@server.tool(description=(
    "DART 에 접속할 수 있는지, API 키가 설정돼 있는지 확인합니다.\n"
    "다른 도구를 쓰기 전에 먼저 부르세요. 실패하면 몇 분씩 멈추는 대신 "
    "바로 원인을 알려줍니다.\n"
    "DART 는 해외 IP 를 차단하므로, 국내가 아닌 곳에서 실행하면 여기서 걸립니다."
))
def check_dart_access() -> dict:
    reachable, reason = check_dart_reachable()
    return {
        'reachable': reachable,
        'reason': reason,
        'api_key_configured': api_key_configured(),
        'note': ('국내에서 실행해야 합니다. DART 가 해외 접속을 제한합니다.'
                 if not reachable else None),
    }


@server.tool(description=(
    "지금 시점에 공시가 끝난 가장 최근 (연도, 분기) 를 알려줍니다.\n"
    "기준일을 정하기 전에 부르세요. 아직 공시되지 않은 분기를 넣으면 재무·주가가 "
    "비어서 결과가 0 으로 나옵니다."
))
def latest_filed_period() -> dict:
    year, qtr = get_latest_filed_period()
    return {
        'year': year,
        'quarter': qtr,
        'period': f'{year}.{qtr}',
        'base_date': get_base_date_str(year, qtr),
    }


@server.tool(description=(
    "KRX 업종과 그 업종의 상장사 후보를 봅니다.\n"
    "sector 를 비우면 업종 목록만 돌려줍니다.\n\n"
    "peer 선정 자체는 dcfpeer 도구가 평가기준일 기준으로 더 정확하게 합니다. "
    "이 도구는 종목코드를 눈으로 확인하거나, 12월 결산이 아닌 회사를 걸러낼 때 쓰세요 "
    "— 결산월이 다르면 기간 비교가 어긋납니다."
))
def list_peer_candidates(sector: str = "") -> dict:
    df_ind = get_krx_industry_listing()
    if df_ind.empty:
        return {'sectors': [], 'candidates': [],
                'note': 'KRX 업종 목록을 받지 못했습니다. 종목코드를 직접 넣어주세요.'}
    sectors = sorted(s for s in df_ind['Sector'].dropna().unique() if str(s).strip())
    if not sector:
        return {'sectors': sectors, 'candidates': []}
    candidates = peer_candidate_rows(df_ind, sector)
    return {
        'sectors': sectors,
        'sector': sector,
        'count': len(candidates),
        'candidates': candidates,
        'note': ('FiscalNot12 가 true 인 회사는 12월 결산이 아닙니다. '
                 '같은 기간으로 비교하면 어긋나니 확인하세요.'),
    }


@server.tool(description=(
    "GPCM 배수와 WACC 를 계산해 엑셀로 만듭니다. 금액 단위는 억원입니다.\n\n"
    "tickers 는 6자리 종목코드 목록입니다 (예: ['005930','000660']).\n"
    "기간은 start_year/start_quarter ~ end_year/end_quarter 이고, 마지막 기간이 "
    "기준일이 됩니다. end_year 를 비우면 공시가 끝난 최신 기간을 씁니다.\n"
    "cycle='annual' 이면 각 연도의 4Q(사업보고서)만 씁니다.\n\n"
    "WACC 가정 기본값은 앱과 같습니다 — rf 3.3%, mrp 8%, kd 3.5%, 법인세 26.4%.\n"
    "size_premium 은 한국공인회계사회 기준이며 회사 규모로 고릅니다:\n"
    "  3분위 — 시총 2,000억 미만 0.0402 / 2,000~20,000억 0.0137 / 20,000억 초과 -0.0036\n"
    "  5분위 — 2,000억 미만 0.0466 / ~5,000억 0.0302 / ~20,000억 0.0121 / "
    "~50,000억 0.0006 / 50,000억 초과 -0.0058\n"
    "  적용하지 않으려면 0.\n"
    "beta_type 은 '5Y'(5년 월간) 또는 '2Y'(2년 주간) 입니다.\n\n"
    "회사 수 × 기간 수만큼 조회하므로 오래 걸립니다. wait_seconds 안에 끝나면 결과를, "
    "안 끝나면 job_id 를 돌려줍니다."
))
def gpcm_valuation(tickers: list[str], start_year: int, start_quarter: str = "1Q",
                   end_year: int = 0, end_quarter: str = "",
                   cycle: str = "quarterly",
                   rf: float = 0.033, mrp: float = 0.08,
                   size_premium: float = 0.0402, kd_pretax: float = 0.035,
                   target_tax_rate: float = 0.264, beta_type: str = "5Y",
                   wait_seconds: int = 25) -> dict:
    latest_year, latest_qtr = get_latest_filed_period()
    end_year = end_year or latest_year
    end_quarter = end_quarter or latest_qtr

    try:
        periods = build_target_periods(start_year, start_quarter, end_year,
                                       end_quarter, cycle=cycle)
    except ValueError as exc:
        return {'status': 'failed', 'error': f'분기 표기가 잘못됐습니다: {exc}'}

    unfiled = [p for p in periods if not is_period_filed(*parse_period(p))]

    def work(job):
        result = runner.run_gpcm(
            tickers, periods, rf=rf, mrp=mrp, size_premium=size_premium,
            kd_pretax=kd_pretax, target_tax_rate=target_tax_rate,
            beta_type=beta_type, progress=JobProgress(job))
        if unfiled:
            result.setdefault('warnings', []).insert(0, (
                f"아직 공시되지 않은 기간이 포함돼 있습니다: {', '.join(unfiled)}. "
                f"해당 기간은 값이 비어 결과가 왜곡됩니다. "
                f"종료 기간을 {latest_year}년 {latest_qtr} 이하로 맞추세요."))
        return result

    try:
        job = jobs.STORE.submit('gpcm', f"GPCM {len(tickers)}개사 × {len(periods)}기간", work)
    except runner.InputError as exc:
        return {'status': 'failed', 'error': str(exc)}
    return _wait(job, wait_seconds)


@server.tool(description=(
    "여러 회사의 과거 재무제표(재무상태표·손익계산서·현금흐름표)를 한 번에 조회해 "
    "엑셀로 정리합니다. 금액 단위는 백만원입니다 (GPCM 은 억원이니 섞지 마세요).\n\n"
    "quarters 를 비우면 연간(사업보고서)만 조회합니다. "
    "quarters=['1Q','4Q'] 처럼 주면 시작·종료 분기로 해석합니다.\n\n"
    "현금흐름표는 영업·투자·재무 대분류만 Summary 에 실리고, 나머지 계정은 "
    "회사별 시트에만 남습니다.\n"
    "배수나 WACC 가 필요하면 이 도구가 아니라 gpcm_valuation 을 쓰세요."
))
def historical_financials(tickers: list[str], start_year: int, end_year: int = 0,
                          quarters: list[str] = None,
                          wait_seconds: int = 25) -> dict:
    end_year = end_year or get_latest_filed_period()[0]
    quarters = quarters or []
    start_qtr = quarters[0] if quarters else None
    end_qtr = quarters[-1] if quarters else None

    try:
        periods = build_periods_to_fetch(start_year, end_year, start_qtr, end_qtr)
    except ValueError as exc:
        return {'status': 'failed', 'error': f'분기 표기가 잘못됐습니다: {exc}'}

    def work(job):
        return runner.run_historical(tickers, periods, progress=JobProgress(job))

    try:
        job = jobs.STORE.submit(
            'hist', f"재무제표 {len(tickers)}개사 × {len(periods)}기간", work)
    except runner.InputError as exc:
        return {'status': 'failed', 'error': str(exc)}
    return _wait(job, wait_seconds)


@server.tool(description=(
    "조회 작업의 진행 상황을 봅니다. 끝났으면 결과를 그대로 돌려줍니다.\n"
    "wait_seconds 를 주면 그만큼 더 기다려 봅니다."
))
def gpcm_job_status(job_id: str, wait_seconds: int = 0) -> dict:
    job = jobs.STORE.get(job_id)
    if job is None:
        return {'status': 'not_found', 'job_id': job_id,
                'error': '그런 작업이 없습니다. 서버가 다시 시작됐을 수 있습니다. '
                         '이미 만들어진 엑셀 파일은 저장 폴더에 그대로 있습니다.'}
    return _wait(job, wait_seconds)


@server.tool(description=(
    "돌고 있는 조회를 멈춥니다. 다음 회사로 넘어가는 시점에 멈추므로 "
    "즉시 끝나지는 않습니다."
))
def gpcm_job_cancel(job_id: str) -> dict:
    job = jobs.STORE.cancel(job_id)
    if job is None:
        return {'status': 'not_found', 'job_id': job_id}
    return job.snapshot()


def main():
    if "--http" in sys.argv:
        server.run(transport="streamable-http")
    else:
        server.run(transport="stdio")


if __name__ == "__main__":
    main()
