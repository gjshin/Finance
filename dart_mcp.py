"""DART 자연어 분석용 MCP 서버.

Claude가 DART(금융감독원 전자공시)를 직접 조회·분석할 수 있게 해주는 도구 모음.

실행:
    set OPENDART_API_KEY=발급받은키
    python dart_mcp.py              # 로컬(stdio) — Claude 데스크톱 앱, Claude Code
    python dart_mcp.py --http       # 원격(HTTP) — 나중에 웹에서 쓸 때

실제 로직은 dart_tools.py에 있다. 이 파일은 도구 설명과 연결만 담당한다.
"""

import os
import sys

from mcp.server.mcpserver import MCPServer

import dart_tools as T

server = MCPServer(
    name="opendart",
    instructions=(
        "금융감독원 전자공시(DART) 조회·분석 도구입니다. 한국 기업의 재무제표, "
        "정기보고서 주요정보 28종, 주요사항보고서 36종, 증권신고서, 지분공시, "
        "공시 원문을 조회할 수 있습니다.\n\n"
        "사용 순서: 먼저 find_companies로 회사를 특정한 뒤 나머지 도구를 부르세요. "
        "비상장사는 회사명으로 찾으면 listed=false 로 함께 나옵니다.\n\n"
        "재무비율은 반드시 compute_ratios를 쓰세요. 직접 나눗셈하지 마세요 — "
        "이 도구는 계산에 쓴 계정과목을 함께 돌려주므로 감사조서에 근거를 남길 수 있습니다.\n\n"
        "값이 null이면 '0'이 아니라 '해당 계정을 찾지 못했다'는 뜻입니다. "
        "not_available 항목에 이유가 적혀 있으니 그대로 사용자에게 알려주세요. "
        "빠진 값을 추측해서 채우지 마세요."
    ),
)

_client = None


def client():
    """API 키는 환경변수에서만 읽는다. 코드나 대화에 남기지 않기 위해서다."""
    global _client
    if _client is None:
        key = os.environ.get('OPENDART_API_KEY', '').strip()
        if not key:
            raise RuntimeError(
                "OPENDART_API_KEY 환경변수가 없습니다. "
                "https://opendart.fss.or.kr 에서 무료 발급 후 설정하세요."
            )
        _client = T.DartClient(key)
    return _client


@server.tool(description=(
    "회사를 찾습니다. 다른 모든 도구를 쓰기 전에 먼저 부르세요.\n"
    "- 회사명이나 6자리 종목코드로 특정 회사 찾기\n"
    "- sector로 업종별 목록 뽑기 (예: '게임', '반도체', '제약')\n"
    "'게임회사들을 비교해줘' 같은 요청은 sector='게임'으로 후보를 먼저 뽑으세요."
))
def find_companies(query: str = "", sector: str = "", market: str = "", limit: int = 30) -> dict:
    return T.find_companies(client(), query or None, sector or None, market or None, limit)


@server.tool(description=(
    "재무제표 전체 계정을 조회합니다. 금액 단위는 원입니다.\n"
    "statement로 'BS'(재무상태표) / 'IS'(손익계산서) / 'CF'(현금흐름표)를 고를 수 있고, "
    "account_filter로 특정 계정만 볼 수 있습니다 (예: '매출채권').\n"
    "summary=True로 부르면 주요계정만 간추려 돌려줍니다 — 여러 회사를 훑을 때 훨씬 가볍습니다.\n"
    "비율을 구할 목적이라면 이 도구 대신 compute_ratios를 쓰세요."
))
def get_financial_statements(companies: list[str], year: int, quarter: str = "4Q",
                             statement: str = "", account_filter: str = "",
                             consolidated: bool = True, summary: bool = False) -> dict:
    return T.get_financial_statements(
        client(), companies, year, quarter,
        fs_div='CFS' if consolidated else 'OFS',
        statement=statement or None, account_filter=account_filter or None,
        summary=summary)


@server.tool(description=(
    "재무비율을 계산합니다. 여러 회사·여러 연도를 한 번에 비교할 수 있습니다.\n"
    "가능한 비율: 매출채권회전율, 재고자산회전율, 총자산회전율, 부채비율, 유동비율, "
    "이자보상배율, ROE, ROA, 매출총이익률, 영업이익률.\n"
    "회전율과 ROE/ROA는 분모에 평균잔액(당기말·전기말)을 씁니다.\n"
    "결과의 sources에 어떤 계정과목을 썼는지 들어 있으니 답변에 함께 알려주세요."
))
def compute_ratios(companies: list[str], years: list[int], quarter: str = "4Q",
                   ratios: list[str] = None, consolidated: bool = True) -> dict:
    return T.compute_ratios(client(), companies, years, quarter, ratios,
                            fs_div='CFS' if consolidated else 'OFS')


@server.tool(description=(
    "정기보고서(사업보고서 등)의 주요정보를 조회합니다.\n"
    "자주 쓰는 item: '타법인출자'(출자 현황), '증자', '채무증권발행', '회사채미상환', "
    "'회계감사'(감사인 이름과 감사의견), '감사용역', '최대주주', '최대주주변동', "
    "'소액주주', '배당', '자기주식', '임원', '주식총수'.\n"
    "전체 28종은 잘못된 item을 넣으면 목록으로 돌려줍니다."
))
def get_report_item(companies: list[str], item: str, years: list[int],
                    quarter: str = "4Q") -> dict:
    return T.get_report_item(client(), companies, item, years, quarter)


@server.tool(description=(
    "공시를 검색합니다. 원문을 읽으려면 여기서 얻은 rcept_no를 get_filing_document에 넘기세요.\n"
    "kind: 'A'정기공시 'B'주요사항보고서 'C'발행공시 'D'지분공시 'E'기타 "
    "'F'외부감사관련(감사보고서) 'I'거래소공시\n"
    "날짜는 'YYYY-MM-DD' 또는 'YYYYMMDD' 형식입니다.\n"
    "감사보고서를 찾으려면 kind='F'를 쓰세요."
))
def search_filings(company: str = "", start: str = "", end: str = "",
                   kind: str = "", keyword: str = "", limit: int = 50) -> dict:
    return T.search_filings(client(), company or None, start or None, end or None,
                            kind or None, keyword or None, limit)


@server.tool(description=(
    "공시 원문에서 필요한 부분을 꺼냅니다.\n"
    "사업보고서·감사보고서는 매우 길어서 통째로 읽을 수 없습니다. "
    "section 없이 먼저 불러 목차를 확인한 뒤, 원하는 절 제목을 section에 넣어 다시 부르세요.\n"
    "예: section='사업의 내용' (피어 비교용), section='주석' (감사보고서 주석 취합용)."
))
def get_filing_document(rcept_no: str, section: str = "", max_chars: int = 12000) -> dict:
    return T.get_filing_document(client(), rcept_no, section or None, max_chars)


@server.tool(description=(
    "기업개황을 조회합니다. 업종코드, 설립일, 대표자, 결산월, 주소, 홈페이지 등.\n"
    "회사가 어떤 업종으로 분류돼 있는지 확인할 때 씁니다."
))
def get_company_profile(companies: list[str]) -> dict:
    return T.get_company_profile(client(), companies)


@server.tool(description=(
    "주요사항보고서를 조회합니다. 정기보고서에는 없는 수시공시 사건들입니다.\n"
    "keyword 예: '소송', '회사합병', '회사분할', '영업양수', '영업양도', '자산양수도', "
    "'타법인증권양수', '타법인증권양도', '전환사채발행', '신주인수권부사채발행', "
    "'유상증자', '감자', '자기주식취득', '자기주식처분', '부도발생', '회생절차', '영업정지'.\n"
    "딜 실사(과거 M&A·양수도 이력)와 감사 위험평가(소송·부도·회생)에 씁니다.\n"
    "전체 36종은 잘못된 keyword를 넣으면 목록으로 돌려줍니다."
))
def get_major_events(companies: list[str], keyword: str,
                     start: str = "", end: str = "", limit: int = 50) -> dict:
    return T.get_major_events(client(), companies, keyword, start or None, end or None, limit)


@server.tool(description=(
    "증권신고서를 조회합니다. keyword: '지분증권', '채무증권', '증권예탁증권', "
    "'합병', '분할', '주식의포괄적교환이전'.\n"
    "합병·분할 신고서에는 평가방법과 합병비율 산정근거가 들어 있어 밸류에이션 참고자료가 됩니다."
))
def get_registration_statements(companies: list[str], keyword: str,
                                start: str = "", end: str = "", limit: int = 50) -> dict:
    return T.get_registration_statements(client(), companies, keyword,
                                         start or None, end or None, limit)


@server.tool(description=(
    "지분공시를 조회합니다.\n"
    "kind='대량보유' — 5% 이상 대량보유상황보고 (지분 변동 추적)\n"
    "kind='임원주요주주' — 임원·주요주주 특정증권등 소유상황보고"
))
def get_shareholding_reports(companies: list[str], kind: str = "대량보유",
                             limit: int = 100) -> dict:
    return T.get_shareholding_reports(client(), companies, kind, limit)


@server.tool(description=(
    "특정 날짜에 접수된 공시를 전부 가져옵니다. 날짜를 비우면 오늘입니다.\n"
    "'어제 감사보고서 낸 회사 알려줘'처럼 회사를 특정하지 않은 전수 조회에 씁니다.\n"
    "keyword로 보고서명을 걸러낼 수 있습니다 (예: '감사보고서')."
))
def list_filings_by_date(date: str = "", keyword: str = "", limit: int = 100) -> dict:
    return T.list_filings_by_date(client(), date or None, limit, keyword or None)


@server.tool(description=(
    "공시에 딸린 하위문서와 첨부파일 목록을 봅니다.\n"
    "사업보고서 안의 감사보고서, 감사보고서에 붙은 재무제표처럼 "
    "본문과 별개로 붙어 있는 문서를 찾을 때 쓰세요."
))
def get_filing_attachments(rcept_no: str, match: str = "") -> dict:
    return T.get_filing_attachments(client(), rcept_no, match or None)


@server.tool(description=(
    "XBRL 표준계정 체계를 조회합니다. 계정과목의 표준명·표준코드를 확인할 때 씁니다.\n"
    "sj_div 예: 'BS1'(재무상태표) 'IS1'(손익계산서) 'CIS1'(포괄손익계산서) 'CF1'(현금흐름표)"
))
def get_xbrl_taxonomy(sj_div: str) -> dict:
    return T.get_xbrl_taxonomy(client(), sj_div)


@server.tool(description=(
    "DART가 직접 산출한 주요 재무지표를 조회합니다.\n"
    "categories: '수익성', '안정성', '성장성', '활동성' (생략하면 전부)\n\n"
    "compute_ratios와의 차이: compute_ratios는 재무제표에서 직접 계산하고 "
    "쓴 계정과목을 알려줍니다. 이 도구는 DART 공식 산출값입니다. "
    "둘을 나란히 놓고 교차검증하면 좋습니다."
))
def get_financial_indicators(companies: list[str], years: list[int],
                             quarter: str = "4Q", categories: list[str] = None) -> dict:
    return T.get_financial_indicators(client(), companies, years, quarter, categories)

if __name__ == "__main__":
    if "--http" in sys.argv:
        server.run(transport="streamable-http")
    else:
        server.run(transport="stdio")
