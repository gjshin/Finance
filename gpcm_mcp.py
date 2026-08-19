"""GPCM을 Claude 데스크톱에서 부르는 로컬 MCP 서버.

피어 스크리닝(채팅) → 종목코드 복사 → Streamlit 앱 따로 실행 → 엑셀. 이 마지막
이음새를 없앤다: 채팅이 스크리닝을 마친 종목코드로 run_gpcm 을 바로 불러 엑셀을
만들게 한다. Streamlit 앱은 그대로 두고, 여기서는 gpcm_kr 을 임포트해 쓰기만 한다.

설치는 install_mcp.ps1, 인증키는 환경변수 DART_API_KEY (mydart 와 같은 방식).

실행이 티커당 수 초~수십 초라 도구 호출 안에서 기다리면 클라이언트가 타임아웃한다.
run_gpcm 은 백그라운드 스레드를 띄우고 즉시 돌아오고, gpcm_status 로 확인한다.
"""

from __future__ import annotations

import contextlib
import io
import os
import re
import sys
import threading
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import MCPServer

mcp = MCPServer("gpcm")

class _ThreadRoutedStdout:
    """sys.stdout 대역 — '조용히' 표시한 스레드의 출력만 stderr 로 보낸다.

    stdout 은 MCP 프로토콜 통로라 gpcm_kr 임포트가 흘리는 안내문(streamlit 의
    "to view this Streamlit app..." 등)이 섞이면 안 된다. 그렇다고
    contextlib.redirect_stdout 을 쓰면 안 된다 — 그건 프로세스 전역이라,
    백그라운드 스레드가 임포트하는 동안 서버가 보내는 JSON-RPC 응답까지 통째로
    stderr 로 새어 클라이언트가 initialize 응답을 못 받고 60초 뒤 끊는다(실측).
    그래서 스레드 단위로만 돌린다. 프로토콜을 쓰는 스레드는 손대지 않는다.
    """

    _muted = threading.local()

    def __init__(self, real, alt):
        self._real, self._alt = real, alt

    def _target(self):
        return self._alt if getattr(self._muted, "on", False) else self._real

    def write(self, s):
        return self._target().write(s)

    def flush(self):
        self._target().flush()

    def __getattr__(self, name):  # encoding·buffer·fileno·reconfigure 등은 원본 것
        return getattr(self._real, name)


@contextlib.contextmanager
def _muted_stdout():
    """이 스레드에서만 print 를 stderr 로 보낸다 (다른 스레드는 그대로)."""
    local = _ThreadRoutedStdout._muted
    previous = getattr(local, "on", False)
    local.on = True
    try:
        yield
    finally:
        local.on = previous


# 임포트 시점에 갈아 끼운다 — gpcm_kr 로드가 main() 을 거치지 않고 일어나는
# 경로(테스트·직접 임포트)에서도 _muted_stdout 이 실제로 듣게 하려면 여기여야 한다.
sys.stdout = _ThreadRoutedStdout(sys.stdout, sys.stderr)


# gpcm_kr 은 임포트가 무겁다: streamlit·pandas·scipy 를 끌고 오고, 사이드바 블록이
# 실행되며 KRX 목록까지 조회한다 — 합쳐서 수십 초. Claude 는 initialize 응답을
# 60초만 기다리므로, 서버 기동 시 이걸 먼저 하면 접속 자체가 끊긴다(실측 80초).
# 그래서 기동은 빈손으로 즉시 하고, 모듈은 뒤에서(main 의 예열 스레드) 또는
# 첫 도구 호출에서 로드한다.
M = None
_load_lock = threading.Lock()


def _load():
    """gpcm_kr 을 (한 번만) 로드하고 돌려준다. 끝날 때까지 막힌다."""
    global M
    with _load_lock:
        if M is None:
            print("gpcm-mcp: 계산 모듈 로드 중...", file=sys.stderr)
            with _muted_stdout():
                import gpcm_kr as _M
            M = _M
            print("gpcm-mcp: 계산 모듈 준비 완료", file=sys.stderr)
    return M

OUTPUT_DIR = Path.home() / "Documents" / "GPCM"

# 사이드바 기본값과 동일하게 유지한다 (gpcm_kr.py 의 위젯 기본값)
BETA_TYPES = ("5Y", "2Y")
MAX_TICKERS = 30  # 티커당 수 초~수십 초 — 이 이상은 앱에서 돌리는 게 낫다
PREHEAT_DELAY_SEC = 5  # 접속 handshake 가 끝난 뒤에 예열을 시작한다

# 원본은 gpcm_kr.py UI 블록의 notes_list (Valuation Methodology Notes).
# 그쪽 파일은 수정하지 않기로 해 사본을 둔다 — 저쪽이 바뀌면 여기도 맞출 것.
# 첫 줄(Base Date)은 실행 시점 값으로 만들어 붙인다.
_NOTES_STATIC = [
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
    '• MRP: 한국공인회계사회 시장위험 프리미엄 가이던스 (2026.06.05) 7~9% 범위 내 선택',
    '• Size Premium: 한국공인회계사회 기업규모위험 프리미엄 연구결과 (2026.06.05) 시가총액 분위 기준',
    '• Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1 (조정을 Unlevered 계산 前에 적용)',
    '• Beta 벤치마크: 전 종목 KOSPI(^KS11) 단일 기준. 코스닥 종목도 KOSPI 대비로 산출한다 —',
    '  피어 무차입베타를 평균해 하나의 WACC 을 만들므로 모든 β 가 같은 지수 기준이어야 하고,',
    '  MRP 도 시장 전체(KOSPI) 기준 추정치라 β 와 기준을 맞춘 것이다',
    '• Beta 수익률: 배당 미반영 가격수익률 기준 (종목·지수 모두 동일 기준)',
    '• Beta 평균 대상: 조정베타가 0 초과인 회사 (엑셀 GPCM 시트 Mean 행과 동일 모집단)',
    '• D/E Ratio = IBD / (Market Cap + 우선주 + NCI)',
    '• Debt Ratio (D/V) = IBD / (Market Cap + 우선주 + IBD + NCI)',
    '• 우선주: BS의 우선주자본금(액면) 기준. 시가총액은 보통주만 반영하므로 자기자본가치에 가산',
    '• Unlevered Beta = Levered Beta / (1 + (1 - Tax Rate) × D/E Ratio)',
    '• Tax Rate: 한국 법인세 한계세율 (지방소득세 포함, 세전순이익 기준, 사업연도별 세율표 적용)',
    '   - FY2023~2025: 2억 이하 9.9% | 2~200억 20.9% | 200~3,000억 23.1% | 3,000억 초과 26.4%',
    '   - FY2026~    : 2억 이하 11.0% | 2~200억 22.0% | 200~3,000억 24.2% | 3,000억 초과 27.5% (2025년 세법개정)',
]

_PERIOD_RE = re.compile(r"^\d{4}\.[1-4]Q$")

# 이 서버가 내보내야 할 도구. 파이썬 3.14 에서 pydantic 이 깨져 등록이 중간에
# 멈춘 적이 있어(첫 도구 하나만 남았다), gpcm_doctor 가 실제 등록분과 대조한다.
EXPECTED_TOOLS = ("get_wacc_inputs", "run_gpcm", "gpcm_status", "gpcm_review",
                  "list_krx_companies", "check_trading_gaps", "gpcm_doctor")

# 설치 경로가 둘이라 버전이 어긋날 수 있다: 확장(.mcpb)은 설치 시점의 사본을 쓰고,
# install_mcp.ps1 는 내려받은 폴더의 파일을 그대로 돌린다(git pull 을 안 하면 옛 판).
# 어느 쪽이 돌고 있는지 채팅에서 바로 보이도록 버전과 경로를 읽어 둔다.
SERVER_DIR = Path(__file__).resolve().parent


def _server_version() -> str:
    try:
        import json
        return json.loads((SERVER_DIR / "manifest.json").read_text(encoding="utf-8"))["version"]
    except Exception:
        return "unknown"


def _api_key() -> str:
    """DART 인증키. 확장(.mcpb)으로 설치할 때 입력칸이 비면 환경변수에
    `${user_config.dart_api_key}` 리터럴이 그대로 들어온다 — 없음으로 취급한다.
    (mydart·myacc에서 실제로 겪은 함정과 같은 처방)
    """
    value = (os.environ.get("DART_API_KEY") or "").strip()
    if value.startswith("${") and value.endswith("}"):
        return ""
    return value


def _parse_period(p: str) -> tuple[int, str]:
    # gpcm_kr.parse_period 와 동일 — 모듈 로드 전에도 입력 검증이 되도록 여기 둔다
    year, qtr = p.strip().split(".")
    return int(year), qtr


def _build_periods(start_period: str, end_period: str | None = None) -> list[str]:
    """UI 의 기간 조립 로직 그대로: 시작~끝 분기를 "YYYY.NQ" 목록으로 편다."""
    end_period = end_period or start_period
    for p in (start_period, end_period):
        if not _PERIOD_RE.match(p):
            raise ValueError(f'기간은 "2025.4Q" 형식입니다: {p}')
    qtrs = ["1Q", "2Q", "3Q", "4Q"]
    sy, sq = _parse_period(start_period)
    ey, eq = _parse_period(end_period)
    if (ey, qtrs.index(eq)) < (sy, qtrs.index(sq)):
        raise ValueError(f"종료 기간({end_period})이 시작 기간({start_period})보다 빠릅니다.")
    periods = []
    for y in range(sy, ey + 1):
        s_idx = qtrs.index(sq) if y == sy else 0
        e_idx = qtrs.index(eq) if y == ey else 3
        periods.extend(f"{y}.{qtrs[i]}" for i in range(s_idx, e_idx + 1))
    return periods


class Recorder:
    """Streamlit 의 status_container/progress_bar 자리에 들어가 진행을 기록한다."""

    def __init__(self, job: dict[str, Any]):
        self.job = job

    def write(self, text: Any = "") -> None:
        self.job["log"].append(str(text)[:200])
        del self.job["log"][:-50]  # 마지막 50줄만 유지

    def update(self, *args: Any, **kwargs: Any) -> None:
        label = kwargs.get("label")
        if label:
            self.job["log"].append(str(label)[:200])

    def progress(self, value: Any) -> None:
        try:
            self.job["pct"] = round(float(value) * 100)
        except (TypeError, ValueError):
            pass


_jobs: dict[str, dict[str, Any]] = {}
_lock = threading.Lock()
_counter = 0


def _work(job: dict[str, Any]) -> None:
    p = job["params"]
    with _muted_stdout():  # 수집 중 라이브러리가 흘리는 출력이 프로토콜에 섞이지 않게
        _run_job(job, p)


def _run_job(job: dict[str, Any], p: dict[str, Any]) -> None:
    try:
        dart = M.get_dart_reader(p["api_key"])
        rec = Recorder(job)

        (raw_bs, raw_pl, all_mkt, names, summary, base_year, base_qtr,
         base_date_str, all_multiples, quality) = M.fetch_financial_data(
            p["api_key"], p["tickers"], p["periods"], dart, rec, rec)

        # 타겟 법인세율 — 비워 두면 피평가회사 세전이익으로 정한다.
        # 세율표는 사업연도별(FY2026 부터 인상)이라 기준일 연도를 함께 넣는다.
        tax_rate = p["tax_rate"]
        if tax_rate is None:
            row = next((s for s in summary if s.get("Ticker") == p["target"]), None)
            pretax = (row or {}).get("Pretax_Income")
            tax_rate = M.get_korean_marginal_tax_rate(pretax, base_year)
            job["tax_basis"] = {
                "ticker": p["target"],
                "company": (row or {}).get("Company", ""),
                "pretaxIncome100M": None if pretax is None else round(float(pretax), 1),
                "fiscalYear": base_year,
                "rate_pct": round(tax_rate * 100, 2),
            }
            quality.add(M.SEV_INFO, p["target"], (row or {}).get("Company", ""), 'Tax Rate',
                        f'타겟 법인세율을 세전이익 {pretax if pretax is not None else "미상"}억원과 '
                        f'FY{base_year} 한계세율표로 {tax_rate*100:.2f}% 로 정했습니다 '
                        f'(한국 법인 전제, 지방소득세 포함).')

        wacc_data, avg_debt_ratio = M.calculate_wacc_and_beta(
            p["tickers"], summary, tax_rate, p["rf"], p["mrp"],
            p["size_premium"], p["kd_pretax"], p["beta_type"], fiscal_year=base_year,
            quality=quality)  # 베타가 빠지거나 기본값이 쓰이면 Data_Quality 에 남는다

        base_period = p["periods"][-1]
        if wacc_data["Target_WACC"] <= p["rf"]:
            # 시가총액이 0으로 수집돼 자본구조가 무너진 경우 — 앱과 같은 경고
            quality.add(M.SEV_ERROR, "", "", "WACC",
                        f"계산된 WACC({wacc_data['Target_WACC']*100:.2f}%)이 무위험이자율보다 "
                        "낮습니다. 시가총액 수집 실패 여부를 Data_Quality에서 확인하세요.")

        notes = [f"• Base Date: {base_period} ({base_date_str}) | Unit: 억원 (KRW 100M)"]
        if p.get("rate_source"):
            # 조서에서 "이 rf 는 어디서 왔나"에 답할 근거를 산출물 안에 남긴다
            notes.append(f'• Rf/Kd 출처: {p["rate_source"]}')
        notes += _NOTES_STATIC
        book: io.BytesIO = M.export_gpcm_excel(
            base_period, base_qtr, p["tickers"], summary, raw_bs, raw_pl, all_mkt,
            names, wacc_data, p["beta_type"], notes, avg_debt_ratio, base_date_str,
            M.pd.DataFrame(all_multiples), p["periods"], quality)

        out = Path(job["file"])
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(book.getvalue())

        job["dq"] = {
            "errors": sum(1 for r in quality.rows if r["Level"] == M.SEV_ERROR),
            "warnings": sum(1 for r in quality.rows if r["Level"] == M.SEV_WARN),
            "top": [f'{r["Level"]} [{r["Ticker"]}] {r["Item"]}' for r in quality.rows[:5]],
        }
        job["target_wacc"] = round(wacc_data["Target_WACC"] * 100, 2)
        job["review"] = _summarize(all_multiples, summary, p["tickers"],
                                   p["periods"][-1], p["beta_type"])
        job["state"] = "done"
        job["pct"] = 100
    except Exception:
        job["state"] = "failed"
        job["error"] = traceback.format_exc(limit=3)


@mcp.tool()
def get_wacc_inputs(as_of: str = "", bond_grade: str = "AA-",
                    rf_maturity: str = "5년", kd_maturity: str = "3년") -> dict[str, Any]:
    """기준일 시장금리를 조회해 WACC 입력값(무위험이자율·타인자본비용)을 근거와 함께 제시한다.

    run_gpcm 을 부르기 전에 이걸 먼저 불러 rf·kd_pretax 후보와 출처를 확인한다.

    Args:
        as_of: **평가기준일**. "2026-06-30" 또는 GPCM 기간 표기 "2026.2Q" 둘 다 된다
            (분기 표기는 분기말로 읽는다). 비우면 오늘. 그 날짜 이전의 최근
            고시치를 쓴다. run_gpcm 의 기준일과 반드시 같은 날로 맞춘다 —
            다르면 금리만 다른 시점이 되어 조서가 어긋난다.
        bond_grade: 회사채 신용등급 (AA- 또는 BBB-). 평가대상의 신용도에 맞춘다.
        rf_maturity: 국고채 만기. "5년"·"10Y"·"10" 형식 모두 된다. 기본 5년.
            평가 대상 현금흐름의 기간에 맞춘다 — 영구현금흐름 전제면 10년·20년을
            쓰기도 한다. 무키 경로는 1·3·5·10년만 되고, 그 밖(20·30년)은
            ECOS 인증키가 있어야 한다.
        kd_maturity: 회사채 만기. 기본 3년 (ECOS 시장금리 통계표에 실린 만기).

    [쓰는 규칙]
    - 여기서 나온 값을 **그대로 run_gpcm 에 넣지 않는다.** 사용자에게 값과 출처를
      보여주고 확정을 받는다. 시장금리는 후보일 뿐 기준일·만기·등급 선택은 판단이다.
    - mrp(시장위험프리미엄)·size_premium(규모프리미엄)은 여기서 안 나온다. 시장에서
      관측되는 값이 아니라 판단 항목이라, 종전처럼 사용자에게 개별로 묻는다.
    - 조회에 실패하면 값을 지어내지 않고 실패를 알린다. 그때는 사용자가 직접 넣는다.
    - 확정 후 run_gpcm(rate_source=<citation>) 로 넘기면 엑셀 Notes 에 근거가 남는다.
    """
    _load()  # 금리 조회 로직은 gpcm_kr 에 있다 — Streamlit 앱과 같은 코드를 쓴다
    grade = (bond_grade or "").strip().upper()
    if grade not in M.BOND_GRADES:
        raise ValueError(f"bond_grade 는 {' 또는 '.join(M.BOND_GRADES)} 입니다: {bond_grade!r}")
    rf_term = M.maturity(rf_maturity, "rf")
    kd_term = M.maturity(kd_maturity, "kd")
    asof = M.as_of_date(as_of)
    key = M.ecos_key_from_env()

    rf, rf_tried = M.fetch_market_rate("rf", asof, None, rf_term, key)
    kd, kd_tried = M.fetch_market_rate("kd", asof, grade, kd_term, key)
    if rf is None and kd is None:
        raise RuntimeError(
            "시장금리를 한 곳에서도 못 받았습니다. rf·kd_pretax 를 직접 넣어야 합니다.\n"
            "시도한 경로 — " + " / ".join(rf_tried + kd_tried)
            + ("\n한국은행 ECOS 인증키(무료, ecos.bok.or.kr)를 확장 설정에 넣으면 "
               "이 경로가 열립니다." if not key else ""))

    parts: list[str] = []
    result: dict[str, Any] = {
        "asOf": asof.strftime("%Y-%m-%d"),
        "fetchedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    for key, got, label, tried in (("rf", rf, f"국고채 {rf_term}", rf_tried),
                                   ("kd_pretax", kd, f"회사채 {kd_term} {grade}", kd_tried)):
        if got:
            result[key] = {"label": label, **got}
            parts.append(f'{label} {got["value"]}% ({got["rateDate"]}, {got["source"]})')
        else:
            result[key] = {"label": label, "value": None,
                           "failed": " / ".join(tried),
                           "note": "못 구했습니다. 이 값은 사용자가 직접 넣어야 합니다."}
    result["citation"] = " | ".join(parts)
    result["judgment"] = {
        "mrp": "시장위험프리미엄 — 관측치가 아니라 판단 항목이라 조회하지 않는다. 사용자에게 묻는다.",
        "size_premium": "규모프리미엄 — 위와 같다. 참고 관행값은 있으나 대상 규모에 따라 달라진다.",
        "tax_rate": "한계세율 — 사업연도 세율표에 따르되 대상의 과세표준 구간 판단이 필요하다.",
        "maturity": f"만기 선택도 판단이다 (지금 rf={rf_term}, kd={kd_term}). "
                    "평가 대상 현금흐름의 기간에 맞춘다 — 필요하면 rf_maturity 를 바꿔 다시 부른다.",
    }
    result["note"] = ("값을 그대로 쓰지 말고 사용자에게 확정을 받는다. "
                      "확정 후 run_gpcm(rate_source=citation) 으로 넘기면 엑셀 Notes 에 근거가 남는다.")
    return result


@mcp.tool()
def run_gpcm(
    tickers: list[str],
    start_period: str,
    end_period: str | None = None,
    rf: float = 3.3,
    mrp: float = 8.0,
    size_premium: float = 4.02,
    kd_pretax: float = 3.5,
    tax_rate: float | None = None,
    beta_type: str = "5Y",
    rate_source: str = "",
    target_ticker: str = "",
) -> dict[str, Any]:
    """GPCM 밸류에이션을 백그라운드로 실행해 엑셀 파일을 만든다.

    [호출 전 확인 — 반드시 지킬 것]
    사용자가 WACC 파라미터를 직접 말하지 않았다면, 호출하기 전에 아래 6개를
    **하나씩 개별 질문**해 확정받는다. 각 질문에 기본값을 제시하고, 사용자가
    "기본값" 또는 값으로 답하면 그것을 쓴다. 기본값을 임의로 적용하지 않는다.
    rf·kd_pretax 는 먼저 get_wacc_inputs 로 기준일 시장금리를 조회해, 조회된
    값과 출처를 보여주며 묻는다(조회값을 확정 없이 그대로 쓰지 않는다).
      1. 무위험이자율 rf (기본 3.3%)
      2. 시장위험프리미엄 mrp (기본 8.0%)
      3. 사이즈 프리미엄 size_premium — 시총 구간을 안내하며 묻는다
         (한국공인회계사회 2026.06.05 연구결과, 3분위 기준)
         2,583억 이하 4.02% / 2,585~8,671억 1.19% / 8,679억 초과 -0.45% / 미적용 0
         5분위가 필요하면 1,759억 이하 4.86% / ~3,135억 2.67% / ~6,600억 0.97% /
         ~18,794억 -0.06% / 초과 -0.51%
      4. 세전 타인자본비용 kd_pretax (기본 3.5%)
      5. 타겟 법인세율 tax_rate — **묻지 않는다.** 비워 두면 피평가회사의
         세전이익과 사업연도 세율표로 자동 산출한다(대상이 한국 법인이라는 전제).
         사용자가 값을 지정하면 그것을 쓴다.
      6. 베타 종류 beta_type — 5Y(5년 월간) 또는 2Y(2년 주간)
    사용자가 "전부 기본값으로"라고 하면 개별 질문을 생략해도 된다.

    티커당 수 초~수십 초 걸리므로 이 도구는 시작만 하고 바로 돌아온다.
    진행과 결과는 gpcm_status 로 확인한다. **결과를 사용자에게 전할 때는
    파일 경로와 함께 Data_Quality 요약을 반드시 같이 전한다.**

    Args:
        tickers: 종목코드 6자리 목록 (피평가회사 포함). 최대 30개.
        start_period: 시작 기간, "2025.4Q" 형식. 연간이면 4Q 하나만 지정.
        end_period: 종료 기간 (생략하면 start_period 단일 기간). 기준일은
            마지막 기간의 분기말이다.
        rf: 무위험이자율 %. 기본 3.3
        mrp: 시장위험프리미엄 %. 기본 8.0
        size_premium: 사이즈 프리미엄 %. 기본 4.02 (3분위 Micro, 시총 2,583억 이하).
            2,585~8,671억은 1.19, 8,679억 초과는 -0.45, 미적용은 0.
            한국공인회계사회 기업규모위험 프리미엄 연구결과(2026.06.05) 기준.
        kd_pretax: 세전 타인자본비용 %. 기본 3.5
        tax_rate: 타겟 법인세율 %. **비우면 자동** — 피평가회사의 세전이익(LTM)을
            사업연도 한계세율표(지방소득세 포함)에 넣어 정한다. 한국 법인 전제.
            직접 넣으면 그 값을 쓰고, 자동 산출은 하지 않는다.
        target_ticker: 세율을 정할 피평가회사 종목코드. 비우면 tickers 의 첫 번째.
        beta_type: WACC 에 쓸 베타. "5Y"(5년 월간) 또는 "2Y"(2년 주간).
        rate_source: rf·kd 의 근거 문장. get_wacc_inputs 의 citation 을 그대로
            넣으면 엑셀 Notes 에 출처·금리기준일이 남는다. 손으로 정한 값이면
            그 근거를 적는다. 비우면 Notes 에 아무것도 안 붙는다.
    """
    global _counter

    codes = [t.strip() for t in tickers if t and t.strip()]
    if not codes:
        raise RuntimeError("tickers가 비어 있습니다.")
    bad = [t for t in codes if not re.match(r"^\d{6}$", t)]
    if bad:
        raise RuntimeError(f"종목코드는 6자리 숫자입니다: {', '.join(bad)}")
    if len(codes) > MAX_TICKERS:
        raise RuntimeError(
            f"{len(codes)}개는 너무 많습니다(상한 {MAX_TICKERS}). 후보를 좁힌 뒤 돌리세요.")
    if beta_type not in BETA_TYPES:
        raise RuntimeError(f'beta_type은 "5Y" 또는 "2Y"입니다: {beta_type}')
    target = (target_ticker or "").strip() or codes[0]
    if target not in codes:
        raise RuntimeError(
            f"target_ticker({target})가 tickers 안에 없습니다. "
            "피평가회사도 목록에 넣어야 세율을 그 회사 기준으로 정합니다.")

    periods = _build_periods(start_period, end_period)

    api_key = _api_key()
    if not api_key:
        raise RuntimeError(
            "DART 인증키가 없습니다. 확장 설정에서 인증키를 입력하거나, "
            "install_mcp.ps1 로 설치했다면 다시 실행해 등록하세요.")

    _load()  # 예열이 아직이면 여기서 마저 로드한다 (백그라운드 실행이라 도구 시간엔 여유가 있다)

    ok, reason = M.check_dart_reachable()
    if not ok:
        raise RuntimeError(
            f"DART에 연결할 수 없습니다({reason}). DART는 해외 접속을 제한하므로 "
            "국내 네트워크인지, 방화벽이 opendart.fss.or.kr 을 막지 않는지 확인하세요.")

    with _lock:
        running = [j for j in _jobs.values() if j["state"] == "running"]
        if running:
            raise RuntimeError(
                f'이미 실행 중인 작업이 있습니다({running[0]["id"]}). '
                "gpcm_status 로 완료를 확인한 뒤 다시 부르세요.")
        _counter += 1
        job_id = f"job-{_counter}"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        job: dict[str, Any] = {
            "id": job_id,
            "state": "running",
            "pct": 0,
            "log": [],
            "file": str(OUTPUT_DIR / f"GPCM_{periods[-1].replace('.', '_')}_{stamp}.xlsx"),
            "params": {
                "api_key": api_key, "tickers": codes, "periods": periods,
                "rf": rf / 100, "mrp": mrp / 100, "size_premium": size_premium / 100,
                "kd_pretax": kd_pretax / 100,
                "tax_rate": None if tax_rate is None else tax_rate / 100,
                "target": target,
                "beta_type": beta_type, "rate_source": rate_source.strip(),
            },
        }
        _jobs[job_id] = job

    unfiled = [p for p in periods
               if not M.is_period_filed(*_parse_period(p))]

    thread = threading.Thread(target=_work, args=(job,), daemon=True)
    job["thread"] = thread
    thread.start()

    result: dict[str, Any] = {
        "job_id": job_id,
        "state": "running",
        "tickers": len(codes),
        "periods": periods,
        "output_file": job["file"],
        "note": f"약 {len(codes) * len(periods) * 10}초 내외 걸립니다. gpcm_status로 확인하세요.",
    }
    if unfiled:
        result["warning"] = (
            f"아직 공시되지 않은 기간이 있습니다: {', '.join(unfiled)} — "
            "해당 기간은 0으로 나오고 Data_Quality에 기록됩니다.")
    return result


def _find_job(job_id: str | None) -> dict[str, Any]:
    """id 로 작업을 찾는다. 생략하면 가장 최근 것 (status·review 가 함께 쓴다)."""
    with _lock:
        if not _jobs:
            raise RuntimeError("실행한 작업이 없습니다. run_gpcm 부터 부르세요.")
        job = _jobs.get(job_id) if job_id else _jobs[f"job-{_counter}"]
    if job is None:
        raise RuntimeError(f"작업을 찾을 수 없습니다: {job_id}")
    return job


@mcp.tool()
def gpcm_status(job_id: str | None = None) -> dict[str, Any]:
    """run_gpcm 작업의 진행 상황과 결과를 확인한다.

    Args:
        job_id: run_gpcm이 돌려준 작업 id. 생략하면 가장 최근 작업.
    """
    job = _find_job(job_id)

    result: dict[str, Any] = {
        "job_id": job["id"],
        "state": job["state"],
        "progress_pct": job["pct"],
        "recent": job["log"][-5:],
    }
    if job["state"] == "done":
        result["file"] = job["file"]
        result["data_quality"] = job["dq"]
        result["target_wacc_pct"] = job["target_wacc"]
        if job.get("tax_basis"):
            # 세율을 자동으로 정했으면 무엇을 근거로 했는지 함께 보고한다
            result["tax_rate_basis"] = job["tax_basis"]
        result["note"] = (
            "숫자를 쓰기 전에 엑셀의 Data_Quality 시트를 먼저 확인하세요. "
            "ERROR가 있으면 해당 값은 수집 실패로 0이 들어가 있습니다. "
            "D&A는 자동 수집이 안 되므로 EBITDA는 엑셀에서 수기 입력이 필요합니다.")
        if job.get("review", {}).get("toReview"):
            result["to_review"] = len(job["review"]["toReview"])
            result["note"] += (f' 배수·베타에서 재검토가 필요한 회사가 '
                               f'{len(job["review"]["toReview"])}곳 있습니다 — gpcm_review 로 확인하세요.')
    elif job["state"] == "failed":
        result["error"] = job["error"]
    return result


# 명단 메모이즈 — st.cache_resource 는 bare mode 에서 신뢰할 수 없어 자체로 든다
_roster_memo: dict[str, Any] = {"at": 0.0, "df": None}
_ROSTER_TTL_SEC = 600
MAX_COMPANIES = 400


def _roster():
    import time as _time

    now = _time.monotonic()
    if _roster_memo["df"] is not None and now - _roster_memo["at"] < _ROSTER_TTL_SEC:
        return _roster_memo["df"]
    df = _load().get_krx_industry_listing()
    if df is None or df.empty:
        raise RuntimeError(
            "KRX 상장회사목록을 받지 못했습니다. 잠시 후 다시 시도하세요 "
            "(사내망이면 kind.krx.co.kr 차단 여부 확인).")
    _roster_memo.update(at=now, df=df)
    return df


# --- 결과 검토 (배수 이상치 · 베타 신뢰도) ------------------------------------
# 산출 엑셀을 다시 읽어 요약할 수는 없다 — 배수 셀은 전부 엑셀 수식이고 openpyxl 은
# 계산 결과를 저장하지 않아 읽으면 비어 있다. 그래서 실행이 끝나는 자리에서,
# 아직 손에 있는 수집 결과로 요약을 만들어 job 에 넣어 둔다.

OUTLIER_FACTOR = 3.0   # 중앙값 대비 몇 배 벗어나면 눈에 띄게 할지 (배제 기준 아님)
LOW_R2 = 0.1           # 시장 설명력이 이보다 낮으면 주의 — 관행적 참고선
LOW_N_RATIO = 0.6      # 관측치가 기간 최대치의 이 비율에 못 미치면 주의
MAX_OBS = {"5Y": 60, "2Y": 104}  # 5년 월간 60개, 2년 주간 104개가 상한

MULTIPLE_KEYS = ("EV/Revenue", "EV/EBITDA", "EV/EBIT", "PER", "PBR")


def _numbers(rows: list[dict[str, Any]], key: str) -> list[tuple[str, float]]:
    out = []
    for r in rows:
        v = r.get(key)
        if v is None:
            continue
        try:
            f = float(v)
        except (TypeError, ValueError):
            continue
        if f == f and abs(f) != float("inf"):  # NaN·inf 제외
            out.append((r.get("Ticker", ""), f))
    return out


def _quantile(values: list[float], q: float) -> float:
    if not values:
        return 0.0
    s = sorted(values)
    pos = (len(s) - 1) * q
    lo, hi = int(pos), min(int(pos) + 1, len(s) - 1)
    return s[lo] + (s[hi] - s[lo]) * (pos - lo)


def _summarize(all_multiples: list[dict[str, Any]], summary: list[dict[str, Any]],
               tickers: list[str], base_period: str, beta_type: str) -> dict[str, Any]:
    """기준기간 배수 분포와 베타 신뢰도를 정리한다. 판단은 하지 않는다."""
    base_rows = [m for m in (all_multiples or []) if m.get("Period") == base_period]
    names = {s.get("Ticker"): s.get("Company", "") for s in (summary or [])}

    multiples: dict[str, Any] = {}
    flagged: dict[str, list[str]] = {}
    for key in MULTIPLE_KEYS:
        pairs = _numbers(base_rows, key)
        if not pairs:
            continue
        values = [v for _, v in pairs]
        med = _quantile(values, 0.5)
        entry = {
            "n": len(values),
            "median": round(med, 2),
            "q1": round(_quantile(values, 0.25), 2),
            "q3": round(_quantile(values, 0.75), 2),
        }
        negatives = [t for t, v in pairs if v <= 0]
        if negatives:
            entry["nonPositive"] = negatives  # 적자·자본잠식 — 배수 자체가 의미 없다
            for t in negatives:
                flagged.setdefault(t, []).append(f"{key} 가 0 이하")
        if med > 0:
            far = [t for t, v in pairs
                   if v > 0 and (v > med * OUTLIER_FACTOR or v < med / OUTLIER_FACTOR)]
            if far:
                entry["farFromMedian"] = far
                for t in far:
                    flagged.setdefault(t, []).append(f"{key} 가 중앙값의 {OUTLIER_FACTOR}배 밖")
        missing = [t for t in tickers if t not in {p[0] for p in pairs}]
        if missing:
            entry["missing"] = missing
            for t in missing:
                flagged.setdefault(t, []).append(f"{key} 없음(수집 실패 또는 계산 불가)")
        multiples[key] = entry

    # 베타 신뢰도 — R² 는 시장 설명력, n 은 회귀에 들어간 관측치 수
    cap = MAX_OBS.get(beta_type, 60)
    betas = []
    for s in (summary or []):
        ticker = s.get("Ticker")
        if ticker not in tickers:
            continue
        r2, n = s.get(f"Beta_{beta_type}_R2"), s.get(f"Beta_{beta_type}_N")
        row: dict[str, Any] = {"code": ticker, "name": s.get("Company", ""),
                               "r2": None if r2 is None else round(float(r2), 3),
                               "n": n, "maxN": cap}
        reasons = []
        if r2 is None or n is None:
            reasons.append("베타를 산출하지 못했습니다 (Data_Quality 의 Beta 경고 참조)")
        else:
            if r2 < LOW_R2:
                reasons.append(f"시장 설명력이 낮습니다 (R² {r2:.3f} < {LOW_R2}) — "
                               "주가가 시장과 거의 같이 움직이지 않아 기울기의 근거가 약합니다")
            if n < cap * LOW_N_RATIO:
                reasons.append(f"관측치가 {n}개로 기간 최대치({cap})의 "
                               f"{LOW_N_RATIO*100:.0f}% 에 못 미칩니다 — 상장이 늦었거나 "
                               "거래정지 구간이 있는지 확인하세요")
        if reasons:
            row["caution"] = reasons
            flagged.setdefault(ticker, []).extend(reasons)
        betas.append(row)

    return {
        "basePeriod": base_period,
        "betaType": beta_type,
        "multiples": multiples,
        "beta": betas,
        "toReview": [{"code": t, "name": names.get(t, ""), "reasons": rs}
                     for t, rs in flagged.items()],
        "note": ("재검토 권고일 뿐 배제 기준이 아닙니다. 어느 회사를 뺄지는 사용자가 정합니다. "
                 f"기준선(중앙값 {OUTLIER_FACTOR}배, R² {LOW_R2}, 관측치 {LOW_N_RATIO*100:.0f}%)은 "
                 "관행적 참고치이며 사안에 따라 조정하십시오."),
    }


@mcp.tool()
def gpcm_review(job_id: str | None = None) -> dict[str, Any]:
    """끝난 run_gpcm 결과에서 배수 이상치와 베타 신뢰도를 짚어 재검토 대상을 알려준다.

    엑셀을 열어 눈으로 훑는 대신 쓴다. 배수는 기준기간 기준이다.

    Args:
        job_id: 생략하면 가장 최근 작업.

    [읽는 법]
    - toReview 에 오른 회사를 **자동으로 빼지 않는다.** 사용자에게 사유와 함께 보이고
      판단을 받는다.
    - 베타의 R² 는 주가 변동 중 시장으로 설명되는 비중이다. 낮으면 그 회사의 베타는
      시장과의 관계가 아니라 개별 사정을 담고 있어 자본비용 근거로 약하다.
    - n 은 회귀에 실제로 들어간 수익률 개수다. 적으면 기간을 조금만 옮겨도 값이 흔들린다.
    - 배수가 0 이하인 회사는 그 배수를 쓸 수 없다 (적자·자본잠식).
    """
    job = _find_job(job_id)
    if job["state"] != "done":
        raise ValueError(f'작업이 아직 {job["state"]} 입니다. gpcm_status 로 완료를 확인하세요.')
    review = job.get("review")
    if not review:
        raise ValueError("이 작업에는 검토 자료가 없습니다 (이전 버전에서 실행된 작업).")
    return {"job_id": job["id"], "file": job["file"], **review}


@mcp.tool()
def list_krx_companies(query: str = "", december_only: bool = False) -> dict[str, Any]:
    """오늘 상장돼 있는 회사 명단을 KRX 에서 직접 받아 업종·주요제품과 함께 준다.

    이 PC 에서 KRX 에 바로 물어보므로 **상장폐지가 반영된 진짜 오늘 명단**이다.
    dcfpeer 의 peergroup_get_population_latest 가 KRX 403 으로 폴백 중일 때
    이 도구를 대신 쓴다. 역할 구분: **모집단 명단은 여기(실시간)**, 분기말 스냅샷
    재현성·사업내용 요약은 dcfpeer, 재무·공시는 mydart.

    재현되지 않는다 — 조서에 쓸 때는 조회일과 목록 자체를 기록한다.

    Args:
        query: 업종명·주요제품·회사명에 부분일치하는 검색어 (예: "반도체", "이차전지").
            **비워서 부르면 업종별 종목수 집계만** 돌려준다 — 무엇으로 좁힐지 고르는 용도.
        december_only: True면 12월 결산 회사만 (결산월이 다르면 비교기간이 어긋난다).
    """
    df = _roster()
    meta: dict[str, Any] = {
        "retrievedAt": datetime.now().strftime("%Y-%m-%d"),
        "listedUniverse": int(len(df)),
        "basis": "오늘 기준 KRX 상장회사목록 — 재현되지 않는다. 조회일과 목록을 기록할 것",
    }

    if december_only:
        df = df[df["SettleMonth"].astype(str).str.contains("12", na=False)]

    q = (query or "").strip()
    if not q:
        counts = df["Sector"].fillna("(업종 미상)").value_counts()
        return {
            "meta": meta,
            "sectors": [{"sector": s, "count": int(n)} for s, n in counts.items()],
            "note": "업종을 고른 뒤 query에 넣어 다시 부르면 회사 목록이 나온다.",
        }

    def _col(name):
        return df[name].fillna("").astype(str) if name in df.columns else ""

    hit = (
        _col("Sector").str.contains(q, regex=False)
        | _col("Industry").str.contains(q, regex=False)
        | _col("Name").str.contains(q, regex=False)
    )
    sub = df[hit].sort_values("Code")
    if len(sub) > MAX_COMPANIES:
        raise RuntimeError(
            f'"{q}" 매칭이 {len(sub)}개입니다(상한 {MAX_COMPANIES}). '
            "업종명을 더 구체적으로 넣거나 december_only 로 좁히세요.")

    companies = []
    for _, r in sub.iterrows():
        settle = str(r.get("SettleMonth") or "").strip()
        companies.append({
            "code": str(r.get("Code") or "").zfill(6),
            "name": str(r.get("Name") or ""),
            "market": str(r.get("Market") or ""),
            "sector": str(r.get("Sector") or ""),
            "product": str(r.get("Industry") or "").strip(),
            "settleMonth": settle,
            "fiscalMonthNot12": bool(settle) and "12" not in settle,
        })
    return {"meta": {**meta, "query": q, "count": len(companies)}, "companies": companies}


# 지수 거래일 기준 연속 결측이 이 이상이면 거래정지 의심 구간으로 본다
GAP_MIN_TRADING_DAYS = 5


@mcp.tool()
def check_trading_gaps(tickers: list[str], years: int = 5) -> dict[str, Any]:
    """베타 관측기간 중 거래정지 이력이 있는지 종목별로 점검한다 — 판단 재료다.

    KOSPI 지수 거래일을 기준으로, 종목 시세가 연속 5거래일 이상 비어 있는
    구간을 거래정지 의심으로 보고한다. 거래정지 구간이 있으면 그 기간의
    수익률이 빠져 베타 신뢰도가 떨어지므로, 피어 선정 때 표시하고 배제 여부는
    사용자가 판단한다. **자동으로 배제하지 않는다.**

    상장이 관측 창보다 늦은 회사(신규상장)는 정지와 구분해 observedFrom 으로
    표시한다. 결측은 거래정지 외에 데이터 누락일 수도 있으니, 이력이 나온
    회사는 공시(매매거래정지)로 확인하라고 안내한다.

    Args:
        tickers: 종목코드 6자리 목록. 최대 30개.
        years: 점검 창(년). 기본 5 — Weekly-2Y만 쓰면 2로 줄여도 된다.
    """
    from datetime import timedelta

    codes = [t.strip() for t in tickers if t and t.strip()]
    if not codes:
        raise RuntimeError("tickers가 비어 있습니다.")
    bad = [t for t in codes if not re.match(r"^\d{6}$", t)]
    if bad:
        raise RuntimeError(f"종목코드는 6자리 숫자입니다: {', '.join(bad)}")
    if len(codes) > MAX_TICKERS:
        raise RuntimeError(f"{len(codes)}개는 너무 많습니다(상한 {MAX_TICKERS}).")

    m = _load()
    end = datetime.now()
    start = end - timedelta(days=365 * max(1, int(years)) + 20)

    market = m.fdr.DataReader("KS11", start, end)
    if market is None or market.empty:
        raise RuntimeError("KOSPI 지수 시계열을 받지 못해 거래일 기준을 만들 수 없습니다.")
    market_days = sorted(d.normalize() for d in market.index)

    results: list[dict[str, Any]] = []
    failed: list[str] = []
    for code in codes:
        try:
            px = m.fdr.DataReader(code, start, end)
        except Exception:
            px = None
        if px is None or px.empty:
            failed.append(code)
            continue
        stock_days = {d.normalize() for d in px.index}
        first = min(stock_days)

        gaps: list[dict[str, Any]] = []
        run: list[Any] = []
        for d in market_days:
            if d < first:
                continue  # 상장 전 — 정지가 아니다
            if d in stock_days:
                if len(run) >= GAP_MIN_TRADING_DAYS:
                    gaps.append({"from": run[0].strftime("%Y-%m-%d"),
                                 "to": run[-1].strftime("%Y-%m-%d"),
                                 "tradingDays": len(run)})
                run = []
            else:
                run.append(d)
        currently_suspended = False
        if len(run) >= GAP_MIN_TRADING_DAYS:  # 창 끝까지 이어진 결측 = 현재 정지 중
            gaps.append({"from": run[0].strftime("%Y-%m-%d"),
                         "to": run[-1].strftime("%Y-%m-%d"),
                         "tradingDays": len(run)})
            currently_suspended = True

        results.append({
            "code": code,
            "observedFrom": first.strftime("%Y-%m-%d"),
            "suspectedHalts": gaps,
            "currentlySuspended": currently_suspended,
            "flag": bool(gaps),
        })

    output: dict[str, Any] = {
        "meta": {
            "windowYears": int(years),
            "threshold": f"KOSPI 지수 거래일 기준 연속 {GAP_MIN_TRADING_DAYS}일 이상 결측",
            "note": "판단 재료다 — 자동 배제하지 말고 사용자에게 구간과 함께 표시하라. "
                    "결측은 데이터 누락일 수도 있으니, 이력이 나온 회사는 매매거래정지 "
                    "공시로 교차 확인하라.",
        },
        "results": results,
    }
    if failed:
        output["failed"] = failed
    return output


def _registered_tools() -> list[str]:
    """서버에 실제로 등록된 도구 이름. 선언(EXPECTED_TOOLS)과 대조해 누락을 잡는다."""
    try:
        import asyncio
        tools = asyncio.run(mcp.list_tools())
        return [t.name for t in tools]
    except Exception:
        # list_tools 를 못 부르는 판이면 등록 여부를 확인할 수 없다 — 빈 목록으로
        # 두면 "전부 빠짐"으로 잘못 보고하므로, 선언분을 그대로 돌려주고 넘어간다.
        return list(EXPECTED_TOOLS)


def _check(name: str, ok: bool | None, detail: str, fix: str = "") -> dict[str, Any]:
    entry = {"항목": name, "결과": {True: "정상", False: "실패", None: "주의"}[ok], "내용": detail}
    if fix and ok is not True:
        entry["조치"] = fix
    return entry


def _diagnose() -> list[dict[str, Any]]:
    """설치·환경을 한 줄씩 점검한다. 인증키 값은 절대 담지 않는다 (길이만)."""
    checks: list[dict[str, Any]] = []

    version = _server_version()
    checks.append(_check(
        "버전", version != "unknown", f"gpcm {version}  ({SERVER_DIR})",
        "manifest.json 을 못 읽었습니다. 설치 폴더가 온전한지 확인하세요."))

    # 파이썬 — 3.14 에서 pydantic 이 깨져 도구가 하나만 등록된 적이 있다
    v = sys.version_info
    checks.append(_check(
        "파이썬", (3, 10) <= (v.major, v.minor) < (3, 14),
        f"{v.major}.{v.minor}.{v.micro}",
        "manifest 의 --python 3.12 지정이 무시되고 있습니다. 확장을 재설치하세요."))

    try:
        import importlib.metadata as _md
        checks.append(_check("MCP SDK", True, _md.version("mcp")))
    except Exception as e:
        checks.append(_check("MCP SDK", False, str(e), "확장을 재설치하세요."))

    registered = _registered_tools()
    missing = [t for t in EXPECTED_TOOLS if t not in registered]
    checks.append(_check(
        "등록된 도구", not missing,
        f"{len(registered)}/{len(EXPECTED_TOOLS)}개"
        + (f" — 빠짐: {', '.join(missing)}" if missing else ""),
        "파이썬 버전을 먼저 보세요. 3.14 면 등록이 중간에 멈춥니다."))

    key = _api_key()
    checks.append(_check("DART 인증키", bool(key),
                         f"있음 ({len(key)}자)" if key else "없음",
                         "확장 설정에서 OpenDART 인증키를 넣으세요 (opendart.fss.or.kr 무료)."))
    ecos = (os.environ.get("ECOS_API_KEY") or "").strip()
    if ecos.startswith("${") and ecos.endswith("}"):
        ecos = ""  # 확장 입력칸이 비면 리터럴이 그대로 들어온다
    checks.append(_check("ECOS 인증키(선택)", True if ecos else None,
                         f"있음 ({len(ecos)}자)" if ecos else "없음 — 회사채 금리 조회 불가",
                         "국고채만 쓰면 없어도 됩니다. 필요하면 ecos.bok.or.kr 에서 무료 발급."))

    out = OUTPUT_DIR
    try:
        out.mkdir(parents=True, exist_ok=True)
        probe = out / ".write-test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        checks.append(_check("출력 폴더", True, str(out)))
    except Exception as e:
        checks.append(_check("출력 폴더", False, f"{out} — {e}",
                             "폴더 권한을 확인하거나 동기화 프로그램을 잠시 멈추세요."))

    checks.append(_check("계산 모듈", M is not None,
                         "예열 완료" if M is not None else "아직 로드 전 (첫 호출 때 수십 초 걸립니다)",
                         "그대로 두면 곧 준비됩니다."))

    if key and M is not None:
        ok, reason = M.check_dart_reachable()
        checks.append(_check("DART 연결", ok, "정상" if ok else f"실패 ({reason})",
                             "사내망·방화벽에서 opendart.fss.or.kr 을 막고 있는지 확인하세요."))
        if ok:
            try:
                corp = M.get_dart_reader(key).find_corp_code("005930")
                good = corp == "00126380"
                checks.append(_check("인증키 조회", good,
                                     "정상 (삼성전자 조회 성공)" if good else f"이상한 응답: {corp}",
                                     "인증키가 승인 대기 중이거나 잘못되었을 수 있습니다."))
            except Exception as e:
                checks.append(_check("인증키 조회", False, str(e)[:120],
                                     "인증키를 다시 확인하세요."))
        try:
            df = M.get_krx_listing()
            n = 0 if df is None else len(df)
            checks.append(_check("KRX 시세 목록", n > 0, f"{n}종목" if n else "조회 실패",
                                 "KRX 접속이 막히면 주가·베타가 안 나옵니다. 잠시 뒤 다시 시도하세요."))
        except Exception as e:
            checks.append(_check("KRX 시세 목록", False, str(e)[:120],
                                 "잠시 뒤 다시 시도하세요."))
    return checks


@mcp.tool()
def gpcm_doctor() -> dict[str, Any]:
    """gpcm 이 제대로 돌 준비가 됐는지 한 번에 점검한다 (설치 직후·문제 발생 시).

    파이썬 버전, MCP SDK, 등록된 도구 수, 인증키, 출력 폴더 쓰기 권한, DART·KRX
    연결을 순서대로 확인하고 실패한 항목에는 조치를 붙여 돌려준다.

    인증키 값은 돌려주지 않는다 — 유무와 길이만 확인한다.
    """
    checks = _diagnose()
    bad = [c for c in checks if c["결과"] == "실패"]
    return {
        "버전": _server_version(),
        "실행 위치": str(SERVER_DIR),
        "판정": "정상" if not bad else f"{len(bad)}개 항목 실패",
        "점검": checks,
        "다음": ("run_gpcm 을 써도 됩니다." if not bad else
                 "실패 항목의 조치를 먼저 처리하세요. 그래도 안 되면 이 결과를 그대로 보여주세요."),
    }


def _selftest() -> int:
    """설치 직후 Claude 없이 점검한다: install_mcp.ps1 이 부른다."""
    print("gpcm-mcp 자체점검")
    _load()
    bad = 0
    for c in _diagnose():
        print(f'  [{c["결과"]}] {c["항목"]}: {c["내용"]}' + (f'  → {c["조치"]}' if c.get("조치") else ''))
        bad += c["결과"] == "실패"
    print("전부 통과했습니다." if not bad else f"{bad}개 항목이 실패했습니다.")
    return 1 if bad else 0


def _prepare_stdio() -> None:
    """mydart 에서 Windows 실증된 stdio 대비책: utf-8 강제 + 라인 버퍼링.

    stdout이 파이프에 물리면 블록 버퍼링이 걸려 initialize 응답이 버퍼에 갇히고,
    한국어 Windows 기본 인코딩(cp949)은 도구 설명의 한글을 깨뜨린다.
    """
    for stream in (sys.stdin, sys.stdout):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")
    if getattr(sys.stdout, "reconfigure", None) is not None:
        sys.stdout.reconfigure(line_buffering=True)


def main() -> None:
    if "--selftest" in sys.argv:
        sys.exit(_selftest())
    _prepare_stdio()
    # 접속(initialize·tools/list)은 즉시 응답하고, 무거운 모듈은 뒤에서 예열한다.
    # 보통 사용자가 첫 도구를 부를 때쯤엔 준비가 끝나 있다.
    # 예열을 몇 초 늦추는 이유: 임포트가 GIL 을 오래 잡아 handshake 응답이
    # 밀리면 클라이언트가 접속을 포기한다. 손잡이부터 잡고 짐을 든다.
    threading.Timer(PREHEAT_DELAY_SEC, _load).start()
    mcp.run()


if __name__ == "__main__":
    main()
