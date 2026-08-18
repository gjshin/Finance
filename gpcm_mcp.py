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
    '• Adjusted Beta = 2/3 × Raw Beta + 1/3 × 1',
    '• D/E Ratio = IBD / (Market Cap + 우선주 + NCI)',
    '• Debt Ratio (D/V) = IBD / (Market Cap + 우선주 + IBD + NCI)',
    '• 우선주: BS의 우선주자본금(액면) 기준. 시가총액은 보통주만 반영하므로 자기자본가치에 가산',
    '• Unlevered Beta = Levered Beta / (1 + (1 - Tax Rate) × D/E Ratio)',
    '• Tax Rate: 한국 법인세 한계세율 (지방소득세 포함, 세전순이익 기준, 사업연도별 세율표 적용)',
    '   - FY2023~2025: 2억 이하 9.9% | 2~200억 20.9% | 200~3,000억 23.1% | 3,000억 초과 26.4%',
    '   - FY2026~    : 2억 이하 11.0% | 2~200억 22.0% | 200~3,000억 24.2% | 3,000억 초과 27.5% (2025년 세법개정)',
]

_PERIOD_RE = re.compile(r"^\d{4}\.[1-4]Q$")


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

        wacc_data, avg_debt_ratio = M.calculate_wacc_and_beta(
            p["tickers"], summary, p["tax_rate"], p["rf"], p["mrp"],
            p["size_premium"], p["kd_pretax"], p["beta_type"], fiscal_year=base_year)

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
        job["state"] = "done"
        job["pct"] = 100
    except Exception:
        job["state"] = "failed"
        job["error"] = traceback.format_exc(limit=3)


# --- 시장금리 (WACC 입력값의 근거) ------------------------------------------
# rf·Kd 는 판단이 아니라 기준일의 시장 관측치다. 값만 손으로 넣으면 나중에
# "이 3.3%는 어디서 왔나"에 답할 근거가 산출물에 남지 않는다. 조회해서 값과 함께
# 출처·금리기준일을 돌려주고, 그 문장을 그대로 엑셀 Notes 에 실을 수 있게 한다.
# MRP·size premium 은 조회 대상이 아니다 — 시장에서 관측되지 않는 판단 항목이다.

ECOS_MARKET_RATES = "817Y002"  # 한국은행 ECOS 통계표: 시장금리(일별)
BOND_GRADES = ("AA-", "BBB-")
# 항목코드는 하드코딩하지 않는다. ECOS 가 코드를 바꾸면 조용히 엉뚱한 금리를 물어오게
# 되므로, 통계표의 항목 목록을 받아 이름으로 찾는다.
RATE_KINDS = {
    "rf": ("국고채", "5년"),
    "kd": ("회사채", "3년"),
}
FDR_SYMBOLS = {"rf": "KR5YT=RR"}  # 무키 경로 — 회사채는 FDR 에 없어 rf 만 시도한다


def _ecos_key() -> str:
    """한국은행 ECOS 인증키(선택). 미입력 시 확장이 `${...}` 리터럴을 넘긴다."""
    value = (os.environ.get("ECOS_API_KEY") or "").strip()
    if value.startswith("${") and value.endswith("}"):
        return ""
    return value


def _get_fdr():
    """FinanceDataReader. gpcm_kr 이 이미 올라와 있으면 그것을 쓴다(테스트도 여기 문다)."""
    if M is not None:
        return M.fdr
    with _muted_stdout():
        import FinanceDataReader as fdr
    return fdr


def _rate_from_fdr(kind: str, asof: datetime) -> dict[str, Any] | None:
    """무키 경로. 못 구하면 None — 여기서 조용히 기본값을 만들지 않는다."""
    symbol = FDR_SYMBOLS.get(kind)
    if symbol is None:
        return None
    try:
        df = _get_fdr().DataReader(symbol, (asof - timedelta(days=30)).strftime("%Y-%m-%d"),
                                   asof.strftime("%Y-%m-%d"))
    except Exception:
        return None
    if df is None or len(df) == 0 or "Close" not in getattr(df, "columns", []):
        return None
    row = df.dropna(subset=["Close"]).tail(1)
    if len(row) == 0:
        return None
    return {
        "value": round(float(row["Close"].iloc[0]), 3),
        "rateDate": str(row.index[0])[:10],
        "source": f"FinanceDataReader {symbol}",
    }


def _ecos_get(path: str) -> Any:
    import requests
    response = requests.get(f"https://ecos.bok.or.kr/api/{path}", timeout=20)
    response.raise_for_status()
    payload = response.json()
    if "RESULT" in payload:  # ECOS 는 오류도 200 으로 준다
        raise RuntimeError(payload["RESULT"].get("MESSAGE", "ECOS 오류"))
    return payload


def _ecos_item_code(words: tuple[str, ...], grade: str | None) -> tuple[str, str]:
    """통계표 항목 목록에서 이름으로 찾는다 — 코드를 박아두지 않는다."""
    key = _ecos_key()
    payload = _ecos_get(f"StatisticItemList/{key}/json/kr/1/500/{ECOS_MARKET_RATES}")
    rows = payload.get("StatisticItemList", {}).get("row", [])
    wanted = list(words) + ([grade] if grade else [])
    hits = [r for r in rows if all(w in r.get("ITEM_NAME", "") for w in wanted)]
    if not hits:
        raise RuntimeError(
            f"ECOS 통계표 {ECOS_MARKET_RATES} 에서 '{' '.join(wanted)}' 항목을 못 찾았습니다.")
    hit = min(hits, key=lambda r: len(r.get("ITEM_NAME", "")))
    return hit["ITEM_CODE"], hit["ITEM_NAME"]


def _rate_from_ecos(kind: str, asof: datetime, grade: str | None) -> dict[str, Any] | None:
    if not _ecos_key():
        return None
    code, name = _ecos_item_code(RATE_KINDS[kind], grade)
    start = (asof - timedelta(days=30)).strftime("%Y%m%d")
    payload = _ecos_get(
        f"StatisticSearch/{_ecos_key()}/json/kr/1/100/{ECOS_MARKET_RATES}/D/"
        f"{start}/{asof.strftime('%Y%m%d')}/{code}")
    rows = [r for r in payload.get("StatisticSearch", {}).get("row", [])
            if (r.get("DATA_VALUE") or "").strip()]
    if not rows:
        raise RuntimeError(f"ECOS 에 {start}~{asof:%Y%m%d} 구간 '{name}' 값이 없습니다.")
    last = rows[-1]
    return {
        "value": round(float(last["DATA_VALUE"]), 3),
        "rateDate": f'{last["TIME"][:4]}-{last["TIME"][4:6]}-{last["TIME"][6:8]}',
        "source": f"한국은행 ECOS {ECOS_MARKET_RATES} {name}",
    }


def _fetch_rate(kind: str, asof: datetime, grade: str | None) -> tuple[dict[str, Any] | None, list[str]]:
    """무키(FDR) → ECOS 순으로 시도하고, 실패한 경로를 그대로 남긴다."""
    tried: list[str] = []
    for label, getter in (("FinanceDataReader", lambda: _rate_from_fdr(kind, asof)),
                          ("한국은행 ECOS", lambda: _rate_from_ecos(kind, asof, grade))):
        try:
            got = getter()
        except Exception as exc:
            tried.append(f"{label}: {exc}")
            continue
        if got:
            return got, tried
        tried.append(f"{label}: 값 없음"
                     + ("" if label != "한국은행 ECOS" or _ecos_key() else " (인증키 미설정)"))
    return None, tried


@mcp.tool()
def get_wacc_inputs(as_of: str = "", bond_grade: str = "AA-") -> dict[str, Any]:
    """기준일 시장금리를 조회해 WACC 입력값(무위험이자율·타인자본비용)을 근거와 함께 제시한다.

    run_gpcm 을 부르기 전에 이걸 먼저 불러 rf·kd_pretax 후보와 출처를 확인한다.

    Args:
        as_of: 기준일 YYYY-MM-DD. 비우면 오늘. 그 날짜 이전의 최근 고시치를 쓴다.
        bond_grade: 회사채 신용등급 (AA- 또는 BBB-). 평가대상의 신용도에 맞춘다.

    [쓰는 규칙]
    - 여기서 나온 값을 **그대로 run_gpcm 에 넣지 않는다.** 사용자에게 값과 출처를
      보여주고 확정을 받는다. 시장금리는 후보일 뿐 기준일·만기·등급 선택은 판단이다.
    - mrp(시장위험프리미엄)·size_premium(규모프리미엄)은 여기서 안 나온다. 시장에서
      관측되는 값이 아니라 판단 항목이라, 종전처럼 사용자에게 개별로 묻는다.
    - 조회에 실패하면 값을 지어내지 않고 실패를 알린다. 그때는 사용자가 직접 넣는다.
    - 확정 후 run_gpcm(rate_source=<citation>) 로 넘기면 엑셀 Notes 에 근거가 남는다.
    """
    grade = (bond_grade or "").strip().upper()
    if grade not in BOND_GRADES:
        raise ValueError(f"bond_grade 는 {' 또는 '.join(BOND_GRADES)} 입니다: {bond_grade!r}")
    try:
        asof = datetime.strptime(as_of.strip(), "%Y-%m-%d") if as_of.strip() else datetime.now()
    except ValueError:
        raise ValueError(f"as_of 는 YYYY-MM-DD 형식입니다: {as_of!r}")

    rf, rf_tried = _fetch_rate("rf", asof, None)
    kd, kd_tried = _fetch_rate("kd", asof, grade)
    if rf is None and kd is None:
        raise RuntimeError(
            "시장금리를 한 곳에서도 못 받았습니다. rf·kd_pretax 를 직접 넣어야 합니다.\n"
            "시도한 경로 — " + " / ".join(rf_tried + kd_tried)
            + ("\n한국은행 ECOS 인증키(무료, ecos.bok.or.kr)를 확장 설정에 넣으면 "
               "이 경로가 열립니다." if not _ecos_key() else ""))

    parts: list[str] = []
    result: dict[str, Any] = {
        "asOf": asof.strftime("%Y-%m-%d"),
        "fetchedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    for key, got, label, tried in (("rf", rf, "국고채 5년", rf_tried),
                                   ("kd_pretax", kd, f"회사채 3년 {grade}", kd_tried)):
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
    tax_rate: float = 26.4,
    beta_type: str = "5Y",
    rate_source: str = "",
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
         (2천억 미만 4.02% / 2천억~2조 1.37% / 2조 초과 -0.36% / 미적용 0)
      4. 세전 타인자본비용 kd_pretax (기본 3.5%)
      5. 타겟 법인세율 tax_rate (기본 26.4%)
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
        size_premium: 사이즈 프리미엄 %. 기본 4.02 (3분위 Micro, 시총 2천억 미만).
            2,000~20,000억은 1.37, 2조 초과는 -0.36, 미적용은 0.
        kd_pretax: 세전 타인자본비용 %. 기본 3.5
        tax_rate: 타겟 법인세율 %. 기본 26.4
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
                "kd_pretax": kd_pretax / 100, "tax_rate": tax_rate / 100,
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


@mcp.tool()
def gpcm_status(job_id: str | None = None) -> dict[str, Any]:
    """run_gpcm 작업의 진행 상황과 결과를 확인한다.

    Args:
        job_id: run_gpcm이 돌려준 작업 id. 생략하면 가장 최근 작업.
    """
    with _lock:
        if not _jobs:
            raise RuntimeError("실행한 작업이 없습니다. run_gpcm 부터 부르세요.")
        job = _jobs.get(job_id) if job_id else _jobs[f"job-{_counter}"]
    if job is None:
        raise RuntimeError(f"작업을 찾을 수 없습니다: {job_id}")

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
        result["note"] = (
            "숫자를 쓰기 전에 엑셀의 Data_Quality 시트를 먼저 확인하세요. "
            "ERROR가 있으면 해당 값은 수집 실패로 0이 들어가 있습니다. "
            "D&A는 자동 수집이 안 되므로 EBITDA는 엑셀에서 수기 입력이 필요합니다.")
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


def _selftest() -> int:
    """설치 직후 Claude 없이 연결을 점검한다: install_mcp.ps1 이 부른다."""
    print("gpcm-mcp 자체점검")
    _load()
    key = _api_key()
    print(f"1. DART_API_KEY: {'있음 (' + str(len(key)) + '자)' if key else '없음 ← 실패'}")
    if not key:
        return 1
    ok, reason = M.check_dart_reachable()
    print(f"2. DART 연결: {'정상' if ok else f'실패 ({reason})'}")
    if not ok:
        return 1
    try:
        dart = M.get_dart_reader(key)
        corp = dart.find_corp_code("005930")
        good = corp == "00126380"
        print(f"3. 인증키 조회(삼성전자): {'정상' if good else f'이상한 응답: {corp}'}")
        if not good:
            return 1
    except Exception as e:
        print(f"3. 인증키 조회 실패: {e}")
        return 1
    df = M.get_krx_listing()
    print(f"4. KRX 시세 목록: {'정상 (' + str(len(df)) + '종목)' if df is not None and not df.empty else '실패 — 주가·베타가 안 나올 수 있음'}")
    print("전부 통과했습니다. Claude Desktop을 켜고 run_gpcm 을 써보세요.")
    return 0


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
