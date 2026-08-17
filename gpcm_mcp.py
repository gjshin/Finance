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
from datetime import datetime
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import MCPServer

# gpcm_kr 은 임포트만 해도 Streamlit 사이드바 블록이 실행된다(bare mode 경고 + KRX
# 목록 1회 조회). MCP 는 stdout 이 프로토콜 통로라, 임포트 중 새는 출력이 한 글자만
# 있어도 클라이언트 파서가 깨진다. stderr 로 돌려서 임포트한다.
with contextlib.redirect_stdout(sys.stderr):
    import gpcm_kr as M

mcp = MCPServer("gpcm")

OUTPUT_DIR = Path.home() / "Documents" / "GPCM"

# 사이드바 기본값과 동일하게 유지한다 (gpcm_kr.py 의 위젯 기본값)
BETA_TYPES = ("5Y", "2Y")
MAX_TICKERS = 30  # 티커당 수 초~수십 초 — 이 이상은 앱에서 돌리는 게 낫다

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


def _build_periods(start_period: str, end_period: str | None = None) -> list[str]:
    """UI 의 기간 조립 로직 그대로: 시작~끝 분기를 "YYYY.NQ" 목록으로 편다."""
    end_period = end_period or start_period
    for p in (start_period, end_period):
        if not _PERIOD_RE.match(p):
            raise ValueError(f'기간은 "2025.4Q" 형식입니다: {p}')
    qtrs = ["1Q", "2Q", "3Q", "4Q"]
    sy, sq = M.parse_period(start_period)
    ey, eq = M.parse_period(end_period)
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

        notes = [f"• Base Date: {base_period} ({base_date_str}) | Unit: 억원 (KRW 100M)"] + _NOTES_STATIC
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
) -> dict[str, Any]:
    """GPCM 밸류에이션을 백그라운드로 실행해 엑셀 파일을 만든다.

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

    api_key = (os.environ.get("DART_API_KEY") or "").strip()
    if not api_key:
        raise RuntimeError(
            "DART_API_KEY 환경변수가 없습니다. install_mcp.ps1 을 다시 실행해 "
            "인증키를 등록하세요.")

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
                "beta_type": beta_type,
            },
        }
        _jobs[job_id] = job

    unfiled = [p for p in periods
               if not M.is_period_filed(*M.parse_period(p))]

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


def _selftest() -> int:
    """설치 직후 Claude 없이 연결을 점검한다: install_mcp.ps1 이 부른다."""
    print("gpcm-mcp 자체점검")
    key = (os.environ.get("DART_API_KEY") or "").strip()
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
    mcp.run()


if __name__ == "__main__":
    main()
