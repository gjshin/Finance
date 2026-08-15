# GPCM MCP 서버 (국내 상장사)

`gpcm_kr.py` 가 브라우저에서 하던 계산을 **Claude 가 직접 부를 수 있게** 만든 것.

배수를 계산하려면 지금은 브라우저를 띄우고, 어디선가 고른 종목코드를 손으로 옮겨 적고,
결과가 쓸 만한지는 엑셀을 열어 `Data_Quality` 시트를 읽어야 안다. 이 서버는 그 가운데를
도구로 만들어 다음처럼 이어 붙인다.

```
dcfpeer(peer 선정) → gpcm-mcp(계산) → mydart(보완 리서치)
```

**Streamlit 앱은 그대로 있다.** `run_kr.bat` 을 쓰던 사람은 아무것도 달라지지 않는다.

---

## 계산은 앱과 같다

계산 로직을 다시 짜지 않았다. `gpcm_kr.py` 의 줄 범위를 그대로 떠서 모듈로 나눴고,
바꾼 것은 import 와 `@st.cache_resource` 세 개뿐이다.

같은 숫자가 나온다는 걸 말로 하지 않고 기계로 확인한다. `tests/test_parity.py` 가
원본과 이식본에 **같은 가짜 DART 응답**을 먹여 이걸 전부 대조한다.

- 수집한 재무상태표·손익계산서·시가총액 행 전부
- 품질 기록 (순서까지 — Data_Quality 행 순서가 여기서 정해진다)
- 주가 시계열 (5년 월간, 2년 주간)
- WACC 14개 항목
- **두 워크북의 모든 셀** (값과 수식 문자열)

원본을 부르는 방법이 조금 특이하다. `gpcm_kr.py` 는 import 만 해도 `st.set_page_config` 를
부르고 KRX 목록을 내려받는다. 그래서 `with st.sidebar:` 앞까지만 잘라서 exec 한다
(`tests/reference.py`). streamlit 도, 네트워크도 필요 없다. **원본 파일은 읽기만 한다.**

---

## 설치

### 윈도우 — `run_mcp_setup.bat` 더블클릭

저장소 맨 위(`run_kr.bat` 옆)에 있다. 준비 도구·설치·키 저장·Claude 연결까지 한 번에 한다.
브라우저 앱은 건드리지 않는다.

**설치가 끝나면 창을 껐다 켜야 한다.** 저장한 키는 그 뒤에 열리는 창부터 적용된다.

### 직접 설치할 때

`mcp` 패키지가 요구하는 starlette 버전이 streamlit 과 충돌한다. **앱과 다른 환경**에
설치해야 한다. 같은 곳에 넣으면 브라우저 앱이 깨진다.

```bash
cd gpcm-mcp
uv venv .venv-mcp
uv pip install --python .venv-mcp -e .
```

DART 인증키를 환경변수로 설정한다. <https://opendart.fss.or.kr> 에서 무료로 발급받는다.

```bash
# Windows (영구 저장 — 새 창부터 적용)
setx OPENDART_API_KEY 발급받은키

# macOS / Linux
export OPENDART_API_KEY=발급받은키
```

**키는 환경변수로만 받는다.** 도구 인자로 받으면 대화 기록과 세션 로그에 영구히 남는다.

### Claude 에 등록

저장소 맨 위에 `.mcp.json` 을 만든다 (`.mcp.json.example` 참고). `run_mcp_setup.bat` 은
이 파일을 알아서 만든다.

```json
{
  "mcpServers": {
    "gpcm-kr": {
      "command": "C:/GPCM/gpcm-mcp/.venv-mcp/Scripts/python.exe",
      "args": ["-m", "gpcm_mcp.server"]
    }
  }
}
```

경로는 본인 PC 것으로 바꾼다. **JSON 에서는 역슬래시 대신 슬래시(`/`)를 쓴다.**
macOS·Linux 는 `.venv-mcp/bin/python` 이다.

키를 이 파일에 적지 않는다 — 환경변수로 둬야 파일이 새어 나가도 키는 남지 않는다.

---

## 국내에서만 동작한다

**DART 가 해외 IP 를 차단한다.** 회사 PC 나 국내 개인 PC 에서는 문제없다.

`check_dart_access` 를 먼저 부르면 몇 분씩 멈추는 대신 바로 알려준다.

---

## 도구

| 도구 | 하는 일 |
|---|---|
| `check_dart_access` | DART 도달 여부와 키 설정 여부. **제일 먼저 부른다** |
| `latest_filed_period` | 공시가 끝난 최신 분기. 기준일을 잘못 잡는 것 방지 |
| `list_peer_candidates` | KRX 업종·후보 조회 (결산월 확인용) |
| `gpcm_valuation` | **모드 1** — 배수와 WACC. 단위 억원 |
| `historical_financials` | **모드 2** — 다기간 재무제표 요약. 단위 백만원 |
| `gpcm_job_status` | 진행 상황 / 결과 |
| `gpcm_job_cancel` | 조회 중단 |

### 오래 걸리는 조회

회사 수 × 기간 수만큼 DART 를 두드린다. 10 개사 × 8 분기면 수백 번이고 몇 분이 걸린다.

`gpcm_valuation` 은 작업을 띄우고 `wait_seconds`(기본 25초) 만큼 기다린다. 그 안에 끝나면
결과를 바로 주고 — 2~3 개사 조회는 보통 여기서 끝난다 — 안 끝나면 `job_id` 를 준다.
그 뒤로는 `gpcm_job_status` 로 확인한다.

작업자는 하나뿐이다. 계산 코드가 회사마다 `time.sleep(0.5)` 로 DART 호출 간격을 벌리는데,
둘이 동시에 돌면 그 간격이 반으로 준다.

---

## 산출물

엑셀 파일은 `~/Documents/GPCM_Reports/` 에 저장된다 (`GPCM_MCP_OUTPUT_DIR` 로 바꿀 수 있다).
도구는 절대경로를 돌려준다.

**워크북 안의 숫자는 엑셀 수식이다.** openpyxl 은 계산 결과를 저장하지 않으므로 엑셀로
열어야 값이 보인다. 그래서 도구 응답의 배수·WACC 는 워크북이 아니라 **파이썬이 계산한 값**이다.
같은 계산이라 둘이 다르지 않다.

---

## 원본과 다른 곳 (Known divergences)

숫자를 바꿀 수 있는 것은 전부 보존했다. 크래시와 숫자에 영향 없는 것만 고쳤다.

### 고친 것

| 무엇 | 원본 동작 | 왜 |
|---|---|---|
| **FY2026 세율 수식** | GPCM 시트 31열이 FY2023~25 세율(9.9/20.9/23.1/26.4%)을 하드코딩 | 파이썬은 사업연도별 세율표를 쓴다. **FY2026 이후 기준일로 돌리면 엑셀 세율 열과 파이썬 WACC 가 서로 모순**되고 Unlevered β 열도 어긋난다. 지금이 2026년이라 바로 걸린다. **FY2025 이하에서는 원본과 글자까지 같다** |
| 시트명 정제 | `comp[:31]` 그대로 사용 | 회사명에 `/ [ ] : * ? \` 가 있으면 `ValueError` 로 죽는다 (재현 확인). 앞 31자가 같은 두 회사는 시트가 겹친다. 시트를 만들 때와 Summary 의 SUMIFS 수식이 **같은 이름**을 쓰도록 매핑을 한 번만 만든다 |
| 빈 입력 | `ZeroDivisionError` | 종목이나 기간이 비면 죽는다. 조회 전에 걸러낸다 |
| `beta_type` 오타 | `'5y'` 를 넣으면 조용히 2Y 베타 사용 | 도구 경계에서 `5Y`/`2Y` 만 받는다. 계산 함수는 그대로 |
| 죽은 인자 | `export_gpcm_excel` 의 `base_qtr`·`avg_debt_ratio`·`base_date_str`, 모드2의 `api_key`·`df_krx` | 본문에서 전혀 읽지 않는다 (확인 완료) |
| 모드2 실패 행 키 | 실패 행은 `Equity`, 성공 행은 `Equity_Total` | 유령 all-NaN 열이 생긴다. 실패 행은 어차피 전부 NaN 이라 값은 안 바뀐다 |

### 일부러 그대로 둔 것 (고치면 숫자가 바뀐다)

- **`Equity_Total` + `Equity_P` 이중계상** — 둘 다 보고된 공시에서 `Equity` 버킷에 합산된다.
  파이썬 `Equity` 는 `> 0` 가드로만 쓰이고 엑셀은 `"Equity_Total"` 만 참조하므로
  둘은 이미 서로 다른 값이다. 고치면 WACC 가 달라진다.
- **`_parse_amount` 가 0 을 `None` 으로 취급** — 어떤 행이 상세 시트에 남는지가 달라진다.
- **BS 집계에서 `IBD(Option)`·`NOA(Option)` 제외** — EV 정의가 달라진다.
- **LTM 부호 규약** `당기누계 + 전기연간 − 전년동기` (4Q 는 연간 단독).
- **모드2 보고서코드 폴백** — 요청한 분기 보고서가 없으면 다른 분기로 조용히 대체한다.
  `Report` 열에 실제 쓴 코드가 남으므로 추적은 된다.
- **단위 차이** — 모드1 억원, 모드2 백만원.

---

## 코드가 두 벌이라는 것

`gpcm_kr.py` 를 고치지 않기로 했으므로 계산 로직이 두 곳에 있다. 한쪽 버그 수정이
다른 쪽에 자동으로 가지 않는다.

`tests/test_parity.py` 가 그 격차를 잡는 유일한 장치다. **`gpcm_kr.py` 를 고치면 이 테스트가
깨지고 어느 함수인지 알려준다.** 그때 이식본도 함께 고친다.

포팅한 줄 범위는 `PORTED_FROM.md` 에 적혀 있다.

---

## 개발

```bash
uv run --no-project --with-requirements requirements-dev.txt python -m pytest tests/ -q
```

이 저장소(해외)에서는 실제 DART 조회를 할 수 없다. 모든 테스트가 가짜 응답으로 돈다.
실기기 검증 절차는 `CHECKLIST_KR.md` 에 있다.
