# 국내 PC 검증 절차

DART 가 해외 IP 를 막아서, 개발 환경에서는 **실제 조회를 한 번도 하지 못했다.**
가짜 응답으로 계산이 원본과 일치하는 것까지만 확인했다.

아래는 국내 PC 에서 한 번 해봐야 하는 것들이다. 3번이 합격 기준이다.

---

## 0. 준비

```bash
cd gpcm-mcp
uv venv .venv-mcp
uv pip install --python .venv-mcp -e .
set OPENDART_API_KEY=발급받은키
claude mcp add gpcm-kr -- <경로>\.venv-mcp\Scripts\python.exe -m gpcm_mcp.server
```

- [ ] **Claude 가 이 PC 에서 돌고 있는가** — Claude 에게 작업 폴더를 물어 `C:\...` 로
      시작하면 로컬, `/home/...` 이면 클라우드다. 클라우드면 이 PC 의 서버를 볼 수 없고
      DART 도 막힌다. (데스크톱 앱에서 열어도 클라우드 세션일 수 있으니 반드시 확인)
- [ ] 새 창에서 `echo %OPENDART_API_KEY%` → 키가 찍히는가 (설치 직후 창을 껐다 켰는가)
- [ ] `check_dart_access` → `reachable: true`, `api_key_configured: true`

여기서 막히면 방화벽에서 `opendart.fss.or.kr` 을 열어야 한다.

---

## 1. 모드 1 — 앱과 같은 입력

앱의 기본값을 그대로 쓴다.

- 종목코드 `000250`, `039030`, `005290`
- 기간 2024.4Q ~ 2025.4Q, 분기별
- Rf 3.3% / MRP 8% / 규모 3분위 Micro(0.0402) / Beta 5Y / Kd 3.5% / 세율 26.4%

- [ ] 엑셀이 만들어지고 경로가 돌아온다
- [ ] `quality` 에 ERROR·WARN 이 실려 나온다 (있다면)
- [ ] 도구가 돌려준 `wacc.Target_WACC` 가 상식적인 범위다

## 2. 같은 입력으로 Streamlit 앱 실행

`run_kr.bat` 을 띄워 **똑같은 값**을 넣고 엑셀을 받는다.

> 같은 날 실행해야 한다. 베타는 5년/2년 주가 창의 끝이 "오늘" 이라 날짜가 바뀌면 값이 조금 움직인다.

## 3. ★ 두 엑셀 대조 — 합격 기준

```bash
uv run --no-project --with openpyxl==3.1.5 python tools/diff_workbooks.py 앱.xlsx MCP.xlsx
```

- [ ] **FY2025 이하 기준일: 전 항목 일치**
- [ ] FY2026 이후 기준일: `GPCM` 시트 **31열(Tax Rate)과 34·35열(Unlevered β)만** 다르고
      나머지 일치 — 의도한 수정이다 (README 의 "원본과 다른 곳" 참고)

여기서 예상 못 한 차이가 나오면 **거기서 멈추고** 어느 시트·어느 셀인지 알려주시라.

---

## 4. 모드 2

- [ ] 3개사 × 3개년 연간 조회 → `Summary` 시트와 회사별 시트가 만들어진다
- [ ] 앱 결과와 `Summary` 시트 대조
- [ ] **회사명에 `/` 나 `:` 가 든 회사**를 하나 넣어 본다 (앱은 여기서 죽는다)
- [ ] 앞 31자가 같은 회사 둘을 넣어 본다 → 시트가 둘 다 만들어지는가
- [ ] 회사별 시트를 열어 Summary 의 SUMIFS 가 `#REF!` 가 아닌지 확인

## 5. 잘못된 입력

- [ ] 종목코드 없이 호출 → 조회를 시작하지 않고 바로 안내가 나온다
- [ ] `AAPL` 같은 형식 → 6자리 안내
- [ ] 아직 공시 안 된 분기 → 경고가 함께 나온다
- [ ] `beta_type='5y'` → 거부된다

## 6. 소요 시간 기록

`wait_seconds` 기본값과 사용자 안내를 조정하는 데 쓴다.

| 규모 | 시간 |
|---|---|
| 1개사 × 1기간 | |
| 3개사 × 4기간 | |
| 10개사 × 8기간 | |

- [ ] 오래 걸리는 조회에서 `gpcm_job_status` 가 진행률을 제대로 보여주는가
- [ ] `gpcm_job_cancel` 이 실제로 멈추는가

## 7. 앱이 그대로인지

- [ ] `git status` — `gpcm_kr.py` 에 변경이 없다
- [ ] `run_kr.bat` 이 예전처럼 뜬다
