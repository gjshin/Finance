# 어디서 옮겨왔는가

이 패키지는 `gpcm_kr.py` 의 복사본이다. 어느 줄을 어디로 옮겼는지 적어 둔다.
원본을 고칠 일이 생기면 이 표에서 대응하는 파일을 찾아 함께 고치고,
`tests/test_parity.py` 를 돌려 확인한다.

- 원본: `../gpcm_kr.py`
- 포팅 기준 커밋: `4602adf` (Say which companies were considered and which numbers could not be filled)

| 이식본 | 원본 줄 | 내용 |
|---|---|---|
| `constants.py` | 39–52 | RCODE_MAP, QUARTER_INFO, BETA_*_DAYS, MIN_*_PTS, SEV_* |
| `quality.py` | 55–74 | QualityLog |
| `tax.py` | 130–185 | 법인세 구간표, 한계세율, 하마다 언레버링 |
| `periods.py` | 187–224 | parse_period ~ get_ltm_required_periods |
| `listings.py` | 226–326 | KRX 상장·업종 목록, peer 후보, 회사 식별 |
| `prices.py` | 80–128, 328–338 | 주가·지수 시계열, 기준일 종가 |
| `shares.py` | 340–466 | 발행·유통주식수 (DART stockTotqySttus) |
| `accounts.py` | 468–592 | 계정과목 매칭 (IBD/Cash/NOA/PL) |
| `dartio.py` | 593–640, 2585–2601 | DART 호출 래퍼, 도달 확인, 리더 생성 |
| `historical.py` | 667–851 | 모드 2 수집·지표 |
| `gpcm.py` | 1101–1532 | 모드 1 수집, WACC·베타 |
| `excel/styles.py` | 641–662, 1053–1096 | 색·폰트·서식, sc / style_range / add_gpcm_section_row |
| `excel/historical_book.py` | 853–1051 | 모드 2 워크북 |
| `excel/gpcm_book.py` | 1533–2331 | 모드 1 워크북 |
| `orphans.py` | 2377–2388, 2549–2569, 2643, 2645–2661, 2672–2676 | **UI 블록 안에 있던 계산 4토막** |

## 새로 쓴 것 (원본에 대응하는 코드 없음)

`cache.py` `progress.py` `output.py` `summarize.py` `jobs.py` `runner.py`
`server.py` `excel/sheetnames.py`

## 옮기면서 바꾼 것

바꾼 곳은 이게 전부다. 나머지는 줄 단위로 같다.

1. import 문
2. `@st.cache_resource(ttl=3600)` → `@ttl_cached(3600)` (`listings.py` 2곳)
3. `@st.cache_resource` → 지연 싱글턴 (`dartio.get_dart_reader`)
4. 죽은 인자 제거 — 본문에서 읽지 않는 것만
5. GPCM 31열 세율 수식을 사업연도별로 (README 참고, **의도한 유일한 숫자 변경**)
6. 모드 2 시트명 정제 (`excel/sheetnames.py`)
7. 모드 2 실패 행의 `Equity` → `Equity_Total` (값은 전부 NaN 이라 변화 없음)
