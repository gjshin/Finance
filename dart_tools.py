"""DART 조회·계산 도구.

MCP 서버(dart_mcp.py)가 쓰는 실제 로직. Streamlit에 의존하지 않아 단독으로 테스트할 수 있다.
금액 파싱과 재시도는 gpcm_kr.py에서 검증된 방식을 그대로 옮겨왔다.

설계 원칙 하나: 비율은 여기서 계산한다. Claude에게 산수를 시키면 실행할 때마다 값이
달라질 수 있어 감사조서에 쓸 수 없다. 대신 어떤 계정과목을 썼는지 함께 돌려준다.
"""

import re
import time
from datetime import datetime, timedelta

import pandas as pd

# 분기 → DART 보고서 코드. 11013=1분기, 11012=반기, 11014=3분기, 11011=사업보고서
RCODE_MAP = {'1Q': '11013', '2Q': '11012', '3Q': '11014', '4Q': '11011'}
QUARTER_END = {'1Q': '03-31', '2Q': '06-30', '3Q': '09-30', '4Q': '12-31'}
FILING_DEADLINE_DAYS = {'1Q': 45, '2Q': 45, '3Q': 45, '4Q': 90}

# report() 로 조회 가능한 정기보고서 주요정보 항목
REPORT_ITEMS = [
    '증자', '배당', '자기주식', '최대주주', '최대주주변동', '소액주주', '임원', '직원',
    '임원개인보수', '임원전체보수', '개인별보수', '타법인출자', '주식총수',
    '회계감사', '감사용역', '회계감사용역계약', '사외이사',
    '채무증권발행', '기업어음미상환', '단기사채미상환', '회사채미상환',
    '신종자본증권미상환', '조건부자본증권미상환',
    '공모자금사용', '사모자금사용', '미등기임원보수',
    '임원전체보수승인', '임원전체보수유형',
]


# event() 로 조회 가능한 주요사항보고서 항목
EVENT_KEYWORDS = [
    '부도발생', '영업정지', '회생절차', '해산사유', '관리절차개시', '관리절차중단',
    '유상증자', '무상증자', '유무상증자', '감자',
    '전환사채발행', '신주인수권부사채발행', '교환사채발행', '조건부자본증권발행',
    '사채권양수', '사채권양도결정',
    '자기주식취득', '자기주식처분', '자기주식취득신탁계약체결', '자기주식취득신탁계약해지',
    '회사합병', '회사분할', '회사분할합병', '주식교환',
    '영업양수', '영업양도', '자산양수도', '유형자산양수', '유형자산양도',
    '타법인증권양수', '타법인증권양도',
    '소송',
    '해외상장결정', '해외상장폐지결정', '해외상장', '해외상장폐지',
]

# regstate() 로 조회 가능한 증권신고서 항목
REGSTATE_KEYWORDS = ['지분증권', '채무증권', '증권예탁증권', '합병', '분할', '주식의포괄적교환이전']


def _rows(df, limit, drop=('corp_code',)):
    """DataFrame → 레코드 목록. 내부 식별자만 걷어내고 접수번호는 남긴다
    (접수번호가 있어야 원문을 열어 근거를 확인할 수 있다)."""
    if df is None or not hasattr(df, 'empty') or df.empty:
        return [], 0
    recs = df.head(limit).to_dict('records')
    return [{k: v for k, v in r.items() if k not in drop} for r in recs], int(len(df))


def _parse_amount(x):
    """DART 금액 문자열 → float. 빈 값/0은 None(미확인)으로 돌려준다.

    조회 결과 표시용. 0원 행은 대부분 의미 없는 빈 줄이라 걸러낸다.
    """
    v = pd.to_numeric(str(x).replace(',', ''), errors='coerce')
    if pd.isna(v) or v == 0:
        return None
    return float(v)


def _parse_amount_strict(x):
    """계산용 파서. 0을 진짜 0으로 남긴다.

    자본잠식으로 자본총계가 0인 회사를 '계정을 찾지 못함'으로 처리하면 안 되기 때문이다.
    빈 값만 None(미확인)이다.
    """
    v = pd.to_numeric(str(x).replace(',', ''), errors='coerce')
    if pd.isna(v):
        return None
    return float(v)


def is_period_filed(year, qtr, asof=None):
    """해당 분기 정기보고서 제출기한이 지났는지 (= 조회 가능한지)"""
    asof = asof or datetime.now()
    qtr_end = pd.to_datetime(f"{year}-{QUARTER_END[qtr]}")
    return (qtr_end + timedelta(days=FILING_DEADLINE_DAYS[qtr])) <= asof


def latest_filed_period(asof=None):
    """지금 시점에서 조회 가능한 가장 최근 (연도, 분기)"""
    asof = asof or datetime.now()
    y, qi = asof.year, 3
    order = ['1Q', '2Q', '3Q', '4Q']
    for _ in range(12):
        if is_period_filed(y, order[qi], asof):
            return y, order[qi]
        qi -= 1
        if qi < 0:
            y, qi = y - 1, 3
    return asof.year - 1, '4Q'


# 계산에 쓰는 계정과목 정의.
# XBRL account_id를 우선 보고, 없으면 계정명으로 찾는다. account_id가 더 안정적이다.
ACCOUNT_SPEC = {
    'Revenue':          (['ifrs-full_Revenue'], ['매출액', '수익(매출액)', '영업수익', '매출']),
    'CostOfSales':      (['ifrs-full_CostOfSales'], ['매출원가']),
    'GrossProfit':      (['ifrs-full_GrossProfit'], ['매출총이익']),
    'OperatingIncome':  (['dart_OperatingIncomeLoss', 'ifrs-full_ProfitLossFromOperatingActivities'], ['영업이익']),
    'NetIncome':        (['ifrs-full_ProfitLoss'], ['당기순이익', '분기순이익', '반기순이익']),
    'InterestExpense':  (['ifrs-full_InterestExpense', 'dart_InterestExpense'], ['이자비용']),
    'Assets':           (['ifrs-full_Assets'], ['자산총계']),
    'CurrentAssets':    (['ifrs-full_CurrentAssets'], ['유동자산']),
    'Liabilities':      (['ifrs-full_Liabilities'], ['부채총계']),
    'CurrentLiabilities': (['ifrs-full_CurrentLiabilities'], ['유동부채']),
    'Equity':           (['ifrs-full_Equity'], ['자본총계']),
    'Receivables':      (['ifrs-full_TradeAndOtherCurrentReceivables', 'dart_ShortTermTradeReceivable'],
                         ['매출채권및기타채권', '매출채권', '매출채권및기타유동채권']),
    'Inventories':      (['ifrs-full_Inventories'], ['재고자산']),
    'Cash':             (['ifrs-full_CashAndCashEquivalents'], ['현금및현금성자산']),
}

# 비율 정의: (표시명, 분자, 분모, 분모에 평균잔액을 쓰는지)
# 평균잔액 = (당기말 + 전기말) / 2. 회전율은 기간 평균을 쓰는 것이 실무 관행이다.
RATIO_SPEC = {
    '매출채권회전율':   ('Revenue', 'Receivables', True),
    '재고자산회전율':   ('CostOfSales', 'Inventories', True),
    '총자산회전율':     ('Revenue', 'Assets', True),
    '부채비율':         ('Liabilities', 'Equity', False),
    '유동비율':         ('CurrentAssets', 'CurrentLiabilities', False),
    '이자보상배율':     ('OperatingIncome', 'InterestExpense', False),
    'ROE':              ('NetIncome', 'Equity', True),
    'ROA':              ('NetIncome', 'Assets', True),
    '매출총이익률':     ('GrossProfit', 'Revenue', False),
    '영업이익률':       ('OperatingIncome', 'Revenue', False),
}

# 비율(%)로 표시할 항목. 나머지는 배수(회/배).
PERCENT_RATIOS = {'부채비율', '유동비율', 'ROE', 'ROA', '매출총이익률', '영업이익률'}


class DartClient:
    """OpenDartReader 래퍼. 재시도와 캐시를 담당한다."""

    def __init__(self, api_key, reader=None):
        self.api_key = api_key          # OpenDartReader가 지원하지 않는 API 직접 호출용
        if reader is not None:
            self.dart = reader
        else:
            import OpenDartReader
            self.dart = OpenDartReader(api_key)
        self._cache = {}
        self._listing = None

    def _call(self, fn, *args, retries=2, **kwargs):
        """일시적 오류는 재시도. 끝내 실패하면 (None, 오류)를 돌려준다."""
        last = None
        for attempt in range(retries + 1):
            try:
                return fn(*args, **kwargs), None
            except Exception as e:
                last = e
                if attempt < retries:
                    time.sleep(0.4)
        return None, last

    def cached(self, key, producer):
        """같은 (회사, 연도, 보고서) 반복 조회를 막는다. 오류는 캐시하지 않는다."""
        if key in self._cache:
            return self._cache[key]
        value, err = producer()
        if err is None:
            self._cache[key] = value
        return value

    def listing(self):
        """KRX 상장종목 목록 (업종 정보 포함). 실패해도 조회는 계속되도록 빈 표를 돌려준다."""
        if self._listing is None:
            try:
                import FinanceDataReader as fdr
                df = fdr.StockListing('KRX')
                self._listing = df if df is not None else pd.DataFrame()
            except Exception:
                self._listing = pd.DataFrame()
        return self._listing

    def corp_code(self, company):
        """회사명 또는 종목코드 → DART 고유번호(corp_code)"""
        key = ('corp_code', str(company))
        if key in self._cache:
            return self._cache[key]
        code, err = self._call(self.dart.find_corp_code, str(company))
        if err is None and code:
            self._cache[key] = code
        return code


# ---------------------------------------------------------------- 도구 1

def _listing_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _search_unlisted(client, query, limit, already):
    """DART 전체 회사목록에서 이름으로 검색 (비상장 포함)."""
    try:
        codes = client.dart.corp_codes
        hit = codes[codes['corp_name'].astype(str).str.contains(str(query), na=False)]
    except Exception:
        return []
    out = []
    for _, r in hit.head(limit * 3).iterrows():
        name = str(r.get('corp_name', ''))
        if name in already:
            continue
        stock = str(r.get('stock_code', '') or '').strip()
        if stock:      # 상장사는 위에서 이미 처리했다
            continue
        out.append({'ticker': None, 'name': name, 'market': None,
                    'sector': None, 'listed': False,
                    'corp_code': str(r.get('corp_code', ''))})
        if len(out) >= limit:
            break
    return out

def find_companies(client, query=None, sector=None, market=None, limit=30):
    """이름·종목코드·업종으로 회사를 찾는다. 다른 모든 도구의 출발점.

    query  : 회사명 일부 또는 6자리 종목코드
    sector : 업종 키워드 (예: '게임', '반도체') — KRX 업종명에서 부분일치
    market : 'KOSPI' | 'KOSDAQ'
    """
    df = client.listing()
    if df.empty:
        # 상장목록을 못 받았으면 DART 회사명 검색으로만 처리
        if not query:
            return {'companies': [], 'note': 'KRX 상장목록을 가져오지 못해 업종 검색을 쓸 수 없습니다.'}
        code = client.corp_code(query)
        return {'companies': ([{'corp_code': code, 'name': query}] if code else []),
                'note': 'KRX 상장목록 없이 DART 이름검색만 사용했습니다.'}

    col_code = _listing_col(df, 'Code', 'Symbol')
    col_name = _listing_col(df, 'Name')
    col_mkt = _listing_col(df, 'Market')
    col_sec = _listing_col(df, 'Sector', 'Industry', 'IndustryName')

    out = df
    if market and col_mkt:
        out = out[out[col_mkt].astype(str).str.upper() == market.upper()]
    if sector and col_sec:
        out = out[out[col_sec].astype(str).str.contains(sector, na=False)]
    if query:
        q = str(query).strip()
        if re.fullmatch(r'\d{6}', q) and col_code:
            out = out[out[col_code].astype(str) == q]
        elif col_name:
            out = out[out[col_name].astype(str).str.contains(q, na=False)]

    hits = out.head(limit)
    companies = []
    for _, r in hits.iterrows():
        companies.append({
            'ticker': str(r[col_code]) if col_code else None,
            'name': str(r[col_name]) if col_name else None,
            'market': str(r[col_mkt]) if col_mkt else None,
            'sector': str(r[col_sec]) if col_sec else None,
            'listed': True,
        })

    # KRX 목록은 상장사만 담고 있다. 딜 업무에서는 비상장 타겟도 찾아야 하므로
    # 이름으로 찾는 경우엔 DART 전체 회사목록(비상장 포함)까지 뒤진다.
    if query and not sector and len(companies) < limit:
        companies.extend(_search_unlisted(client, query, limit - len(companies),
                                          {c['name'] for c in companies}))

    note = None
    if sector and not col_sec:
        note = '상장목록에 업종 컬럼이 없어 업종 조건은 무시했습니다.'
    elif len(out) > limit:
        note = f'조건에 맞는 회사가 {len(out)}개라 상위 {limit}개만 돌려줍니다.'
    return {'companies': companies, 'matched': int(len(out)), 'note': note}


# ---------------------------------------------------------------- 도구 2

def _fetch_statements(client, company, year, quarter, fs_div='CFS'):
    """한 회사·한 기간의 전체 계정 재무제표. 연결(CFS) 없으면 별도(OFS)로 넘어간다."""
    corp = client.corp_code(company)
    if not corp:
        return None, f"'{company}'의 DART 고유번호를 찾지 못했습니다."
    if not is_period_filed(year, quarter):
        ly, lq = latest_filed_period()
        return None, f"{year}년 {quarter}는 아직 공시 전입니다. 조회 가능한 최근 기간은 {ly}년 {lq}입니다."

    rcode = RCODE_MAP[quarter]

    def produce():
        for div in ([fs_div] if fs_div else ['CFS', 'OFS']) + (['OFS'] if fs_div == 'CFS' else []):
            df, err = client._call(client.dart.finstate_all, corp, year, reprt_code=rcode, fs_div=div)
            if err is None and df is not None and not df.empty:
                df = df.copy()
                df['_fs_div'] = div
                return df, None
        return pd.DataFrame(), None

    df = client.cached(('fs', corp, year, rcode, fs_div), produce)
    if df is None or df.empty:
        return None, f"{year}년 {quarter} 재무제표를 찾지 못했습니다."
    return df, None


def get_financial_statements(client, companies, year, quarter='4Q', fs_div='CFS',
                             statement=None, account_filter=None, max_rows=200,
                             summary=False):
    """여러 회사의 전체 계정 재무제표.

    statement      : 'BS'(재무상태표) | 'IS'(손익계산서) | 'CF'(현금흐름표) — 생략하면 전부
    account_filter : 계정명 키워드로 걸러내기 (예: '매출채권')
    """
    if summary:
        return _summary_statements(client, companies, year, quarter, max_rows)

    results = []
    for comp in companies:
        df, err = _fetch_statements(client, comp, year, quarter, fs_div)
        if err:
            results.append({'company': comp, 'error': err})
            continue

        sub = df
        if statement:
            want = {'BS': '상태', 'IS': '손익', 'CF': '현금'}.get(statement.upper(), statement)
            sub = sub[sub['sj_nm'].astype(str).str.contains(want, na=False)]
        if account_filter:
            sub = sub[sub['account_nm'].astype(str).str.contains(account_filter, na=False)]

        rows = []
        for _, r in sub.head(max_rows).iterrows():
            rows.append({
                'statement': str(r.get('sj_nm', '')),
                'account': str(r.get('account_nm', '')).strip(),
                'account_id': str(r.get('account_id', '')).strip(),
                'amount': _parse_amount(r.get('thstrm_amount', '')),
                'amount_cumulative': _parse_amount(r.get('thstrm_add_amount', '')),
                'amount_prior': _parse_amount(r.get('frmtrm_amount', '')),
            })

        results.append({
            'company': comp,
            'year': year, 'quarter': quarter,
            'basis': '연결' if df['_fs_div'].iloc[0] == 'CFS' else '별도',
            'rows': rows,
            'truncated': int(len(sub)) > max_rows,
            'total_rows': int(len(sub)),
        })
    return {'unit': '원', 'results': results}


# ---------------------------------------------------------------- 도구 3

def _pick(df, canonical, quarter):
    """계정과목 하나를 찾아 (당기값, 전기값, 사용한 계정명) 을 돌려준다."""
    ids, names = ACCOUNT_SPEC[canonical]

    def value_of(row):
        # 손익·현금흐름은 분기보고서에서 누적금액 칸을 봐야 한다
        sj = str(row.get('sj_nm', ''))
        if '상태' in sj or quarter == '4Q':
            order = ['thstrm_amount', 'thstrm_add_amount']
        else:
            order = ['thstrm_add_amount', 'thstrm_amount']
        for c in order:
            v = _parse_amount_strict(row.get(c, ''))
            if v is not None:
                return v
        return None

    for aid in ids:
        hit = df[df['account_id'].astype(str).str.strip() == aid]
        if not hit.empty:
            r = hit.iloc[0]
            return value_of(r), _parse_amount_strict(r.get('frmtrm_amount', '')), str(r.get('account_nm', '')).strip()

    for nm in names:
        hit = df[df['account_nm'].astype(str).str.strip() == nm]
        if not hit.empty:
            r = hit.iloc[0]
            return value_of(r), _parse_amount_strict(r.get('frmtrm_amount', '')), str(r.get('account_nm', '')).strip()

    return None, None, None


def compute_ratios(client, companies, years, quarter='4Q', ratios=None, fs_div='CFS'):
    """재무비율을 코드로 계산한다.

    값을 못 구하면 0이 아니라 null을 돌려주고 이유를 남긴다 — 감사에서는
    '0'과 '확인 못 함'이 전혀 다른 의미이기 때문이다.
    사용한 계정과목명을 함께 돌려주므로 조서에 근거를 남길 수 있다.
    """
    wanted = ratios or list(RATIO_SPEC.keys())
    unknown = [r for r in wanted if r not in RATIO_SPEC]
    wanted = [r for r in wanted if r in RATIO_SPEC]

    results = []
    for comp in companies:
        for year in years:
            df, err = _fetch_statements(client, comp, year, quarter, fs_div)
            if err:
                results.append({'company': comp, 'year': year, 'error': err})
                continue

            needed = set()
            for r in wanted:
                num, den, _ = RATIO_SPEC[r]
                needed.update([num, den])
            picked = {c: _pick(df, c, quarter) for c in needed}

            values, sources, missing = {}, {}, {}
            for r in wanted:
                num_key, den_key, use_avg = RATIO_SPEC[r]
                num, _, num_nm = picked[num_key]
                den_cur, den_prior, den_nm = picked[den_key]

                if num is None or den_cur is None:
                    values[r] = None
                    missing[r] = f"{'분자 ' + num_key if num is None else ''}{' / ' if num is None and den_cur is None else ''}{'분모 ' + den_key if den_cur is None else ''} 계정을 찾지 못했습니다."
                    continue

                if use_avg and den_prior is not None:
                    den = (den_cur + den_prior) / 2
                    basis = f'{den_nm} 평균(당기말·전기말)'
                else:
                    den = den_cur
                    basis = f'{den_nm} 기말'

                if den == 0:
                    values[r] = None
                    missing[r] = f'{den_nm}이 0이라 계산할 수 없습니다.'
                    continue

                raw = num / den
                values[r] = round(raw * 100, 2) if r in PERCENT_RATIOS else round(raw, 2)
                sources[r] = {'분자': num_nm, '분모': basis,
                              '단위': '%' if r in PERCENT_RATIOS else ('회' if '회전율' in r else '배')}

            results.append({
                'company': comp, 'year': year, 'quarter': quarter,
                'basis': '연결' if df['_fs_div'].iloc[0] == 'CFS' else '별도',
                'ratios': values,
                'sources': sources,
                'not_available': missing or None,
            })

    out = {'results': results}
    if unknown:
        out['ignored'] = f"모르는 비율은 건너뛰었습니다: {', '.join(unknown)}. 가능한 항목: {', '.join(RATIO_SPEC)}"
    return out


# ---------------------------------------------------------------- 도구 4

def get_report_item(client, companies, item, years, quarter='4Q', max_rows=100):
    """정기보고서 주요정보. 타법인출자·증자·회계감사(감사인/의견)·최대주주 등 29종."""
    if item not in REPORT_ITEMS:
        return {'error': f"'{item}'은(는) 지원하지 않습니다.", 'available': REPORT_ITEMS}

    results = []
    for comp in companies:
        corp = client.corp_code(comp)
        if not corp:
            results.append({'company': comp, 'error': f"'{comp}'의 DART 고유번호를 찾지 못했습니다."})
            continue
        for year in years:
            if not is_period_filed(year, quarter):
                results.append({'company': comp, 'year': year,
                                'error': f'{year}년 {quarter}는 아직 공시 전입니다.'})
                continue
            rcode = RCODE_MAP[quarter]
            df = client.cached(
                ('report', corp, item, year, rcode),
                lambda: client._call(client.dart.report, corp, item, year, rcode))

            if df is None or (hasattr(df, 'empty') and df.empty):
                results.append({'company': comp, 'year': year, 'rows': [],
                                'note': '해당 연도에 공시된 내용이 없습니다.'})
                continue

            recs, total = _rows(df, max_rows)
            results.append({'company': comp, 'year': year, 'rows': recs,
                            'total_rows': total, 'truncated': total > max_rows})
    return {'item': item, 'results': results}


# ---------------------------------------------------------------- 도구 5

def search_filings(client, company=None, start=None, end=None, kind=None, keyword=None, limit=50):
    """공시 검색.

    kind : 'A'정기공시 'B'주요사항보고 'C'발행공시 'D'지분공시 'E'기타
           'F'외부감사관련(감사보고서) 'G'펀드 'H'자산유동화 'I'거래소 'J'공정위
    """
    corp = client.corp_code(company) if company else None
    if company and not corp:
        return {'error': f"'{company}'의 DART 고유번호를 찾지 못했습니다."}

    df, err = client._call(client.dart.list, corp, start=start, end=end, kind=(kind or ''))
    if err is not None:
        return {'error': f'공시 검색에 실패했습니다: {err}'}
    if df is None or df.empty:
        return {'filings': [], 'note': '조건에 맞는 공시가 없습니다.'}

    if keyword:
        df = df[df['report_nm'].astype(str).str.contains(keyword, na=False)]

    rows = []
    for _, r in df.head(limit).iterrows():
        rows.append({
            'rcept_no': str(r.get('rcept_no', '')),
            'date': str(r.get('rcept_dt', '')),
            'company': str(r.get('corp_name', '')),
            'title': str(r.get('report_nm', '')),
            'filer': str(r.get('flr_nm', '')),
            'remark': str(r.get('rm', '')),
        })
    return {'filings': rows, 'total': int(len(df)), 'truncated': int(len(df)) > limit}


# ---------------------------------------------------------------- 도구 6

_SECTION_PAT = re.compile(
    r'(?:^|\n)\s*((?:[IVXⅠ-Ⅹ]+\s*\.|제?\s*\d+\s*[\.장절]|주석\s*\d+|\d+\s*\.)\s*[^\n]{2,60})')


def _document_text(client, rcept_no):
    """공시 원문을 평문으로. 태그를 걷어내고 공백을 정리한다."""
    raw, err = client._call(client.dart.document, str(rcept_no))
    if err is not None or not raw:
        return None, f'원문을 가져오지 못했습니다: {err}'
    text = raw if isinstance(raw, str) else str(raw)
    text = re.sub(r'<[^>]+>', '\n', text)
    text = re.sub(r'&nbsp;?', ' ', text)
    text = re.sub(r'&[a-zA-Z]+;', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n', text)
    return text.strip(), None


def get_filing_document(client, rcept_no, section=None, max_chars=12000):
    """공시 원문에서 필요한 부분만 꺼낸다.

    section 없이 부르면 목차를 돌려준다 — 사업보고서·감사보고서는 수백 페이지라
    통째로 넘기면 쓸 수 없기 때문이다. 목차를 보고 원하는 절을 지정해 다시 부르면 된다.
    """
    text, err = _document_text(client, rcept_no)
    if err:
        return {'error': err}

    heads = [(m.start(), ' '.join(m.group(1).split())) for m in _SECTION_PAT.finditer(text)]

    if not section:
        seen, toc = set(), []
        for _, h in heads:
            if h not in seen:
                seen.add(h)
                toc.append(h)
        return {
            'rcept_no': str(rcept_no),
            'total_chars': len(text),
            'sections': toc[:120],
            'hint': "원하는 절의 제목 일부를 section 인자로 넣어 다시 부르세요. 예: section='사업의 내용', section='주석'",
            'preview': text[:1500],
        }

    starts = [i for i, (_, h) in enumerate(heads) if section in h]
    if not starts:
        pos = text.find(section)
        if pos < 0:
            return {'error': f"'{section}'을(를) 원문에서 찾지 못했습니다. section 없이 불러 목차를 먼저 확인하세요."}
        chunk = text[pos:pos + max_chars]
        return {'rcept_no': str(rcept_no), 'section': section, 'matched_by': '본문 검색',
                'content': chunk, 'truncated': len(text) - pos > max_chars}

    i = starts[0]
    begin = heads[i][0]
    end = heads[i + 1][0] if i + 1 < len(heads) else len(text)
    chunk = text[begin:end]
    return {
        'rcept_no': str(rcept_no),
        'section': heads[i][1],
        'matched_by': '목차',
        'other_matches': [heads[j][1] for j in starts[1:6]] or None,
        'content': chunk[:max_chars],
        'truncated': len(chunk) > max_chars,
    }


# ---------------------------------------------------------------- 도구 7~13

def get_company_profile(client, companies):
    """기업개황 — 업종코드, 설립일, 대표자, 결산월, 주소, 홈페이지."""
    out = []
    for comp in companies:
        corp = client.corp_code(comp)
        if not corp:
            out.append({'company': comp, 'error': f"'{comp}'의 DART 고유번호를 찾지 못했습니다."})
            continue
        info = client.cached(('company', corp), lambda: client._call(client.dart.company, corp))
        if not info:
            out.append({'company': comp, 'error': '기업개황을 가져오지 못했습니다.'})
            continue
        d = dict(info) if isinstance(info, dict) else info
        d.pop('status', None)
        d.pop('message', None)
        out.append({'company': comp, 'profile': d})
    return {'results': out}


def get_major_events(client, companies, keyword, start=None, end=None, limit=50):
    """주요사항보고서 — 부도·회생절차·소송·합병·분할·영업양수도·전환사채발행 등 36종.

    수시공시라 정기보고서에는 없는 사건들이다. 딜 실사와 감사 위험평가에서 쓴다.
    """
    if keyword not in EVENT_KEYWORDS:
        return {'error': f"'{keyword}'은(는) 지원하지 않습니다.", 'available': EVENT_KEYWORDS}

    out = []
    for comp in companies:
        corp = client.corp_code(comp)
        if not corp:
            out.append({'company': comp, 'error': f"'{comp}'의 DART 고유번호를 찾지 못했습니다."})
            continue
        df, err = client._call(client.dart.event, corp, keyword, start, end)
        if err is not None:
            out.append({'company': comp, 'error': f'조회에 실패했습니다: {err}'})
            continue
        rows, total = _rows(df, limit)
        out.append({'company': comp, 'rows': rows, 'total': total,
                    'truncated': total > limit,
                    'note': None if rows else '해당 기간에 공시된 내용이 없습니다.'})
    return {'event': keyword, 'results': out}


def get_registration_statements(client, companies, keyword, start=None, end=None, limit=50):
    """증권신고서 — 지분증권·채무증권·증권예탁증권·합병·분할·주식의포괄적교환이전.

    합병·분할 신고서에는 평가방법과 합병비율 산정근거가 들어 있다.
    """
    if keyword not in REGSTATE_KEYWORDS:
        return {'error': f"'{keyword}'은(는) 지원하지 않습니다.", 'available': REGSTATE_KEYWORDS}

    out = []
    for comp in companies:
        corp = client.corp_code(comp)
        if not corp:
            out.append({'company': comp, 'error': f"'{comp}'의 DART 고유번호를 찾지 못했습니다."})
            continue
        df, err = client._call(client.dart.regstate, corp, keyword, start, end)
        if err is not None:
            out.append({'company': comp, 'error': f'조회에 실패했습니다: {err}'})
            continue
        rows, total = _rows(df, limit)
        out.append({'company': comp, 'rows': rows, 'total': total,
                    'truncated': total > limit,
                    'note': None if rows else '해당 기간에 공시된 내용이 없습니다.'})
    return {'statement': keyword, 'results': out}


def get_shareholding_reports(client, companies, kind='대량보유', limit=100):
    """지분공시.

    kind='대량보유' : 5% 이상 대량보유상황보고 (주식등의 대량보유상황보고서)
    kind='임원주요주주' : 임원·주요주주 특정증권등 소유상황보고
    """
    fn_map = {'대량보유': client.dart.major_shareholders,
              '임원주요주주': client.dart.major_shareholders_exec}
    if kind not in fn_map:
        return {'error': f"'{kind}'은(는) 지원하지 않습니다.", 'available': list(fn_map)}

    out = []
    for comp in companies:
        corp = client.corp_code(comp)
        if not corp:
            out.append({'company': comp, 'error': f"'{comp}'의 DART 고유번호를 찾지 못했습니다."})
            continue
        df, err = client._call(fn_map[kind], corp)
        if err is not None:
            out.append({'company': comp, 'error': f'조회에 실패했습니다: {err}'})
            continue
        rows, total = _rows(df, limit)
        out.append({'company': comp, 'rows': rows, 'total': total, 'truncated': total > limit,
                    'note': None if rows else '공시된 내용이 없습니다.'})
    return {'kind': kind, 'results': out}


def list_filings_by_date(client, date=None, limit=100, keyword=None):
    """특정 날짜에 접수된 공시 전체. 날짜를 비우면 오늘.

    '어제 감사보고서 낸 회사들' 같은 전수 조회에 쓴다.
    """
    df, err = client._call(client.dart.list_date_ex, date)
    if err is not None:
        return {'error': f'조회에 실패했습니다: {err}'}
    if df is None or not hasattr(df, 'empty') or df.empty:
        return {'filings': [], 'note': '해당 날짜에 접수된 공시가 없습니다.'}
    if keyword:
        col = 'report_nm' if 'report_nm' in df.columns else df.columns[0]
        df = df[df[col].astype(str).str.contains(keyword, na=False)]
    rows, total = _rows(df, limit)
    return {'date': date or '오늘', 'filings': rows, 'total': total, 'truncated': total > limit}


def get_filing_attachments(client, rcept_no, match=None):
    """공시에 딸린 하위문서와 첨부파일 목록.

    사업보고서 안의 감사보고서, 감사보고서에 붙은 재무제표처럼
    본문과 별개로 붙어 있는 문서를 찾을 때 쓴다.
    """
    subs, err1 = client._call(client.dart.sub_docs, str(rcept_no), match=match)
    files, err2 = client._call(client.dart.attach_files, str(rcept_no))

    def to_list(df):
        if df is None or not hasattr(df, 'empty') or df.empty:
            return []
        return df.head(50).to_dict('records')

    result = {'rcept_no': str(rcept_no),
              'sub_documents': to_list(subs),
              'attached_files': to_list(files)}
    if err1 is not None and err2 is not None:
        return {'error': f'첨부문서를 가져오지 못했습니다: {err1}'}
    if not result['sub_documents'] and not result['attached_files']:
        result['note'] = '하위문서나 첨부파일이 없습니다.'
    return result


def get_xbrl_taxonomy(client, sj_div):
    """XBRL 표준계정 체계. 계정과목 표준명을 확인할 때 쓴다.

    sj_div 예: 'BS1'(재무상태표) 'IS1'(손익계산서) 'CIS1'(포괄손익) 'CF1'(현금흐름표)
    """
    df, err = client._call(client.dart.xbrl_taxonomy, sj_div)
    if err is not None:
        return {'error': f'조회에 실패했습니다: {err}'}
    rows, total = _rows(df, 300)
    return {'sj_div': sj_div, 'accounts': rows, 'total': total}


def _summary_statements(client, companies, year, quarter, max_rows):
    """요약 재무제표(주요계정). 전체 계정보다 훨씬 가벼워 여러 회사 비교에 적합하다."""
    if not is_period_filed(year, quarter):
        ly, lq = latest_filed_period()
        return {'error': f'{year}년 {quarter}는 아직 공시 전입니다. '
                         f'조회 가능한 최근 기간은 {ly}년 {lq}입니다.'}

    results = []
    for comp in companies:
        corp = client.corp_code(comp)
        if not corp:
            results.append({'company': comp, 'error': f"'{comp}'의 DART 고유번호를 찾지 못했습니다."})
            continue
        rcode = RCODE_MAP[quarter]
        df = client.cached(('finstate', corp, year, rcode),
                           lambda: client._call(client.dart.finstate, corp, year, rcode))
        rows, total = _rows(df, max_rows)
        clean = [{'statement': r.get('sj_nm'), 'account': r.get('account_nm'),
                  'amount': _parse_amount(r.get('thstrm_amount', '')),
                  'amount_prior': _parse_amount(r.get('frmtrm_amount', ''))}
                 for r in rows]
        results.append({'company': comp, 'year': year, 'quarter': quarter,
                        'rows': clean, 'total_rows': total,
                        'note': None if clean else '요약 재무제표를 찾지 못했습니다.'})
    return {'unit': '원', 'summary': True, 'results': results}


# ---------------------------------------------------------------- 도구 14

# OpenDartReader가 지원하지 않는 API는 직접 호출한다.
# (gpcm_kr.py가 주식수 조회를 같은 방식으로 처리한다)
DART_API_BASE = "https://opendart.fss.or.kr/api/"

# 지표 분류 코드. 공식 문서 접근이 막혀 있어 두 개의 독립 구현체를 대조해 정했다.
# 코드가 틀려도 결과는 어긋나지 않는다 — 분류명은 응답에 담겨 오는 idx_cl_nm을 그대로 쓴다.
INDICATOR_CATEGORIES = {
    '수익성': 'M210000',
    '안정성': 'M220000',
    '성장성': 'M230000',
    '활동성': 'M240000',
}


def _dart_get(api_key, endpoint, params, timeout=10):
    """OpenDART API 직접 호출."""
    import requests
    try:
        p = dict(params)
        p['crtfc_key'] = api_key
        r = requests.get(DART_API_BASE + endpoint, params=p, timeout=timeout)
        r.raise_for_status()
        js = r.json()
        status = js.get('status')
        if status == '013':                       # 조회된 데이터 없음
            return [], None
        if status != '000':
            return None, f"DART 응답 {status}: {js.get('message', '')}"
        return js.get('list', []), None
    except Exception as e:
        return None, str(e)


def get_financial_indicators(client, companies, years, quarter='4Q', categories=None):
    """DART가 직접 산출한 주요 재무지표 (수익성·안정성·성장성·활동성).

    compute_ratios가 재무제표에서 직접 계산하는 것과 달리, 이쪽은 DART 공식 산출값이다.
    둘을 나란히 놓고 교차검증할 수 있다.

    회사가 여러 곳이면 다중회사 API(fnlttCmpnyIndx)를, 한 곳이면 단일회사
    API(fnlttSinglIndx)를 쓴다.
    """
    if not client.api_key:
        return {'error': 'API 키가 없어 이 기능을 쓸 수 없습니다.'}

    wanted = categories or list(INDICATOR_CATEGORIES)
    unknown = [c for c in wanted if c not in INDICATOR_CATEGORIES]
    wanted = [c for c in wanted if c in INDICATOR_CATEGORIES]
    if not wanted:
        return {'error': f"모르는 분류입니다: {', '.join(unknown)}",
                'available': list(INDICATOR_CATEGORIES)}

    corp_codes, unresolved = [], []
    for comp in companies:
        code = client.corp_code(comp)
        (corp_codes.append((comp, code)) if code else unresolved.append(comp))

    results = []
    for year in years:
        if not is_period_filed(year, quarter):
            ly, lq = latest_filed_period()
            results.append({'year': year,
                            'error': f'{year}년 {quarter}는 아직 공시 전입니다. '
                                     f'조회 가능한 최근 기간은 {ly}년 {lq}입니다.'})
            continue

        rcode = RCODE_MAP[quarter]
        multi = len(corp_codes) > 1
        endpoint = 'fnlttCmpnyIndx.json' if multi else 'fnlttSinglIndx.json'
        corp_param = ','.join(c for _, c in corp_codes)

        rows, errors = [], []
        for cat in wanted:
            key = ('indx', endpoint, corp_param, year, rcode, cat)
            got = client.cached(key, lambda: _dart_get(
                client.api_key, endpoint,
                {'corp_code': corp_param, 'bsns_year': str(year),
                 'reprt_code': rcode, 'idx_cl_code': INDICATOR_CATEGORIES[cat]}))
            if got is None:
                errors.append(f'{cat} 조회 실패')
                continue
            for r in got:
                rows.append({
                    'company': r.get('corp_name'),
                    # 분류명은 DART 응답값을 그대로 쓴다 (코드 매핑이 틀려도 안전)
                    'category': r.get('idx_cl_nm') or cat,
                    'indicator': r.get('idx_nm'),
                    'value': r.get('idx_val'),
                })

        results.append({'year': year, 'quarter': quarter, 'rows': rows,
                        'errors': errors or None,
                        'note': None if rows else '공시된 지표가 없습니다.'})

    out = {'source': 'DART 공식 산출 재무지표', 'results': results}
    if unresolved:
        out['not_found'] = f"DART 고유번호를 찾지 못한 회사: {', '.join(unresolved)}"
    if unknown:
        out['ignored'] = f"모르는 분류는 건너뛰었습니다: {', '.join(unknown)}"
    return out
