"""네트워크 없이 계산 경로를 통째로 돌리기 위한 가짜 데이터.

DART 는 해외 IP 를 막는다. 그래서 이 저장소에서는 실제 조회를 한 번도 할 수 없고,
모든 검증은 가짜 응답으로 해야 한다. gpcm_kr.py 의 기존 테스트가 쓰던 방식을
그대로 넓혔다 — 손으로 계산하기 쉬운 숫자, 조용히 실패하는 경로를 일부러 만들기.

원본과 이식본 **양쪽에 같은 것을 먹여야** 패리티를 볼 수 있으므로, 여기 있는
것들은 어느 쪽 모듈에도 붙일 수 있도록 모듈 객체를 인자로 받는다.
"""

import numpy as np
import pandas as pd

# 계산 경로에서 갈아끼워야 하는 지점. 원본(gpcm_kr)과 이식본(gpcm_mcp.gpcm)이
# 같은 이름을 자기 네임스페이스에 두고 있어야 이 목록 하나로 양쪽을 몰 수 있다.
SEAMS = (
    'get_krx_listing',
    'get_stock_price',
    'get_outstanding_shares',
    'fetch_pl_df',
    '_get_market_index_data',
    'fdr',
)


# --- 재무제표 --------------------------------------------------------------

def _bs_row(account, aid, amount):
    return {'sj_nm': '연결재무상태표', 'account_nm': account,
            'account_id': aid, 'thstrm_amount': str(amount)}


def _pl_row(account, aid, amount, add=None):
    return {'sj_nm': '연결포괄손익계산서', 'account_nm': account, 'account_id': aid,
            'thstrm_amount': str(amount),
            'thstrm_add_amount': str(add if add is not None else amount)}


# 억원 환산이 깔끔하게 떨어지는 숫자로 잡았다 (1억 = 1e8).
BS = pd.DataFrame([
    _bs_row('현금및현금성자산', 'ifrs-full_CashAndCashEquivalents', 100_000_000_000),   # 1,000억
    _bs_row('단기차입금', 'ifrs-full_ShorttermBorrowings', 30_000_000_000),             #   300억
    _bs_row('사채', 'ifrs-full_BondsIssued', 20_000_000_000),                           #   200억
    _bs_row('비지배지분', 'ifrs-full_NoncontrollingInterests', 5_000_000_000),          #    50억
    _bs_row('자본총계', 'ifrs-full_Equity', 600_000_000_000),                           # 6,000억
])

PL = pd.DataFrame([
    _pl_row('매출액', 'ifrs-full_Revenue', 800_000_000_000),                            # 8,000억
    _pl_row('영업이익', 'dart_OperatingIncomeLoss', 80_000_000_000),                    #   800억
    _pl_row('당기순이익', 'ifrs-full_ProfitLoss', 60_000_000_000),                      #   600억
    _pl_row('법인세비용차감전순이익', 'ifrs-full_ProfitLossBeforeTax', 75_000_000_000),  #   750억
])

# 현금흐름표는 모드 2 에서만 쓰인다. 대분류 셋만 M_Key 가 붙고 나머지는 회사별 시트에만 남는다.
CF = pd.DataFrame([
    {'sj_nm': '연결현금흐름표', 'account_nm': '영업활동현금흐름',
     'account_id': 'ifrs-full_CashFlowsFromUsedInOperatingActivities',
     'thstrm_amount': '90000000000', 'thstrm_add_amount': ''},
    {'sj_nm': '연결현금흐름표', 'account_nm': '투자활동현금흐름',
     'account_id': 'ifrs-full_CashFlowsFromUsedInInvestingActivities',
     'thstrm_amount': '-40000000000', 'thstrm_add_amount': ''},
    {'sj_nm': '연결현금흐름표', 'account_nm': '재무활동현금흐름',
     'account_id': 'ifrs-full_CashFlowsFromUsedInFinancingActivities',
     'thstrm_amount': '-20000000000', 'thstrm_add_amount': ''},
    {'sj_nm': '연결현금흐름표', 'account_nm': '감가상각비',
     'account_id': 'ifrs-full_DepreciationExpense',
     'thstrm_amount': '15000000000', 'thstrm_add_amount': ''},
])

FULL_FS = pd.concat([BS, PL, CF], ignore_index=True)


class FakeDart:
    """조회가 전부 성공하는 DART.

    corp_codes / find_corp_code / company / finstate / finstate_all 만 쓰인다.
    어떤 (연도, 보고서코드, fs_div) 로 물어봤는지 기록해 둔다 — 보고서코드 매핑이
    한 번 뒤집혀서 '연간'을 고르면 1분기 숫자가 나온 적이 있다.
    """

    def __init__(self, names=None, fs=None):
        names = names or {'005930': '테스트전자'}
        self._names = names
        self._fs = FULL_FS if fs is None else fs
        self.asked = []
        codes = sorted(names)
        self.corp_codes = pd.DataFrame({
            'corp_code': [f'{i:08d}' for i in range(1, len(codes) + 1)],
            'corp_name': [names[c] for c in codes],
            'stock_code': codes,
        })
        self._by_ticker = {c: f'{i:08d}' for i, c in enumerate(codes, start=1)}
        self._by_corp = {v: names[k] for k, v in self._by_ticker.items()}

    def find_corp_code(self, ticker):
        return self._by_ticker.get(str(ticker))

    def company(self, corp_code):
        return {'corp_name': self._by_corp.get(corp_code, '알수없음')}

    def finstate(self, corp_code, year, reprt_code='11011'):
        self.asked.append((int(year), str(reprt_code), 'finstate'))
        return self._fs.copy()

    def finstate_all(self, corp_code, year, reprt_code='11011', fs_div='CFS'):
        self.asked.append((int(year), str(reprt_code), fs_div))
        return self._fs.copy() if fs_div == 'CFS' else pd.DataFrame()


def fake_krx_listing(names=None):
    names = names or {'005930': '테스트전자'}
    codes = sorted(names)
    return pd.DataFrame({
        'Code': codes,
        'Name': [names[c] for c in codes],
        'Stocks': [100_000_000] * len(codes),
    })


INDUSTRY = pd.DataFrame([
    {'Code': '005930', 'Name': '삼성전자', 'Market': 'KOSPI', 'Sector': '반도체 제조업',
     'Industry': 'DRAM, NAND Flash', 'SettleMonth': '12월'},
    {'Code': '000660', 'Name': 'SK하이닉스', 'Market': 'KOSPI', 'Sector': '반도체 제조업',
     'Industry': 'DRAM', 'SettleMonth': '12월'},
    {'Code': '111111', 'Name': '삼월결산㈜', 'Market': 'KOSDAQ', 'Sector': '반도체 제조업',
     'Industry': '반도체 장비', 'SettleMonth': '3월'},
    {'Code': '222222', 'Name': '딴업종㈜', 'Market': 'KOSDAQ', 'Sector': '음식료품',
     'Industry': '라면', 'SettleMonth': '12월'},
])


# --- 주가 시계열 ------------------------------------------------------------
#
# 베타는 값 자체보다 "원본과 이식본이 같은 값을 내는가" 가 중요하다. 그래서
# 난수 씨앗을 고정한 결정론적 시계열을 쓴다. 같은 입력이면 언제 돌려도 같은 베타가 나온다.

def price_frame(end_date, days, seed, drift=0.0004, vol=0.012, beta=1.0):
    """거래일 기준 종가 프레임. 시장 대비 beta 배로 움직이는 종목을 만든다."""
    idx = pd.bdate_range(end=pd.to_datetime(end_date), periods=days)
    rng = np.random.RandomState(seed)
    market_returns = rng.normal(drift, vol, size=days)
    returns = market_returns * beta
    close = 100.0 * np.cumprod(1.0 + returns)
    return pd.DataFrame({'Close': close}, index=idx)


def fake_price_readers(end_date, days=1300, beta=1.2):
    """(fdr 대역, _get_market_index_data 대역) 을 만들어 돌려준다.

    시장지수와 종목이 같은 난수열에서 나오므로 관계가 고정된다.
    """
    market = price_frame(end_date, days, seed=7, beta=1.0)
    stock = price_frame(end_date, days, seed=7, beta=beta)

    class FakeFdr:
        @staticmethod
        def DataReader(ticker, start=None, end=None, *a, **k):
            df = stock
            if start is not None:
                df = df[df.index >= pd.to_datetime(start)]
            if end is not None:
                df = df[df.index <= pd.to_datetime(end)]
            return df.copy()

        @staticmethod
        def StockListing(market_name):
            return fake_krx_listing()

    def fake_market(market_idx, start, end, cache):
        key = (market_idx, start, end)
        if key not in cache:
            df = market
            if start is not None:
                df = df[df.index >= pd.to_datetime(start)]
            if end is not None:
                df = df[df.index <= pd.to_datetime(end)]
            cache[key] = df.copy()
        return cache[key]

    return FakeFdr, fake_market


# --- 진행 표시 --------------------------------------------------------------

class Silent:
    """status_container / progress_bar 자리에 넣는 아무것도 안 하는 객체."""

    def write(self, *a, **k):
        pass

    def update(self, *a, **k):
        pass

    def progress(self, *a, **k):
        pass


# --- 갈아끼우기 -------------------------------------------------------------

def install(*modules, names=None, shares=1_000_000, price=70_000,
            end_date='2025-03-31', beta=1.2, pl=None):
    """계산 모듈의 네트워크 지점을 전부 가짜로 바꾼다. 되돌리기 함수를 반환한다.

    모듈을 여러 개 받는 이유가 있다. 원본은 모든 함수가 한 모듈에 있어서
    gpcm_kr.get_krx_listing 하나만 갈아끼우면 resolve_company_info 안에서 부르는
    것까지 같이 바뀐다. 이식본은 모듈이 나뉘어 있어서, 이름을 가져다 쓰는 쪽
    (gpcm) 과 정의한 쪽(listings) 을 **둘 다** 갈아끼워야 한다. 한쪽만 바꾸면
    resolve_company_info 가 진짜 KRX 로 나간다 — 이 저장소에서는 그대로 멈춘다.

    각 모듈에 실제로 있는 이름만 바꾸므로, 어느 조합을 넘겨도 안전하다.
    """
    names = names or {'005930': '테스트전자'}
    fake_fdr, fake_market = fake_price_readers(end_date, beta=beta)

    replacements = {
        'get_krx_listing': lambda: fake_krx_listing(names),
        'get_stock_price': lambda t, d: (price, d),
        'get_outstanding_shares': lambda *a, **k: (
            shares,
            'KRX' if shares else 'N/A',
            {'status': '013', 'message': '조회된 데이타가 없습니다.'},
        ),
        'fetch_pl_df': pl or (lambda dart, corp, year, rcode: (PL.copy(), 'CFS', None)),
        '_get_market_index_data': fake_market,
        'fdr': fake_fdr,
    }

    saved = []
    for module in modules:
        for name, value in replacements.items():
            if hasattr(module, name):
                saved.append((module, name, getattr(module, name)))
                setattr(module, name, value)

    def restore():
        for module, name, value in reversed(saved):
            setattr(module, name, value)

    return restore
