"""계정과목 매칭. gpcm_kr.py 에서 그대로 옮겼다.

이 파일의 키워드 집합이 어떤 계정을 IBD/Cash/NOA 로 볼지 결정한다.
한 글자만 달라져도 EV 가 달라지므로 원본과 대조하지 않고는 손대지 않는다.
"""

import re

import pandas as pd


# --- BS Matching Logic ---
IBD_AID_ALWAYS = {
    'ifrs-full_CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings',
    'ifrs-full_LongtermBorrowings',
    'ifrs-full_CurrentLeaseLiabilities',
    'ifrs-full_CurrentPortionOfLongtermBorrowings',
    'ifrs-full_ShorttermBorrowings',
    'ifrs-full_NoncurrentLeaseLiabilities',
    'dart_CurrentPortionOfBonds',
    'ifrs-full_BondsIssued',
    'ifrs-full_Borrowings',
}
IBD_AID_PATTERN = re.compile(r'(Borrowings|Bonds|LeaseLiabilit)', re.IGNORECASE)
MEZZ_KW_KR = ['전환사채', '교환사채', '신주인수권부사채', 'BW', 'CB', 'EB', '전환', '상환', '신주인수', '교환']
MEZZ_KW_EN = ['convertible', 'exchangeable', 'bond with warrant', 'bonds with warrants', 'warrant']
IBD_KW_NAME = ['차입금', '사채', '리스부채', 'Borrowings', 'Bond', 'Bonds', 'LeaseLiabilit', 'Lease Liability']
IBD_EXCLUDE = [
    '매입채무', '미지급', '충당', '선수', '예수', '보증금',
    '자산', '대여금', '미수', '매출채권', '미수금', '미수수익',
    '선급', '선급금', '선급비용', '예치금', '보증금',
    '리스채권', '대여', '대출금(자산)',
]

def _norm(s):
    s = "" if s is None else str(s)
    return re.sub(r"\s+", "", s).strip()

def match_bs_ev_component(account_nm, account_id):
    acct = "" if account_nm is None else str(account_nm).strip()
    aid = "" if account_id is None else str(account_id).strip()
    acct_n = _norm(acct)
    acct_u = acct_n.upper()
    acct_l = acct_n.lower()

    if aid in ['ifrs-full_CashAndCashEquivalents', 'ifrs-full_ShorttermDepositsNotClassifiedAsCashEquivalents']:
        return 'Cash', '현금및단기예금'
    if aid == 'ifrs-full_Equity':
        return 'Equity_Total', '자본총계'
    if aid == 'ifrs-full_EquityAttributableToOwnersOfParent':
        return 'Equity_P', '지배기업지분'
    # 계정명 표기가 회사마다 달라(비지배지분/비지배주주지분/소수주주지분) 표준계정코드를 우선 사용
    if aid == 'ifrs-full_NoncontrollingInterests':
        return 'NCI', '비지배지분'
    # 우선주 자본금: 시가총액(보통주)에 잡히지 않으므로 자기자본가치에 별도 가산
    # '자본금'을 요구해 부채로 분류된 상환우선주(상환전환우선주부채 등)의 중복계상을 방지
    if '우선주' in acct_n and '자본금' in acct_n and '부채' not in acct_n:
        return 'Preferred', acct
    if aid == 'dart_ElementsOfOtherStockholdersEquity':
        return None, None

    if '우선주' not in acct_n:
        mezz_hit = False
        for kw in MEZZ_KW_KR:
            if kw.replace(" ", "") in acct_n: mezz_hit = True; break
        if (not mezz_hit) and any(kw in acct_l for kw in MEZZ_KW_EN): mezz_hit = True
        if (not mezz_hit) and re.search(r'(\bCB\b|\bEB\b|\bBW\b)', acct_u): mezz_hit = True
        if mezz_hit: return 'IBD(Option)', acct

    if not any(ex.replace(" ", "") in acct_n for ex in IBD_EXCLUDE):
        if aid in IBD_AID_ALWAYS: return 'IBD', acct
        if aid and IBD_AID_PATTERN.search(aid): return 'IBD', acct

    if any(k.replace(" ", "") in acct_n for k in IBD_KW_NAME):
        if not any(ex.replace(" ", "") in acct_n for ex in IBD_EXCLUDE):
            return 'IBD', acct

    if (('비지배' in acct_n and '지분' in acct_n) or '소수주주지분' in acct_n) and ('귀속' not in acct):
        return 'NCI', '비지배지분'

    noa_keywords = ['관계기업', '지분법', '공동기업', '종속기업', '금융자산', '금융상품']
    noa_exclude = ['단기', '현금', '매출', '보증금', '미수', '대여금', '예치금', '부채', '충당', '손실', '리스채권']
    if any(kw in acct for kw in noa_keywords) and not any(ex in acct for ex in noa_exclude):
        if aid not in ['ifrs-full_CashAndCashEquivalents', 'ifrs-full_ShorttermDepositsNotClassifiedAsCashEquivalents']:
            return 'NOA(Option)', acct
    return None, None

# --- PL Logic ---
PL_REVENUE = {
    '매출액', '수익(매출액)', '수익(매출)', '영업수익',
    '수익', '매출', '총매출액', '총수익', '영업수익',
    '매출액합계', '수익합계', '총영업수익'
}
PL_EBIT    = {'영업이익', '영업이익(손실)', '영업손실', '영업손익'}
PL_NI      = {
    '당기순이익', '당기순이익(손실)', '당기순손실', '당기순손익',
    '분기순이익', '분기순이익(손실)', '분기순손실', '분기순손익',
    '반기순이익', '반기순이익(손실)', '반기순손실', '반기순손익',
    '연결당기순이익', '연결당기순이익(손실)', '연결당기순손실', '연결당기순손익',
    'ProfitLoss', 'ifrs-full_ProfitLoss'
}
PL_PRETAX_INCOME = {
    '법인세비용차감전순이익', '법인세비용차감전순이익(손실)', '법인세차감전순이익',
    '법인세비용차감전계속사업이익', '법인세비용차감전이익', '세전순이익',
    '법인세비용차감전순손실', '세전이익', '법인세차감전이익'
}

_norm_pl = _norm

def match_pl_core_only(account_nm, aid=None):
    if aid == 'ifrs-full_ProfitLoss': return 'NI'
    a = _norm_pl(account_nm)
    if '지배' in a: return None # Exclude subset (지배기업, 비지배기업)
    if '포괄' in a: return None # Exclude Comprehensive Income
    if a in PL_REVENUE: return 'Revenue'
    if a in PL_EBIT:    return 'EBIT'
    if a in PL_NI:      return 'NI'
    if a in PL_PRETAX_INCOME: return 'Pretax_Income'
    return None

def _parse_amount(x):
    v = pd.to_numeric(str(x).replace(',', ''), errors='coerce')
    if pd.isna(v) or v == 0: return None
    return float(v)

def pick_pl_value(row: pd.Series, qtr: str):
    if qtr == '4Q':
        for col in ['thstrm_amount', 'thstrm_add_amount']:
            v = _parse_amount(row.get(col, ''))
            if v is not None: return v
    else:
        for col in ['thstrm_add_amount', 'thstrm_amount']:
            v = _parse_amount(row.get(col, ''))
            if v is not None: return v
    return None

