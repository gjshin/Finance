"""모드 1 — GPCM 배수와 WACC. gpcm_kr.py 에서 그대로 옮겼다.

조회 지점(get_stock_price, get_outstanding_shares, fetch_pl_df,
_get_market_index_data, fdr)은 일부러 이 모듈의 이름으로 바인딩한다.
원본과 같은 자리를 갈아끼워야 패리티 테스트가 양쪽을 똑같이 몰 수 있다.
"""

import time
from datetime import datetime, timedelta

import FinanceDataReader as fdr
import numpy as np
import pandas as pd

from .accounts import match_bs_ev_component, match_pl_core_only, pick_pl_value
from .constants import (BETA_2Y_DAYS, BETA_5Y_DAYS, MIN_MONTHLY_PTS,
                        MIN_WEEKLY_PTS, RCODE_MAP, SEV_ERROR, SEV_WARN)
from .dartio import fetch_pl_df, filter_income_statement
from .listings import get_krx_listing, resolve_company_info
from .prices import _get_market_index_data, _slice_from, _to_price_series, get_market_index, get_stock_price
from .periods import get_base_date_str, get_ltm_required_periods, parse_period
from .quality import QualityLog
from .shares import get_outstanding_shares
from .tax import get_korean_marginal_tax_rate


def fetch_financial_data(api_key_input, target_code_list, target_periods, dart, status_container, progress_bar):
    df_krx = get_krx_listing()
    
    # 변수 초기화
    base_period_str = target_periods[-1]
    base_year, base_qtr = parse_period(base_period_str)
    base_date_str = get_base_date_str(base_year, base_qtr)

    raw_bs_rows = []
    raw_pl_rows = []
    all_mkt = []
    ticker_to_name = {}

    screen_summary_data = []
    all_multiples = []
    quality = QualityLog()

    total_tickers = len(target_code_list)
    dart_fs_cache = {}  # DART API Call 최소화를 위한 캐시 (ticker 포함 키로 충돌 방지)
    market_idx_cache = {}  # 시장지수 시계열 캐시 (전 종목 공통)

    for idx, ticker in enumerate(target_code_list):
        status_container.write(f"Processing [{ticker}] ({idx+1}/{total_tickers})...")
        progress_bar.progress((idx) / total_tickers)

        corp_code, krx_name = resolve_company_info(dart, ticker)
        if not corp_code:
            status_container.write(f"❌ [{ticker}] DART 고유번호 조회 실패")
            quality.add(SEV_ERROR, ticker, '', '회사 조회',
                        'DART 고유번호를 찾지 못해 이 회사는 분석에서 통째로 빠졌습니다. '
                        '종목코드 6자리가 맞는지, 상장사가 맞는지 확인하세요.')
            continue

        display_name = krx_name if krx_name else f"Company_{ticker}"
        ticker_to_name[ticker] = display_name

        # 임시 저장소 (화면 출력용) - 최신 기준일 데이터
        temp_metrics = {
            'Company': display_name, 'Ticker': ticker,
            'Market_Cap': 0, 'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA': 0, 'Equity': 0, 'Preferred': 0,
            'Revenue': 0, 'EBIT': 0, 'NI': 0, 'Pretax_Income': 0,
            'Stock_Monthly_Prices_5Y': None, 'Market_Monthly_Prices_5Y': None,
            'Stock_Weekly_Prices_2Y': None, 'Market_Weekly_Prices_2Y': None,
            'Exchange': 'KRX', 'Market_Index': 'KS11',
        }

        for tp in target_periods:
            tyear, tqtr = parse_period(tp)
            required_periods = get_ltm_required_periods(tyear, tqtr)
            
            period_metrics = {
                'Market_Cap': 0, 'Cash': 0, 'IBD': 0, 'NCI': 0, 'NOA': 0, 'Equity': 0, 'Preferred': 0,
                'Revenue': 0, 'EBIT': 0, 'NI': 0, 'Pretax_Income': 0
            }

            for year, qtr, role in required_periods:
                r_code = RCODE_MAP[qtr]
                bds = get_base_date_str(year, qtr)

                # 1) Market Cap (기준시점만)
                if role in ('current_cum', 'annual'):
                    price, price_date = get_stock_price(ticker, bds)
                    shares, shares_src, sh_meta = get_outstanding_shares(api_key_input, corp_code, ticker, year, r_code, df_krx, cache=dart_fs_cache)

                    mkt_100m = 0
                    if price is not None and shares is not None and shares > 0:
                        mkt_100m = round((price * shares) / 1e8, 1)
                    else:
                        # 시가총액이 0이면 EV·자본구조·베타가 한꺼번에 무너진다
                        missing = []
                        if price is None:
                            missing.append(f'{bds} 종가')
                        if not shares:
                            missing.append('발행주식수')
                        quality.add(SEV_ERROR, ticker, display_name, f'시가총액 {tp}',
                                    f"{' · '.join(missing)}를 못 가져와 시가총액이 0입니다. "
                                    f"EV와 자본구조가 왜곡되니 Market_Cap 시트에서 사유를 확인하세요."
                                    + (f" (DART: {sh_meta.get('message')})" if sh_meta.get('message') else ''))

                    period_metrics['Market_Cap'] = mkt_100m

                    all_mkt.append({
                        'Company': display_name, 'Ticker': ticker, 'Period': tp,
                        'Price_Date': price_date or bds, 'Close': price,
                        'Outstanding_Shares': shares, 'Market_Cap_100M': mkt_100m,
                        'Shares_Source': shares_src, 'Shares_RceptNo': sh_meta.get('rcept_no'),
                        'Shares_StlmDt': sh_meta.get('stlm_dt'), 'Shares_Se': sh_meta.get('se'),
                        'DART_Status': sh_meta.get('status'), 'DART_Message': sh_meta.get('message'),
                    })

                # 2) BS Fetch (finstate_all: 상세) - CFS 우선 → OFS
                if role in ('current_cum', 'annual'):
                    df_all = None
                    cache_key = f"all_{ticker}_{year}_{r_code}"
                    if cache_key in dart_fs_cache:
                        df_all = dart_fs_cache[cache_key]
                    else:
                        for fs in ['CFS', 'OFS']:
                            try:
                                _df = dart.finstate_all(corp_code, year, reprt_code=r_code, fs_div=fs)
                                if _df is not None and not _df.empty:
                                    df_all = _df
                                    dart_fs_cache[cache_key] = _df
                                    break
                            except Exception:
                                continue

                    if df_all is not None and not df_all.empty:
                        df_bs = df_all[df_all['sj_nm'].astype(str).str.contains('상태표|재정상태', na=False)].copy()
                        for _, row in df_bs.iterrows():
                            amt = pd.to_numeric(str(row.get('thstrm_amount', '')).replace(',', ''), errors='coerce')
                            if pd.isna(amt) or amt == 0: continue

                            acct = str(row.get('account_nm', '')).strip()
                            aid = str(row.get('account_id', '')).strip()
                            ev_comp, _ = match_bs_ev_component(acct, aid)

                            if ev_comp:
                                # 화면 출력용 집계
                                val_100m = amt / 1e8
                                if ev_comp == 'Cash': period_metrics['Cash'] += val_100m
                                elif ev_comp == 'IBD': period_metrics['IBD'] += val_100m
                                elif ev_comp == 'NCI': period_metrics['NCI'] += val_100m
                                elif ev_comp == 'NOA': period_metrics['NOA'] += val_100m
                                elif ev_comp == 'Preferred': period_metrics['Preferred'] += val_100m
                                elif ev_comp in ['Equity_Total', 'Equity_P']: period_metrics['Equity'] += val_100m

                            raw_bs_rows.append({
                                'Company': display_name, 'Ticker': ticker, 'Period': tp,
                                'sj_nm': row.get('sj_nm', ''), 'account_nm': acct, 'account_id': aid,
                                'EV_Component': ev_comp or '', 'Amount_100M': amt / 1e8,
                            })
                    else:
                        # 재무상태표를 못 받으면 현금·차입금·자본이 전부 0으로 남는다
                        quality.add(SEV_ERROR, ticker, display_name, f'재무상태표 {tp}',
                                    f'{year}년 {qtr} 재무상태표를 연결·별도 모두 가져오지 못했습니다. '
                                    f'현금·이자부부채·비지배지분·자본이 0으로 처리됩니다.')

                # 3) PL Fetch
                df_is = None
                cache_key_pl = f"pl_{ticker}_{year}_{r_code}"
                if cache_key_pl in dart_fs_cache:
                    df_is, pl_src = dart_fs_cache[cache_key_pl]
                else:
                    df_pl_raw, pl_src, _ = fetch_pl_df(dart, corp_code, year, r_code)
                    if df_pl_raw is not None and not df_pl_raw.empty:
                        df_is = filter_income_statement(df_pl_raw)
                        dart_fs_cache[cache_key_pl] = (df_is, pl_src)
                    
                # LTM = 당기누계 + 전기연간 - 전년동기. 셋 중 하나만 빠져도 합계가
                # 그럴듯하게 틀리므로, 어느 조각이 빠졌는지 남긴다.
                ltm_part = {'current_cum': '당기누계', 'prior_annual': '전기연간',
                            'prior_same_q': '전년동기'}.get(role)
                if df_is is None or df_is.empty:
                    if ltm_part:
                        quality.add(SEV_ERROR, ticker, display_name, f'LTM {tp}',
                                    f'{ltm_part}({year} {qtr}) 손익계산서를 못 가져왔습니다. '
                                    f'이 조각을 뺀 채로 합산되어 매출·영업이익이 실제와 다릅니다.')
                    else:
                        quality.add(SEV_ERROR, ticker, display_name, f'손익계산서 {tp}',
                                    f'{year}년 {qtr} 손익계산서를 가져오지 못해 매출·영업이익이 0입니다.')
                    continue

                wanted = {'Revenue', 'EBIT', 'NI', 'Pretax_Income'}
                picked = set()

                for _, row in df_is.iterrows():
                    acct = str(row.get('account_nm', '')).strip()
                    aid = str(row.get('account_id', '')).strip()
                    calc_key = match_pl_core_only(acct, aid)
                    if not calc_key or calc_key not in wanted: continue
                    if calc_key in picked: continue

                    val = pick_pl_value(row, qtr)
                    if val is None: continue

                    amt_100m = val / 1e8
                    raw_pl_rows.append({
                        'Company': display_name, 'Ticker': ticker, 'Period': tp,
                        'Role': role, 'PL_Source': pl_src, 'account_nm': acct,
                        'calc_key': calc_key, 'Amount_100M': amt_100m,
                    })

                    if role in ('current_cum', 'annual'):
                        period_metrics[calc_key] += amt_100m
                    elif role == 'prior_annual':
                        period_metrics[calc_key] += amt_100m
                    elif role == 'prior_same_q':
                        period_metrics[calc_key] -= amt_100m

                    picked.add(calc_key)
                    if picked == wanted: break

                # 계정과목 표기가 회사마다 달라 매칭에서 빠질 수 있다. 그 계정만 0이 된다.
                unmatched = wanted - picked
                if unmatched:
                    label = {'Revenue': '매출액', 'EBIT': '영업이익',
                             'NI': '당기순이익', 'Pretax_Income': '법인세비용차감전순이익'}
                    quality.add(SEV_WARN, ticker, display_name, f'계정 매칭 {tp}',
                                f"{year} {qtr} 손익계산서에서 "
                                f"{', '.join(label[k] for k in sorted(unmatched))}을(를) "
                                f"찾지 못했습니다. 해당 계정은 0으로 집계됩니다 — "
                                f"PL_Data 시트에서 실제 계정과목명을 확인하세요.")

            # Period loop ends, append to all_multiples
            all_multiples.append({
                'Company': display_name, 'Ticker': ticker, 'Period': tp,
                **period_metrics
            })
            
            # If this is the main base period, update temp_metrics
            if tp == base_period_str:
                temp_metrics.update(period_metrics)

        # 4) Beta Calculation (5Y Monthly, 2Y Weekly)
        exchange, market_idx = get_market_index(ticker)
        temp_metrics['Exchange'] = exchange
        temp_metrics['Market_Index'] = market_idx

        try:
            end_date = base_date_str
            start_5y = (pd.to_datetime(base_date_str) - timedelta(days=BETA_5Y_DAYS)).strftime('%Y-%m-%d')
            start_2y = (pd.to_datetime(base_date_str) - timedelta(days=BETA_2Y_DAYS)).strftime('%Y-%m-%d')

            # 5년 월간 베타 데이터
            # 시장지수는 모든 종목이 동일하므로 종목마다 다시 받지 않고 캐시에서 재사용
            stock_data_5y = fdr.DataReader(ticker, start_5y, end_date)
            market_data_5y = _get_market_index_data(market_idx, start_5y, end_date, market_idx_cache)

            if stock_data_5y is not None and not stock_data_5y.empty and market_data_5y is not None and not market_data_5y.empty:
                stock_prices_5y = _to_price_series(stock_data_5y)
                market_prices_5y = _to_price_series(market_data_5y)

                if not isinstance(stock_prices_5y.index, pd.DatetimeIndex):
                    stock_prices_5y.index = pd.to_datetime(stock_prices_5y.index)
                if stock_prices_5y.index.tz is not None:
                    stock_prices_5y.index = stock_prices_5y.index.tz_localize(None)
                if not isinstance(market_prices_5y.index, pd.DatetimeIndex):
                    market_prices_5y.index = pd.to_datetime(market_prices_5y.index)
                if market_prices_5y.index.tz is not None:
                    market_prices_5y.index = market_prices_5y.index.tz_localize(None)

                stock_monthly_prices = stock_prices_5y.resample('ME').last().dropna()
                market_monthly_prices = market_prices_5y.resample('ME').last().dropna()

                if len(stock_monthly_prices) >= MIN_MONTHLY_PTS and len(market_monthly_prices) >= MIN_MONTHLY_PTS:
                    temp_metrics['Stock_Monthly_Prices_5Y'] = stock_monthly_prices
                    temp_metrics['Market_Monthly_Prices_5Y'] = market_monthly_prices

            # 2년 주간 베타 데이터
            # 2년 구간은 위에서 받은 5년 구간의 부분집합이므로 잘라 쓴다 (재조회 불필요)
            stock_data_2y = _slice_from(stock_data_5y, start_2y)
            market_data_2y = _slice_from(market_data_5y, start_2y)

            if stock_data_2y is not None and not stock_data_2y.empty and market_data_2y is not None and not market_data_2y.empty:
                stock_prices_2y = _to_price_series(stock_data_2y)
                market_prices_2y = _to_price_series(market_data_2y)

                if not isinstance(stock_prices_2y.index, pd.DatetimeIndex):
                    stock_prices_2y.index = pd.to_datetime(stock_prices_2y.index)
                if stock_prices_2y.index.tz is not None:
                    stock_prices_2y.index = stock_prices_2y.index.tz_localize(None)
                if not isinstance(market_prices_2y.index, pd.DatetimeIndex):
                    market_prices_2y.index = pd.to_datetime(market_prices_2y.index)
                if market_prices_2y.index.tz is not None:
                    market_prices_2y.index = market_prices_2y.index.tz_localize(None)

                stock_weekly_prices = stock_prices_2y.resample('W-FRI').last().dropna()
                market_weekly_prices = market_prices_2y.resample('W-FRI').last().dropna()

                if len(stock_weekly_prices) >= MIN_WEEKLY_PTS and len(market_weekly_prices) >= MIN_WEEKLY_PTS:
                    temp_metrics['Stock_Weekly_Prices_2Y'] = stock_weekly_prices
                    temp_metrics['Market_Weekly_Prices_2Y'] = market_weekly_prices

        except Exception as e:
            beta_failed = True
            quality.add(SEV_WARN, ticker, display_name, '베타',
                        f'주가 시계열을 받지 못해 베타를 계산할 수 없습니다 ({type(e).__name__}). '
                        f'이 회사는 WACC 평균에서 빠집니다.')
        else:
            beta_failed = False

        # 조회는 됐지만 관측치가 모자라 시계열을 담지 못한 경우도 조용히 넘어간다
        if not beta_failed and temp_metrics['Stock_Monthly_Prices_5Y'] is None:
            quality.add(SEV_WARN, ticker, display_name, '베타 5Y',
                        f'월간 관측치가 {MIN_MONTHLY_PTS}개에 못 미쳐 5년 월간 베타를 산출하지 않았습니다. '
                        f'상장한 지 얼마 안 된 회사에서 주로 발생합니다.')
        if not beta_failed and temp_metrics['Stock_Weekly_Prices_2Y'] is None:
            quality.add(SEV_WARN, ticker, display_name, '베타 2Y',
                        f'주간 관측치가 {MIN_WEEKLY_PTS}개에 못 미쳐 2년 주간 베타를 산출하지 않았습니다.')

        screen_summary_data.append(temp_metrics)
        time.sleep(0.5) # API 호출 간격 조절

    progress_bar.progress(1.0)
    status_container.update(label="분석 완료!", state="complete", expanded=False)

    # --- 결과 처리 및 엑셀 생성 ---

    return raw_bs_rows, raw_pl_rows, all_mkt, ticker_to_name, screen_summary_data, base_year, base_qtr, base_date_str, all_multiples, quality

def calculate_wacc_and_beta(target_code_list, screen_summary_data, target_tax_rate_input, rf_input, mrp_input, size_premium_input, kd_pretax_input, beta_type_input, fiscal_year=None):
    # 1.5. WACC Calculation (Target 기업용)
    # Beta 시트에서 계산될 Unlevered Beta를 엑셀에서 참조할 것이므로,
    # 여기서는 대략적인 값만 계산 (정확한 값은 엑셀 수식 기반)

    # 피어들의 평균 계산을 위한 준비
    avg_debt_ratios = []
    avg_unlevered_betas_5y = []
    avg_unlevered_betas_2y = []

    for ticker in target_code_list:
        comp_data = next((item for item in screen_summary_data if item["Ticker"] == ticker), None)
        if not comp_data:
            continue

        mkt_cap = comp_data.get('Market_Cap', 0)
        ibd = comp_data.get('IBD', 0)
        nci = comp_data.get('NCI', 0)
        pref = comp_data.get('Preferred', 0)  # 우선주 자본금 (시가총액은 보통주만 반영)
        equity = comp_data.get('Equity', 0)
        pretax_income = comp_data.get('Pretax_Income', 0)

        # Debt Ratio (D/V) = IBD / (Mkt Cap + 우선주 + IBD + NCI)
        total_value = mkt_cap + pref + ibd + nci
        if total_value > 0:
            debt_ratio = ibd / total_value
            avg_debt_ratios.append(debt_ratio)

        # D/E Ratio = IBD / (Mkt Cap + 우선주 + NCI)
        equity_value = mkt_cap + pref + nci
        de_ratio = ibd / equity_value if equity_value > 0 else 0

        # 한계세율 계산 (사업연도별 한국 법인세율표, 지방소득세 포함)
        tax_rate = get_korean_marginal_tax_rate(pretax_income, fiscal_year)
        comp_data['Tax_Rate'] = tax_rate  # 저장 (나중에 Excel 출력용)

        # Beta 계산 (간단히 수익률 기반)
        stock_monthly_5y = comp_data.get('Stock_Monthly_Prices_5Y')
        market_monthly_5y = comp_data.get('Market_Monthly_Prices_5Y')
        stock_weekly_2y = comp_data.get('Stock_Weekly_Prices_2Y')
        market_weekly_2y = comp_data.get('Market_Weekly_Prices_2Y')

        # 5Y Monthly Beta
        if stock_monthly_5y is not None and market_monthly_5y is not None and not stock_monthly_5y.empty and not market_monthly_5y.empty:
            try:
                common_dates = stock_monthly_5y.index.intersection(market_monthly_5y.index)
                if len(common_dates) > MIN_MONTHLY_PTS:
                    stock_ret = stock_monthly_5y.loc[common_dates].pct_change().dropna()
                    market_ret = market_monthly_5y.loc[common_dates].pct_change().dropna()
                    common_idx = stock_ret.index.intersection(market_ret.index)
                    if len(common_idx) > 10:
                        stock_ret_aligned = stock_ret.loc[common_idx]
                        market_ret_aligned = market_ret.loc[common_idx]
                        cov_matrix = np.cov(stock_ret_aligned, market_ret_aligned)
                        beta_raw = cov_matrix[0, 1] / cov_matrix[1, 1] if cov_matrix[1, 1] != 0 else np.nan
                        beta_adj = (2/3) * beta_raw + (1/3) * 1

                        if not np.isnan(beta_adj) and equity > 0:
                            unlevered_beta_5y = beta_adj / (1 + (1 - tax_rate) * de_ratio)
                            avg_unlevered_betas_5y.append(unlevered_beta_5y)
            except Exception:
                pass

        # 2Y Weekly Beta
        if stock_weekly_2y is not None and market_weekly_2y is not None and not stock_weekly_2y.empty and not market_weekly_2y.empty:
            try:
                common_dates = stock_weekly_2y.index.intersection(market_weekly_2y.index)
                if len(common_dates) > MIN_WEEKLY_PTS:
                    stock_ret = stock_weekly_2y.loc[common_dates].pct_change().dropna()
                    market_ret = market_weekly_2y.loc[common_dates].pct_change().dropna()
                    common_idx = stock_ret.index.intersection(market_ret.index)
                    if len(common_idx) > 20:
                        stock_ret_aligned = stock_ret.loc[common_idx]
                        market_ret_aligned = market_ret.loc[common_idx]
                        cov_matrix = np.cov(stock_ret_aligned, market_ret_aligned)
                        beta_raw = cov_matrix[0, 1] / cov_matrix[1, 1] if cov_matrix[1, 1] != 0 else np.nan
                        beta_adj = (2/3) * beta_raw + (1/3) * 1

                        if not np.isnan(beta_adj) and equity > 0:
                            unlevered_beta_2y = beta_adj / (1 + (1 - tax_rate) * de_ratio)
                            avg_unlevered_betas_2y.append(unlevered_beta_2y)
            except Exception:
                pass

    # 평균값 계산
    avg_debt_ratio = np.mean(avg_debt_ratios) if avg_debt_ratios else 0.3

    # Beta Type에 따라 선택
    if beta_type_input == "5Y":
        avg_unlevered_beta = np.mean(avg_unlevered_betas_5y) if avg_unlevered_betas_5y else 0.8
    else:
        avg_unlevered_beta = np.mean(avg_unlevered_betas_2y) if avg_unlevered_betas_2y else 0.8

    # Target D/E Ratio 계산
    target_de_ratio = avg_debt_ratio / (1 - avg_debt_ratio) if avg_debt_ratio < 1 else 0

    # Relevered Beta 계산
    target_relevered_beta = avg_unlevered_beta * (1 + (1 - target_tax_rate_input) * target_de_ratio)

    # Ke (자기자본비용) 계산
    target_ke = rf_input + mrp_input * target_relevered_beta + size_premium_input

    # Kd (타인자본비용, 세후)
    kd_aftertax = kd_pretax_input * (1 - target_tax_rate_input)

    # E/V, D/V
    equity_weight = 1 - avg_debt_ratio
    debt_weight = avg_debt_ratio

    # Target WACC
    target_wacc = equity_weight * target_ke + debt_weight * kd_aftertax

    # WACC 데이터 저장
    target_wacc_data = {
        'Rf': rf_input,
        'MRP': mrp_input,
        'Size_Premium': size_premium_input,
        'Avg_Unlevered_Beta': avg_unlevered_beta,
        'Target_Tax_Rate': target_tax_rate_input,
        'Avg_Debt_Ratio': avg_debt_ratio,
        'Target_DE_Ratio': target_de_ratio,
        'Target_Relevered_Beta': target_relevered_beta,
        'Target_Ke': target_ke,
        'Kd_Pretax': kd_pretax_input,
        'Kd_Aftertax': kd_aftertax,
        'Equity_Weight': equity_weight,
        'Debt_Weight': debt_weight,
        'Target_WACC': target_wacc
    }
    return target_wacc_data, avg_debt_ratio

