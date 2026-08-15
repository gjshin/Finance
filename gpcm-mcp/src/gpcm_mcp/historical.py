"""모드 2 — 다기간 재무제표 요약. gpcm_kr.py 에서 그대로 옮겼다."""

import numpy as np
import pandas as pd

from .accounts import (PL_REVENUE, _norm_pl, _parse_amount,
                       match_bs_ev_component, pick_pl_value)
from .constants import RCODE_MAP
from .dartio import safe_finstate_all
from .listings import resolve_company_info


def fetch_historical_financials(api_key, target_code_list, periods_to_fetch, dart, status_container, progress_bar, df_krx):
    total = len(target_code_list) * len(periods_to_fetch)
    cnt = 0
    hist_summary = []
    hist_details = []

    for ticker in target_code_list:
        corp_code, _ = resolve_company_info(dart, ticker)
        if not corp_code:
            cnt += len(periods_to_fetch); progress_bar.progress(cnt/total)
            continue
        
        comp_name = dart.company(corp_code).get('corp_name', ticker)

        for p in periods_to_fetch:
            year = p['year']
            hist_qtr = p['qtr']
            plabel = p['label']
            
            if not hist_qtr: 
                req_periods = [(year, '4Q', 'annual')]
            else:
                req_periods = [(year, hist_qtr, 'current_cum')]
            
            # BS & EV Components
            ast, liab, eq = np.nan, np.nan, np.nan
            cash, ibd, noa, nci = 0.0, 0.0, 0.0, 0.0
            
            # PL & CF LTM Aggregator
            pl_agg = {'Revenue': 0.0, 'GrossProfit': 0.0, 'EBIT': 0.0, 'NI': 0.0, 'CFO': 0.0, 'CFI': 0.0, 'CFF': 0.0}
            valid_pl_flags = {'Revenue': False, 'GrossProfit': False, 'EBIT': False, 'NI': False, 'CFO': False, 'CFI': False, 'CFF': False}
            
            used_code_current = 'N/A'
            df_fs_current = None
            
            for req_year, req_qtr, role in req_periods:
                primary = RCODE_MAP.get(req_qtr, '11013')
                fallbacks = [c for c in ['11013', '11014', '11012', '11011'] if c != primary]
                target_qtrs = [primary] + fallbacks
                
                df_fs = None
                used_code = None
                for rcode in target_qtrs:
                    df_fs, _ = safe_finstate_all(dart, corp_code, req_year, rcode, fs_div='CFS')
                    if df_fs is None or df_fs.empty:
                        df_fs, _ = safe_finstate_all(dart, corp_code, req_year, rcode, fs_div='OFS')
                    if df_fs is not None and not df_fs.empty:
                        used_code = rcode
                        break
                
                if df_fs is None or df_fs.empty:
                    continue # Skip if data is missing
                    
                if role in ('current_cum', 'annual'):
                    used_code_current = used_code
                    df_fs_current = df_fs
                    
                temp_pl = {'Revenue': np.nan, 'GrossProfit': np.nan, 'EBIT': np.nan, 'NI': np.nan, 'CFO': np.nan, 'CFI': np.nan, 'CFF': np.nan}
                
                for row_idx, row in df_fs.iterrows():
                    sj = str(row.get('sj_nm', ''))
                    acc = str(row.get('account_nm', '')).strip()
                    aid = str(row.get('account_id', '')).strip()
                    _raw = _parse_amount(str(row.get('thstrm_amount', '')))
                    val_1m = (_raw / 1000000) if _raw is not None else np.nan
                    
                    if pd.isna(val_1m): continue
                    
                    m_key = ""
                    if '상태' in sj and role in ('current_cum', 'annual'):
                        if acc == '자산총계': m_key = 'Assets'
                        elif acc == '부채총계': m_key = 'Liabilities'
                        elif acc == '자본총계': m_key = 'Equity_Total'
                        ev_comp, _ = match_bs_ev_component(acc, aid)
                        if ev_comp:
                            m_key = ev_comp # 'Cash', 'Cash(Option)', 'IBD', 'IBD(Option)', 'NOA', 'NOA(Option)', 'NCI'
                            
                    elif '손익' in sj and role in ('current_cum', 'annual'):
                        n_acc = _norm_pl(acc)
                        if '지배' not in n_acc and '포괄' not in n_acc:
                            if n_acc in PL_REVENUE: m_key = 'Revenue'
                            elif '매출총이익' in acc: m_key = 'GrossProfit'
                            elif '영업이익' in acc: m_key = 'EBIT'
                            elif '당기순이익' in acc or '분기순이익' in acc or '반기순이익' in acc or aid == 'ifrs-full_ProfitLoss': m_key = 'NI'
                            
                    elif '현금' in sj and role in ('current_cum', 'annual'):
                        if '영업활동' in acc and '흐름' in acc: m_key = 'CFO'
                        elif '투자활동' in acc and '흐름' in acc: m_key = 'CFI'
                        elif '재무활동' in acc and '흐름' in acc: m_key = 'CFF'

                    # Store Raw Data for Details Sheet (Only for current period)
                    if role in ('current_cum', 'annual') and val_1m != 0:
                        hist_details.append({
                            'Company': comp_name, 'Ticker': ticker, 'Period': plabel, 'Report': used_code_current,
                            'M_Key': m_key, 'Type': sj, 'Account_ID': aid, 'Account_NM': acc, 
                            'Amount': val_1m, 'Row_Idx': row_idx
                        })
                    
                    if '상태' in sj and role in ('current_cum', 'annual'):
                        if acc == '자산총계': ast = val_1m
                        elif acc == '부채총계': liab = val_1m
                        elif acc == '자본총계': eq = val_1m
                        
                        ev_comp, _ = match_bs_ev_component(acc, aid)
                        if ev_comp:
                            if ev_comp == 'Cash': cash += val_1m
                            elif ev_comp == 'IBD': ibd += val_1m
                            elif ev_comp == 'NCI': nci += val_1m
                            elif ev_comp == 'NOA': noa += val_1m
                            
                    elif '손익' in sj:
                        n_acc = _norm_pl(acc)
                        _raw_pl = pick_pl_value(row, req_qtr)
                        val_pl = (_raw_pl / 1000000) if _raw_pl is not None else np.nan
                        if not pd.isna(val_pl) and '지배' not in n_acc and '포괄' not in n_acc:
                            if pd.isna(temp_pl['Revenue']) and n_acc in PL_REVENUE: temp_pl['Revenue'] = val_pl
                            if pd.isna(temp_pl['GrossProfit']) and '매출총이익' in acc: temp_pl['GrossProfit'] = val_pl
                            if pd.isna(temp_pl['EBIT']) and '영업이익' in acc: temp_pl['EBIT'] = val_pl
                            if pd.isna(temp_pl['NI']) and '당기순이익' in acc: temp_pl['NI'] = val_pl
                            
                    elif '현금' in sj:
                        if pd.isna(temp_pl['CFO']) and '영업활동' in acc and '흐름' in acc: temp_pl['CFO'] = val_1m
                        if pd.isna(temp_pl['CFI']) and '투자활동' in acc and '흐름' in acc: temp_pl['CFI'] = val_1m
                        if pd.isna(temp_pl['CFF']) and '재무활동' in acc and '흐름' in acc: temp_pl['CFF'] = val_1m
                
                # Apply to aggregator
                for k in temp_pl:
                    v = temp_pl[k]
                    if pd.notna(v):
                        pl_agg[k] += v
                        valid_pl_flags[k] = True

            if df_fs_current is None or df_fs_current.empty:
                hist_summary.append({
                    'Company': comp_name, 'Ticker': ticker, 'Period': plabel, 'Report': 'N/A',
                    'Revenue': np.nan, 'GrossProfit': np.nan, 'EBIT': np.nan, 'NI': np.nan,
                    'Assets': np.nan, 'Liabilities': np.nan, 'Equity': np.nan,
                    'CFO': np.nan, 'CFI': np.nan, 'CFF': np.nan,
                    'Cash': np.nan, 'IBD': np.nan, 'NOA': np.nan, 'NCI': np.nan,
                    'Shares': np.nan, 'Price': np.nan, 'MarketCap': np.nan
                })
                cnt += 1; progress_bar.progress(cnt/total)
                continue

            hist_summary.append({
                'Company': comp_name, 'Ticker': ticker, 'Period': plabel, 'Report': used_code_current,
                'Revenue': pl_agg['Revenue'] if valid_pl_flags['Revenue'] else np.nan, 
                'GrossProfit': pl_agg['GrossProfit'] if valid_pl_flags['GrossProfit'] else np.nan, 
                'EBIT': pl_agg['EBIT'] if valid_pl_flags['EBIT'] else np.nan, 
                'NI': pl_agg['NI'] if valid_pl_flags['NI'] else np.nan,
                'Assets': ast, 'Liabilities': liab, 'Equity_Total': eq,
                'CFO': pl_agg['CFO'] if valid_pl_flags['CFO'] else np.nan, 
                'CFI': pl_agg['CFI'] if valid_pl_flags['CFI'] else np.nan, 
                'CFF': pl_agg['CFF'] if valid_pl_flags['CFF'] else np.nan,
                'Cash': cash, 'IBD': ibd, 'NOA': noa, 'NCI': nci
            })
            
            status_container.update(label=f"다기간 재무데이터 수집 중... {comp_name} ({plabel})")
            cnt += 1; progress_bar.progress(cnt/total)
            
    return pd.DataFrame(hist_summary), pd.DataFrame(hist_details)

def calculate_historical_metrics(df_summ):
    if df_summ.empty: return df_summ
    
    for col in ['OPM', 'GPM', 'ROE', 'DebtRatio', 'NetDebt']:
        df_summ[col] = np.nan
        
    for i, row in df_summ.iterrows():
        rev = row.get('Revenue'); ebit = row.get('EBIT'); gp = row.get('GrossProfit'); ni = row.get('NI')
        eq = row.get('Equity_Total'); liab = row.get('Liabilities')
        cash = row.get('Cash', 0.0); ibd = row.get('IBD', 0.0)
        noa = row.get('NOA', 0.0); nci = row.get('NCI', 0.0)
        
        if rev and rev > 0:
            df_summ.at[i, 'OPM'] = ebit / rev if pd.notna(ebit) else np.nan
            df_summ.at[i, 'GPM'] = gp / rev if pd.notna(gp) else np.nan
        if eq and eq > 0:
            df_summ.at[i, 'ROE'] = ni / eq if pd.notna(ni) else np.nan
            if pd.notna(liab): df_summ.at[i, 'DebtRatio'] = liab / eq
                
        nd = (ibd if pd.notna(ibd) else 0) - (cash if pd.notna(cash) else 0)
        df_summ.at[i, 'NetDebt'] = nd

    return df_summ
