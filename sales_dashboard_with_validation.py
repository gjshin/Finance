import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from io import BytesIO
import plotly.colors

# ==================== 오류 검증 시스템 추가 ====================
class DataValidator:
    """
    대시보드 데이터 오류 검증 클래스
    - 원본 데이터 반영 여부 확인
    - 연도별/제품군별 매출 합계 Cross-check
    - 데이터 완전성 검증
    """
    def __init__(self, df):
        self.df = df
        self.validation_results = {}
        self.errors = []
        self.warnings = []

    def validate_all(self):
        """모든 검증 실행"""
        self.check_data_completeness()
        self.check_revenue_totals()
        self.check_quantity_totals()
        self.check_duplicate_records()
        self.check_negative_values()
        self.check_zero_revenue_quantity()
        self.check_asp_consistency()
        return self.get_validation_report()

    def check_data_completeness(self):
        """데이터 완전성 검증"""
        total_rows = len(self.df)

        # 필수 컬럼 결측치 확인
        required_cols = ['Date', 'Year', 'Product_Group', 'Category', 'Revenue', 'Quantity']
        missing_data = {}

        for col in required_cols:
            if col in self.df.columns:
                missing_count = self.df[col].isna().sum()
                if missing_count > 0:
                    missing_data[col] = missing_count
                    self.warnings.append(f"⚠️ {col} 컬럼에 {missing_count}개 결측치 발견")

        self.validation_results['data_completeness'] = {
            'total_rows': total_rows,
            'missing_data': missing_data,
            'status': 'PASS' if not missing_data else 'WARNING'
        }

    def check_revenue_totals(self):
        """연도별/제품군별 매출 합계 일관성 검증"""
        # 전체 매출
        total_revenue = self.df['Revenue'].sum()

        # 연도별 매출 합계
        yearly_revenue = self.df.groupby('Year')['Revenue'].sum()
        yearly_total = yearly_revenue.sum()

        # 카테고리별 매출 합계
        category_revenue = self.df.groupby('Category')['Revenue'].sum()
        category_total = category_revenue.sum()

        # 제품군별 매출 합계
        product_revenue = self.df.groupby('Product_Group')['Revenue'].sum()
        product_total = product_revenue.sum()

        # 허용 오차 (부동소수점 오차 고려)
        tolerance = 0.01

        # 검증
        revenue_check = {
            'total_revenue': total_revenue,
            'yearly_total': yearly_total,
            'category_total': category_total,
            'product_total': product_total,
            'yearly_match': abs(total_revenue - yearly_total) < tolerance,
            'category_match': abs(total_revenue - category_total) < tolerance,
            'product_match': abs(total_revenue - product_total) < tolerance
        }

        if not revenue_check['yearly_match']:
            self.errors.append(f"❌ 연도별 매출 합계 불일치: 전체 {total_revenue:,.0f} vs 연도별 합계 {yearly_total:,.0f}")

        if not revenue_check['category_match']:
            self.errors.append(f"❌ 카테고리별 매출 합계 불일치: 전체 {total_revenue:,.0f} vs 카테고리별 합계 {category_total:,.0f}")

        if not revenue_check['product_match']:
            self.errors.append(f"❌ 제품군별 매출 합계 불일치: 전체 {total_revenue:,.0f} vs 제품군별 합계 {product_total:,.0f}")

        self.validation_results['revenue_totals'] = revenue_check
        self.validation_results['revenue_totals']['status'] = 'PASS' if all([
            revenue_check['yearly_match'],
            revenue_check['category_match'],
            revenue_check['product_match']
        ]) else 'FAIL'

    def check_quantity_totals(self):
        """수량 합계 일관성 검증"""
        total_quantity = self.df['Quantity'].sum()

        yearly_quantity = self.df.groupby('Year')['Quantity'].sum().sum()
        category_quantity = self.df.groupby('Category')['Quantity'].sum().sum()
        product_quantity = self.df.groupby('Product_Group')['Quantity'].sum().sum()

        tolerance = 0.01

        quantity_check = {
            'total_quantity': total_quantity,
            'yearly_total': yearly_quantity,
            'category_total': category_quantity,
            'product_total': product_quantity,
            'yearly_match': abs(total_quantity - yearly_quantity) < tolerance,
            'category_match': abs(total_quantity - category_quantity) < tolerance,
            'product_match': abs(total_quantity - product_quantity) < tolerance
        }

        if not quantity_check['yearly_match']:
            self.errors.append(f"❌ 연도별 수량 합계 불일치: 전체 {total_quantity:,.0f} vs 연도별 합계 {yearly_quantity:,.0f}")

        if not quantity_check['category_match']:
            self.errors.append(f"❌ 카테고리별 수량 합계 불일치")

        if not quantity_check['product_match']:
            self.errors.append(f"❌ 제품군별 수량 합계 불일치")

        self.validation_results['quantity_totals'] = quantity_check
        self.validation_results['quantity_totals']['status'] = 'PASS' if all([
            quantity_check['yearly_match'],
            quantity_check['category_match'],
            quantity_check['product_match']
        ]) else 'FAIL'

    def check_duplicate_records(self):
        """중복 레코드 검증"""
        # 날짜, 제품군, 카테고리 조합이 중복되는 경우 확인
        key_cols = ['Date', 'Product_Group', 'Category', 'Sub_Product', 'Customer']
        duplicates = self.df[self.df.duplicated(subset=key_cols, keep=False)]

        if len(duplicates) > 0:
            self.warnings.append(f"⚠️ {len(duplicates)}개 중복 레코드 발견 (동일 날짜/제품군/고객)")

        self.validation_results['duplicates'] = {
            'count': len(duplicates),
            'status': 'PASS' if len(duplicates) == 0 else 'WARNING'
        }

    def check_negative_values(self):
        """음수 값 검증"""
        negative_revenue = (self.df['Revenue'] < 0).sum()
        negative_quantity = (self.df['Quantity'] < 0).sum()

        if negative_revenue > 0:
            self.warnings.append(f"⚠️ {negative_revenue}개 음수 매출 발견")

        if negative_quantity > 0:
            self.warnings.append(f"⚠️ {negative_quantity}개 음수 수량 발견")

        self.validation_results['negative_values'] = {
            'negative_revenue': negative_revenue,
            'negative_quantity': negative_quantity,
            'status': 'PASS' if (negative_revenue == 0 and negative_quantity == 0) else 'WARNING'
        }

    def check_zero_revenue_quantity(self):
        """0 값 검증"""
        zero_revenue = (self.df['Revenue'] == 0).sum()
        zero_quantity = (self.df['Quantity'] == 0).sum()
        both_zero = ((self.df['Revenue'] == 0) & (self.df['Quantity'] == 0)).sum()

        if both_zero > 0:
            self.warnings.append(f"⚠️ {both_zero}개 레코드가 매출과 수량이 모두 0 (이미 필터링됨)")

        self.validation_results['zero_values'] = {
            'zero_revenue': zero_revenue,
            'zero_quantity': zero_quantity,
            'both_zero': both_zero,
            'status': 'PASS'
        }

    def check_asp_consistency(self):
        """ASP 일관성 검증 (Revenue / Quantity)"""
        df_nonzero = self.df[self.df['Quantity'] > 0].copy()
        calculated_asp = df_nonzero['Revenue'] / df_nonzero['Quantity']
        stored_asp = df_nonzero['ASP']

        tolerance = 0.01
        asp_mismatch = (abs(calculated_asp - stored_asp) > tolerance).sum()

        if asp_mismatch > 0:
            self.warnings.append(f"⚠️ {asp_mismatch}개 레코드의 ASP 계산 불일치")

        self.validation_results['asp_consistency'] = {
            'mismatch_count': asp_mismatch,
            'status': 'PASS' if asp_mismatch == 0 else 'WARNING'
        }

    def get_validation_report(self):
        """검증 리포트 반환"""
        return {
            'results': self.validation_results,
            'errors': self.errors,
            'warnings': self.warnings,
            'overall_status': 'PASS' if len(self.errors) == 0 else 'FAIL'
        }


def display_validation_report(report):
    """검증 리포트 표시"""
    with st.expander("🔍 데이터 검증 리포트 (Data Validation Report)", expanded=(report['overall_status'] == 'FAIL')):
        # 전체 상태
        if report['overall_status'] == 'PASS':
            st.success("✅ **전체 검증 통과** - 모든 데이터가 정상적으로 반영되었습니다.")
        else:
            st.error("❌ **검증 실패** - 데이터 오류가 발견되었습니다.")

        # 에러 표시
        if report['errors']:
            st.markdown("### ❌ 오류 (Errors)")
            for error in report['errors']:
                st.error(error)

        # 경고 표시
        if report['warnings']:
            st.markdown("### ⚠️ 경고 (Warnings)")
            for warning in report['warnings']:
                st.warning(warning)

        # 상세 검증 결과
        st.markdown("### 📊 상세 검증 결과")

        # 1. 데이터 완전성
        if 'data_completeness' in report['results']:
            result = report['results']['data_completeness']
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**1. 데이터 완전성**: {result['total_rows']:,}개 레코드")
            with col2:
                status_color = "🟢" if result['status'] == 'PASS' else "🟡"
                st.markdown(f"{status_color} {result['status']}")

        # 2. 매출 합계 검증
        if 'revenue_totals' in report['results']:
            result = report['results']['revenue_totals']
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**2. 매출 합계 일관성**: R$ {result['total_revenue']:,.0f}")
                st.caption(f"   - 연도별 합계: R$ {result['yearly_total']:,.0f} {'✓' if result['yearly_match'] else '✗'}")
                st.caption(f"   - 카테고리별 합계: R$ {result['category_total']:,.0f} {'✓' if result['category_match'] else '✗'}")
                st.caption(f"   - 제품군별 합계: R$ {result['product_total']:,.0f} {'✓' if result['product_match'] else '✗'}")
            with col2:
                status_color = "🟢" if result['status'] == 'PASS' else "🔴"
                st.markdown(f"{status_color} {result['status']}")

        # 3. 수량 합계 검증
        if 'quantity_totals' in report['results']:
            result = report['results']['quantity_totals']
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**3. 수량 합계 일관성**: {result['total_quantity']:,.0f}")
                st.caption(f"   - 연도별: {result['yearly_total']:,.0f} {'✓' if result['yearly_match'] else '✗'}")
                st.caption(f"   - 카테고리별: {result['category_total']:,.0f} {'✓' if result['category_match'] else '✗'}")
                st.caption(f"   - 제품군별: {result['product_total']:,.0f} {'✓' if result['product_match'] else '✗'}")
            with col2:
                status_color = "🟢" if result['status'] == 'PASS' else "🔴"
                st.markdown(f"{status_color} {result['status']}")

        # 4. 중복 레코드
        if 'duplicates' in report['results']:
            result = report['results']['duplicates']
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**4. 중복 레코드**: {result['count']}개")
            with col2:
                status_color = "🟢" if result['status'] == 'PASS' else "🟡"
                st.markdown(f"{status_color} {result['status']}")

        # 5. 음수 값
        if 'negative_values' in report['results']:
            result = report['results']['negative_values']
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**5. 음수 값**: 매출 {result['negative_revenue']}개, 수량 {result['negative_quantity']}개")
            with col2:
                status_color = "🟢" if result['status'] == 'PASS' else "🟡"
                st.markdown(f"{status_color} {result['status']}")

        # 6. ASP 일관성
        if 'asp_consistency' in report['results']:
            result = report['results']['asp_consistency']
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**6. ASP 계산 일관성**: {result['mismatch_count']}개 불일치")
            with col2:
                status_color = "🟢" if result['status'] == 'PASS' else "🟡"
                st.markdown(f"{status_color} {result['status']}")


def validate_filtered_data(filtered_df, original_df, selected_years, sel_cat, sel_prod):
    """필터링된 데이터의 일관성 검증 (각 탭에서 사용)"""
    issues = []

    # 필터 조건에 맞는 원본 데이터 추출
    expected_df = original_df[original_df['Year'].isin(selected_years)]
    if sel_cat != "All Categories":
        expected_df = expected_df[expected_df['Category'] == sel_cat]
    if sel_prod != "All Products":
        expected_df = expected_df[expected_df['Product_Group'] == sel_prod]

    # 매출 합계 비교
    expected_revenue = expected_df['Revenue'].sum()
    actual_revenue = filtered_df['Revenue'].sum()

    if abs(expected_revenue - actual_revenue) > 0.01:
        issues.append(f"필터링 후 매출 불일치: 예상 {expected_revenue:,.0f} vs 실제 {actual_revenue:,.0f}")

    # 레코드 수 비교
    expected_count = len(expected_df)
    actual_count = len(filtered_df)

    if expected_count != actual_count:
        issues.append(f"레코드 수 불일치: 예상 {expected_count} vs 실제 {actual_count}")

    return issues

# ==================== 오류 검증 시스템 끝 ====================

# --- 1. 디자인 시스템 & 설정 ---
st.set_page_config(layout="wide", page_title="Integrated Sales Dashboard (MBR)", page_icon="📈")
COLORS = {
    'primary': '#00338D',    # KPMG Blue
    'secondary': '#0091DA',  # Light Blue
    'accent': '#00A3A1',     # Teal
    'negative': '#D32929',   # Red
    'positive': '#009944',   # Green
    'neutral': '#6D6E71',    # Grey
    'sequence': plotly.colors.qualitative.Prism
}
# --- 2. 데이터 로드 및 통합 전처리 ---
@st.cache_data(ttl=3600)
def load_all_data(file_obj):
    try:
        all_dfs = []

        # 1) Bumper
        cols_bumper = "A, C, R, Y, AA, AB, AI"
        df_b = pd.read_excel(file_obj, sheet_name='bumper', usecols=cols_bumper)
        df_b.columns = ['Year_Raw', 'Date', 'Product_Group', 'Sub_Product', 'Customer', 'Quantity', 'Revenue']
        df_b['Category'] = 'Bumper'
        all_dfs.append(df_b)

        # 2) Sill side
        cols_sill = "A, E, I, N, U"
        df_s = pd.read_excel(file_obj, sheet_name='Sill side', usecols=cols_sill)
        df_s.columns = ['Year_Raw', 'Date', 'Product_Group', 'Quantity', 'Revenue']
        df_s['Category'] = 'Sill side'
        df_s['Sub_Product'] = '-'
        df_s['Customer'] = '-'
        all_dfs.append(df_s)

        # 3) Carrier
        cols_carrier = "A, E, I, N, U"
        df_c = pd.read_excel(file_obj, sheet_name='Carrier', usecols=cols_carrier)
        df_c.columns = ['Year_Raw', 'Date', 'Product_Group', 'Quantity', 'Revenue']
        df_c['Category'] = 'Carrier'
        df_c['Sub_Product'] = '-'
        df_c['Customer'] = '-'
        all_dfs.append(df_c)

        # 통합
        df = pd.concat(all_dfs, ignore_index=True)

        # 공통 전처리
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date'])
        df['Year'] = df['Date'].dt.year

        df['Quantity'] = df['Quantity'].fillna(0).astype(float)
        df['Revenue'] = df['Revenue'].fillna(0).astype(float)
        df['Product_Group'] = df['Product_Group'].fillna("Unknown")

        # Clean Data
        df = df[~((df['Revenue'] == 0) & (df['Quantity'] == 0))]

        # 파생 변수
        df['Month_Dt'] = df['Date'].dt.to_period('M').dt.to_timestamp()
        df['Quarter_Str'] = df['Date'].dt.to_period('Q').astype(str)

        # ASP
        df['ASP'] = np.where(df['Quantity'] > 0, df['Revenue'] / df['Quantity'], 0)

        return df

    except ValueError as ve:
        st.error(f"❌ 엑셀 구조 오류: {ve}")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ 데이터 로드 오류: {e}")
        return pd.DataFrame()
# --- 3. 분석 로직 (P/V/M Bridge - Multi Step 수정됨) ---
def calculate_bridge_mix_steps(df, years, group_col):
    """
    Price / Volume / Mix 3단계 Bridge 계산 (다중 연도 루프 지원)
    """
    all_steps = []
    if len(years) < 2: return []
    years = sorted(years)

    # [수정] 연도별 루프 복구 (i, i+1 비교)
    for i in range(len(years) - 1):
        y0, y1 = years[i], years[i+1]

        # 1. 데이터 집계
        d0 = df[df['Year'] == y0].groupby(group_col).agg({'Revenue':'sum', 'Quantity':'sum'}).reset_index()
        d1 = df[df['Year'] == y1].groupby(group_col).agg({'Revenue':'sum', 'Quantity':'sum'}).reset_index()

        # 2. Total 집계
        total_q0 = d0['Quantity'].sum()
        total_q1 = d1['Quantity'].sum()
        total_r0 = d0['Revenue'].sum()
        total_r1 = d1['Revenue'].sum()

        avg_p0 = total_r0 / total_q0 if total_q0 > 0 else 0

        # 3. Merge & Effect Calculation
        merged = pd.merge(d0, d1, on=group_col, how='outer', suffixes=('_0', '_1')).fillna(0)

        merged['P0'] = np.where(merged['Quantity_0'] > 0, merged['Revenue_0'] / merged['Quantity_0'], 0)
        merged['P1'] = np.where(merged['Quantity_1'] > 0, merged['Revenue_1'] / merged['Quantity_1'], 0)

        # (1) Volume Effect
        vol_effect = (total_q1 - total_q0) * avg_p0

        # (2) Price Effect
        merged['Price_Impact'] = (merged['P1'] - merged['P0']) * merged['Quantity_1']
        price_effect = merged['Price_Impact'].sum()

        # (3) Mix Effect
        total_variance = total_r1 - total_r0
        mix_effect = total_variance - vol_effect - price_effect

        all_steps.append({
            'label': f"{y0}→{y1}",
            'start_year': str(y0),
            'end_year': str(y1),
            'volume': vol_effect,
            'price': price_effect,
            'mix': mix_effect,
            'start_val': total_r0,
            'end_val': total_r1
        })

    return all_steps
def create_waterfall_fig_multi(steps):
    """
    다중 단계 Waterfall 차트 생성기 (22->23->24 모두 표시)
    """
    if not steps: return go.Figure()

    # 초기값 (첫 연도 시작)
    first_step = steps[0]

    x_labels = [first_step['start_year']]
    y_vals = [first_step['start_val']]
    measures = ["absolute"]
    text_vals = [f"R$ {first_step['start_val']/1000000:,.1f}M"]

    # 각 스텝별 변동 추가
    for step in steps:
        # Vol / Mix / Price 순서
        for eff_name, eff_val in [("Volume", step['volume']), ("Mix", step['mix']), ("Price", step['price'])]:
            x_labels.append(f"{eff_name}<br>({step['label']})") # 라벨에 연도 표기
            y_vals.append(eff_val)
            measures.append("relative")
            text_vals.append(f"{eff_val/1000000:+,.1f}M")

        # 해당 구간 끝 (다음 구간 시작)
        x_labels.append(step['end_year'])
        y_vals.append(step['end_val'])
        measures.append("absolute") # total 대신 absolute 사용하여 중간 합계 표시
        text_vals.append(f"R$ {step['end_val']/1000000:,.1f}M")

    # 마지막은 total로 처리해도 되지만, 연속성을 위해 absolute로 유지함

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=measures,
        x=x_labels,
        y=y_vals,
        text=text_vals,
        textposition="outside",
        decreasing={"marker":{"color": COLORS['negative']}},
        increasing={"marker":{"color": COLORS['accent']}},
        totals={"marker":{"color": COLORS['primary']}},
        connector={"line":{"color":"#666"}}
    ))

    fig.update_layout(title_text="Multi-Year Price-Volume-Mix Bridge", xaxis_type='category')
    return fig
def detect_outliers_iqr(df, col='ASP', factor=1.5):
    if df.empty: return df
    q1 = df[col].quantile(0.25)
    q3 = df[col].quantile(0.75)
    iqr = q3 - q1
    lower, upper = q1 - (factor * iqr), q3 + (factor * iqr)
    return df[(df[col] < lower) | (df[col] > upper)]
# --- 4. 메인 UI ---
def main():
    st.sidebar.markdown(f"<h1 style='color: {COLORS['primary']};'>KPMG Analysis (MBR)</h1>", unsafe_allow_html=True)

    st.sidebar.markdown("### 📂 Data Upload")
    uploaded_file = st.sidebar.file_uploader("업로드: 매출데이터_MBR.xlsx", type=['xlsx'])

    if uploaded_file is None:
        st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드해주세요.")
        return
    df = load_all_data(uploaded_file)
    if df.empty: return

    # ==================== 데이터 검증 실행 ====================
    if 'validation_report' not in st.session_state:
        validator = DataValidator(df)
        validation_report = validator.validate_all()
        st.session_state['validation_report'] = validation_report
        st.session_state['original_df'] = df.copy()  # 원본 데이터 저장
    else:
        validation_report = st.session_state['validation_report']

    # 검증 리포트 표시
    display_validation_report(validation_report)
    # ==================== 데이터 검증 끝 ====================

    # [사이드바 필터]
    st.sidebar.header("🔍 Global Filters")

    all_years = sorted(df['Year'].unique())
    selected_years = st.sidebar.multiselect(
        "Analysis Years", all_years,
        default=all_years if len(all_years) <= 3 else all_years[-3:], # 기본값 전체 혹은 최근 3개
        key="year_multi"
    )

    st.sidebar.markdown("---")
    st.sidebar.header("📂 Category Filter")

    cat_options = ["All Categories"] + sorted(df['Category'].unique().tolist())
    sel_cat = st.sidebar.selectbox("Major Category", cat_options, key="cat_select")

    if sel_cat == "All Categories":
        temp_df = df
    else:
        temp_df = df[df['Category'] == sel_cat]

    prod_options = ["All Products"] + sorted(temp_df['Product_Group'].unique().tolist())
    sel_prod = st.sidebar.selectbox("Product Group", prod_options, key="prod_select")

    # [참조표]
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ℹ️ Product Group Reference")
    ref_data = {
        'Product': ['GH', 'JB', 'CV', 'IJ'],
        'Car Model': ['HB20 (소형)', '크레타', '크레타(소형)', 'HB12']
    }
    st.sidebar.dataframe(pd.DataFrame(ref_data), hide_index=True, use_container_width=True)
    # 필터링 데이터
    filtered_df = df[df['Year'].isin(selected_years)]
    if sel_cat != "All Categories":
        filtered_df = filtered_df[filtered_df['Category'] == sel_cat]
    if sel_prod != "All Products":
        filtered_df = filtered_df[filtered_df['Product_Group'] == sel_prod]

    # ==================== 필터링 데이터 검증 ====================
    filter_issues = validate_filtered_data(filtered_df, df, selected_years, sel_cat, sel_prod)
    if filter_issues:
        st.warning(f"⚠️ 필터링 데이터 검증: {len(filter_issues)}개 이슈 발견")
        for issue in filter_issues:
            st.caption(f"  • {issue}")
    # ==================== 필터링 데이터 검증 끝 ====================

    # KPI Summary
    st.markdown("### 📊 Executive Summary (Filtered Scope)")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    f_rev = filtered_df['Revenue'].sum()
    f_qty = filtered_df['Quantity'].sum()
    f_asp = f_rev / f_qty if f_qty > 0 else 0

    if len(selected_years) > 0:
        max_y = max(selected_years)
        curr = filtered_df[filtered_df['Year'] == max_y]['Revenue'].sum()
        prev = filtered_df[filtered_df['Year'] == (max_y - 1)]['Revenue'].sum()
        yoy = ((curr - prev) / prev * 100) if prev > 0 else 0
    else:
        max_y, yoy = "-", 0
    kpi1.metric("Revenue", f"R$ {f_rev:,.0f}")
    kpi2.metric("Quantity", f"{f_qty:,.0f}")
    kpi3.metric("Avg ASP", f"R$ {f_asp:,.0f}")
    kpi4.metric(f"YoY Growth ('{max_y})", f"{yoy:,.1f}%", delta_color="normal")

    # 탭 구성
    tab0, tab1, tab2, tab3, tab4 = st.tabs(["📑 전사 요약", "📈 Growth & Share", "📅 Time Series", "💰 P/Q/Mix Bridge", "🚨 Outlier"])
    # === Tab 0: 전사 요약 ===
    with tab0:
        st.subheader("🏢 Major Category Overview")
        df_overview = df[df['Year'].isin(selected_years)]
        ov_grp = df_overview.groupby('Category').agg({'Revenue': 'sum', 'Quantity': 'sum'}).reset_index()
        ov_grp['ASP'] = ov_grp['Revenue'] / ov_grp['Quantity']

        total_rev_ov = ov_grp['Revenue'].sum()
        total_qty_ov = ov_grp['Quantity'].sum()
        ov_grp['Rev_Label'] = ov_grp.apply(lambda x: f"R$ {x['Revenue']:,.0f}<br>({x['Revenue']/total_rev_ov:.1%})", axis=1)
        ov_grp['Qty_Label'] = ov_grp.apply(lambda x: f"{x['Quantity']:,.0f}<br>({x['Quantity']/total_qty_ov:.1%})", axis=1)
        c1, c2, c3 = st.columns(3)
        with c1:
            fig_rev = px.bar(ov_grp, x='Category', y='Revenue', text='Rev_Label', color='Category',
                             title="Revenue & Share", color_discrete_sequence=COLORS['sequence'])
            fig_rev.update_traces(textposition='outside')
            st.plotly_chart(fig_rev, use_container_width=True)
        with c2:
            fig_qty = px.bar(ov_grp, x='Category', y='Quantity', text='Qty_Label', color='Category',
                             title="Sales Volume & Share", color_discrete_sequence=COLORS['sequence'])
            fig_qty.update_traces(textposition='outside')
            st.plotly_chart(fig_qty, use_container_width=True)
        with c3:
            st.plotly_chart(px.bar(ov_grp, x='Category', y='ASP', color='Category', title="Avg ASP (R$)", text_auto='.0f', color_discrete_sequence=COLORS['sequence']), use_container_width=True)
        st.markdown("---")
        st.subheader("📈 Yearly Trend by Category")
        trend_ov = df_overview.groupby(['Year', 'Category']).agg({'Revenue':'sum', 'Quantity':'sum'}).reset_index()
        trend_ov['ASP'] = trend_ov['Revenue'] / trend_ov['Quantity']

        c_tr1, c_tr2, c_tr3 = st.columns(3)
        with c_tr1:
            fig_l1 = px.line(trend_ov, x='Year', y='Revenue', color='Category', markers=True, title="Revenue Trend", color_discrete_sequence=COLORS['sequence'])
            fig_l1.update_xaxes(dtick=1)
            st.plotly_chart(fig_l1, use_container_width=True)
        with c_tr2:
            fig_l2 = px.line(trend_ov, x='Year', y='Quantity', color='Category', markers=True, title="Quantity Trend", color_discrete_sequence=COLORS['sequence'])
            fig_l2.update_xaxes(dtick=1)
            st.plotly_chart(fig_l2, use_container_width=True)
        with c_tr3:
            fig_l3 = px.line(trend_ov, x='Year', y='ASP', color='Category', markers=True, title="ASP Trend (R$)", color_discrete_sequence=COLORS['sequence'])
            fig_l3.update_xaxes(dtick=1)
            st.plotly_chart(fig_l3, use_container_width=True)
    # === Tab 1: Growth & Share ===
    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader(f"[{sel_cat}] Yearly Revenue")
            y_df = filtered_df.groupby('Year')['Revenue'].sum().reset_index()
            y_df['YoY'] = y_df['Revenue'].pct_change() * 100

            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(x=y_df['Year'], y=y_df['Revenue'], name="Revenue", marker_color=COLORS['primary'],
                                 text=y_df['Revenue'].apply(lambda x: f"R$ {x:,.0f}"), textposition='auto'), secondary_y=False)
            fig.add_trace(go.Scatter(x=y_df['Year'], y=y_df['YoY'], name="YoY %", mode='lines+markers+text',
                                     line=dict(color=COLORS['secondary'], width=3), text=y_df['YoY'].apply(lambda x: f"{x:.1f}%" if pd.notnull(x) else ""), textposition="top center"), secondary_y=True)
            fig.update_xaxes(dtick=1)
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            st.subheader(f"[{sel_cat}] Revenue Share")
            comp_df = filtered_df.groupby(['Year', 'Product_Group'])['Revenue'].sum().reset_index()
            comp_df['Total'] = comp_df.groupby('Year')['Revenue'].transform('sum')
            comp_df['Share'] = (comp_df['Revenue'] / comp_df['Total'] * 100)
            comp_df['Label'] = comp_df.apply(lambda x: f"R$ {x['Revenue']:,.0f}<br>({x['Share']:.1f}%)", axis=1)
            fig = px.bar(comp_df, x='Year', y='Revenue', color='Product_Group',
                         text='Label', color_discrete_sequence=COLORS['sequence'])
            fig.update_traces(textposition='inside')
            fig.update_xaxes(dtick=1)
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        st.subheader(f"[{sel_cat}] Volume & Price Trend Analysis")
        group_col = 'Product_Group' if sel_prod == "All Products" else 'Sub_Product'
        trend_detailed = filtered_df.groupby(['Year', group_col]).agg({'Quantity': 'sum','Revenue': 'sum'}).reset_index()
        trend_detailed['ASP'] = trend_detailed['Revenue'] / trend_detailed['Quantity']
        trend_detailed['Total_Qty_Year'] = trend_detailed.groupby('Year')['Quantity'].transform('sum')
        trend_detailed['Qty_Label'] = trend_detailed.apply(lambda x: f"{x['Quantity']:,.0f}<br>({x['Quantity']/x['Total_Qty_Year']:.1%})", axis=1)

        c3, c4 = st.columns(2)
        with c3:
            fig_qty = px.bar(trend_detailed, x='Year', y='Quantity', color=group_col, text='Qty_Label',
                             title="Quantity by Product Group", color_discrete_sequence=COLORS['sequence'])
            fig_qty.update_traces(textposition='inside')
            fig_qty.update_xaxes(dtick=1)
            st.plotly_chart(fig_qty, use_container_width=True)
        with c4:
            fig_asp = px.line(trend_detailed, x='Year', y='ASP', color=group_col,
                              title="ASP by Product Group", markers=True, color_discrete_sequence=COLORS['sequence'])
            fig_asp.update_yaxes(tickprefix="R$ ")
            fig_asp.update_xaxes(dtick=1)
            st.plotly_chart(fig_asp, use_container_width=True)
    # === Tab 2: 시계열 ===
    with tab2:
        mix_col = 'Category' if sel_cat == "All Categories" else 'Product_Group'
        st.subheader(f"Revenue Mix by {mix_col}")

        mode = st.radio("Time Unit", ["Monthly", "Quarterly"], horizontal=True, key="ts_mode")
        t_col = 'Month_Dt' if mode == "Monthly" else 'Quarter_Str'

        t_df = filtered_df.groupby([t_col, mix_col])['Revenue'].sum().reset_index()
        if mode == "Quarterly": t_df = t_df.sort_values(t_col)
        t_df['Total'] = t_df.groupby(t_col)['Revenue'].transform('sum')
        t_df['Share'] = t_df['Revenue'] / t_df['Total']
        c1, c2 = st.columns([2, 1])
        with c1:
            fig = px.area(t_df, x=t_col, y='Revenue', color=mix_col, title="Revenue Trend", color_discrete_sequence=COLORS['sequence'])
            if mode == "Monthly": fig.update_xaxes(tickformat="%Y-%m")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.bar(t_df, x=t_col, y='Share', color=mix_col, title="Revenue Mix", color_discrete_sequence=COLORS['sequence'])
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        st.subheader(f"Quantity vs ASP Trend ({mode})")
        if mode == "Monthly":
            agg_df = filtered_df.groupby('Month_Dt').agg({'Revenue':'sum', 'Quantity':'sum'}).reset_index()
            x_ax = 'Month_Dt'
        else:
            agg_df = filtered_df.groupby('Quarter_Str').agg({'Revenue':'sum', 'Quantity':'sum'}).reset_index().sort_values('Quarter_Str')
            x_ax = 'Quarter_Str'
        agg_df['ASP'] = np.where(agg_df['Quantity']>0, agg_df['Revenue']/agg_df['Quantity'], 0)
        fig_dual = make_subplots(specs=[[{"secondary_y": True}]])
        fig_dual.add_trace(go.Bar(x=agg_df[x_ax], y=agg_df['Quantity'], name="Quantity", marker_color='#B0BEC5', opacity=0.6), secondary_y=False)
        fig_dual.add_trace(go.Scatter(x=agg_df[x_ax], y=agg_df['ASP'], name="ASP", mode='lines+markers', line=dict(color=COLORS['primary'], width=3)), secondary_y=True)
        fig_dual.update_layout(hovermode="x unified")
        fig_dual.update_yaxes(title_text="Quantity", showgrid=False, secondary_y=False)
        fig_dual.update_yaxes(title_text="ASP (R$)", showgrid=False, secondary_y=True)
        if mode == "Monthly": fig_dual.update_xaxes(tickformat="%Y-%m")
        st.plotly_chart(fig_dual, use_container_width=True)
    # === Tab 3: P/Q/Mix Bridge (Multi-Step Fixed) ===
    with tab3:
        st.subheader("💰 Price-Volume-Mix Bridge Analysis")
        st.info("매출 변동을 **물량(Volume)**, **가격(Price)**, **믹스(Mix)** 효과로 분해합니다.")

        target_years = sorted(list(set(selected_years)))

        if len(target_years) >= 2:
            # 1. Company-Wide (All Data, Group by Category)
            st.markdown("#### 1. Company-Wide Bridge (by Major Category)")
            st.caption("전사 기준: 대분류(Category) 믹스 효과 분석")
            df_company = df[df['Year'].isin(selected_years)]
            res1 = calculate_bridge_mix_steps(df_company, target_years, 'Category')
            if res1:
                st.plotly_chart(create_waterfall_fig_multi(res1), use_container_width=True)

            st.markdown("---")

            # 2. Major Category Level (Group by Product Group)
            if sel_cat == "All Categories":
                st.markdown("#### 2. Major Category Bridge (All -> Product Group)")
                st.caption("전체 카테고리 내 제품군(Product Group) 믹스 효과 분석")
                res2 = calculate_bridge_mix_steps(df_company, target_years, 'Product_Group')
            else:
                st.markdown(f"#### 2. Major Category Bridge: {sel_cat} (by Product Group)")
                st.caption(f"선택된 {sel_cat} 내 제품군 믹스 효과 분석")
                df_cat_bridge = df_company[df_company['Category'] == sel_cat]
                res2 = calculate_bridge_mix_steps(df_cat_bridge, target_years, 'Product_Group')

            if res2:
                st.plotly_chart(create_waterfall_fig_multi(res2), use_container_width=True)

            st.markdown("---")

            # 3. Detailed Level (Major & Product Selected -> Customer)
            st.markdown("#### 3. Detailed Bridge (Specific Product Group)")
            if sel_cat != "All Categories" and sel_prod != "All Products":
                st.success(f"Selected: {sel_cat} > {sel_prod}")
                st.caption(f"선택된 제품군 내 거래처(Customer) 믹스 효과 상세 분석")

                # Bumper는 Sub_Product, 나머지는 Customer로 Breakdown
                is_bumper = (sel_cat == 'Bumper')
                detail_col = 'Sub_Product' if is_bumper else 'Customer'

                res3 = calculate_bridge_mix_steps(filtered_df, target_years, detail_col)
                if res3:
                    st.plotly_chart(create_waterfall_fig_multi(res3), use_container_width=True)
                else:
                    st.warning("분석할 데이터가 부족합니다.")
            else:
                st.info("💡 **Major Category**와 **Product Group**을 모두 선택해야 최하단 상세 브릿지가 표시됩니다.")

        else:
            st.info("비교할 연도를 2개 이상 선택해주세요.")
    # === Tab 4: Outlier ===
    with tab4:
        st.subheader("Statistical Outlier Analysis")
        c1, c2 = st.columns([2, 1])
        with c1:
            fig_box = px.box(filtered_df, x='Product_Group', y='ASP', color='Product_Group',
                             points="outliers", notched=True, color_discrete_sequence=COLORS['sequence'])
            st.plotly_chart(fig_box, use_container_width=True)
        with c2:
            stats = filtered_df.groupby('Product_Group')['ASP'].agg(['count', 'mean', 'min', 'max']).reset_index()
            st.dataframe(stats.style.format({'mean': 'R$ {:,.0f}', 'min': 'R$ {:,.0f}', 'max': 'R$ {:,.0f}'}), use_container_width=True)
        st.markdown("---")
        st.markdown("#### 🚨 Top 50 Outliers")
        outliers = filtered_df.groupby('Product_Group', group_keys=False).apply(lambda x: detect_outliers_iqr(x))
        if not outliers.empty:
            outlier_view = outliers[['Date', 'Category', 'Product_Group', 'Sub_Product', 'Quantity', 'Revenue', 'ASP']].sort_values('ASP', ascending=False).head(50)
            st.dataframe(outlier_view.style.format({'Revenue': 'R$ {:,.0f}', 'ASP': 'R$ {:,.0f}', 'Quantity': '{:,.0f}'}).bar(subset=['ASP'], color='#ffcccc'), use_container_width=True)
        else:
            st.success("이상치 없음")
if __name__ == "__main__":
    main()
