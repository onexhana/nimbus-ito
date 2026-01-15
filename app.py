import streamlit as st
import pandas as pd
import plotly.express as px

# 페이지 설정
st.set_page_config(page_title="Deal-ito 통합 실적 대시보드", layout="wide")

st.title("📊 Deal-ito 통합 실적/이익 대시보드")
st.markdown("""
엑셀 파일을 업로드하면 **매출**과 **이익**에 대한 실적 반영율을 자동 계산합니다.
- **매출 비율**: 고객(40%), 관리(30%), 소싱(30%)
- **이익 비율**: 고객(40%), 관리(30%), 소싱(30%)
""")

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"])

if uploaded_file:
    # 데이터 로드
    df = pd.read_excel(uploaded_file)
    
    # 2. 사이드바 필터 (조회 기간 & 담당자 선택)
    st.sidebar.header("🔍 조회 조건 설정")
    
    # 기간 선택 옵션
    period_map = {
        "전체 (1-12월)": list(range(1, 13)),
        "1분기 (1-3월)": list(range(1, 4)),
        "2분기 (4-6월)": list(range(4, 7)),
        "3분기 (7-9월)": list(range(7, 10)),
        "4분기 (10-12월)": list(range(10, 13)),
        "상반기 (1-6월)": list(range(1, 7)),
        "하반기 (7-12월)": list(range(7, 13))
    }
    selected_period_label = st.sidebar.selectbox("조회 기간을 선택하세요", list(period_map.keys()))
    selected_months_ints = period_map[selected_period_label]
    selected_months = [f"{m:02d}" for m in selected_months_ints]

    # 모든 담당자 추출 (필터용)
    all_managers = sorted(list(set(
        df['Deal - 담당자_고객'].dropna().unique().tolist() + 
        df['Deal - 담당자_관리'].dropna().unique().tolist() + 
        df['Deal - 담당자_소싱'].dropna().unique().tolist()
    )))
    
    selected_managers = st.sidebar.multiselect("조회할 담당자를 선택하세요", all_managers, default=all_managers)
    
    # 3. 데이터 전처리
    sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
    profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
    
    def clean_currency(column):
        if column in df.columns:
            # 문자열 변환 후 숫자 이외의 문자 제거
            s = df[column].astype(str).str.replace(r'[^0-9.-]', '', regex=True)
            return pd.to_numeric(s, errors='coerce').fillna(0)
        return 0

    for col in sales_cols + profit_cols:
        df[col] = clean_currency(col)
    
    df['선택기간_총매출'] = df[sales_cols].sum(axis=1)
    df['선택기간_총이익'] = df[profit_cols].sum(axis=1)

    # 4. 데이터 집계 함수 (역할별 40/30/30 합산)
    def calculate_consolidated_results(target_col):
        # 각 역할별 기여도 계산
        role_configs = [
            ('고객(40%)', 0.4, 'Deal - 담당자_고객'),
            ('관리(30%)', 0.3, 'Deal - 담당자_관리'),
            ('소싱(30%)', 0.3, 'Deal - 담당자_소싱')
        ]
        
        individual_results = []
        for role_label, ratio, manager_col in role_configs:
            temp = df[['Deal - 이름', manager_col, target_col]].copy()
            temp['반영실적'] = temp[target_col] * ratio
            temp['역할'] = role_label
            temp.columns = ['Deal명', '담당자', '원금액', '반영실적', '역할']
            individual_results.append(temp)
            
        combined = pd.concat(individual_results)
        
        # 선택된 담당자만 필터링
        if selected_managers:
            combined = combined[combined['담당자'].isin(selected_managers)]
            
        return combined

    # 5. 탭 구성
    tab1, tab2 = st.tabs(["💰 매출 분석", "📉 이익 분석"])

    with tab1:
        st.subheader(f"📅 매출 조회 기간: {selected_period_label}")
        m_df = calculate_consolidated_results('선택기간_총매출')
        
        # 담당자별 합산
        summary_m = m_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
        
        st.write("#### 👤 담당자별 합산 실적 (40%+30%+30%)")
        st.dataframe(
            summary_m.style.format({'반영실적': '{:,.0f}원'}), 
            use_container_width=True,
            hide_index=True
        )

        # 상세 내역 (탭1 내부)
        st.write("---")
        st.write("#### 📋 매출 상세 기여 내역")
        st.dataframe(
            m_df[m_df['반영실적'] > 0].style.format({'원금액': '{:,.0f}원', '반영실적': '{:,.0f}원'}), 
            use_container_width=True,
            hide_index=True
        )

    with tab2:
        st.subheader(f"📅 이익 조회 기간: {selected_period_label}")
        p_df = calculate_consolidated_results('선택기간_총이익')
        
        # 담당자별 합산
        summary_p = p_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
        
        st.write("#### 👤 담당자별 합산 실적 (40%+30%+30%)")
        st.dataframe(
            summary_p.style.format({'반영실적': '{:,.0f}원'}), 
            use_container_width=True,
            hide_index=True
        )

        # 상세 내역 (탭2 내부)
        st.write("---")
        st.write("#### 📋 이익 상세 기여 내역")
        st.dataframe(
            p_df[p_df['반영실적'] > 0].style.format({'원금액': '{:,.0f}원', '반영실적': '{:,.0f}원'}), 
            use_container_width=True,
            hide_index=True
        )

else:
    st.info("왼쪽 사이드바에서 엑셀 파일을 업로드해주세요.")
