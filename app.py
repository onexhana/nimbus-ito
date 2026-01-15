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
    
    # 2. 사이드바 필터 (월 범위 & 역할 선택)
    st.sidebar.header("🔍 조회 조건 설정")
    month_range = st.sidebar.text_input("조회할 월 범위를 입력하세요 (예: 1-3, 4-9)", value="1-12")
    
    # 역할 선택 필터 추가
    role_options = ["전체 합산", "고객(40%)", "관리(30%)", "소싱(30%)"]
    selected_role = st.sidebar.selectbox("보고 싶은 역할을 선택하세요", role_options)
    
    try:
        if '-' in month_range:
            start_m, end_m = map(int, month_range.split('-'))
        else:
            start_m = end_m = int(month_range)
        
        selected_months = [f"{m:02d}" for m in range(start_m, end_m + 1)]
    except:
        st.error("월 범위 형식이 올바르지 않습니다. '1-3' 형태로 입력해주세요.")
        st.stop()

    # 3. 데이터 전처리
    sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
    profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
    
    def clean_currency(column):
        if column in df.columns:
            return pd.to_numeric(df[column].astype(str).str.replace(',', '').str.replace('₩', '').str.replace(' ', ''), errors='coerce').fillna(0)
        return 0

    for col in sales_cols + profit_cols:
        df[col] = clean_currency(col)
    
    df['선택기간_총매출'] = df[sales_cols].sum(axis=1)
    df['선택기간_총이익'] = df[profit_cols].sum(axis=1)

    # 4. 탭 구성 (매출 분석 / 이익 분석)
    tab1, tab2 = st.tabs(["💰 매출 분석", "📉 이익 분석"])

    def process_data(target_col, ratios):
        results = []
        for role_label, ratio, manager_col in ratios:
            temp = df[[manager_col, target_col, 'Deal - 이름']].copy()
            temp['반영실적'] = temp[target_col] * ratio
            temp['구분'] = role_label
            temp.columns = ['담당자', '금액', 'Deal명', '반영실적', '구분']
            results.append(temp)
        
        combined = pd.concat(results)
        
        # 역할 필터 적용
        if selected_role != "전체 합산":
            combined = combined[combined['구분'] == selected_role]
            
        return combined

    # 매출/이익 공통 비율 (고객 40, 관리 30, 소싱 30)
    common_ratios = [
        ('고객(40%)', 0.4, 'Deal - 담당자_고객'),
        ('관리(30%)', 0.3, 'Deal - 담당자_관리'),
        ('소싱(30%)', 0.3, 'Deal - 담당자_소싱')
    ]

    with tab1:
        st.subheader(f"📅 매출 조회 기간: {month_range}월 ({selected_role})")
        m_df = process_data('선택기간_총매출', common_ratios)
        summary_m = m_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
        
        c1, c2 = st.columns([1, 1])
        with c1:
            st.write(f"#### 👤 담당자별 {selected_role} 실적")
            st.dataframe(summary_m.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True)
        with c2:
            fig_m = px.pie(summary_m, values='반영실적', names='담당자', hole=0.3, title=f"매출 {selected_role} 비중")
            st.plotly_chart(fig_m, use_container_width=True)

    with tab2:
        st.subheader(f"📅 이익 조회 기간: {month_range}월 ({selected_role})")
        p_df = process_data('선택기간_총이익', common_ratios)
        summary_p = p_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
        
        c3, c4 = st.columns([1, 1])
        with c3:
            st.write(f"#### 👤 담당자별 {selected_role} 실적")
            st.dataframe(summary_p.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True)
        with c4:
            fig_p = px.pie(summary_p, values='반영실적', names='담당자', hole=0.3, title=f"이익 {selected_role} 비중", color_discrete_sequence=px.colors.sequential.RdBu)
            st.plotly_chart(fig_p, use_container_width=True)

    # 상세 데이터 확인
    st.write("---")
    st.write(f"### 📋 상세 내역 - {selected_role} 기준")
    if selected_role == "전체 합산":
        st.dataframe(df[['Deal - 이름', 'Deal - 담당자_고객', 'Deal - 담당자_관리', 'Deal - 담당자_소싱', '선택기간_총매출', '선택기간_총이익']].style.format({'선택기간_총매출': '{:,.0f}원', '선택기간_총이익': '{:,.0f}원'}))
    else:
        # 특정 역할 선택 시 해당 역할의 상세 데이터만 표시
        filtered_detail = m_df[['Deal명', '담당자', '금액', '반영실적']].copy()
        filtered_detail.columns = ['프로젝트명', '담당자', '원금액', '내반영실적']
        st.dataframe(filtered_detail.style.format({'원금액': '{:,.0f}원', '내반영실적': '{:,.0f}원'}))

else:
    st.info("왼쪽 사이드바 또는 상단에서 엑셀 파일을 업로드해주세요.")
