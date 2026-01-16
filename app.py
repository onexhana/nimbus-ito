import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import os
import io
import time
import re

# 파일 경로 설정
TARGETS_FILE = "targets.json"
DASHBOARD_CACHE_FILE = "dashboard_cache.pkl"

def load_dashboard_data():
    if 'dashboard_df' in st.session_state:
        return st.session_state.dashboard_df
    if os.path.exists(DASHBOARD_CACHE_FILE):
        try:
            df = pd.read_pickle(DASHBOARD_CACHE_FILE)
            st.session_state.dashboard_df = df
            return df
        except:
            return None
    return None

def save_dashboard_data(df):
    st.session_state.dashboard_df = df
    df.to_pickle(DASHBOARD_CACHE_FILE)
    # 데이터 저장 시 담당자 명단 즉시 동기화
    targets_data = load_targets()
    excel_managers = sorted(list(set(
        df['Deal - 담당자_고객'].dropna().unique().tolist() + 
        df['Deal - 담당자_관리'].dropna().unique().tolist() + 
        df['Deal - 담당자_소싱'].dropna().unique().tolist()
    )))
    updated = False
    for manager in excel_managers:
        if manager not in targets_data:
            targets_data[manager] = {f"q{i}": {"mm": 0.0, "sales": 0.0, "profit": 0.0} for i in range(1, 5)}
            updated = True
    if updated:
        save_targets(targets_data)

def delete_dashboard_data():
    if 'dashboard_df' in st.session_state:
        del st.session_state.dashboard_df
    if os.path.exists(DASHBOARD_CACHE_FILE):
        os.remove(DASHBOARD_CACHE_FILE)

def load_targets():
    if os.path.exists(TARGETS_FILE):
        with open(TARGETS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_targets(targets):
    with open(TARGETS_FILE, "w", encoding="utf-8") as f:
        json.dump(targets, f, ensure_ascii=False, indent=4)

# 엑셀 템플릿 생성 함수
def create_excel_template(targets_data):
    output = io.BytesIO()
    rows = []
    managers = sorted(targets_data.keys()) if targets_data else ["고봉수", "김길래", "박승수", "손병희", "이민지"]
    for mgr in managers:
        m_data = targets_data.get(mgr, {f"q{i}": {"mm": 0, "sales": 0, "profit": 0} for i in range(1, 5)})
        for category, label in [("mm", "MM"), ("sales", "매출"), ("profit", "매출이익")]:
            row = {
                "성명": mgr if label == "MM" else "",
                "내용": label,
                "년 목표": sum(m_data[f"q{i}"][category] for i in range(1, 5)),
                "1/4분기 목표": m_data["q1"][category],
                "2/4분기 목표": m_data["q2"][category],
                "3/4분기 목표": m_data["q3"][category],
                "4/4분기 목표": m_data["q4"][category]
            }
            rows.append(row)
    df_template = pd.DataFrame(rows)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='목표설정')
    return output.getvalue()

# 페이지 설정
st.set_page_config(page_title="Deal-ito 통합 실적 대시보드", layout="wide")

# 세션 상태 초기화
if 'page' not in st.session_state:
    st.session_state.page = "dashboard"

# 사이드바 메뉴 구성
st.sidebar.title("📌 메뉴")
if st.sidebar.button("📊 실적 대시보드", use_container_width=True):
    st.session_state.page = "dashboard"
if st.sidebar.button("🎯 목표 설정하기", use_container_width=True):
    st.session_state.page = "targets"
if st.sidebar.button("📈 목표 달성률 확인하기", use_container_width=True):
    st.session_state.page = "achievement"

st.sidebar.write("---")

# 1. 실적 데이터 업로드 (대시보드 페이지용)
if st.session_state.page == "dashboard":
    with st.sidebar.expander("📁 실적 데이터 업로드", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"], key="dashboard_uploader")
        
        if uploaded_file is not None:
            # 파일이 새로 올라왔을 때만 처리하기 위해 체크
            if 'last_uploaded_file' not in st.session_state or st.session_state.last_uploaded_file != uploaded_file.name:
                with st.status("🚀 실적 데이터 로드 중...", expanded=True) as status:
                    try:
                        st.write("📂 엑셀 파일 읽는 중...")
                        df_loaded = pd.read_excel(uploaded_file)
                        st.write("💾 데이터 저장 및 명단 동기화 중...")
                        save_dashboard_data(df_loaded)
                        st.session_state.last_uploaded_file = uploaded_file.name
                        status.update(label="✅ 로드 완료!", state="complete", expanded=False)
                        st.success("데이터가 로드되었습니다! 페이지를 업데이트합니다.")
                        st.rerun()
                    except Exception as e:
                        status.update(label="❌ 오류 발생", state="error")
                        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        
        if load_dashboard_data() is not None:
            if st.button("🗑️ 업로드된 데이터 삭제", use_container_width=True):
                delete_dashboard_data()
                st.info("데이터가 삭제되었습니다. 새로운 파일을 업로드해 주세요.")
                st.rerun()
    st.sidebar.write("---")

# 2. 목표 관리 메뉴 (목표 설정 페이지 전용)
if st.session_state.page == "targets":
    targets_data = load_targets()
    with st.sidebar.expander("📂 엑셀 일괄 관리", expanded=False):
        template_excel = create_excel_template(targets_data)
        st.download_button("📥 양식 다운로드", data=template_excel, file_name="target_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        uploaded_target_file = st.file_uploader("엑셀 업로드", type=["xlsx"], key="sidebar_target_uploader")
        if uploaded_target_file:
            # 파일이 새로 올라왔을 때만 처리하기 위해 체크
            if 'last_target_file' not in st.session_state or st.session_state.last_target_file != uploaded_target_file.name:
                with st.status("🚀 데이터 반영 중...", expanded=True) as status:
                    try:
                        st.write("📂 파일 읽는 중...")
                        up_df = pd.read_excel(uploaded_target_file)
                        up_df['성명'] = up_df['성명'].ffill()
                        new_targets = targets_data.copy()
                        for mgr in up_df['성명'].unique():
                            if pd.isna(mgr): continue
                            if mgr not in new_targets:
                                new_targets[mgr] = {f"q{i}": {"mm": 0, "sales": 0, "profit": 0} for i in range(1, 5)}
                            mgr_rows = up_df[up_df['성명'] == mgr]
                            for _, row in mgr_rows.iterrows():
                                cat_label = row['내용']
                                category = "mm" if cat_label == "MM" else "sales" if "매출" == cat_label else "profit" if "매출이익" == cat_label else None
                                if category:
                                    for i in range(1, 5):
                                        val = row[f'{i}/4분기 목표'] if f'{i}/4분기 목표' in row else 0
                                        new_targets[mgr][f"q{i}"][category] = float(val) if not pd.isna(val) else 0.0
                        save_targets(new_targets)
                        st.session_state.last_target_file = uploaded_target_file.name
                        status.update(label="✅ 반영 완료!", state="complete", expanded=False)
                        st.success("데이터가 반영되었습니다!")
                        st.rerun()
                    except Exception as e:
                        status.update(label="❌ 오류 발생", state="error")
                        st.error(f"오류: {e}")
        st.write("---")
        if st.button("🚨 모든 데이터 초기화", use_container_width=True):
            st.session_state.show_reset_confirm = True
        if st.session_state.get('show_reset_confirm', False):
            st.warning("⚠️ 모든 데이터를 삭제할까요?")
            c1, c2 = st.columns(2)
            if c1.button("✅ 예", use_container_width=True):
                save_targets({})
                st.session_state.show_reset_confirm = False
                st.rerun()
            if c2.button("❌ 아니오", use_container_width=True):
                st.session_state.show_reset_confirm = False
                st.rerun()

    with st.sidebar.expander("👤 개별 담당자 추가 및 삭제", expanded=False):
        st.write("**담당자 추가**")
        new_mgr_name = st.text_input("추가할 이름 입력", key="sidebar_mgr_input")
        if st.button("➕ 추가하기", use_container_width=True):
            if new_mgr_name:
                if new_mgr_name not in targets_data:
                    targets_data[new_mgr_name] = {f"q{i}": {"mm": 0.0, "sales": 0.0, "profit": 0.0} for i in range(1, 5)}
                    save_targets(targets_data)
                    st.success(f"'{new_mgr_name}' 추가 완료!")
                    st.rerun()
                else:
                    st.warning("이미 있는 이름입니다.")
            else:
                st.error("이름을 입력하세요.")
        
        st.write("---")
        st.write("**담당자 삭제**")
        if targets_data:
            mgr_to_delete = st.selectbox("삭제할 담당자 선택", options=["선택하세요"] + sorted(targets_data.keys()), key="mgr_del_select")
            if mgr_to_delete != "선택하세요":
                if st.button(f"🗑️ '{mgr_to_delete}' 삭제", use_container_width=True, type="secondary"):
                    del targets_data[mgr_to_delete]
                    save_targets(targets_data)
                    st.success(f"'{mgr_to_delete}' 삭제 완료!")
                    st.rerun()
        else:
            st.info("삭제할 담당자가 없습니다.")

def clean_currency_val(val):
    """금액 데이터에서 숫자만 추출하는 강력한 함수"""
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    # 숫자, 마침표(.), 마이너스(-)만 남기고 모든 문자 제거
    cleaned = re.sub(r'[^0-9.-]', '', str(val))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

# 게이지 차트 생성 함수
def draw_gauge(current_val, target_val, title):
    percentage = (current_val / target_val * 100) if target_val > 0 else 0
    
    # 사용자의 요청대로 목표 금액을 100% 지점(Max)으로 설정
    max_range = target_val if target_val > 0 else max(current_val, 100)
    
    fig = go.Figure(go.Indicator(
        mode = "gauge", # redundant number 제거
        value = min(current_val, max_range), # 바는 일단 max_range까지만
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': f"<b>{title}</b>", 'font': {'size': 20}},
        gauge = {
            'axis': {
                'range': [0, max_range], 
                'tickmode': 'array',
                'tickvals': [0, max_range * 0.25, max_range * 0.5, max_range * 0.75, max_range],
                'ticktext': ['0%', '25%', '50%', '75%', '100%'],
                'tickfont': {'size': 12}
            },
            'bar': {'color': "#636EFA"},
            'bgcolor': "white",
            'borderwidth': 1,
            'bordercolor': "gray",
        }
    ))
    
    # 중앙에 핵심 정보 배치 (달성률, 목표, 실적)
    fig.add_annotation(
        x=0.5, y=0.35,
        text=f"<span style='font-size:26px; font-weight:bold; color:#636EFA;'>{percentage:.1f}% 달성</span><br><br>" +
             f"<span style='font-size:15px; color:gray;'>목표: {target_val:,.0f}원</span><br>" +
             f"<span style='font-size:18px; font-weight:bold;'>실적: {current_val:,.0f}원</span>",
        showarrow=False,
        align="center"
    )
    
    fig.update_layout(height=400, margin=dict(l=50, r=50, t=80, b=20))
    return fig

# 메인 화면 - 목표 설정
if st.session_state.page == "targets":
    st.title("🎯 담당자별 목표 설정")
    st.markdown("분기별 목표를 입력하세요. (단위: 천원)")
    targets_data = load_targets()
    all_mgrs = sorted(targets_data.keys())
    st.write("### 👥 담당자 선택")
    def on_select_all_change():
        for mgr in all_mgrs:
            st.session_state[f"sel_{mgr}"] = st.session_state.select_all_key
    col_sel1, col_sel2 = st.columns([1, 5])
    select_all = col_sel1.checkbox("전체 선택", key="select_all_key", on_change=on_select_all_change)
    selected_m_list = []
    if all_mgrs:
        mgr_cols = st.columns(5)
        for i, mgr in enumerate(all_mgrs):
            with mgr_cols[i % 5]:
                is_selected = st.checkbox(mgr, key=f"sel_{mgr}")
                if is_selected:
                    selected_m_list.append(mgr)
    if not targets_data and not all_mgrs:
        st.warning("등록된 담당자가 없습니다. 사이드바에서 개별 추가하거나 엑셀을 업로드하세요.")
    if selected_m_list:
        st.write("---")
        st.write(f"### 📋 선택된 담당자 목표 현황 요약")
        summary_rows = []
        for mgr in selected_m_list:
            m_data = targets_data.get(mgr, {f"q{i}": {"mm": 0, "sales": 0, "profit": 0} for i in range(1, 5)})
            for category, label in [("mm", "MM"), ("sales", "매출"), ("profit", "매출이익")]:
                row = {"성명": mgr if label == "MM" else "", "내용": label, "1/4분기": float(m_data["q1"][category]), "2/4분기": float(m_data["q2"][category]), "3/4분기": float(m_data["q3"][category]), "4/4분기": float(m_data["q4"][category])}
                row["년 합계"] = row["1/4분기"] + row["2/4분기"] + row["3/4분기"] + row["4/4분기"]
                summary_rows.append(row)
        if len(selected_m_list) > 1:
            summary_rows.append({"성명": "---", "내용": "---", "1/4분기": 0, "2/4분기": 0, "3/4분기": 0, "4/4분기": 0, "년 합계": 0})
            for category, label in [("mm", "MM"), ("sales", "매출"), ("profit", "매출이익")]:
                q1_sum = sum(float(targets_data[m]["q1"][category]) for m in selected_m_list); q2_sum = sum(float(targets_data[m]["q2"][category]) for m in selected_m_list); q3_sum = sum(float(targets_data[m]["q3"][category]) for m in selected_m_list); q4_sum = sum(float(targets_data[m]["q4"][category]) for m in selected_m_list)
                total_row = {"성명": "★ 전체 합계" if label == "MM" else "", "내용": label, "1/4분기": q1_sum, "2/4분기": q2_sum, "3/4분기": q3_sum, "4/4분기": q4_sum, "년 합계": q1_sum + q2_sum + q3_sum + q4_sum}
                summary_rows.append(total_row)
        if summary_rows:
            df_summary = pd.DataFrame(summary_rows)
            # 모든 수치 컬럼을 반올림하여 소수점 없이 표시
            format_mapping = {
                col: lambda x: f"{round(x):,.0f}" if isinstance(x, (int, float)) else x 
                for col in ["1/4분기", "2/4분기", "3/4분기", "4/4분기", "년 합계"]
            }
            st.dataframe(
                df_summary.style.format(format_mapping), 
                use_container_width=True, 
                hide_index=True
            )
        st.write("---")
        if len(selected_m_list) == 1:
            st.write(f"### 📝 담당자별 목표 수정")
            selected_m = selected_m_list[0]; m_data = targets_data[selected_m]
            with st.container(border=True):
                st.subheader(f"👤 {selected_m}님의 목표 설정")
                cols = st.columns(4); updated_m_data = {}
                for i, q in enumerate(["q1", "q2", "q3", "q4"]):
                    with cols[i]:
                        st.markdown(f"**{i+1}/4분기**")
                        mm = st.number_input(f"MM", value=float(m_data[q]["mm"]), key=f"{selected_m}_{q}_mm", step=0.1)
                        sales = st.number_input(f"매출(천원)", value=float(m_data[q]["sales"]), key=f"{selected_m}_{q}_sales", step=1000.0)
                        profit = st.number_input(f"이익(천원)", value=float(m_data[q]["profit"]), key=f"{selected_m}_{q}_profit", step=1000.0)
                        updated_m_data[q] = {"mm": mm, "sales": sales, "profit": profit}
                total_mm = sum(q_val["mm"] for q_val in updated_m_data.values()); total_sales = sum(q_val["sales"] for q_val in updated_m_data.values()); total_profit = sum(q_val["profit"] for q_val in updated_m_data.values())
                c1, c2, c3 = st.columns(3); c1.metric("연간 총 MM", f"{total_mm:.2f}"); c2.metric("연간 총 매출", f"{total_sales:,.0f}천원"); c3.metric("연간 총 이익", f"{total_profit:,.0f}천원")
                if st.button(f"💾 {selected_m}님 목표 저장", use_container_width=True, type="primary"):
                    targets_data[selected_m] = updated_m_data; save_targets(targets_data); st.success(f"{selected_m}님의 목표가 저장되었습니다!"); st.rerun()
                if st.button(f"🗑️ {selected_m}님 삭제", key=f"del_btn_{selected_m}"):
                    del targets_data[selected_m]; save_targets(targets_data); st.rerun()
    if st.button("📊 대시보드로 돌아가기"):
        st.session_state.page = "dashboard"; st.rerun()

# 메인 화면 - 목표 달성률 확인
elif st.session_state.page == "achievement":
    st.title("📈 목표 달성률 확인하기")
    
    df = load_dashboard_data()
    targets_data = load_targets()
    
    if df is None:
        st.warning("📊 실적 대시보드에서 먼저 엑셀 파일을 업로드해 주세요.")
        if st.button("📊 실적 대시보드로 이동"):
            st.session_state.page = "dashboard"
            st.rerun()
        st.stop()
        
    if not targets_data:
        st.warning("🎯 목표 설정하기에서 먼저 목표를 설정해 주세요.")
        if st.button("🎯 목표 설정으로 이동"):
            st.session_state.page = "targets"
            st.rerun()
        st.stop()

    # 조회 조건 설정
    st.sidebar.header("🔍 달성률 조회 조건")
    period_map = {
        "1분기 (1-3월)": (list(range(1, 4)), ["q1"]),
        "2분기 (4-6월)": (list(range(4, 7)), ["q2"]),
        "3분기 (7-9월)": (list(range(7, 10)), ["q3"]),
        "4분기 (10-12월)": (list(range(10, 13)), ["q4"]),
        "상반기 (1-6월)": (list(range(1, 7)), ["q1", "q2"]),
        "하반기 (7-12월)": (list(range(7, 13)), ["q3", "q4"]),
        "전체 (1-12월)": (list(range(1, 13)), ["q1", "q2", "q3", "q4"])
    }
    selected_period_label = st.sidebar.selectbox("조회 기간 선택", list(period_map.keys()))
    months, quarters = period_map[selected_period_label]
    selected_months = [f"{m:02d}" for m in months]
    
    all_managers = sorted(targets_data.keys())
    selected_manager = st.sidebar.selectbox("조회할 담당자 선택", all_managers)
    
    st.write(f"### 👤 {selected_manager}님의 {selected_period_label} 달성 현황")
    
    # 1. 목표 데이터 계산 (천원 단위 -> 원 단위로 변환)
    m_target_data = targets_data[selected_manager]
    target_sales = sum(float(m_target_data[q]["sales"]) for q in quarters) * 1000
    target_profit = sum(float(m_target_data[q]["profit"]) for q in quarters) * 1000
    
    # 2. 실제 실적 데이터 계산 (40/30/30 로직 적용)
    # 중복 제거 강화: 'Deal - 이름'이 비어있거나 합계인 행 제외
    df = df[df['Deal - 이름'].notna() & (df['Deal - 이름'].astype(str).str.strip() != "")]
    exclude_keywords = ['합계', '소계', 'total', 'sum']
    df = df[~df['Deal - 이름'].astype(str).str.lower().str.contains('|'.join(exclude_keywords), na=False)]

    for col in ['Deal - 담당자_고객', 'Deal - 담당자_관리', 'Deal - 담당자_소싱']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
    profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
    
    role_configs = [
        (0.4, 'Deal - 담당자_고객', '고객'),
        (0.3, 'Deal - 담당자_관리', '관리'),
        (0.3, 'Deal - 담당자_소싱', '소싱')
    ]
    
    actual_sales = 0.0
    actual_profit = 0.0
    detail_records = []
    
    for ratio, mgr_col, role_name in role_configs:
        if mgr_col in df.columns:
            mgr_mask = df[mgr_col] == selected_manager
            matched_df = df[mgr_mask].copy()
            if not matched_df.empty:
                for idx, row in matched_df.iterrows():
                    deal_name = row['Deal - 이름']
                    # 각 월별 실적을 개별적으로 수집
                    monthly_values = {m: clean_currency_val(row[f"Deal - @월별매출 ({m})"]) for m in selected_months if f"Deal - @월별매출 ({m})" in row}
                    monthly_profit_values = {m: clean_currency_val(row[f"Deal - @월별이익 ({m})"]) for m in selected_months if f"Deal - @월별이익 ({m})" in row}
                    
                    row_sales = sum(monthly_values.values())
                    row_profit = sum(monthly_profit_values.values())
                    
                    if row_sales > 0 or row_profit > 0:
                        actual_sales += row_sales * ratio
                        actual_profit += row_profit * ratio
                        
                        # 상세 내역 기록 생성
                        record = {
                            "Deal명": deal_name,
                            "역할": role_name,
                            "비중": f"{int(ratio*100)}%",
                            "원매출(합계)": row_sales,
                            "반영매출": row_sales * ratio
                        }
                        # 각 월별 실적 컬럼 추가
                        for m, val in monthly_values.items():
                            record[f"{int(m)}월 매출"] = val
                        
                        detail_records.append(record)

    # 화면 표시
    col1, col2 = st.columns(2)
    
    with col1:
        if target_sales > 0:
            st.plotly_chart(draw_gauge(actual_sales, target_sales, "💰 매출 달성 현황"), use_container_width=True)
        else:
            st.info("설정된 매출 목표가 없습니다.")
            st.metric("실제 매출 실적", f"{actual_sales:,.0f}원")
            
    with col2:
        if target_profit > 0:
            st.plotly_chart(draw_gauge(actual_profit, target_profit, "📉 이익 달성 현황"), use_container_width=True)
        else:
            st.info("설정된 이익 목표가 없습니다.")
            st.metric("실제 이익 실적", f"{actual_profit:,.0f}원")

    # 상세 표 추가
    st.write("---")
    st.write("#### 📊 요약 데이터")
    summary_data = {
        "구분": ["매출", "이익"],
        "목표 금액": [f"{target_sales:,.0f}원", f"{target_profit:,.0f}원"],
        "실제 실적": [f"{actual_sales:,.0f}원", f"{actual_profit:,.0f}원"],
        "달성률": [
            f"{(actual_sales/target_sales*100):.1f}%" if target_sales > 0 else "-",
            f"{(actual_profit/target_profit*100):.1f}%" if target_profit > 0 else "-"
        ]
    }
    st.table(pd.DataFrame(summary_data))

    with st.expander(f"📋 {selected_manager}님의 기여 내역 상세 확인", expanded=False):
        if detail_records:
            df_detail = pd.DataFrame(detail_records)
            st.dataframe(
                df_detail.style.format({
                    "원매출": "{:,.0f}원", "반영매출": "{:,.0f}원",
                    "원이익": "{:,.0f}원", "반영이익": "{:,.0f}원"
                }),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("선택된 기간 내 실적 내역이 없습니다.")

    if st.button("📊 실적 대시보드로 돌아가기"):
        st.session_state.page = "dashboard"
        st.rerun()

# 메인 화면 - 실적 대시보드
else:
    st.title("📊 Deal-ito 통합 실적/이익 대시보드")
    st.markdown("좌측 사이드바에서 엑셀 파일을 업로드하면 실적을 자동 계산합니다.")
    df = load_dashboard_data()
    if df is not None:
        targets_data = load_targets()
        excel_managers = sorted(list(set(df['Deal - 담당자_고객'].dropna().unique().tolist() + df['Deal - 담당자_관리'].dropna().unique().tolist() + df['Deal - 담당자_소싱'].dropna().unique().tolist())))
        updated = False
        for manager in excel_managers:
            if manager not in targets_data:
                targets_data[manager] = {f"q{i}": {"mm": 0.0, "sales": 0.0, "profit": 0.0} for i in range(1, 5)}
                updated = True
        if updated:
            save_targets(targets_data); st.toast("새로운 담당자가 목표 명단에 추가되었습니다!")
        st.sidebar.header("🔍 조회 조건 설정")
        period_map = {"전체 (1-12월)": list(range(1, 13)), "1분기 (1-3월)": list(range(1, 4)), "2분기 (4-6월)": list(range(4, 7)), "3분기 (7-9월)": list(range(7, 10)), "4분기 (10-12월)": list(range(10, 13)), "상반기 (1-6월)": list(range(1, 7)), "하반기 (7-12월)": list(range(7, 13))}
        selected_period_label = st.sidebar.selectbox("조회 기간을 선택하세요", list(period_map.keys()))
        selected_months = [f"{m:02d}" for m in period_map[selected_period_label]]
        
        # 담당자 선택을 체크박스 형태로 변경
        st.sidebar.write("👤 **조회할 담당자 선택**")
        all_managers = sorted(list(set(df['Deal - 담당자_고객'].dropna().unique().tolist() + df['Deal - 담당자_관리'].dropna().unique().tolist() + df['Deal - 담당자_소싱'].dropna().unique().tolist())))
        
        # 사이드바 전체 선택 기능
        def on_dash_select_all_change():
            for mgr in all_managers:
                st.session_state[f"dash_sel_{mgr}"] = st.session_state.dash_select_all_key

        dash_select_all = st.sidebar.checkbox("전체 선택", key="dash_select_all_key", value=True, on_change=on_dash_select_all_change)
        
        selected_managers = []
        for mgr in all_managers:
            # 개별 체크박스 (기본값 True)
            if st.sidebar.checkbox(mgr, key=f"dash_sel_{mgr}", value=True):
                selected_managers.append(mgr)
        
        st.sidebar.write("---")
        # 0. 데이터 전처리: 합계/소계 행 제외 및 담당자 공백 제거
        exclude_keywords = ['합계', '소계', 'total', 'sum']
        df = df[~df['Deal - 이름'].astype(str).str.lower().str.contains('|'.join(exclude_keywords), na=False)]
        
        for col in ['Deal - 담당자_고객', 'Deal - 담당자_관리', 'Deal - 담당자_소싱']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
        profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
        
        def clean_currency_final(val):
            if pd.isna(val): return 0.0
            if isinstance(val, (int, float)): return float(val)
            cleaned = re.sub(r'[^0-9.-]', '', str(val))
            try: return float(cleaned)
            except: return 0.0

        for col in sales_cols + profit_cols:
            if col in df.columns:
                df[col] = df[col].apply(clean_currency_val) # 기존 정의된 함수 사용

        df['선택기간_총매출'] = df[sales_cols].sum(axis=1)
        df['선택기간_총이익'] = df[profit_cols].sum(axis=1)

        def calculate_consolidated_results(target_col):
            role_configs = [
                ('고객', 0.4, 'Deal - 담당자_고객'), 
                ('관리', 0.3, 'Deal - 담당자_관리'), 
                ('소싱', 0.3, 'Deal - 담당자_소싱')
            ]
            individual_results = []
            for role_label, ratio, manager_col in role_configs:
                if manager_col in df.columns:
                    temp = df[['Deal - 이름', manager_col, target_col]].copy()
                    temp['반영비율'] = f"{int(ratio*100)}%"
                    temp['반영실적'] = temp[target_col] * ratio
                    temp['역할'] = role_label
                    temp.columns = ['Deal명', '담당자', '원금액', '반영비율', '반영실적', '역할']
                    individual_results.append(temp)
            
            combined = pd.concat(individual_results)
            # 담당자 이름 기준 필터링
            if selected_managers:
                combined = combined[combined['담당자'].isin(selected_managers)]
            
            return combined

        tab1, tab2 = st.tabs(["💰 매출 분석", "📉 이익 분석"])
        with tab1:
            st.subheader(f"📅 매출 조회 기간: {selected_period_label}")
            m_df = calculate_consolidated_results('선택기간_총매출')
            summary_m = m_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
            
            st.write("#### 👤 담당자별 합산 실적 요약")
            st.dataframe(summary_m.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
            
            with st.expander("📋 매출 상세 기여 내역 확인 (어떻게 계산되었나요?)", expanded=False):
                st.markdown("""
                **계산 규칙**: 각 Deal의 해당 월 매출에 대해 **고객(40%), 관리(30%), 소싱(30%)** 비율을 적용하여 합산합니다.
                """)
                st.dataframe(
                    m_df[m_df['반영실적'] > 0].style.format({
                        '원금액': '{:,.0f}원', 
                        '반영실적': '{:,.0f}원'
                    }), 
                    use_container_width=True, 
                    hide_index=True
                )

        with tab2:
            st.subheader(f"📅 이익 조회 기간: {selected_period_label}")
            p_df = calculate_consolidated_results('선택기간_총이익')
            summary_p = p_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
            
            st.write("#### 👤 담당자별 합산 실적 요약")
            st.dataframe(summary_p.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
            
            with st.expander("📋 이익 상세 기여 내역 확인 (어떻게 계산되었나요?)", expanded=False):
                st.markdown("""
                **계산 규칙**: 각 Deal의 해당 월 이익에 대해 **고객(40%), 관리(30%), 소싱(30%)** 비율을 적용하여 합산합니다.
                """)
                st.dataframe(
                    p_df[p_df['반영실적'] > 0].style.format({
                        '원금액': '{:,.0f}원', 
                        '반영실적': '{:,.0f}원'
                    }), 
                    use_container_width=True, 
                    hide_index=True
                )
    else:
        st.info("좌측 사이드바에서 실적 엑셀 파일을 업로드해 주세요.")
