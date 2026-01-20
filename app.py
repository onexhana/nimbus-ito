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
def draw_gauge(current_val, target_val, title, color="#636EFA", bg_color="#F0F2F6"):
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
            'bar': {'color': color},
            'bgcolor': bg_color, # 배경색을 파라미터로 받음
            'borderwidth': 0,
        }
    ))
    
    # 중앙에 핵심 정보 배치 (달성률, 목표, 실적)
    fig.add_annotation(
        x=0.5, y=0.35,
        text=f"<span style='font-size:26px; font-weight:bold; color:{color};'>{percentage:.1f}% 달성</span><br><br>" +
             f"<span style='font-size:15px; color:gray;'>목표: {target_val:,.0f}원</span><br>" +
             f"<span style='font-size:18px; font-weight:bold;'>실적: {current_val:,.0f}원</span>",
        showarrow=False,
        align="center"
    )
    
    fig.update_layout(height=400, margin=dict(l=50, r=50, t=80, b=20))
    return fig

# 페이지 설정
st.set_page_config(page_title="Deal-ito 통합 실적 대시보드", layout="wide")

# 세션 상태 초기화
if 'page' not in st.session_state:
    st.session_state.page = "dashboard"

# 사이드바 메뉴 구성
st.sidebar.title("📌 메뉴")
if st.sidebar.button("👥 인력 명단 관리", use_container_width=True):
    st.session_state.page = "personnel"
if st.sidebar.button("📊 실적 대시보드", use_container_width=True):
    st.session_state.page = "dashboard"
if st.sidebar.button("🎯 목표 설정하기", use_container_width=True):
    st.session_state.page = "targets"
if st.sidebar.button("📈 목표 달성률 확인하기", use_container_width=True):
    st.session_state.page = "achievement"

st.sidebar.write("---")

# 1. 실적 데이터 업로드 (대시보드 페이지 전용 사이드바 메뉴)
if st.session_state.page == "dashboard":
    with st.sidebar.expander("📁 실적 데이터 업로드 및 관리", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"], key="dashboard_uploader")
        if uploaded_file:
            df_loaded = pd.read_excel(uploaded_file)
            save_dashboard_data(df_loaded)
            st.sidebar.success("로드 완료!")
            st.rerun()
        if st.button("🗑️ 업로드된 데이터 삭제", use_container_width=True):
            delete_dashboard_data()
            st.rerun()

# --- 메인 화면 로직 ---

# 1. 인력 명단 관리
if st.session_state.page == "personnel":
    st.title("👥 인력 명단 관리")
    st.markdown("담당자를 '내부' 또는 '외부' 인력으로 분류합니다. 분류된 정보는 목표 설정 및 실적 계산의 기준이 됩니다.")
    
    targets_data = load_targets()
    
    if not targets_data:
        st.warning("등록된 담당자가 없습니다. '실적 대시보드'에서 엑셀을 업로드하거나 '목표 설정하기'에서 담당자를 추가해 주세요.")
    else:
        # 데이터 구조 보정 (type 필드 없는 경우 기본값 '내부' 부여)
        updated = False
        for mgr in targets_data:
            if "type" not in targets_data[mgr]:
                targets_data[mgr]["type"] = "내부"
                updated = True
        if updated:
            save_targets(targets_data)

        with st.form("personnel_form"):
            st.subheader("📋 담당자 분류 설정")
            
            h1, h2, h3 = st.columns([2, 3, 2])
            h1.markdown("**성명**")
            h2.markdown("**분류 (내부/외부)**")
            h3.markdown("**현재 상태**")
            st.write("---")

            new_classifications = {}
            for mgr in sorted(targets_data.keys()):
                c1, c2, c3 = st.columns([2, 3, 2])
                c1.write(f"**{mgr}**")
                
                current_type = targets_data[mgr].get("type", "내부")
                selected_type = c2.radio(
                    f"분류_{mgr}", 
                    options=["내부", "외부"], 
                    index=0 if current_type == "내부" else 1,
                    horizontal=True,
                    label_visibility="collapsed"
                )
                new_classifications[mgr] = selected_type
                
                status_color = "blue" if selected_type == "내부" else "orange"
                c3.markdown(f":{status_color}[{selected_type} 인력]")

            st.write("---")
            save_btn = st.form_submit_button("💾 분류 정보 저장", use_container_width=True, type="primary")
            
            if save_btn:
                for mgr, p_type in new_classifications.items():
                    targets_data[mgr]["type"] = p_type
                save_targets(targets_data)
                st.success("인력 분류 정보가 저장되었습니다!")
                time.sleep(1)
                st.rerun()

# 2. 목표 설정
elif st.session_state.page == "targets":
    st.title("🎯 담당자별 목표 설정")
    st.markdown("분기별 목표를 입력하세요. (단위: 만원)")
    
    targets_data = load_targets()
    # 인력 데이터 구조 보정
    for mgr in targets_data:
        if "type" not in targets_data[mgr]:
            targets_data[mgr]["type"] = "내부"

    all_mgrs = sorted(targets_data.keys())
    
    if all_mgrs:
        team_total_mm = sum(float(targets_data[m][q]["mm"]) for m in all_mgrs for q in ["q1", "q2", "q3", "q4"])
        team_total_sales = sum(float(targets_data[m][q]["sales"]) for m in all_mgrs for q in ["q1", "q2", "q3", "q4"])
        team_total_profit = sum(float(targets_data[m][q]["profit"]) for m in all_mgrs for q in ["q1", "q2", "q3", "q4"])
        
        with st.expander("📊 우리 팀 전체 연간 목표 합계 확인", expanded=False):
            tc1, tc2, tc3 = st.columns(3)
            tc1.metric("팀 전체 총 MM", f"{team_total_mm:.1f}")
            tc2.metric("팀 전체 총 매출", f"{team_total_sales:,.0f}만원")
            tc3.metric("팀 전체 총 이익", f"{team_total_profit:,.0f}만원")
            
    st.write("### 👥 담당자 선택")
    
    def on_select_all_change():
        for mgr in all_mgrs:
            st.session_state[f"sel_{mgr}"] = st.session_state.select_all_key
            
    def on_select_internal_change():
        for mgr in internal_mgrs:
            st.session_state[f"sel_{mgr}"] = st.session_state.select_internal_key
            
    def on_select_external_change():
        for mgr in external_mgrs:
            st.session_state[f"sel_{mgr}"] = st.session_state.select_external_key
    
    col_sel1, col_sel2, col_sel3 = st.columns(3)
    select_all = col_sel1.checkbox("모든 인력 전체 선택", key="select_all_key", on_change=on_select_all_change)
    select_internal = col_sel2.checkbox("내부 인력 전체 선택", key="select_internal_key", on_change=on_select_internal_change)
    select_external = col_sel3.checkbox("외부 인력 전체 선택", key="select_external_key", on_change=on_select_external_change)
    
    selected_m_list = []
    internal_mgrs = [m for m in all_mgrs if targets_data[m].get("type") == "내부"]
    external_mgrs = [m for m in all_mgrs if targets_data[m].get("type") == "외부"]

    if internal_mgrs:
        st.markdown("#### 🏠 내부 인력")
        mgr_cols = st.columns(5)
        for i, mgr in enumerate(internal_mgrs):
            with mgr_cols[i % 5]:
                if st.checkbox(mgr, key=f"sel_{mgr}"):
                    selected_m_list.append(mgr)
    
    if external_mgrs:
        st.write("")
        st.markdown("#### 🌐 외부 인력")
        mgr_cols = st.columns(5)
        for i, mgr in enumerate(external_mgrs):
            with mgr_cols[i % 5]:
                if st.checkbox(mgr, key=f"sel_{mgr}"):
                    selected_m_list.append(mgr)
    
    if selected_m_list:
        st.write("---")
        st.write(f"### 📋 선택된 담당자 목표 현황 요약")
        
        total_sum_mm = sum(float(targets_data[m][q]["mm"]) for m in selected_m_list for q in ["q1", "q2", "q3", "q4"])
        total_sum_sales = sum(float(targets_data[m][q]["sales"]) for m in selected_m_list for q in ["q1", "q2", "q3", "q4"])
        total_sum_profit = sum(float(targets_data[m][q]["profit"]) for m in selected_m_list for q in ["q1", "q2", "q3", "q4"])
        
        c1, c2, c3 = st.columns(3)
        c1.metric("선택 인원 총 MM", f"{total_sum_mm:.1f}")
        c2.metric("선택 인원 총 매출", f"{total_sum_sales:,.0f}만원")
        c3.metric("선택 인원 총 이익", f"{total_sum_profit:,.0f}만원")
        
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
                q1_sum = sum(float(targets_data[m]["q1"][category]) for m in selected_m_list)
                q2_sum = sum(float(targets_data[m]["q2"][category]) for m in selected_m_list)
                q3_sum = sum(float(targets_data[m]["q3"][category]) for m in selected_m_list)
                q4_sum = sum(float(targets_data[m]["q4"][category]) for m in selected_m_list)
                total_row = {"성명": "★ 전체 합계" if label == "MM" else "", "내용": label, "1/4분기": q1_sum, "2/4분기": q2_sum, "3/4분기": q3_sum, "4/4분기": q4_sum, "년 합계": q1_sum + q2_sum + q3_sum + q4_sum}
                summary_rows.append(total_row)
        
        if summary_rows:
            df_summary = pd.DataFrame(summary_rows)
            format_mapping = {col: lambda x: f"{round(x):,.0f}" if isinstance(x, (int, float)) else x for col in ["1/4분기", "2/4분기", "3/4분기", "4/4분기", "년 합계"]}
            st.dataframe(df_summary.style.format(format_mapping), use_container_width=True, hide_index=True)
        
        st.write("---")
        if len(selected_m_list) == 1:
            selected_m = selected_m_list[0]
            m_data = targets_data[selected_m]
            with st.container(border=True):
                st.subheader(f"👤 {selected_m}님의 목표 설정")
                cols = st.columns(4); updated_m_data = {}
                for i, q in enumerate(["q1", "q2", "q3", "q4"]):
                    with cols[i]:
                        st.markdown(f"**{i+1}/4분기**")
                        mm = st.number_input(f"MM", value=float(m_data[q]["mm"]), key=f"{selected_m}_{q}_mm", step=0.1)
                        sales = st.number_input(f"매출(만원)", value=float(m_data[q]["sales"]), key=f"{selected_m}_{q}_sales", step=100.0)
                        profit = st.number_input(f"이익(만원)", value=float(m_data[q]["profit"]), key=f"{selected_m}_{q}_profit", step=100.0)
                        updated_m_data[q] = {"mm": mm, "sales": sales, "profit": profit}
                total_mm = sum(q_val["mm"] for q_val in updated_m_data.values())
                total_sales = sum(q_val["sales"] for q_val in updated_m_data.values())
                total_profit = sum(q_val["profit"] for q_val in updated_m_data.values())
                c1, c2, c3 = st.columns(3); c1.metric("연간 총 MM", f"{total_mm:.2f}"); c2.metric("연간 총 매출", f"{total_sales:,.0f}만원"); c3.metric("연간 총 이익", f"{total_profit:,.0f}만원")
                if st.button(f"💾 {selected_m}님 목표 저장", use_container_width=True, type="primary"):
                    targets_data[selected_m].update(updated_m_data)
                    save_targets(targets_data); st.success("저장되었습니다!"); st.rerun()

    with st.sidebar.expander("📂 목표 데이터 일괄 관리", expanded=False):
        template_excel = create_excel_template(targets_data)
        st.download_button("📥 양식 다운로드", data=template_excel, file_name="target_template.xlsx", use_container_width=True)
        uploaded_target_file = st.file_uploader("엑셀 업로드 (사업목표 양식)", type=["xlsx"], key="target_excel_uploader")
        
        if uploaded_target_file:
            with st.status("📊 엑셀 분석 중...", expanded=True) as status:
                try:
                    up_df = pd.read_excel(uploaded_target_file)
                    
                    # 컬럼명 정리 (공백 제거)
                    up_df.columns = [str(c).strip() for c in up_df.columns]
                    
                    # 필수 컬럼 존재 확인
                    required_cols = ['성명', '내용', '1/4분기 목표', '2/4분기 목표', '3/4분기 목표', '4/4분기 목표']
                    if not all(col in up_df.columns for col in required_cols):
                        st.error("엑셀 양식이 올바르지 않습니다. '성명', '내용', '1/4분기 목표' 등의 컬럼이 필요합니다.")
                        st.stop()

                    new_targets = targets_data.copy()
                    current_type = "내부" # 기본값
                    
                    # 데이터 행 순회
                    # 성명 컬럼 ffill 처리 (MM, 매출, 매출이익 3줄을 하나로 묶기 위함)
                    up_df['성명_fill'] = up_df['성명'].fillna(method='ffill').str.replace(" ", "")
                    
                    for mgr_name in up_df['성명_fill'].unique():
                        if pd.isna(mgr_name) or mgr_name in ['전체', '내용', 'nan']: continue
                        
                        # 구분선 감지
                        if '내부' in mgr_name:
                            current_type = "내부"
                            continue
                        if '외부' in mgr_name:
                            current_type = "외부"
                            continue
                            
                        # 실제 인력 데이터 처리
                        mgr_rows = up_df[up_df['성명_fill'] == mgr_name]
                        if mgr_name not in new_targets:
                            new_targets[mgr_name] = {f"q{i}": {"mm": 0, "sales": 0, "profit": 0} for i in range(1, 5)}
                        
                        new_targets[mgr_name]["type"] = current_type
                        
                        for _, row in mgr_rows.iterrows():
                            content = str(row['내용']).strip()
                            cat = "mm" if "MM" in content else "sales" if "매출" == content else "profit" if "매출이익" == content else None
                            
                            if cat:
                                for i in range(1, 5):
                                    val = row[f'{i}/4분기 목표']
                                    new_targets[mgr_name][f"q{i}"][cat] = float(val) if pd.notna(val) and str(val).strip() != "-" else 0.0
                    
                    save_targets(new_targets)
                    status.update(label="✅ 반영 완료!", state="complete", expanded=False)
                    st.success(f"엑셀 데이터를 기반으로 인력 분류 및 목표치가 업데이트되었습니다.")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    status.update(label="❌ 오류 발생", state="error")
                    st.error(f"엑셀 처리 중 오류가 발생했습니다: {e}")
        
        st.write("---")
        if st.button("🚨 모든 데이터 초기화", use_container_width=True):
            st.session_state.show_reset_confirm = True
        
        if st.session_state.get('show_reset_confirm', False):
            st.warning("⚠️ 모든 목표 데이터를 삭제하시겠습니까?")
            c1, c2 = st.columns(2)
            if c1.button("✅ 예", use_container_width=True):
                save_targets({})
                st.session_state.show_reset_confirm = False
                st.rerun()
            if c2.button("❌ 아니오", use_container_width=True):
                st.session_state.show_reset_confirm = False
                st.rerun()

# 3. 목표 달성률 확인
elif st.session_state.page == "achievement":
    st.title("📈 목표 달성률 확인하기")
    df = load_dashboard_data()
    targets_data = load_targets()
    
    if df is None: st.warning("📊 실적 대시보드에서 먼저 엑셀 파일을 업로드해 주세요."); st.stop()
    if not targets_data: st.warning("🎯 목표 설정하기에서 먼저 목표를 설정해 주세요."); st.stop()

    st.sidebar.header("🔍 달성률 조회 조건")
    period_map = {"1분기 (1-3월)": (list(range(1, 4)), ["q1"]), "2분기 (4-6월)": (list(range(4, 7)), ["q2"]), "3분기 (7-9월)": (list(range(7, 10)), ["q3"]), "4분기 (10-12월)": (list(range(10, 13)), ["q4"]), "1-9월": (list(range(1, 10)), ["q1", "q2", "q3"]), "상반기 (1-6월)": (list(range(1, 7)), ["q1", "q2"]), "하반기 (7-12월)": (list(range(7, 13)), ["q3", "q4"]), "전체 (1-12월)": (list(range(1, 13)), ["q1", "q2", "q3", "q4"])}
    selected_period_label = st.sidebar.selectbox("조회 기간 선택", list(period_map.keys()))
    months, quarters = period_map[selected_period_label]
    selected_months = [f"{m:02d}" for m in months]
    
    all_managers = sorted(targets_data.keys())
    internal_managers = [m for m in all_managers if targets_data[m].get("type") == "내부"]
    external_managers = [m for m in all_managers if targets_data[m].get("type") == "외부"]
    
    selected_manager = st.sidebar.selectbox(
        "조회할 담당자 선택", 
        ["선택하세요", "★ 전체 담당자 한눈에 보기", "🏠 내부 인력 전체보기", "🌐 외부 인력 전체보기"] + all_managers
    )
    
    if selected_manager == "선택하세요": st.info("👈 좌측 사이드바에서 조회할 담당자를 선택해 주세요."); st.stop()
    
    st.write(f"### 👤 {selected_manager}님의 {selected_period_label} 달성 현황")
    
    if selected_manager == "★ 전체 담당자 한눈에 보기":
        target_sales = sum(float(targets_data[m][q]["sales"]) for m in all_managers for q in quarters) * 10000
        target_profit = sum(float(targets_data[m][q]["profit"]) for m in all_managers for q in quarters) * 10000
        managers_to_check = all_managers
    elif selected_manager == "🏠 내부 인력 전체보기":
        target_sales = sum(float(targets_data[m][q]["sales"]) for m in internal_managers for q in quarters) * 10000
        target_profit = sum(float(targets_data[m][q]["profit"]) for m in internal_managers for q in quarters) * 10000
        managers_to_check = internal_managers
    elif selected_manager == "🌐 외부 인력 전체보기":
        target_sales = sum(float(targets_data[m][q]["sales"]) for m in external_managers for q in quarters) * 10000
        target_profit = sum(float(targets_data[m][q]["profit"]) for m in external_managers for q in quarters) * 10000
        managers_to_check = external_managers
    else:
        target_sales = sum(float(targets_data[selected_manager][q]["sales"]) for q in quarters) * 10000
        target_profit = sum(float(targets_data[selected_manager][q]["profit"]) for q in quarters) * 10000
        managers_to_check = [selected_manager]

    df = df[df['Deal - 이름'].notna() & (df['Deal - 이름'].astype(str).str.strip() != "")]
    df = df[~df['Deal - 이름'].astype(str).str.contains('합계|소계|total|sum', na=False)]
    if 'People - 이름' in df.columns: df['Deal - 이름'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"
    for col in ['Deal - 담당자_고객', 'Deal - 담당자_관리', 'Deal - 담당자_소싱']:
        if col in df.columns: df[col] = df[col].astype(str).str.strip()

    actual_sales = 0.0; actual_profit = 0.0; detail_records = []
    role_configs = [(0.4, 'Deal - 담당자_고객', '고객'), (0.3, 'Deal - 담당자_관리', '관리'), (0.3, 'Deal - 담당자_소싱', '소싱')]
    
    for current_mgr in managers_to_check:
        for ratio, mgr_col, role_name in role_configs:
            if mgr_col in df.columns:
                matched_df = df[df[mgr_col] == current_mgr]
                for idx, row in matched_df.iterrows():
                    m_vals = {m: clean_currency_val(row[f"Deal - @월별매출 ({m})"]) for m in selected_months if f"Deal - @월별매출 ({m})" in row}
                    p_vals = {m: clean_currency_val(row[f"Deal - @월별이익 ({m})"]) for m in selected_months if f"Deal - @월별이익 ({m})" in row}
                    row_s = sum(m_vals.values()); row_p = sum(p_vals.values())
                    if row_s > 0 or row_p > 0:
                        actual_sales += row_s * ratio; actual_profit += row_p * ratio
                        record = {"Deal명": row['Deal - 이름'], "담당자": current_mgr, "역할": role_name, "비중": f"{int(ratio*100)}%", "원매출(합계)": row_s, "반영매출": row_s * ratio}
                        for m, val in m_vals.items(): record[f"{int(m)}월 매출"] = val
                        detail_records.append(record)

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(draw_gauge(actual_sales, target_sales, "💰 매출 달성 현황", color="#EF553B", bg_color="#FCEAE8"), use_container_width=True)
    with c2: st.plotly_chart(draw_gauge(actual_profit, target_profit, "📉 이익 달성 현황", color="#636EFA", bg_color="#EBEDFE"), use_container_width=True)
    
    summary_data = {"구분": ["매출", "이익"], "목표 금액": [f"{target_sales:,.0f}원", f"{target_profit:,.0f}원"], "실제 실적": [f"{actual_sales:,.0f}원", f"{actual_profit:,.0f}원"], "달성률": [f"{(actual_sales/target_sales*100):.1f}%" if target_sales > 0 else "-", f"{(actual_profit/target_profit*100):.1f}%" if target_profit > 0 else "-"]}
    st.table(pd.DataFrame(summary_data))

    with st.expander(f"📋 {selected_manager}님의 기여 내역 상세 확인", expanded=False):
        if detail_records:
            df_detail = pd.DataFrame(detail_records)
            format_dict = {"원매출(합계)": "{:,.0f}원", "반영매출": "{:,.0f}원"}
            for col in df_detail.columns:
                if "월 매출" in col: format_dict[col] = "{:,.0f}원"
            st.dataframe(df_detail.style.format(format_dict), use_container_width=True, hide_index=True)

# 4. 실적 대시보드
else:
    st.title("📊 Deal-ito 통합 실적/이익 대시보드")
    st.markdown("좌측 사이드바에서 엑셀 파일을 업로드하면 실적을 자동 계산합니다.")
    df = load_dashboard_data()
    if df is not None:
        period_map = {"전체 (1-12월)": list(range(1, 13)), "1-9월": list(range(1, 10)), "1분기 (1-3월)": list(range(1, 4)), "2분기 (4-6월)": list(range(4, 7)), "3분기 (7-9월)": list(range(7, 10)), "4분기 (10-12월)": list(range(10, 13))}
        selected_period_label = st.sidebar.selectbox("조회 기간 선택", list(period_map.keys()))
        selected_months = [f"{m:02d}" for m in period_map[selected_period_label]]
        
        all_mgrs = sorted(list(set(df['Deal - 담당자_고객'].dropna().unique().tolist() + df['Deal - 담당자_관리'].dropna().unique().tolist() + df['Deal - 담당자_소싱'].dropna().unique().tolist())))
        selected_managers = [mgr for mgr in all_mgrs if st.sidebar.checkbox(mgr, value=True, key=f"dash_sel_{mgr}")]

        df = df[~df['Deal - 이름'].astype(str).str.contains('합계|소계|total|sum', na=False)]
        if 'People - 이름' in df.columns: df['Deal - 이름'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"
        
        sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
        profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
        for col in sales_cols + profit_cols:
            if col in df.columns: df[col] = df[col].apply(clean_currency_val)
        df['선택기간_총매출'] = df[sales_cols].sum(axis=1)
        df['선택기간_총이익'] = df[profit_cols].sum(axis=1)

        def calc_consolidated(target_col):
            results = []
            for role, ratio, col in [('고객', 0.4, 'Deal - 담당자_고객'), ('관리', 0.3, 'Deal - 담당자_관리'), ('소싱', 0.3, 'Deal - 담당자_소싱')]:
                if col in df.columns:
                    temp = df[['Deal - 이름', col, target_col]].copy()
                    temp['비중_num'] = ratio; temp['반영실적'] = temp[target_col] * ratio; temp['역할'] = role
                    temp.columns = ['Deal명', '담당자', '원금액', '비중_num', '반영실적', '역할']
                    results.append(temp)
            combined = pd.concat(results)
            return combined[combined['담당자'].isin(selected_managers)] if selected_managers else combined

        tab1, tab2 = st.tabs(["💰 매출 분석", "📉 이익 분석"])
        for tab, col_name, label in [(tab1, '선택기간_총매출', "매출"), (tab2, '선택기간_총이익', "이익")]:
            with tab:
                res_df = calc_consolidated(col_name)
                sum_df = res_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values('반영실적', ascending=False)
                if not sum_df.empty: sum_df = pd.concat([sum_df, pd.DataFrame([{'담당자': '★ 전체 합계', '반영실적': sum_df['반영실적'].sum()}])], ignore_index=True)
                st.write(f"#### 👤 담당자별 합산 실적 요약")
                st.dataframe(sum_df.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
                with st.expander(f"📋 {label} 상세 기여 내역 확인", expanded=False):
                    disp_df = res_df[res_df['반영실적'] > 0].groupby(['Deal명', '담당자']).agg({'원금액': 'first', '역할': lambda x: ', '.join(x), '비중_num': 'sum', '반영실적': 'sum'}).reset_index()
                    disp_df['반영비율'] = disp_df['비중_num'].apply(lambda x: f"{int(x*100)}%")
                    st.dataframe(disp_df[['Deal명', '담당자', '역할', '원금액', '반영비율', '반영실적']].style.format({'원금액': '{:,.0f}원', '반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
    else: st.info("좌측 사이드바에서 실적 엑셀 파일을 업로드해 주세요.")
