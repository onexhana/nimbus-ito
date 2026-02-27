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

def delete_dashboard_data():
    if 'dashboard_df' in st.session_state:
        del st.session_state.dashboard_df
    if os.path.exists(DASHBOARD_CACHE_FILE):
        os.remove(DASHBOARD_CACHE_FILE)

def load_targets():
    """데이터 구조 마이그레이션을 포함한 통합 데이터 로드"""
    if os.path.exists(TARGETS_FILE):
        try:
            with open(TARGETS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            # 구조 마이그레이션 확인
            if "personnel" not in data or "targets" not in data:
                new_data = {"personnel": {}, "targets": {"2026": {}}}
                
                # 기존 데이터가 평면 구조인 경우 (담당자명: 데이터)
                for key, value in data.items():
                    # 이미 연도별 구조인 경우 처리
                    if str(key).isdigit():
                        new_data["targets"][str(key)] = value
                        # 연도별 데이터 내부에 type 정보가 있으면 추출
                        for mgr, mgr_data in value.items():
                            if isinstance(mgr_data, dict) and "type" in mgr_data:
                                new_data["personnel"][mgr] = {"type": mgr_data["type"]}
                    else:
                        # 담당자명: 데이터 구조인 경우
                        if isinstance(value, dict):
                            # 인력 분류 정보 추출
                            if "type" in value:
                                new_data["personnel"][key] = {"type": value["type"]}
                            
                            # 목표 정보 추출 (q1~q4가 있으면 목표 데이터로 간주)
                            if any(f"q{i}" in value for i in range(1, 5)):
                                # 2026년으로 할당
                                target_vals = {f"q{i}": value.get(f"q{i}", {"mm": 0, "sales": 0, "profit": 0}) for i in range(1, 5)}
                                new_data["targets"]["2026"][key] = target_vals
                
                data = new_data
                save_targets(data) # 마이그레이션 후 저장
            return data
        except:
            return {"personnel": {}, "targets": {"2026": {}}}
    return {"personnel": {}, "targets": {"2026": {}}}

def save_targets(data):
    with open(TARGETS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def get_personnel_data():
    all_data = load_targets()
    return all_data.get("personnel", {})

def save_personnel_data(personnel_data):
    all_data = load_targets()
    all_data["personnel"] = personnel_data
    save_targets(all_data)

def get_targets_by_year(year):
    all_data = load_targets()
    return all_data.get("targets", {}).get(str(year), {})

def save_targets_by_year(year, year_targets):
    all_data = load_targets()
    if "targets" not in all_data:
        all_data["targets"] = {}
    all_data["targets"][str(year)] = year_targets
    save_targets(all_data)

def save_dashboard_data(df):
    st.session_state.dashboard_df = df
    df.to_pickle(DASHBOARD_CACHE_FILE)
    # 데이터 저장 시 담당자 명단 즉시 동기화
    all_data = load_targets()
    personnel = all_data["personnel"]
    
    excel_managers = sorted(list(set(
        df['Deal - 담당자_고객'].dropna().unique().tolist() + 
        df['Deal - 담당자_관리'].dropna().unique().tolist() + 
        df['Deal - 담당자_소싱'].dropna().unique().tolist()
    )))
    
    updated = False
    for manager in excel_managers:
        if manager not in personnel:
            personnel[manager] = {"type": "내부"}
            updated = True
    
    if updated:
        save_personnel_data(personnel)

# 엑셀 템플릿 생성 함수 (특정 연도 기준)
def create_excel_template(targets_data, selected_year=2026):
    output = io.BytesIO()
    rows = []
    # targets_data는 특정 연도의 데이터임
    managers = sorted(targets_data.keys()) if targets_data else ["고봉수", "김길래", "박승수", "손병희", "이민지"]
    for mgr in managers:
        m_data = targets_data.get(mgr, {f"q{i}": {"mm": 0, "sales": 0, "profit": 0} for i in range(1, 5)})
        # mm, sales, profit 이외의 필드(type 등)는 건너뜀
        if not isinstance(m_data, dict) or "q1" not in m_data:
            continue
            
        for category, label in [("mm", "MM"), ("sales", "매출"), ("profit", "매출이익")]:
            row = {
                "성명": mgr if label == "MM" else "",
                "내용": label,
                "년 목표": sum(m_data.get(f"q{i}", {}).get(category, 0) for i in range(1, 5)),
                "1/4분기 목표": m_data.get("q1", {}).get(category, 0),
                "2/4분기 목표": m_data.get("q2", {}).get(category, 0),
                "3/4분기 목표": m_data.get("q3", {}).get(category, 0),
                "4/4분기 목표": m_data.get("q4", {}).get(category, 0)
            }
            rows.append(row)
    df_template = pd.DataFrame(rows)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name=f'{selected_year}년 목표설정')
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

def parse_period_input(input_str):
    """'1-3' 형태의 입력을 받아 월 리스트와 분기 리스트 반환"""
    try:
        if not input_str or '-' not in input_str:
            if input_str.isdigit():
                start = end = int(input_str)
            else:
                return [f"{m:02d}" for m in range(1, 13)], ["q1", "q2", "q3", "q4"], "전체 (1-12월)"
        else:
            parts = input_str.split('-')
            start, end = int(parts[0]), int(parts[1])
        
        months = list(range(start, end + 1))
        selected_months = [f"{m:02d}" for m in months]
        
        quarters = []
        if any(m in [1, 2, 3] for m in months): quarters.append("q1")
        if any(m in [4, 5, 6] for m in months): quarters.append("q2")
        if any(m in [7, 8, 9] for m in months): quarters.append("q3")
        if any(m in [10, 11, 12] for m in months): quarters.append("q4")
        
        return selected_months, quarters, f"{start}-{end}월"
    except:
        return [f"{m:02d}" for m in range(1, 13)], ["q1", "q2", "q3", "q4"], "전체 (1-12월)"

# 게이지 차트 생성 함수
def draw_gauge(current_val, target_val, title, color="#636EFA", bg_color="#F0F2F6"):
    percentage = (current_val / target_val * 100) if target_val > 0 else 0
    
    # 사용자의 요청대로 목표 금액을 100% 지점(Max)으로 설정
    max_range = target_val if target_val > 0 else max(current_val, 100)
    
    fig = go.Figure(go.Indicator(
        mode = "gauge", # redundant number 제거
        value = min(current_val, max_range), # 바는 일단 max_range까지만
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': f"<b>{title}</b>", 'font': {'size': 20, 'color': color}},
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
        text=f"<span style='font-size:26px; font-weight:bold; color:{color};'>{percentage:.1f}% 달성</span><br><br><br>" +
             f"<span style='font-size:15px; color:gray;'>목표: {target_val:,.0f}원</span><br><br>" +
             f"<span style='font-size:18px; font-weight:bold; color:{color};'>실적: {current_val:,.0f}원</span>",
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

# 현재 페이지 버튼 강조용 키워드
_page_keywords = {
    "personnel": "인력 명단",
    "targets": "목표 설정",
    "achievement": "구성원별 달성률",
    "monthly_sales": "월별 매출",
    "dashboard": "전체 실적 대시보드",
    "rank_customer": "고객사별",
    "rank_endclient": "엔드클라이언트",
}
_active_keyword = _page_keywords.get(st.session_state.page, "")

# 사이드바 버튼 공통 스타일 (색상은 JS로 적용)
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stButton > button {
        white-space: normal !important;
        height: auto !important;
        min-height: 45px !important;
        line-height: 1.2 !important;
        padding: 10px 5px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 사이드바 버튼 색상 - 여러 DOM 구조 대응
st.markdown("""
    <style>
    /* 방식1: .stButton nth-of-type - 3~7번(구성원별~엔드클라이언트) 연한 빨강 */
    [data-testid="stSidebar"] .stButton:nth-of-type(3) > button,
    [data-testid="stSidebar"] .stButton:nth-of-type(4) > button,
    [data-testid="stSidebar"] .stButton:nth-of-type(5) > button,
    [data-testid="stSidebar"] .stButton:nth-of-type(6) > button,
    [data-testid="stSidebar"] .stButton:nth-of-type(7) > button {
        background-color: #FFF0F0 !important;
        color: #B71C1C !important;
        border: 1px solid #FFCDD2 !important;
    }
    /* 방식2: element-container (다른 구조) */
    [data-testid="stSidebar"] .element-container:nth-of-type(4) button,
    [data-testid="stSidebar"] .element-container:nth-of-type(5) button,
    [data-testid="stSidebar"] .element-container:nth-of-type(6) button,
    [data-testid="stSidebar"] .element-container:nth-of-type(7) button,
    [data-testid="stSidebar"] .element-container:nth-of-type(8) button {
        background-color: #FFF0F0 !important;
        color: #B71C1C !important;
        border: 1px solid #FFCDD2 !important;
    }
    /* 방식3: row-widget (다른 구조) */
    [data-testid="stSidebar"] .row-widget.stButton:nth-of-type(3) > button,
    [data-testid="stSidebar"] .row-widget.stButton:nth-of-type(4) > button,
    [data-testid="stSidebar"] .row-widget.stButton:nth-of-type(5) > button,
    [data-testid="stSidebar"] .row-widget.stButton:nth-of-type(6) > button,
    [data-testid="stSidebar"] .row-widget.stButton:nth-of-type(7) > button {
        background-color: #FFF0F0 !important;
        color: #B71C1C !important;
        border: 1px solid #FFCDD2 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 현재 페이지 강조 - JS (DOM 완성 대기)
if _active_keyword:
    _js = f'''<div style="height:0;overflow:hidden;"><script>
(function applySidebarStyles(){{
 var doc=document;var s=doc.querySelector("[data-testid=stSidebar]");
 if(!s){{setTimeout(applySidebarStyles,300);return;}}
 var btns=s.querySelectorAll(".stButton>button");if(btns.length<4){{setTimeout(applySidebarStyles,200);return;}}
 var pink=["구성원별 달성률","월별 매출","고객사별","엔드클라이언트"];
 var active="{_active_keyword}";
 for(var i=0;i<btns.length;i++){{
  var t=(btns[i].textContent||"").replace(/\\s+/g," ");
  for(var j=0;j<pink.length;j++){{if(t.indexOf(pink[j])!==-1){{btns[i].style.setProperty("background-color","#FFF0F0","important");btns[i].style.setProperty("color","#B71C1C","important");btns[i].style.setProperty("border","1px solid #FFCDD2","important");break;}}}}
  if(active&&t.indexOf(active)!==-1){{btns[i].style.setProperty("background-color","#FFCDD2","important");btns[i].style.setProperty("border","2px solid #E57373","important");}}
 }}
}})();
</script></div>'''
    try:
        st.html(_js, width=1, unsafe_allow_javascript=True)
    except Exception:
        pass

if st.sidebar.button("인력 명단 관리", use_container_width=True):
    st.session_state.page = "personnel"
if st.sidebar.button("목표 설정하기", use_container_width=True):
    st.session_state.page = "targets"
if st.sidebar.button("1.구성원별 달성률 조회", use_container_width=True):
    st.session_state.page = "achievement"
if st.sidebar.button("2.월별 매출/이익 조회", use_container_width=True):
    st.session_state.page = "monthly_sales"
if st.sidebar.button("3.전체 실적 대시보드", use_container_width=True):
    st.session_state.page = "dashboard"
if st.sidebar.button("4.고객사별 매출/이익 순위 조회", use_container_width=True):
    st.session_state.page = "rank_customer"
if st.sidebar.button("5.엔드클라이언트 매출/이익 순위 조회", use_container_width=True):
    st.session_state.page = "rank_endclient"

st.sidebar.write("---")

# 1. 실적 데이터 업로드 (대시보드 페이지 전용 사이드바 메뉴)
if st.session_state.page == "dashboard":
    with st.sidebar.expander("📁 엑셀 수동 업로드 및 관리", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"], key="dashboard_uploader")
        if uploaded_file:
            df_loaded = pd.read_excel(uploaded_file)
            save_dashboard_data(df_loaded)
            st.sidebar.success("로드 완료! 새로고침(F5)을 해주세요:)")
            st.rerun()
        if st.button("🗑️ 업로드된 데이터 삭제", use_container_width=True):
            delete_dashboard_data()
            st.rerun()

# --- 메인 화면 로직 ---

# 1. 인력 명단 관리
if st.session_state.page == "personnel":
    st.title("👥 인력 명단 관리")
    st.markdown("담당자를 '내부' 또는 '외부' 인력으로 분류합니다. 분류된 정보는 목표 설정 및 실적 계산의 기준이 됩니다.")
    
    personnel_data = get_personnel_data()
    
    if not personnel_data:
        st.warning("등록된 담당자가 없습니다. '전체 실적 대시보드'에서 엑셀을 업로드하거나 '목표 설정하기'에서 담당자를 추가해 주세요.")
    else:
        # 데이터 구조 보정 (type 필드 없는 경우 기본값 '내부' 부여)
        updated = False
        for mgr in personnel_data:
            if not isinstance(personnel_data[mgr], dict) or "type" not in personnel_data[mgr]:
                personnel_data[mgr] = {"type": "내부"}
                updated = True
        if updated:
            save_personnel_data(personnel_data)

        with st.form("personnel_form"):
            st.subheader("📋 담당자 분류 설정")
            
            h1, h2, h3 = st.columns([2, 3, 2])
            h1.markdown("**성명**")
            h2.markdown("**분류 (내부/외부)**")
            h3.markdown("**현재 상태**")
            st.write("---")

            new_classifications = {}
            for mgr in sorted(personnel_data.keys()):
                c1, c2, c3 = st.columns([2, 3, 2])
                c1.write(f"**{mgr}**")
                
                current_type = personnel_data[mgr].get("type", "내부")
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
                    personnel_data[mgr]["type"] = p_type
                save_personnel_data(personnel_data)
                st.success("인력 분류 정보가 저장되었습니다!")
                time.sleep(1)
                st.rerun()

# 2. 목표 설정
elif st.session_state.page == "targets":
    st.title("🎯 담당자별 목표 설정")
    
    # 연도 선택 필터 추가
    all_data = load_targets()
    available_years = sorted(list(all_data.get("targets", {}).keys()), reverse=True)
    if not available_years: available_years = ["2026"]
    
    col_year, col_empty = st.columns([2, 5])
    selected_year = col_year.selectbox("📅 설정 연도 선택", options=available_years + ["직접 입력"], index=0)
    if selected_year == "직접 입력":
        selected_year = col_year.text_input("연도 입력 (예: 2027)", value="2027")
    
    st.markdown(f"**{selected_year}년** 분기별 목표를 입력하세요. (단위: 만원)")
    
    targets_data = get_targets_by_year(selected_year)
    personnel_data = get_personnel_data()
    
    # 인력 명단에 있는 사람들을 targets_data에 동기화
    for mgr in personnel_data:
        if mgr not in targets_data:
            targets_data[mgr] = {f"q{i}": {"mm": 0.0, "sales": 0.0, "profit": 0.0} for i in range(1, 5)}

    all_mgrs = sorted(targets_data.keys())
    internal_mgrs = [m for m in all_mgrs if personnel_data.get(m, {}).get("type") == "내부"]
    external_mgrs = [m for m in all_mgrs if personnel_data.get(m, {}).get("type") == "외부"]
    
    if all_mgrs:
        team_total_mm = sum(float(targets_data[m][q]["mm"]) for m in all_mgrs for q in ["q1", "q2", "q3", "q4"])
        team_total_sales = sum(float(targets_data[m][q]["sales"]) for m in all_mgrs for q in ["q1", "q2", "q3", "q4"])
        team_total_profit = sum(float(targets_data[m][q]["profit"]) for m in all_mgrs for q in ["q1", "q2", "q3", "q4"])
        
        with st.expander(f"📊 {selected_year}년 우리 팀 전체 연간 목표 합계 확인", expanded=False):
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
        st.write(f"### 📋 선택된 담당자 목표 현황 요약 ({selected_year}년)")
        
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
                if st.button(f"💾 {selected_m}님 {selected_year}년 목표 저장", use_container_width=True, type="primary"):
                    targets_data[selected_m].update(updated_m_data)
                    save_targets_by_year(selected_year, targets_data)
                    st.success(f"{selected_year}년 목표가 저장되었습니다!")
                    time.sleep(1)
                    st.rerun()

    with st.sidebar.expander("📂 목표 데이터 일괄 관리", expanded=False):
        template_excel = create_excel_template(targets_data, selected_year=selected_year)
        st.download_button(f"📥 {selected_year}년 양식 다운로드", data=template_excel, file_name=f"target_template_{selected_year}.xlsx", use_container_width=True)
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
                    personnel_data = get_personnel_data()
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
                        
                        # 인력 분류 정보 업데이트
                        if mgr_name not in personnel_data:
                            personnel_data[mgr_name] = {"type": current_type}
                        else:
                            personnel_data[mgr_name]["type"] = current_type
                        
                        for _, row in mgr_rows.iterrows():
                            content = str(row['내용']).strip()
                            cat = "mm" if "MM" in content else "sales" if "매출" == content else "profit" if "매출이익" == content else None
                            
                            if cat:
                                for i in range(1, 5):
                                    val = row[f'{i}/4분기 목표']
                                    new_targets[mgr_name][f"q{i}"][cat] = float(val) if pd.notna(val) and str(val).strip() != "-" else 0.0
                    
                    # 데이터 저장
                    save_personnel_data(personnel_data)
                    save_targets_by_year(selected_year, new_targets)
                    
                    status.update(label="✅ 반영 완료!", state="complete", expanded=False)
                    st.success(f"{selected_year}년 엑셀 데이터 및 인력 분류 정보가 업데이트되었습니다.")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    status.update(label="❌ 오류 발생", state="error")
                    st.error(f"엑셀 처리 중 오류가 발생했습니다: {e}")
        
        st.write("---")
        if st.button(f"🚨 {selected_year}년 목표 데이터 초기화", use_container_width=True):
            st.session_state.show_reset_confirm = True
        
        if st.session_state.get('show_reset_confirm', False):
            st.warning(f"⚠️ {selected_year}년의 모든 목표 데이터를 삭제하시겠습니까?")
            c1, c2 = st.columns(2)
            if c1.button("✅ 예", use_container_width=True):
                save_targets_by_year(selected_year, {})
                st.session_state.show_reset_confirm = False
                st.rerun()
            if c2.button("❌ 아니오", use_container_width=True):
                st.session_state.show_reset_confirm = False
                st.rerun()

# 3. 목표 달성률 확인
elif st.session_state.page == "achievement":
    st.title("📈 구성원별 달성률 조회")
    df = load_dashboard_data()
    
    if df is None: st.warning("📊 전체 실적 대시보드에서 먼저 엑셀 파일을 업로드해 주세요."); st.stop()

    st.sidebar.header("🔍 달성률 조회 조건")
    
    # 연도 필터: Deal - 연도만 사용 (Deal - @MM (연도별)은 MM 전용)
    year_col = "Deal - 연도" if "Deal - 연도" in df.columns else next((c for c in df.columns if ('연도' in c or '년도' in c) and 'MM' not in c), None)
    if year_col:
        years = sorted(df[year_col].dropna().unique().tolist(), reverse=True)
        selected_year = st.sidebar.selectbox("📅 조회 연도 선택", options=years, index=0)
        df = df[df[year_col] == selected_year]
    else:
        # 연도 컬럼이 없는 경우 기본적으로 목표 데이터의 최신 연도 사용
        all_data = load_targets()
        available_years = sorted(list(all_data.get("targets", {}).keys()), reverse=True)
        if not available_years: available_years = ["2026"]
        selected_year = st.sidebar.selectbox("📅 조회 연도 선택", options=available_years, index=0)

    # 해당 연도의 목표 및 인력 정보 로드
    targets_data = get_targets_by_year(selected_year)
    personnel_data = get_personnel_data()
    
    if not targets_data:
        st.warning(f"🎯 {selected_year}년에 설정된 목표가 없습니다. '목표 설정하기'에서 먼저 목표를 설정해 주세요.")
        st.stop()

    period_input = st.sidebar.text_input("조회 기간 입력 (예: 1-3, 1-9)", value="1-12")
    selected_months, quarters, selected_period_label = parse_period_input(period_input)
    
    all_managers = sorted(list(set(list(targets_data.keys()) + list(personnel_data.keys()))))
    internal_managers = [m for m in all_managers if personnel_data.get(m, {}).get("type") == "내부"]
    external_managers = [m for m in all_managers if personnel_data.get(m, {}).get("type") == "외부"]
    
    selected_manager = st.sidebar.selectbox(
        "조회할 담당자 선택", 
        ["전체 담당자 한눈에 보기", "내부 인력 전체보기", "외부 인력 전체보기"] + all_managers,
        index=0
    )
    
    st.write(f"### 👤 {selected_manager}님의 {selected_year}년 {selected_period_label} 달성 현황")
    
    # 타겟 데이터 로드 시 mgr_data에 q1~q4가 있는지 확인하는 안전한 함수
    def get_safe_target(mgr, q_list, category):
        m_data = targets_data.get(mgr, {})
        return sum(float(m_data.get(q, {}).get(category, 0)) for q in q_list)

    if selected_manager == "전체 담당자 한눈에 보기":
        target_sales = sum(get_safe_target(m, quarters, "sales") for m in all_managers) * 10000
        target_profit = sum(get_safe_target(m, quarters, "profit") for m in all_managers) * 10000
        managers_to_check = all_managers
    elif selected_manager == "내부 인력 전체보기":
        target_sales = sum(get_safe_target(m, quarters, "sales") for m in internal_managers) * 10000
        target_profit = sum(get_safe_target(m, quarters, "profit") for m in internal_managers) * 10000
        managers_to_check = internal_managers
    elif selected_manager == "외부 인력 전체보기":
        target_sales = sum(get_safe_target(m, quarters, "sales") for m in external_managers) * 10000
        target_profit = sum(get_safe_target(m, quarters, "profit") for m in external_managers) * 10000
        managers_to_check = external_managers
    else:
        target_sales = get_safe_target(selected_manager, quarters, "sales") * 10000
        target_profit = get_safe_target(selected_manager, quarters, "profit") * 10000
        managers_to_check = [selected_manager]

    df = df[df['Deal - 이름'].notna() & (df['Deal - 이름'].astype(str).str.strip() != "")]
    df = df[~df['Deal - 이름'].astype(str).str.contains('합계|소계|total|sum', na=False)]
    if 'People - 이름' in df.columns: df['Deal - 이름'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"
    for col in ['Deal - 담당자_고객', 'Deal - 담당자_관리', 'Deal - 담당자_소싱']:
        if col in df.columns: df[col] = df[col].astype(str).str.strip()

    actual_sales = 0.0; actual_profit = 0.0; 
    role_configs = [(0.4, 'Deal - 담당자_고객', '고객'), (0.3, 'Deal - 담당자_관리', '관리'), (0.3, 'Deal - 담당자_소싱', '소싱')]
    
    detail_sales_records = []
    detail_profit_records = []

    for current_mgr in managers_to_check:
        for ratio, mgr_col, role_name in role_configs:
            if mgr_col in df.columns:
                matched_df = df[df[mgr_col] == current_mgr]
                for idx, row in matched_df.iterrows():
                    # 월별 데이터 추출
                    m_vals = {m: clean_currency_val(row[f"Deal - @월별매출 ({m})"]) for m in selected_months if f"Deal - @월별매출 ({m})" in row}
                    p_vals = {m: clean_currency_val(row[f"Deal - @월별이익 ({m})"]) for m in selected_months if f"Deal - @월별이익 ({m})" in row}
                    
                    # 실적 계산 로직 수정: 1-12월 전체 조회 시 무조건 '연도별' 컬럼 값을 그대로 가져옴
                    if len(selected_months) == 12 and "Deal - @매출액 (연도별)" in df.columns and "Deal - @이익 (연도별)" in df.columns:
                        row_s = clean_currency_val(row.get("Deal - @매출액 (연도별)"))
                        row_p = clean_currency_val(row.get("Deal - @이익 (연도별)"))
                    else:
                        row_s = sum(m_vals.values())
                        row_p = sum(p_vals.values())
                    
                    if row_s > 0 or row_p > 0:
                        actual_sales += row_s * ratio; actual_profit += row_p * ratio
                        
                        # 매출 상세 기록
                        if row_s > 0:
                            s_record = {"Deal명": row['Deal - 이름'], "담당자": current_mgr, "역할": role_name, "비중": f"{int(ratio*100)}%", "원매출(합계)": row_s, "반영매출": row_s * ratio}
                            for m, val in m_vals.items(): s_record[f"{int(m)}월 매출"] = val
                            detail_sales_records.append(s_record)
                        
                        # 이익 상세 기록
                        if row_p > 0:
                            p_record = {"Deal명": row['Deal - 이름'], "담당자": current_mgr, "역할": role_name, "비중": f"{int(ratio*100)}%", "원이익(합계)": row_p, "반영이익": row_p * ratio}
                            for m, val in p_vals.items(): p_record[f"{int(m)}월 이익"] = val
                            detail_profit_records.append(p_record)

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(draw_gauge(actual_sales, target_sales, "💰 매출 달성 현황", color="#EF553B", bg_color="#FCEAE8"), use_container_width=True)
    with c2: st.plotly_chart(draw_gauge(actual_profit, target_profit, "📉 이익 달성 현황", color="#636EFA", bg_color="#EBEDFE"), use_container_width=True)
    
    # 요약 데이터 시각화 개선 (HTML/CSS 사용)
    sales_ach = (actual_sales/target_sales*100) if target_sales > 0 else 0
    profit_ach = (actual_profit/target_profit*100) if target_profit > 0 else 0
    
    st.markdown(f"""
    <div style="background-color: #ffffff; padding: 20px; border-radius: 12px; border: 1px solid #e6e9ef; margin-bottom: 20px;">
        <table style="width:100%; border-collapse: collapse; text-align: center;">
            <thead>
                <tr style="border-bottom: 1px solid #dee2e6;">
                    <th style="padding: 10px; font-size: 16px; color: #666;">구분</th>
                    <th style="padding: 10px; font-size: 16px; color: #666;">목표 금액</th>
                    <th style="padding: 10px; font-size: 16px; color: #666;">실제 실적</th>
                    <th style="padding: 10px; font-size: 16px; color: #666;">달성률</th>
                </tr>
            </thead>
            <tbody>
                <tr style="border-bottom: 1px solid #eee;">
                    <td style="padding: 15px 10px; font-size: 18px; font-weight: bold;">💰 매출</td>
                    <td style="padding: 15px 10px; font-size: 17px; color: #666;">{target_sales:,.0f}원</td>
                    <td style="padding: 15px 10px; font-size: 17px; color: #666;">{actual_sales:,.0f}원</td>
                    <td style="padding: 15px 10px; font-size: 26px; font-weight: 900; color: #EF553B;">{sales_ach:.1f}%</td>
                </tr>
                <tr>
                    <td style="padding: 15px 10px; font-size: 18px; font-weight: bold;">📉 이익</td>
                    <td style="padding: 15px 10px; font-size: 17px; color: #666;">{target_profit:,.0f}원</td>
                    <td style="padding: 15px 10px; font-size: 17px; color: #666;">{actual_profit:,.0f}원</td>
                    <td style="padding: 15px 10px; font-size: 26px; font-weight: 900; color: #636EFA;">{profit_ach:.1f}%</td>
                </tr>
            </tbody>
        </table>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"📋 {selected_manager}님의 기여 내역 상세 확인", expanded=False):
        tab_s, tab_p = st.tabs(["💰 매출 상세", "📉 이익 상세"])
        
        with tab_s:
            if detail_sales_records:
                df_s = pd.DataFrame(detail_sales_records)
                fmt_s = {"원매출(합계)": "{:,.0f}원", "반영매출": "{:,.0f}원"}
                for col in df_s.columns:
                    if "월 매출" in col: fmt_s[col] = "{:,.0f}원"
                st.dataframe(df_s.style.format(fmt_s), use_container_width=True, hide_index=True)
            else:
                st.info("매출 내역이 없습니다.")
                
        with tab_p:
            if detail_profit_records:
                df_p = pd.DataFrame(detail_profit_records)
                fmt_p = {"원이익(합계)": "{:,.0f}원", "반영이익": "{:,.0f}원"}
                for col in df_p.columns:
                    if "월 이익" in col: fmt_p[col] = "{:,.0f}원"
                st.dataframe(df_p.style.format(fmt_p), use_container_width=True, hide_index=True)
            else:
                st.info("이익 내역이 없습니다.")

# 4. 월별 매출/이익 조회
elif st.session_state.page == "monthly_sales":
    st.title("📅 월별 매출/이익 조회")
    df = load_dashboard_data()
    
    if df is None:
        st.warning("📊 전체 실적 대시보드에서 먼저 엑셀 파일을 업로드해 주세요.")
        st.stop()

    st.sidebar.header("🔍 조회 조건")
    
    # 연도 필터: Deal - 연도만 사용 (Deal - @MM (연도별)은 MM 전용)
    year_col = "Deal - 연도" if "Deal - 연도" in df.columns else next((c for c in df.columns if ('연도' in c or '년도' in c) and 'MM' not in c), None)
    if year_col:
        years = sorted(df[year_col].dropna().unique().tolist(), reverse=True)
        selected_year = st.sidebar.selectbox("📅 조회 연도 선택", options=years, index=0)
        df = df[df[year_col] == selected_year]
    else:
        all_data = load_targets()
        available_years = sorted(list(all_data.get("targets", {}).keys()), reverse=True)
        if not available_years:
            available_years = ["2026"]
        selected_year = st.sidebar.selectbox("📅 조회 연도 선택", options=available_years, index=0)

    # 월별 계산: 실투입(m)/MM(연도별)*매출액(연도별), 실투입(m)/MM(연도별)*이익(연도별)
    mm_annual_col = "Deal - @MM (연도별)" if "Deal - @MM (연도별)" in df.columns else next((c for c in df.columns if "MM" in str(c) and "연도별" in str(c)), None)
    sales_annual_col = "Deal - @매출액 (연도별)" if "Deal - @매출액 (연도별)" in df.columns else None
    profit_annual_col = "Deal - @이익 (연도별)" if "Deal - @이익 (연도별)" in df.columns else None
    use_new_formula = mm_annual_col and sales_annual_col and profit_annual_col and any(f"Deal - 실투입 ({m:02d})" in df.columns for m in range(1, 13))

    if use_new_formula:
        months = list(range(1, 13))
        monthly_sales = []
        monthly_profit = []
        monthly_mm = []
        mm_vals = df[mm_annual_col].apply(clean_currency_val)
        sales_vals = df[sales_annual_col].apply(clean_currency_val)
        profit_vals = df[profit_annual_col].apply(clean_currency_val)

        for m in months:
            si_col = f"Deal - 실투입 ({m:02d})"
            si_vals = df[si_col].apply(clean_currency_val) if si_col in df.columns else pd.Series(0.0, index=df.index)
            monthly_mm.append(si_vals.sum())
            # m월 매출 = Σ (실투입(m)/MM(연도별)*매출액(연도별)), MM=0이면 0
            divisor = mm_vals.replace(0, float('nan'))
            ratio = (si_vals / divisor).fillna(0)
            s_val = (ratio * sales_vals).sum()
            p_val = (ratio * profit_vals).sum()
            monthly_sales.append(s_val)
            monthly_profit.append(p_val)

        total_sales = sum(monthly_sales)
        total_profit = sum(monthly_profit)
        total_mm_annual = sum(monthly_mm)
        if total_mm_annual == 0 and mm_annual_col:
            monthly_mm = [mm_vals.sum() * (s / total_sales) if total_sales else (mm_vals.sum() / 12) for s in monthly_sales]
            total_mm_annual = sum(monthly_mm)
        monthly_rate = [(p / s * 100) if s else 0 for s, p in zip(monthly_sales, monthly_profit)]
    else:
        sales_cols = [c for c in df.columns if "월별매출" in c and "(" in c]
        profit_cols = [c for c in df.columns if "월별이익" in c and "(" in c]
        if not sales_cols and not profit_cols:
            st.info("엑셀 파일에 **Deal - @MM (연도별)**, **Deal - @매출액 (연도별)**, **Deal - @이익 (연도별)**, **Deal - 실투입 (01)~(12)** 컬럼 또는 월별 매출/이익 컬럼이 필요합니다.")
            st.stop()
        months = list(range(1, 13))
        monthly_sales = []
        monthly_profit = []
        monthly_mm = []
        for m in months:
            s_col = f"Deal - @월별매출 ({m:02d})"
            p_col = f"Deal - @월별이익 ({m:02d})"
            mm_col = f"Deal - 실투입 ({m:02d})"
            s_val = df[s_col].apply(clean_currency_val).sum() if s_col in df.columns else 0
            p_val = df[p_col].apply(clean_currency_val).sum() if p_col in df.columns else 0
            mm_val = df[mm_col].apply(clean_currency_val).sum() if mm_col in df.columns else 0
            monthly_sales.append(s_val)
            monthly_profit.append(p_val)
            monthly_mm.append(mm_val)
        total_sales = sum(monthly_sales)
        total_profit = sum(monthly_profit)
        total_mm_annual = sum(monthly_mm)
        if total_mm_annual == 0 and mm_annual_col:
            annual_mm = df[mm_annual_col].apply(clean_currency_val).sum() if mm_annual_col else 0
            monthly_mm = [annual_mm * (s / total_sales) if total_sales else (annual_mm / 12) for s in monthly_sales]
            total_mm_annual = sum(monthly_mm)
        monthly_rate = [(p / s * 100) if s else 0 for s, p in zip(monthly_sales, monthly_profit)]

    summary_df = pd.DataFrame({
        "월": [f"{m}월" for m in months],
        "MM": monthly_mm,
        "매출액": monthly_sales,
        "이익액": monthly_profit,
        "이익률": monthly_rate
    })

    st.subheader(f"📊 {selected_year}년 월별 매출/이익 현황")
    has_mm = total_mm_annual > 0
    per_capita_sales = total_sales / total_mm_annual if total_mm_annual > 0 else 0
    per_capita_profit = total_profit / total_mm_annual if total_mm_annual > 0 else 0
    if has_mm:
        per_capita_rate = (per_capita_profit / per_capita_sales * 100) if per_capita_sales else 0
        avg_part = f"""<br><span style="color: #E53935;">인당 매출: {per_capita_sales:,.0f}원</span> | 
        <span style="color: #1E88E5;">인당 이익: {per_capita_profit:,.0f}원</span> | 
        <span style="color: #43A047;">이익률: {per_capita_rate:.1f}%</span>"""
    else:
        avg_part = ""
        if not has_mm:
            with st.expander("🔍 인당 매출/이익이 안 나온다면? (로드된 엑셀 컬럼 확인)", expanded=False):
                st.write("인당 매출/이익을 위해 **Deal - 실투입 (01)~(12)**(월별 MM) 또는 **Deal - @MM (연도별)** 컬럼이 필요합니다.")
                st.write("현재 로드된 컬럼:", list(df.columns))
    profit_rate = (total_profit / total_sales * 100) if total_sales else 0
    st.markdown(f"""
    <div style="font-size: 2.0em; font-weight: bold; margin-bottom: 16px;">
        <span style="color: #E53935;">전체 매출: {total_sales:,.0f}원</span> | 
        <span style="color: #1E88E5;">전체 이익: {total_profit:,.0f}원</span> | 
        <span style="color: #43A047;">이익률: {profit_rate:.1f}%</span>{avg_part}
    </div>
    """, unsafe_allow_html=True)

    month_labels = [f"{m}월" for m in range(1, 13)]
    fig = go.Figure()
    fig.add_trace(go.Bar(name="매출", x=month_labels, y=summary_df["매출액"], marker_color="#E53935", marker_line_width=0.5, marker_line_color="#B71C1C"))
    fig.add_trace(go.Bar(name="이익", x=month_labels, y=summary_df["이익액"], marker_color="#1E88E5", marker_line_width=0.5, marker_line_color="#0D47A1"))
    fig.update_layout(
        barmode="group",
        title=f"{selected_year}년 월별 매출/이익",
        xaxis={"title": "", "tickmode": "array", "tickvals": month_labels, "tickangle": -45},
        yaxis={"title": "금액(원)", "tickformat": ",d", "rangemode": "tozero"},
        height=500,
        margin=dict(l=80, r=40, t=60, b=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(summary_df.style.format({
        "MM": "{:,.2f}",
        "매출액": "{:,.0f}",
        "이익액": "{:,.0f}",
        "이익률": "{:.1f}%"
    }), use_container_width=True, hide_index=True)

# 5. 고객사/엔드클라이언트 순위 조회 공통 로직
elif st.session_state.page in ["rank_customer", "rank_endclient"]:
    is_customer = st.session_state.page == "rank_customer"
    title = "🏢 고객사별 매출/이익 순위 조회" if is_customer else "👤 엔드클라이언트 매출/이익 순위 조회"
    st.title(title)
    
    df = load_dashboard_data()
    if df is None:
        st.warning("📊 전체 실적 대시보드에서 먼저 엑셀 파일을 업로드해 주세요.")
        st.stop()

    st.sidebar.header("🔍 순위 조회 조건")
    
    # 연도 필터: Deal - 연도만 사용 (Deal - @MM (연도별)은 MM 전용)
    year_col = "Deal - 연도" if "Deal - 연도" in df.columns else next((c for c in df.columns if ('연도' in c or '년도' in c) and 'MM' not in c), None)
    if year_col:
        all_years = sorted(df[year_col].dropna().unique().tolist(), reverse=True)
        selected_year = st.sidebar.selectbox("조회 연도 선택", all_years, key=f"rank_year_{st.session_state.page}")
        df = df[df[year_col] == selected_year]
    
    criteria = st.sidebar.radio("분석 기준 선택", ["매출 기준 순위", "이익 기준 순위"])
    period_input = st.sidebar.text_input("조회 기간 입력 (예: 1-12)", value="1-12", key=f"rank_period_{st.session_state.page}")
    selected_months, _, selected_period_label = parse_period_input(period_input)

    # 분석 대상 컬럼 설정
    target_col = "Deal - 고객사" if is_customer else "Deal - 엔드 클라이언트"
    if target_col not in df.columns:
        st.error(f"엑셀 파일에 '{target_col}' 컬럼이 없습니다.")
        st.stop()

    # 실적 데이터 정제 - 매출/이익 둘 다 계산
    sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
    profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
    for col in sales_cols + profit_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_currency_val)
    
    # 연도별 컬럼 우선 사용 (12월 전체 조회 시)
    if len(selected_months) == 12 and "Deal - @매출액 (연도별)" in df.columns and "Deal - @이익 (연도별)" in df.columns:
        df['_총매출'] = df['Deal - @매출액 (연도별)'].apply(clean_currency_val)
        df['_총이익'] = df['Deal - @이익 (연도별)'].apply(clean_currency_val)
    else:
        df['_총매출'] = df[[c for c in sales_cols if c in df.columns]].sum(axis=1)
        df['_총이익'] = df[[c for c in profit_cols if c in df.columns]].sum(axis=1)
    
    # MM 컬럼
    mm_col = "Deal - @MM (연도별)" if "Deal - @MM (연도별)" in df.columns else next((c for c in df.columns if "MM" in str(c) and "연도별" in str(c)), None)
    if mm_col:
        df['_mm'] = df[mm_col].apply(clean_currency_val)
    else:
        df['_mm'] = 0.0

    is_sales = "매출" in criteria

    # 그룹화: 항목별 총MM, 총매출, 총이익
    agg_dict = {'_총매출': 'sum', '_총이익': 'sum'}
    if mm_col:
        agg_dict['_mm'] = 'sum'
    rank_df = df.groupby(target_col).agg(agg_dict).reset_index()
    rank_df = rank_df[(rank_df['_총매출'] != 0) | (rank_df['_총이익'] != 0)]
    rank_df['이익률'] = (rank_df['_총이익'] / rank_df['_총매출'] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
    sort_col = '_총매출' if is_sales else '_총이익'
    rank_df = rank_df.sort_values(sort_col, ascending=False).reset_index(drop=True)
    rank_df.index = rank_df.index + 1  # 순위 1, 2, 3...

    if rank_df.empty:
        st.info("조회된 데이터가 없습니다.")
    else:
        label = "매출액" if is_sales else "이익액"
        plot_col = '_총매출' if is_sales else '_총이익'
        color = "#EF553B" if is_sales else "#636EFA"
        st.subheader(f"🏆 {criteria} (Top 5 + 기타)")
        
        # 상위 5위와 나머지(기타) 데이터 처리
        if len(rank_df) > 5:
            top_5 = rank_df.head(5).copy()
            others_sum = rank_df.iloc[5:][plot_col].sum()
            others_row = pd.DataFrame({target_col: ['기타'], plot_col: [others_sum]})
            plot_df = pd.concat([top_5, others_row], ignore_index=True)
        else:
            plot_df = rank_df.copy()
        
        # 세로 막대 그래프 시각화 (Plotly)
        fig = px.bar(
            plot_df, 
            x=target_col, 
            y=plot_col,
            labels={target_col: target_col.replace("Deal - ", ""), plot_col: label},
            color_discrete_sequence=[color]
        )
        
        fig.update_traces(
            hovertemplate="<b>%{y:,.0f}원</b><extra></extra>",
            hoverlabel=dict(bgcolor="white", bordercolor="white", font=dict(size=22, family="Malgun Gothic", color="black"))
        )
        fig.update_layout(
            xaxis_title="", yaxis_title=f"{label} (원)", height=500, hovermode='x',
            xaxis=dict(tickfont=dict(size=16, family="Malgun Gothic", color="black")),
            yaxis=dict(tickfont=dict(size=14), tickformat=",d", range=[0, plot_df[plot_col].max() * 1.2]),
            margin=dict(t=50, b=50)
        )
        st.plotly_chart(fig, use_container_width=True)

        st.write("---")
        st.subheader(f"📋 {criteria} 전체 내역")
        display_rank_df = rank_df[[target_col] + (['_mm'] if mm_col else []) + ['_총매출', '_총이익', '이익률']].copy()
        display_rank_df.columns = ['항목'] + (['총MM'] if mm_col else []) + ['총매출액', '총이익액', '이익률']
        fmt_dict = {'총매출액': '{:,.0f}원', '총이익액': '{:,.0f}원', '이익률': '{:.1f}%'}
        if mm_col:
            fmt_dict['총MM'] = '{:,.2f}'
        st.dataframe(display_rank_df.style.format(fmt_dict), use_container_width=True)

# 5. 전체 실적 대시보드
else:
    st.title("📊 Deal-ito 통합 실적/이익 대시보드")
    st.markdown("좌측 사이드바에서 엑셀 파일을 업로드하면 실적을 자동 계산합니다.")
    df = load_dashboard_data()
    if df is not None:
        st.sidebar.header("🔍 조회 조건 설정")
        
        # 연도 필터: Deal - 연도만 사용 (Deal - @MM (연도별)은 MM 전용)
        year_col = "Deal - 연도" if "Deal - 연도" in df.columns else next((c for c in df.columns if ('연도' in c or '년도' in c) and 'MM' not in c), None)
        if year_col:
            all_years = sorted(df[year_col].dropna().unique().tolist(), reverse=True)
            selected_year = st.sidebar.selectbox("조회 연도 선택", all_years, key="dash_year_select")
            df = df[df[year_col] == selected_year]
            
        period_input = st.sidebar.text_input("조회 기간 입력 (예: 1-3, 1-9)", value="1-12", key="dash_period_input")
        selected_months, _, selected_period_label = parse_period_input(period_input)
        
        all_mgrs = sorted(list(set(df['Deal - 담당자_고객'].dropna().unique().tolist() + df['Deal - 담당자_관리'].dropna().unique().tolist() + df['Deal - 담당자_소싱'].dropna().unique().tolist())))
        
        # 담당자 선택을 드롭다운 형태로 변경
        targets_data = load_targets()
        internal_managers = [m for m in all_mgrs if targets_data.get(m, {}).get("type") == "내부"]
        external_managers = [m for m in all_mgrs if targets_data.get(m, {}).get("type") == "외부"]
        
        selected_option = st.sidebar.selectbox(
            "조회할 담당자 선택", 
            ["전체 담당자 한눈에 보기", "내부 인력 전체보기", "외부 인력 전체보기"] + all_mgrs,
            key="dash_mgr_select"
        )
        
        if selected_option == "전체 담당자 한눈에 보기":
            selected_managers = all_mgrs
        elif selected_option == "내부 인력 전체보기":
            selected_managers = internal_managers
        elif selected_option == "외부 인력 전체보기":
            selected_managers = external_managers
        else:
            selected_managers = [selected_option]

        df = df[~df['Deal - 이름'].astype(str).str.contains('합계|소계|total|sum', na=False)]
        if 'People - 이름' in df.columns: df['Deal - 이름'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"
        
        sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]
        profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
        for col in sales_cols + profit_cols:
            if col in df.columns: df[col] = df[col].apply(clean_currency_val)
        
        # 실적 계산 로직 수정: 1-12월 전체 조회 시 무조건 '연도별' 컬럼 값을 그대로 사용
        if len(selected_months) == 12 and "Deal - @매출액 (연도별)" in df.columns and "Deal - @이익 (연도별)" in df.columns:
            df['선택기간_총매출'] = df['Deal - @매출액 (연도별)'].apply(clean_currency_val)
            df['선택기간_총이익'] = df['Deal - @이익 (연도별)'].apply(clean_currency_val)
        else:
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
            # 필터링 로직 수정: 선택된 담당자 리스트가 있으면 해당 명단만, 없으면 빈 데이터프레임 반환
            return combined[combined['담당자'].isin(selected_managers)]

        mm_col = "Deal - @MM (연도별)" if "Deal - @MM (연도별)" in df.columns else next((c for c in df.columns if "MM" in str(c) and "연도별" in str(c)), None)
        st.markdown("""
            <style>
            [data-testid="stTabs"] button p, .stTabs [data-baseweb="tab-list"] button p { font-size: 1.2rem !important; font-weight: 700 !important; }
            </style>
        """, unsafe_allow_html=True)
        tab1, tab2, tab3 = st.tabs(["📊 전체 분석", "💰 매출 분석", "📉 이익 분석"])
        for tab, col_name, label in [(tab2, '선택기간_총매출', "매출"), (tab3, '선택기간_총이익', "이익")]:
            with tab:
                res_df = calc_consolidated(col_name)
                sum_df = res_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values('반영실적', ascending=False)
                
                st.markdown(f"<h4 style='text-align: center; margin-bottom: 20px;'>👤 담당자별 합산 {label} 요약</h4>", unsafe_allow_html=True)
                
                if not sum_df.empty:
                    total_val = sum_df['반영실적'].sum()
                    color = "#EF553B" if label == "매출" else "#636EFA"
                    
                    rows_html = ""
                    for _, row in sum_df.iterrows():
                        rows_html += f"<tr style='border-bottom: 1px solid #eee;'><td style='padding: 12px; text-align: center; font-weight: bold;'>{row['담당자']}</td><td style='padding: 12px; text-align: center; color: {color}; font-weight: 600;'>{row['반영실적']:,.0f}원</td></tr>"
                    
                    total_html = f"<tr style='background-color: #f8f9fb; font-weight: 900;'><td style='padding: 15px; text-align: center; border-top: 2px solid #ddd;'>★ 전체 합계</td><td style='padding: 15px; text-align: center; color: {color}; border-top: 2px solid #ddd; font-size: 18px;'>{total_val:,.0f}원</td></tr>"
                    
                    table_html = f"""
                    <div style="border: 1px solid #e6e9ef; border-radius: 10px; overflow: hidden; margin: 0 auto; background-color: white; width: 100%;">
                        <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                            <thead>
                                <tr style="background-color: #f1f3f6; border-bottom: 2px solid #dee2e6;">
                                    <th style="padding: 12px; text-align: center; color: #555; width: 50%;">담당자</th>
                                    <th style="padding: 12px; text-align: center; color: #555; width: 50%;">반영 실적 ({label})</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows_html}
                                {total_html}
                            </tbody>
                        </table>
                    </div>
                    """
                    st.markdown(table_html, unsafe_allow_html=True)
                    st.write("") 
                else:
                    st.info("조회된 실적 데이터가 없습니다.")
                with st.expander(f"📋 {label} 상세 기여 내역 확인", expanded=False):
                    disp_df = res_df[res_df['반영실적'] > 0].groupby(['Deal명', '담당자']).agg({'원금액': 'first', '역할': lambda x: ', '.join(x), '비중_num': 'sum', '반영실적': 'sum'}).reset_index()
                    disp_df['반영비율'] = disp_df['비중_num'].apply(lambda x: f"{int(x*100)}%")
                    st.dataframe(disp_df[['Deal명', '담당자', '역할', '원금액', '반영비율', '반영실적']].style.format({'원금액': '{:,.0f}원', '반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)

        with tab1:
            filtered = df[df['Deal - 담당자_고객'].isin(selected_managers) | df['Deal - 담당자_관리'].isin(selected_managers) | df['Deal - 담당자_소싱'].isin(selected_managers)]
            deal_id_col = "Deal - RecordId" if "Deal - RecordId" in df.columns else "Deal - 이름"
            if deal_id_col == "Deal - 이름":
                filtered = filtered.copy()
                filtered["_deal_key"] = filtered["Deal - 이름"].astype(str).str.replace(r"\s*\([^)]*\)$", "", regex=True)
                deal_id_col = "_deal_key"
            deal_count = filtered[deal_id_col].nunique()
            agg_dict = {'선택기간_총매출': 'first', '선택기간_총이익': 'first'}
            if mm_col:
                agg_dict[mm_col] = 'first'
            agg_df = filtered.groupby(deal_id_col).agg(agg_dict).reset_index()

            sales_total = agg_df['선택기간_총매출'].sum()
            sales_avg = agg_df['선택기간_총매출'].mean() if deal_count > 0 else 0
            profit_total = agg_df['선택기간_총이익'].sum()
            profit_avg = agg_df['선택기간_총이익'].mean() if deal_count > 0 else 0

            st.markdown("**1. MM 분석**")
            if mm_col:
                mm_total = agg_df[mm_col].apply(clean_currency_val).sum()
                mm_avg = agg_df[mm_col].apply(clean_currency_val).mean() if deal_count > 0 else 0
                c1, c2 = st.columns(2)
                with c1: st.metric("MM 전체값", f"{mm_total:,.2f}")
                with c2: st.metric("인당 MM평균값", f"{mm_avg:,.2f}")
            else:
                st.caption("'Deal - @MM (연도별)' 컬럼이 없습니다.")

            st.markdown("**2. 매출 분석**")
            c1, c2 = st.columns(2)
            with c1: st.metric("매출 전체값", f"{sales_total:,.0f}원")
            with c2: st.metric("인당 매출 평균값", f"{sales_avg:,.0f}원")

            st.markdown("**3. 이익 분석**")
            c1, c2 = st.columns(2)
            with c1: st.metric("이익 전체값", f"{profit_total:,.0f}원")
            with c2: st.metric("인당 이익 평균값", f"{profit_avg:,.0f}원")

            profit_rate = (profit_total / sales_total * 100) if sales_total > 0 else 0
            st.markdown("**4. 이익률**")
            st.metric("이익률 (이익 전체값 ÷ 매출 전체값 × 100)", f"{profit_rate:,.1f}%")
    else: st.info("좌측 사이드바에서 실적 엑셀 파일을 업로드해 주세요.")
