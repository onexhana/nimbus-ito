import streamlit as st
import pandas as pd
import plotly.express as px
import json
import os
import io

# 파일 경로 설정
TARGETS_FILE = "targets.json"

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

st.sidebar.write("---")

# 1. 실적 데이터 업로드 (대시보드 페이지용)
if st.session_state.page == "dashboard":
    with st.sidebar.expander("📁 실적 데이터 업로드", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"], key="dashboard_uploader")
        
        if uploaded_file is not None:
            try:
                # 파일을 읽어서 세션 스테이트에 저장
                df_loaded = pd.read_excel(uploaded_file)
                st.session_state.dashboard_df = df_loaded
                st.success("데이터가 로드되었습니다!")
            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        
        if 'dashboard_df' in st.session_state:
            if st.button("🗑️ 업로드된 데이터 삭제", use_container_width=True):
                del st.session_state.dashboard_df
                if 'dashboard_uploader' in st.session_state:
                    # uploader 위젯 초기화 (고급 기법: 내부 키 제거는 안되므로 멘트 유도)
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
            st.dataframe(df_summary.style.format({"1/4분기": lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) and x != 0 else f"{x:,.0f}" if isinstance(x, (int, float)) else x, "2/4분기": lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) and x != 0 else f"{x:,.0f}" if isinstance(x, (int, float)) else x, "3/4분기": lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) and x != 0 else f"{x:,.0f}" if isinstance(x, (int, float)) else x, "4/4분기": lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) and x != 0 else f"{x:,.0f}" if isinstance(x, (int, float)) else x, "년 합계": lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) and x != 0 else f"{x:,.0f}" if isinstance(x, (int, float)) else x}), use_container_width=True, hide_index=True)
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

# 메인 화면 - 실적 대시보드
else:
    st.title("📊 Deal-ito 통합 실적/이익 대시보드")
    st.markdown("좌측 사이드바에서 엑셀 파일을 업로드하면 실적을 자동 계산합니다.")
    if 'dashboard_df' in st.session_state:
        df = st.session_state.dashboard_df
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
        sales_cols = [f"Deal - @월별매출 ({m})" for m in selected_months]; profit_cols = [f"Deal - @월별이익 ({m})" for m in selected_months]
        def clean_currency(column):
            if column in df.columns:
                s = df[column].astype(str).str.replace(r'[^0-9.-]', '', regex=True)
                return pd.to_numeric(s, errors='coerce').fillna(0)
            return 0
        for col in sales_cols + profit_cols: df[col] = clean_currency(col)
        df['선택기간_총매출'] = df[sales_cols].sum(axis=1); df['선택기간_총이익'] = df[profit_cols].sum(axis=1)
        def calculate_consolidated_results(target_col):
            role_configs = [('고객(40%)', 0.4, 'Deal - 담당자_고객'), ('관리(30%)', 0.3, 'Deal - 담당자_관리'), ('소싱(30%)', 0.3, 'Deal - 담당자_소싱')]
            individual_results = []
            for role_label, ratio, manager_col in role_configs:
                temp = df[['Deal - 이름', manager_col, target_col]].copy()
                temp['반영실적'] = temp[target_col] * ratio; temp['역할'] = role_label; temp.columns = ['Deal명', '담당자', '원금액', '반영실적', '역할']
                individual_results.append(temp)
            combined = pd.concat(individual_results)
            if selected_managers: combined = combined[combined['담당자'].isin(selected_managers)]
            return combined
        tab1, tab2 = st.tabs(["💰 매출 분석", "📉 이익 분석"])
        with tab1:
            st.subheader(f"📅 매출 조회 기간: {selected_period_label}")
            m_df = calculate_consolidated_results('선택기간_총매출'); summary_m = m_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
            st.write("#### 👤 담당자별 합산 실적"); st.dataframe(summary_m.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
            st.write("---"); st.write("#### 📋 매출 상세 기여 내역"); st.dataframe(m_df[m_df['반영실적'] > 0].style.format({'원금액': '{:,.0f}원', '반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
        with tab2:
            st.subheader(f"📅 이익 조회 기간: {selected_period_label}")
            p_df = calculate_consolidated_results('선택기간_총이익'); summary_p = p_df.groupby('담당자')['반영실적'].sum().reset_index().sort_values(by='반영실적', ascending=False)
            st.write("#### 👤 담당자별 합산 실적"); st.dataframe(summary_p.style.format({'반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
            st.write("---"); st.write("#### 📋 이익 상세 기여 내역"); st.dataframe(p_df[p_df['반영실적'] > 0].style.format({'원금액': '{:,.0f}원', '반영실적': '{:,.0f}원'}), use_container_width=True, hide_index=True)
    else:
        st.info("좌측 사이드바에서 실적 엑셀 파일을 업로드해 주세요.")
