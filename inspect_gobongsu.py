import pandas as pd
import pickle
import re

def clean_currency_val(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    cleaned = re.sub(r'[^0-9.-]', '', str(val))
    try: return float(cleaned)
    except: return 0.0

# 데이터 로드
try:
    with open("dashboard_cache.pkl", "rb") as f:
        df = pickle.load(f)
except:
    print("데이터 파일을 찾을 수 없습니다.")
    exit()

target_mgr = "고봉수"
print(f"--- [{target_mgr}] 실적 상세 분석 (전체 기간 기준) ---")

# 전처리
if 'People - 이름' in df.columns:
    df['Deal - 이름'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"

role_configs = [(0.4, 'Deal - 담당자_고객', '고객'), (0.3, 'Deal - 담당자_관리', '관리'), (0.3, 'Deal - 담당자_소싱', '소싱')]
total_actual_sales = 0.0
total_actual_profit = 0.0

project_details = []

for ratio, mgr_col, role_name in role_configs:
    if mgr_col in df.columns:
        df[mgr_col] = df[mgr_col].astype(str).str.strip()
        matched = df[df[mgr_col] == target_mgr]
        if not matched.empty:
            for idx, row in matched.iterrows():
                # 연도별 컬럼 우선 사용
                row_s = clean_currency_val(row.get("Deal - @매출액 (연도별)", 0))
                row_p = clean_currency_val(row.get("Deal - @이익 (연도별)", 0))
                
                # 연도별 컬럼이 0인 경우 월별 합산 시도
                if row_s == 0:
                    m_cols = [f"Deal - @월별매출 ({m:02d})" for m in range(1, 13)]
                    row_s = sum(clean_currency_val(row.get(c, 0)) for c in m_cols)
                if row_p == 0:
                    p_cols = [f"Deal - @월별이익 ({m:02d})" for m in range(1, 13)]
                    row_p = sum(clean_currency_val(row.get(c, 0)) for c in p_cols)
                
                reflected_s = row_s * ratio
                reflected_p = row_p * ratio
                
                total_actual_sales += reflected_s
                total_actual_profit += reflected_p
                
                project_details.append({
                    "Deal명": row['Deal - 이름'],
                    "역할": role_name,
                    "비중": f"{int(ratio*100)}%",
                    "원매출": row_s,
                    "반영매출": reflected_s,
                    "원이익": row_p,
                    "반영이익": reflected_p
                })

print(f"\n[최종 합계 결과]")
print(f"총 반영 매출: {total_actual_sales:,.0f}원")
print(f"총 반영 이익: {total_actual_profit:,.0f}원")

print("\n[상세 내역 (상위 10건)]")
details_df = pd.DataFrame(project_details)
if not details_df.empty:
    print(details_df.sort_values("반영매출", ascending=False).head(10).to_string(index=False))
