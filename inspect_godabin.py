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

target_mgr = "고다빈"
selected_months = [f"{m:02d}" for m in range(1, 13)]

print(f"--- [{target_mgr}] 실적 상세 분석 (이익 기준) ---")

role_configs = [(0.4, 'Deal - 담당자_고객', '고객'), (0.3, 'Deal - 담당자_관리', '관리'), (0.3, 'Deal - 담당자_소싱', '소싱')]
total_actual_profit = 0.0

# 전처리 (Deal 이름에 이름 추가 등 app.py와 동일하게)
if 'People - 이름' in df.columns:
    df['Deal - 이름'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"

for ratio, mgr_col, role_name in role_configs:
    if mgr_col in df.columns:
        df[mgr_col] = df[mgr_col].astype(str).str.strip()
        matched = df[df[mgr_col] == target_mgr]
        if not matched.empty:
            print(f"\n[역할: {role_name} (지분 {int(ratio*100)}%)]")
            for idx, row in matched.iterrows():
                # 연도별 컬럼 우선 사용 (app.py 최신 로직 반영)
                row_p = clean_currency_val(row.get("Deal - @이익 (연도별)", 0))
                reflected_p = row_p * ratio
                if row_p != 0:
                    print(f"- {row['Deal - 이름']}: 원이익 {row_p:,.0f}원 -> 반영이익 {reflected_p:,.1f}원")
                    total_actual_profit += reflected_p

print("\n" + "="*40)
print(f"최종 합계 계산 결과: {total_actual_profit:,.1f}원")
print("="*40)
