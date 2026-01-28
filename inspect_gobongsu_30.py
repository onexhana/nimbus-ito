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
if 'People - 이름' in df.columns:
    df['Deal - 이름_unique'] = df['Deal - 이름'].astype(str) + " (" + df['People - 이름'].fillna("미지정").astype(str) + ")"
else:
    df['Deal - 이름_unique'] = df['Deal - 이름'].astype(str)

results = []
for idx, row in df.iterrows():
    roles = []
    ratio_sum = 0.0
    if str(row.get('Deal - 담당자_고객', '')).strip() == target_mgr:
        roles.append('고객')
        ratio_sum += 0.4
    if str(row.get('Deal - 담당자_관리', '')).strip() == target_mgr:
        roles.append('관리')
        ratio_sum += 0.3
    if str(row.get('Deal - 담당자_소싱', '')).strip() == target_mgr:
        roles.append('소싱')
        ratio_sum += 0.3
    
    if ratio_sum == 0.3:
        row_s = clean_currency_val(row.get("Deal - @매출액 (연도별)", 0))
        row_p = clean_currency_val(row.get("Deal - @이익 (연도별)", 0))
        if row_s == 0:
            m_cols = [f"Deal - @월별매출 ({m:02d})" for m in range(1, 13)]
            row_s = sum(clean_currency_val(row.get(c, 0)) for c in m_cols)
        if row_p == 0:
            p_cols = [f"Deal - @월별이익 ({m:02d})" for m in range(1, 13)]
            row_p = sum(clean_currency_val(row.get(c, 0)) for c in p_cols)
            
        results.append({
            "Deal명": row['Deal - 이름_unique'],
            "역할": ", ".join(roles),
            "원매출": row_s,
            "반영매출": row_s * 0.3,
            "원이익": row_p,
            "반영이익": row_p * 0.3
        })

print(f"--- [{target_mgr}] 30% 비중 그룹 상세 내역 ---")
print(pd.DataFrame(results).to_string(index=False))
