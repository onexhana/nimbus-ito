import pandas as pd
import pickle

try:
    with open("dashboard_cache.pkl", "rb") as f:
        df = pickle.load(f)
    print("현재 데이터의 컬럼 목록:")
    for col in df.columns:
        print(f"- {col}")
except Exception as e:
    print(f"오류 발생: {e}")
