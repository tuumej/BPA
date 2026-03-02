import os
import pandas as pd
import numpy as np

# 🔹 엑셀 파일 경로 설정
folder_path = "./"  # 👉 실제 경로 수정
output_file = "merged_all_cleaned.xlsx"

merged_df = pd.DataFrame()

# 🔹 폴더 내 모든 .xlsx 파일 읽기
for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path, file)
        print(f"📄 읽는 중: {file_path}")

        # ✅ 3번째 행을 컬럼명으로 인식
        df = pd.read_excel(file_path, header=2)

        # 🔹 전체 44개 컬럼까지만 사용
        df = df.iloc[:, :44]

        # 🔹 컬럼명 공백 제거
        df.columns = df.columns.map(lambda x: str(x).strip())

        # 🔹 빈 행 제거
        df = df.dropna(how="all")

        # 🔹 전체 병합
        merged_df = pd.concat([merged_df, df], ignore_index=True)

# ✅ 이름, 이메일 컬럼 자동 탐지
name_col_candidates = [col for col in merged_df.columns if "이름" in col]
email_col_candidates = [col for col in merged_df.columns if "메일" in col or "email" in col.lower()]

if not name_col_candidates or not email_col_candidates:
    raise ValueError("⚠️ 이름 또는 이메일 컬럼을 3번째 행에서 찾을 수 없습니다. 컬럼명을 확인해주세요.")

name_col = name_col_candidates[0]
email_col = email_col_candidates[0]

print(f"✅ 인식된 컬럼명: 이름 = '{name_col}', 이메일 = '{email_col}'")

# ✅ 병합 규칙 함수 (문자열 > '-' > 공백)
def merge_values(values):
    cleaned = [str(v).strip() for v in values if pd.notna(v)]
    if not cleaned:
        return ""
    meaningful = [v for v in cleaned if v not in ["", "-", "nan", "None"]]
    if meaningful:
        unique = list(dict.fromkeys(meaningful))  # 중복 제거
        return ", ".join(unique)
    else:
        if "-" in cleaned:
            return "-"
        return ""

# ✅ 이름 + 이메일 기준으로 그룹화하고 나머지 컬럼 병합
group_cols = [name_col, email_col]
merged_final = (
    merged_df.groupby(group_cols, dropna=False)
    .agg(lambda x: merge_values(x))
    .reset_index()
)

# ✅ 결과 파일 저장
output_path = os.path.join(folder_path, output_file)
merged_final.to_excel(output_path, index=False)

print(f"\n🎉 모든 파일 병합 및 중복 정리 완료!\n결과 파일: {output_path}")
