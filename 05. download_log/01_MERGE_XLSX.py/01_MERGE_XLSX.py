import os
import pandas as pd

# 1️⃣ 엑셀 파일들이 저장된 폴더 경로 설정
folder_path = "./"  # 👉 여기를 실제 경로로 바꿔주세요
output_file = "merged_columns.xlsx"   # 결과 파일명

# 2️⃣ 결과를 담을 빈 DataFrame 생성
merged_df = pd.DataFrame()

# 3️⃣ 폴더 내의 모든 .xlsx 파일 순회
for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path, file)
        print(f"읽는 중: {file_path}")

        # 4️⃣ 엑셀 파일의 A,B,C 컬럼만 읽기
        df = pd.read_excel(file_path
        , usecols=["NO","승인번호","이름","이메일 주소","내부망 PC IP","내부망 PC MAC","외부망 PC IP","외부망 PC MAC"])

        # 5️⃣ 읽은 데이터를 merged_df에 추가
        merged_df = pd.concat([merged_df, df], ignore_index=True)

# 6️⃣ 합쳐진 데이터를 새로운 엑셀 파일로 저장
output_path = os.path.join(folder_path, output_file)
merged_df.to_excel(output_path, index=False)

print(f"모든 파일의 A,B,C 컬럼을 합친 파일이 생성되었습니다: {output_path}")
