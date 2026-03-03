import pandas as pd
from datetime import datetime, timezone, timedelta

# ---------------------------------------------
# 🔧 사용자 설정
# ---------------------------------------------
csv_path = r""        # CSV 파일 경로
output_path = r""   # 결과 XLSX 저장 경로
# ---------------------------------------------

# 1) CSV 파일 읽기
df = pd.read_csv(csv_path)

# 2) CSV → XLSX 변환 (중간 저장)
intermediate_xlsx = csv_path.replace(".csv", ".xlsx")
df.to_excel(intermediate_xlsx, index=False)

# 3) type 컬럼이 11인 행만 추출
df = df[df["type"] == 11]

# 4) 필요한 컬럼만 선택
df = df[["timestamp", "user_id", "result", "description"]]

# 5) 🟦 Unix timestamp → 한국 시간(KST) 변환 함수
def unix_to_kst(ts):
    try:
        ts = int(ts)
        dt_utc = datetime.fromtimestamp(ts, tz=timezone.utc)
        kst = dt_utc + timedelta(hours=9)
        return kst.strftime("%Y-%m-%d %H:%M")
    except:
        return ts  # 변환 실패 시 원본 값 유지

df["timestamp"] = df["timestamp"].apply(unix_to_kst)

# 6) result 값 변환 (1 → 성공, 2 → 실패)
df["result"] = df["result"].map({1: "성공", 2: "실패"})

# 7) 컬럼명 변경
df = df.rename(columns={
    "timestamp": "일시",
    "user_id": "사용자 아이디",
    "result": "결과",
    "description": "활동내역"
})

# 8) 최종 결과 엑셀 저장
df.to_excel(output_path, index=False)

print("완료되었습니다 →", output_path)
