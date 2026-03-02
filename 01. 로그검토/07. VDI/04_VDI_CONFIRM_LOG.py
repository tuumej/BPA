import pandas as pd

# ---------------------------------------------------
# ① 엑셀 파일 경로 지정
# ---------------------------------------------------
input_path = r""
output_path = r""  # 저장 경로

# ---------------------------------------------------
# ② 엑셀 파일 읽기
# ---------------------------------------------------
df = pd.read_excel(input_path)

# timestamp 컬럼이 datetime이 아니면 변환
df['일시'] = pd.to_datetime(df['일시'], errors='coerce')

# 비고 컬럼 없으면 생성
if '비고' not in df.columns:
    df['비고'] = ""

# ---------------------------------------------------
# ③ 로그인 실패( result == "실패" )만 처리
# ---------------------------------------------------
df_fail = df[df['결과'] == '실패']

# ---------------------------------------------------
# ④ 사용자별로 1분 이내 실패 5회 이상 탐지
# ---------------------------------------------------
for user, group in df_fail.groupby('사용자 아이디'):
    group = group.sort_values('일시')

    # timestamp → seconds 변환
    ts = group['일시'].astype('int64') // 10**9

    # 60초 이내 실패 횟수 계산
    counts = ts.rolling(window=len(ts), min_periods=1).apply(
        lambda x: (x - x.iloc[-1] >= -60).sum(), raw=False
    )

    # 조건 충족 index
    idx_over = group[counts >= 5].index

    # 원본 DataFrame에 비고 텍스트 입력
    df.loc[idx_over, '비고'] = "로그인 실패 시도(1분이내 5회 이상)"

# ---------------------------------------------------
# ⑤ 결과 엑셀 저장
# ---------------------------------------------------
df.to_excel(output_path, index=False)
print("완료되었습니다:", output_path)
