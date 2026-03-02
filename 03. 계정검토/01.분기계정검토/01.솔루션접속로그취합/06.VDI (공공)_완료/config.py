## config.py

# 경로 및 파일
TARGET_FILE = r""
SOURCE_FOLDER = r""

# 워크시트명
TARGET_SHEET = "06.VDI (공공)"
SOURCE_SHEET = "VDI(공공)"

# 생성할 컬럼 (큰따옴표 안 값 그대로)
TARGET_COLUMNS = [
    "사이트","소속사","인력구분","직무","성명",
    "ID","마지막 로그인시간","검토결과","결과사유","비고"
]

# 컬럼 매핑 {source_column: target_column}
COLUMN_MAPPING = {
    "아이디": "ID",
    "이름": "성명",
    "최종": "마지막 로그인시간"
}

# 색상 코드
COLOR_RED = "F4B084"
COLOR_BLUE = "BDD7EE"

# 엑셀 스타일
FONT_SIZE = 8
ALIGN_HORIZONTAL = "left"
ALIGN_VERTICAL = "center"

# 검토 기준 일수
DAYS_STOP = 90
DAYS_DELETE = 180
