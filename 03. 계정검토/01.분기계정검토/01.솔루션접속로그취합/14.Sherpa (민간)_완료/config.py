## config.py

# 경로 및 파일
TARGET_FILE = r""
SOURCE_FOLDER = r""

# 워크시트명
TARGET_SHEET = "14.Sherpa (민간)"
SOURCE_SHEET = "Sherpa(민간)"

# 생성할 컬럼
TARGET_COLUMNS = [
    "사이트","소속사","인력구분","직무","성명",
    "ID","마지막 로그인시간","검토결과","결과사유","비고"
]

# 컬럼 매핑 {source_column: target_column}
COLUMN_MAPPING = {
    "ID": "ID",
    "마지막접근이력": "마지막 로그인시간"
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
