# config.py

# Paths
SOURCE_CSV_DIR = r""  # CSV 파일 경로
TARGET_EXCEL_FILE = r""  # 편집할 엑셀 파일

# Target sheet
TARGET_SHEET_NAME = "04.DB Safer 7.0 (공공)"

# Columns
TARGET_COLUMNS = [
    "사이트",
    "소속사",
    "인력구분",
    "직무",
    "성명",
    "ID",
    "마지막 로그인시간",
    "검토결과",
    "결과사유",
    "비고",
]

# CSV to target mapping
CSV_COLUMN_MAPPING = {
    "[보안계정]": "ID",
    "[사용자명]": "성명",
    "[마지막 로그인 시각]": "마지막 로그인시간",
}

# Review rules
REVIEW_STOP_DAYS = 90
REVIEW_DELETE_DAYS = 180
REVIEW_STOP = "중지"
REVIEW_DELETE = "삭제"
REVIEW_KEEP = "유지"

# Colors
COLOR_STOP = "BDD7EE"    # 파랑
COLOR_DELETE = "F4B084"  # 빨강

# Style
FONT_SIZE = 8
ALIGN_HORIZONTAL = "left"
ALIGN_VERTICAL = "center"

# Date format
OUTPUT_DATE_FORMAT = "%Y-%m-%d %H:%M"
