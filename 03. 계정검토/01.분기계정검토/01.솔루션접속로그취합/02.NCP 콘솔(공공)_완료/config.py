# config.py

# Paths
SOURCE_EXCEL_DIR = r""  # 데이터를 가져올 폴더
TARGET_EXCEL_FILE = r""  # 편집할 엑셀 파일

# Target sheet
TARGET_SHEET_NAME = "02.NCP 콘솔 (공공)"

# Columns
TARGET_COLUMNS = [
    "사이트",
    "소속사",
    "인력구분",
    "직무",
    "성명",
    "ID",
    "구분",
    "마지막 로그인시간",
    "검토결과",
    "결과사유",
    "비고",
]

# Source to target mapping
COLUMN_MAPPING = {
    "로그인": "ID",
    "사용자명": "성명",
    "최종": "마지막 로그인시간",
}

# Sheet filter keywords
INCLUDE_KEYWORDS = ["NCP 콘솔"]
EXCLUDE_KEYWORDS = [""]

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
