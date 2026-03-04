# config.py

# Paths
EXCEL_TAB_NAME_DIR = r""    # 탭 이름을 가져올 폴더 경로
SOURCE_EXCEL_DIR = r""    # 데이터를 가져올 엑셀 파일이 있는 폴더
TARGET_EXCEL_FILE = r""     # 생성할 엑셀 파일

# Target sheet
TARGET_SHEET_NAME = "01.NCP 콘솔 (민간)"

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

# Source to target column mapping
COLUMN_MAPPING = {
    "로그인": "ID",
    "사용자명": "성명",
    "최종": "마지막 로그인시간",
}

# Sheet filter
INCLUDE_KEYWORDS = ["NCP 콘솔", "CLOUDNAME"]
EXCLUDE_KEYWORDS = []

# Review rules
REVIEW_STOP_DAYS = 90
REVIEW_DELETE_DAYS = 180
REVIEW_KEEP = "유지"
REVIEW_STOP = "중지"
REVIEW_DELETE = "삭제"

# Colors
COLOR_STOP = "BDD7EE"    # 파랑
COLOR_DELETE = "F4B084"  # 빨강

# Style
FONT_SIZE = 8
ALIGN_HORIZONTAL = "left"
ALIGN_VERTICAL = "center"

# Date format
OUTPUT_DATE_FORMAT = "%Y-%m-%d %H:%M"
