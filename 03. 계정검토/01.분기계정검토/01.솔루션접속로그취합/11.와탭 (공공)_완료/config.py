# =========================
# Path Configuration
# =========================
BASE_EXCEL_PATH = r""
SOURCE_FOLDER_PATH = r""

# =========================
# Sheet Configuration
# =========================
TARGET_SHEET_NAME = "11.와탭 (공공)"
SOURCE_SHEET_INCLUDE = "Whatap"
SOURCE_SHEET_EXCLUDES = ["온라인클래스", "패밀리사이트"]

# =========================
# Column Configuration
# =========================
TARGET_COLUMNS = [
    "사이트", "소속사", "인력구분", "직무",
    "성명", "ID", "구분",
    "마지막 로그인시간",
    "검토결과", "결과사유", "비고"
]

SOURCE_COLUMN_MAPPING = {
    "로그인": "ID",
    "이름": "성명",
    "최종": "마지막 로그인시간"
}

# =========================
# Style Configuration
# =========================
FONT_SIZE = 8
ALIGN_HORIZONTAL = "left"
ALIGN_VERTICAL = "center"

# requested color mapping
COLOR_RED = "F4B084"
COLOR_BLUE = "BDD7EE"

# =========================
# Review Policy
# =========================
STOP_DAYS = 90
DELETE_DAYS = 180

# =========================
# Date Format
# =========================
OUTPUT_DATE_FORMAT = "%Y-%m-%d %H:%M"
