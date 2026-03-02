# config.py

BASE_EXCEL_PATH = r""
SOURCE_EXCEL_DIR = r""
# config.py

# ===== FILE PATH =====
TARGET_EXCEL_PATH = r""
RESULT_EXCEL_PATH = r""

# ===== SHEET =====
TARGET_SHEET_NAME = "09.NCP-SSLVPN (민간)"
RESULT_SHEET_NAME = "Result"

# ===== COLUMNS =====
TARGET_COLUMNS = [
    "사이트", "소속사", "인력구분", "직무", "성명", "ID",
    "구분", "상세구분", "마지막 로그인시간",
    "검토결과", "결과사유", "비고"
]

RESULT_FILTER_COLUMN = "message.result"
RESULT_FILTER_VALUE = "\"connect\""

RESULT_MAPPING = {
    "message1": "마지막 로그인시간",
    "message.username": "성명",
    "sslvpn.sslVpnName": "상세구분",
    "구분": "구분"
}

# ===== STYLE =====
FONT_SIZE = 8
ALIGN_HORIZONTAL = "left"
ALIGN_VERTICAL = "center"

COLOR_RED = "F4B084"
COLOR_BLUE = "BDD7EE"

# ===== DATE =====
DATE_FORMAT = "%Y-%m-%d %H:%M"

