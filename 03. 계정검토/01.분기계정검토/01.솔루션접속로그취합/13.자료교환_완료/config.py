# ===============================
# File / Path
# ===============================
TARGET_EXCEL_PATH = r""
RESULT_FOLDER_PATH = r""

# ===============================
# Worksheet
# ===============================
TARGET_WORKSHEET_NAME = "13.자료교환"

# ===============================
# Columns
# ===============================
CREATE_COLUMNS = [
    "사이트", "소속사", "인력구분", "직무", "성명",
    "ID", "마지막 로그인시간", "검토결과", "결과사유", "비고"
]

RESULT_SOURCE_COLUMNS = {
    "이름": "성명",
    "아이디": "ID",
    "최종접속일자": "마지막 로그인시간"
}

# ===============================
# Excel Style
# ===============================
FONT_SIZE = 8
ALIGN_HORIZONTAL = "left"
ALIGN_VERTICAL = "center"

COLOR_RED = "F4B084"
COLOR_BLUE = "BDD7EE"

# ===============================
# Review Result Values
# ===============================
RESULT_KEEP = "유지"
RESULT_STOP = "중지"
RESULT_DELETE = "삭제"

# ===============================
# Date Format
# ===============================
DATETIME_FORMAT = "%Y-%m-%d %H:%M"
