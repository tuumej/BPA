# config.py

# NCP Archive Storage 인증 정보
# CLOUD NAME
NCP_ACCESS_KEY = ""
NCP_SECRET_KEY = ""

# Keystone v3
AUTH_URL = "https://kr.archive.ncloudstorage.com:5000/v3"
DOMAIN_ID = "default"
# CLOUD ARCHIVE NAME
PROJECT_ID = ""

# 다운로드 기본 경로
BASE_DOWNLOAD_DIR = "./"

# 다운로드 작업 목록
DOWNLOAD_JOBS = [
    # CLOUD NAME
    {
        "container": "",
        "prefix": ""
    } 
]
