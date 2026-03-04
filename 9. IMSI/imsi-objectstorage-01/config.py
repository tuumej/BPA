# config.py

ENDPOINT_URL = "https://kr.object.ncloudstorage.com"
REGION_NAME = "kr-standard"

BASE_DOWNLOAD_DIR = "./"

# 계정 + 다운로드 작업 세트
DOWNLOAD_SETS = [
    {
        "name": "",
        "NCP_ACCESS_KEY": "",
        "NCP_SECRET_KEY": "",
        "DOWNLOAD_JOBS": [
            {
                "bucket": "",
                "prefix": "",
            },
        ],
    },
]
