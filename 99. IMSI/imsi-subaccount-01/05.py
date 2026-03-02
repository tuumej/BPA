import time
import datetime
import hmac
import hashlib
import base64
import json
import requests
import os


# ==============================
# 설정값
# ==============================

ACCESS_KEY = ""
SECRET_KEY = ""

API_HOST = "https://cloudactivitytracer.apigw.ntruss.com"
REQUEST_URI = "/api/v1/activities"
METHOD = "POST"

# 저장 경로 (원하는 경로로 변경)
OUTPUT_DIR = r""


# ==============================
# Signature 생성
# ==============================
def make_signature(timestamp: str) -> str:
    message = f"{METHOD} {REQUEST_URI}\n{timestamp}\n{ACCESS_KEY}"
    digest = hmac.new(
        SECRET_KEY.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256
    ).digest()
    return base64.b64encode(digest).decode()


# ==============================
# 로그 조회
# ==============================
def get_activity_logs():

    now = datetime.datetime.utcnow()
    one_hour_ago = now - datetime.timedelta(hours=1)

    start_time = one_hour_ago.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_time = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    body = {
        "startTime": start_time,
        "endTime": end_time,
        "pageNo": 1,
        "pageSize": 200
    }

    body_str = json.dumps(body)

    timestamp = str(int(time.time() * 1000))
    signature = make_signature(timestamp)

    headers = {
        "Content-Type": "application/json",
        "x-ncp-apigw-timestamp": timestamp,
        "x-ncp-iam-access-key": ACCESS_KEY,
        "x-ncp-apigw-signature-v2": signature
    }

    url = API_HOST + REQUEST_URI

    response = requests.post(url, headers=headers, data=body_str)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error {response.status_code}: {response.text}")
        return None


# ==============================
# JSON 파일 저장
# ==============================
def save_to_json(data):

    if data is None:
        print("저장할 데이터가 없습니다.")
        return

    # 디렉터리 없으면 생성
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    now_str = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    file_name = f"cloud_activity_{now_str}.json"
    file_path = os.path.join(OUTPUT_DIR, file_name)

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"파일 저장 완료: {file_path}")


# ==============================
# main
# ==============================
def main():
    logs = get_activity_logs()
    save_to_json(logs)


if __name__ == "__main__":
    main()
