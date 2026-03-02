# @@작성자 : tuumej
# @@작성일자 : 251224
# @@참고_01 : 스크립트
# ncp objectstorage 특정 버킷의 특정 경로 내의 파일 다운로드하는 파이썬 코드 작성해줘
# ACCESS_KEY, SECRET_KEY, BUCKET_NAME, PREFIX, DOWNLOAD_DIR 는 하나의 배열로 묶어서 여러 개를 한번에 실행시킬 수 있도록 코드 작성해줘
# DOWNLOAD_JOB 배열은 어떻게 불러와?
# DOWNLOAD_JOBS 에 NAME 값을 추가하고 DOWNLOAD_DIR의 값이 실행한 시점의 년월 + _ + NAME이 되고 해당 폴더가 없을 시 생성하는 코드로 수정해줘
# 최종 파일에 호출 시점의 년 월이 포함되어 있는 파일만 다운로드 해줘. 예를 들어서 25년도 12월에 호출했으면 202512가 파일명에 포함되어 있으면 돼. 25년도 11월에 호출했으면 202511가 파일명에 포함되어 있으면 돼.
# download_dir 경로 안에 PREFIX 경로는 생성하지 말고 파일만 다운로드 해줘
# DOWNLOAD_JOBS 배열에 공공, 민간을 분류해서 민간은 ENDPOINT_URL이 https://kr.object.ncloudstorage.com 이고 공공은 ENDPOINT_URL이 https://kr.object.gov-ncloudstorage.com 을 호출하도록 변경 해줘.
# DOWNLOAD_JOBS 배열에 BUCKET_NAME을 여러개 저장할 수 있도록 배열로 만들고 버킷 개수만큼 호출하도록 변경해줘
# DOWNLOAD_JOBS, BASE_DOWNLOAD_DIR 값을 별도 파일에서 가져올 수 있도록 코드 작성하고 가져오는 방법 알려줘
# 이 프로그램을 윈도우에서 exe 파일로 돌아가도록 만들어주고 배치로 매주 월요일 9시에 돌아가는 방법 알려줘
# 배치클릭했을때 exe 파일 실행이 안되는데?

# @@참고_02 : exe 파일 만드는 방법
# exe 파일 생성을 위한 사전 작업 : pyinstaller --onefile main.py
# exe 파일 생성 : pyinstaller --onefile --add-data "config.py;." main.py

import boto3
import os
from datetime import datetime

# ✅ 설정 파일에서 값 가져오기
from config import DOWNLOAD_JOBS, BASE_DOWNLOAD_DIR


def get_endpoint_url(job_type):
    if job_type == "public":
        return "https://kr.object.gov-ncloudstorage.com"
    elif job_type == "private":
        return "https://kr.object.ncloudstorage.com"
    else:
        raise ValueError(f"알 수 없는 TYPE 값: {job_type}")


def download_prefix(job):
    year_month = datetime.now().strftime("%Y%m")

    download_dir = os.path.join(
        BASE_DOWNLOAD_DIR,
        f"{year_month}_{job['NAME']}"
    )
    os.makedirs(download_dir, exist_ok=True)

    s3 = boto3.client(
        "s3",
        aws_access_key_id=job["ACCESS_KEY"],
        aws_secret_access_key=job["SECRET_KEY"],
        endpoint_url=get_endpoint_url(job["TYPE"])
    )

    for bucket_name in job["BUCKET_NAME"]:
        paginator = s3.get_paginator("list_objects_v2")

        for page in paginator.paginate(
            Bucket=bucket_name,
            Prefix=job["PREFIX"]
        ):
            if "Contents" not in page:
                continue

            for obj in page["Contents"]:
                key = obj["Key"]

                if key.endswith("/"):
                    continue

                filename = os.path.basename(key)

                # YYYYMM 필터
                if year_month not in filename:
                    continue

                local_path = os.path.join(download_dir, filename)

                print(f"[{job['TYPE']} | {job['NAME']} | {bucket_name}] {filename}")
                s3.download_file(bucket_name, key, local_path)


def main():
    for idx, job in enumerate(DOWNLOAD_JOBS, start=1):
        print(f"\n=== Job {idx} ({job['NAME']}) 시작 ===")
        try:
            download_prefix(job)
            print(f"=== Job {idx} ({job['NAME']}) 완료 ===")
        except Exception as e:
            print(f"❌ Job {idx} ({job['NAME']}) 실패: {e}")


if __name__ == "__main__":
    main()
