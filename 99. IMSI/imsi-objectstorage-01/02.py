import os
import boto3
from botocore.exceptions import ClientError

from config import (
    DOWNLOAD_SETS,
    BASE_DOWNLOAD_DIR,
    ENDPOINT_URL,
    REGION_NAME,
)


def download_prefix(s3, bucket, prefix, base_dir):
    paginator = s3.get_paginator("list_objects_v2")

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        if "Contents" not in page:
            print(f"[INFO] No objects: s3://{bucket}/{prefix}")
            continue

        for obj in page["Contents"]:
            key = obj["Key"]

            # 폴더 객체 skip
            if key.endswith("/"):
                continue

            local_path = os.path.join(
                base_dir,
                bucket,
                key.replace(prefix, "", 1)
            )

            os.makedirs(os.path.dirname(local_path), exist_ok=True)

            try:
                print(f"[DOWNLOAD] s3://{bucket}/{key}")
                s3.download_file(bucket, key, local_path)
            except ClientError as e:
                print(f"[ERROR] {key} 다운로드 실패: {e}")


def main():
    for set_idx, account in enumerate(DOWNLOAD_SETS, start=1):
        try:
            print(f"\n[ACCOUNT {set_idx}] {account['name']}")

            s3 = boto3.client(
                service_name="s3",
                aws_access_key_id=account["NCP_ACCESS_KEY"],
                aws_secret_access_key=account["NCP_SECRET_KEY"],
                endpoint_url=ENDPOINT_URL,
                region_name=REGION_NAME,
            )

            for job_idx, job in enumerate(account["DOWNLOAD_JOBS"], start=1):
                try:
                    print(
                        f"[JOB {set_idx}-{job_idx}] "
                        f"bucket={job['bucket']} prefix={job['prefix']}"
                    )

                    download_prefix(
                        s3=s3,
                        bucket=job["bucket"],
                        prefix=job["prefix"],
                        base_dir=os.path.join(
                            BASE_DOWNLOAD_DIR,
                            account["name"]
                        ),
                    )

                except Exception as e:
                    print(f"[JOB {set_idx}-{job_idx} ERROR] {e}")

        except Exception as e:
            print(f"[ACCOUNT {set_idx} ERROR] {e}")


if __name__ == "__main__":
    main()
