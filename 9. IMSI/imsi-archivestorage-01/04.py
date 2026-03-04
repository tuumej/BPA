import os
import requests
from config import (
    DOWNLOAD_JOBS,
    BASE_DOWNLOAD_DIR,
    NCP_ACCESS_KEY,
    NCP_SECRET_KEY,
    AUTH_URL,
    DOMAIN_ID,
    PROJECT_ID
)

ARCHIVE_ENDPOINT = "https://kr.archive.ncloudstorage.com"

def get_token_and_storage_account():
    """
    Keystone v3 인증
    → Token 발급
    → Swift Storage Account (AUTH_xxxx) 안전 추출
    """
    url = f"{AUTH_URL}/auth/tokens"
    headers = {"Content-Type": "application/json"}

    payload = {
        "auth": {
            "identity": {
                "methods": ["password"],
                "password": {
                    "user": {
                        "name": NCP_ACCESS_KEY,
                        "password": NCP_SECRET_KEY,
                        "domain": {"id": DOMAIN_ID}
                    }
                }
            },
            "scope": {
                "project": {"id": PROJECT_ID}
            }
        }
    }

    resp = requests.post(url, json=payload, headers=headers)
    resp.raise_for_status()

    token = resp.headers.get("X-Subject-Token")
    if not token:
        raise RuntimeError("Failed to get X-Subject-Token")

    catalog = resp.json()["token"]["catalog"]
    object_store = next(
        svc for svc in catalog if svc.get("type") == "object-store"
    )

    # public endpoint 우선
    endpoints = object_store["endpoints"]
    public_ep = next(
        ep for ep in endpoints if ep.get("interface") == "public"
    )

    endpoint_url = public_ep["url"]

    # AUTH_xxxx 안전 추출
    storage_account = endpoint_url.rstrip("/").split("/")[-1]

    if not storage_account.startswith("AUTH_"):
        raise RuntimeError(
            f"Invalid storage_account parsed: {storage_account}, endpoint={endpoint_url}"
        )

    return token, storage_account


def build_container_url(storage_account, container):
    """
    Archive Storage 컨테이너 URL 생성 (강제 검증)
    """
    if not storage_account or not storage_account.startswith("AUTH_"):
        raise ValueError(f"Invalid storage_account: {storage_account}")

    return f"{ARCHIVE_ENDPOINT}/v1/{storage_account}/{container}"


def list_objects(token, storage_account, container, prefix):
    """
    prefix 하위 object 목록 조회
    """
    url = build_container_url(storage_account, container)
    headers = {"X-Auth-Token": token}
    params = {
        "prefix": prefix,
        "format": "json"
    }

    print(f"[DEBUG] LIST URL: {url}")

    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()

    return resp.json() if resp.text else []


def download_object(token, storage_account, container, object_name, save_path):
    """
    단일 object 다운로드
    """
    url = f"{build_container_url(storage_account, container)}/{object_name}"
    headers = {"X-Auth-Token": token}

    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    with requests.get(url, headers=headers, stream=True) as r:
        r.raise_for_status()
        with open(save_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)


def run_job(job_index, job, token, storage_account):
    container = job["container"]
    prefix = job["prefix"]

    print(f"[JOB {job_index}] START container={container}, prefix={prefix}")

    objects = list_objects(token, storage_account, container, prefix)

    if not objects:
        print(f"[JOB {job_index}] NO OBJECTS FOUND")
        return

    for obj in objects:
        object_name = obj.get("name")

        # Swift 가상 폴더 제외
        if not object_name or object_name.endswith("/"):
            continue

        local_path = os.path.join(
            BASE_DOWNLOAD_DIR,
            container,
            object_name
        )

        download_object(
            token,
            storage_account,
            container,
            object_name,
            local_path
        )

        print(f"[JOB {job_index}] DOWNLOADED {object_name}")

    print(f"[JOB {job_index}] END")


def main():
    token, storage_account = get_token_and_storage_account()

    print(f"[INFO] STORAGE_ACCOUNT = {storage_account}")

    for idx, job in enumerate(DOWNLOAD_JOBS, start=1):
        try:
            run_job(idx, job, token, storage_account)
        except Exception as e:
            print(f"[JOB {idx} ERROR] {e}")


if __name__ == "__main__":
    main()
