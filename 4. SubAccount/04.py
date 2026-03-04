import time
import hmac
import hashlib
import base64
import requests

from config import ACCESS_KEY, SECRET_KEY, BASE_HOST


def make_signature(method, uri, timestamp):
    """
    NCP API Gateway Signature v2
    message format:
    {method} {uri}\n{timestamp}\n{accessKey}
    """
    message = f"{method} {uri}\n{timestamp}\n{ACCESS_KEY}"
    signing_key = SECRET_KEY.encode("utf-8")

    signature = hmac.new(
        signing_key,
        message.encode("utf-8"),
        hashlib.sha256
    ).digest()

    return base64.b64encode(signature).decode("utf-8")


def fetch_subaccounts(page, size=10):
    method = "GET"

    # ⚠️ page는 0-base
    page_index = page - 1

    uri = f"/api/v1/sub-accounts?page={page_index}&size={size}"
    timestamp = str(int(time.time() * 1000))

    headers = {
        "x-ncp-apigw-timestamp": timestamp,
        "x-ncp-access-key": ACCESS_KEY,
        "x-ncp-apigw-signature-v2": make_signature(method, uri, timestamp),
        "Accept": "application/json"
    }

    response = requests.get(
        BASE_HOST + uri,
        headers=headers,
        timeout=10
    )
    response.raise_for_status()

    return response.json()


def main():
    page = 2
    size = 10

    result = fetch_subaccounts(page, size)

    total_count = result.get("totalCount")
    items = result.get("items", [])

    print(f"전체 계정 수: {total_count}")
    print(f"{page}페이지 계정 수: {len(items)}")

    for acct in items:
        print(
            acct.get("subAccountId"),
            acct.get("loginId"),
            acct.get("name")
        )


if __name__ == "__main__":
    main()
