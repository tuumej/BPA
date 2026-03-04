
import os
from keystoneauth1 import session
from keystoneauth1.identity import v3
import swiftclient

# ===============================
# 인증 정보 (API 인증키)
# ===============================
OS_AUTH_URL      = "https://kr.archive.ncloudstorage.com:5000/v3"
OS_USERNAME      = ""
OS_PASSWORD      = ""
OS_PROJECT_ID    = ""
OS_USER_DOMAIN_ID = "default"

# ===============================
# 업로드할 로컬 경로 & 컨테이너 (버킷)
# ===============================
LOCAL_DIR         = r""
CONTAINER_NAME    = ""
ARCHIVE_PREFIX    = ""  # 업로드될 Archive Storage 경로 (폴더처럼)

# ===============================
# 인증 & 세션 설정
# ===============================
auth = v3.Password(
    auth_url=OS_AUTH_URL,
    username=OS_USERNAME,
    password=OS_PASSWORD,
    project_id=OS_PROJECT_ID,
    user_domain_id=OS_USER_DOMAIN_ID,
)

auth_session = session.Session(auth=auth, timeout=30)

# Swift connection
conn = swiftclient.Connection(
    session=auth_session,
)

# ===============================
# 폴더 전체 업로드
# ===============================
def upload_directory(local_dir, container, prefix):
    for root, dirs, files in os.walk(local_dir):
        for file_name in files:
            local_path = os.path.join(root, file_name)

            # 로컬 기준 상대경로
            rel_path = os.path.relpath(local_path, local_dir)

            # Archive Storage 상의 객체 키
            object_path = f"{prefix}/{rel_path}".replace("\\", "/")

            with open(local_path, "rb") as fp:
                print(f"↑ 업로드: {object_path}")
                conn.put_object(
                    container=container,
                    obj=object_path,
                    contents=fp.read()
                )

if __name__ == "__main__":
    upload_directory(LOCAL_DIR, CONTAINER_NAME, ARCHIVE_PREFIX)
    print("🎉 업로드 완료!")

