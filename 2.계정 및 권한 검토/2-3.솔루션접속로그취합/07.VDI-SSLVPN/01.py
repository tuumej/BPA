import os
import pandas as pd

from config import (
    SOURCE_DIR,
    OUTPUT_XLSX_PATH,
    TARGET_KEYWORD,
    CSV_ENCODING,
    FILTER_COLUMN,
    FILTER_KEYWORD,
    REMAIN_COLUMNS,
    USER_ID_COLUMN,
    TIMESTAMP_COLUMN,
    TIMESTAMP_FORMAT,
)


def collect_csv_files():
    return [
        os.path.join(SOURCE_DIR, f)
        for f in os.listdir(SOURCE_DIR)
        if f.lower().endswith(".csv") and TARGET_KEYWORD in f
    ]


def parse_timestamp(series: pd.Series) -> pd.Series:
    """
    timestamp 컬럼 자동 판별
    - 문자열 datetime
    - Unix timestamp (초 / ms)
    """
    # 숫자인지 판단
    if pd.api.types.is_numeric_dtype(series):
        # 값 크기로 초 / ms 판단
        sample = series.dropna().iloc[0]

        if sample > 1_000_000_000_000:  # ms
            return pd.to_datetime(series, unit="ms", errors="coerce")
        else:  # s
            return pd.to_datetime(series, unit="s", errors="coerce")

    # 문자열인 경우
    return pd.to_datetime(series, errors="coerce")


def merge_filter_and_deduplicate(csv_files):
    merged_df = pd.DataFrame()

    for idx, csv_path in enumerate(csv_files, start=1):
        try:
            print(f"[JOB {idx}] CSV 읽는 중: {csv_path}")
            df = pd.read_csv(csv_path, encoding=CSV_ENCODING)
            merged_df = pd.concat([merged_df, df], ignore_index=True)
            print(f"[JOB {idx}] 병합 완료")
        except Exception as e:
            print(f"[JOB {idx} ERROR] {csv_path}: {e}")

    if merged_df.empty:
        print("[INFO] 병합된 데이터 없음")
        return

    # 1️⃣ description 필터
    if FILTER_COLUMN not in merged_df.columns:
        print(f"[ERROR] '{FILTER_COLUMN}' 컬럼 없음")
        return

    filtered_df = merged_df[
        merged_df[FILTER_COLUMN]
        .astype(str)
        .str.contains(FILTER_KEYWORD, na=False)
    ]

    if filtered_df.empty:
        print("[INFO] 로그인 로그 없음")
        return

    # 2️⃣ 필요한 컬럼만 유지
    trimmed_df = filtered_df[REMAIN_COLUMNS].copy()

    # 3️⃣ timestamp 안전 파싱 (🔥 핵심)
    trimmed_df[TIMESTAMP_COLUMN] = parse_timestamp(
        trimmed_df[TIMESTAMP_COLUMN]
    )

    # 파싱 실패 제거
    trimmed_df = trimmed_df.dropna(subset=[TIMESTAMP_COLUMN])

    # 🚨 1970 값 방어 (epoch 잔존 제거)
    trimmed_df = trimmed_df[
        trimmed_df[TIMESTAMP_COLUMN] > pd.Timestamp("2000-01-01")
    ]

    # 4️⃣ user_id 기준 최신 timestamp만 유지
    dedup_df = (
        trimmed_df
        .sort_values(TIMESTAMP_COLUMN, ascending=False)
        .drop_duplicates(subset=[USER_ID_COLUMN], keep="first")
    )

    # 5️⃣ timestamp 포맷 변경
    dedup_df[TIMESTAMP_COLUMN] = dedup_df[TIMESTAMP_COLUMN].dt.strftime(
        TIMESTAMP_FORMAT
    )

    # 6️⃣ 저장
    os.makedirs(os.path.dirname(OUTPUT_XLSX_PATH), exist_ok=True)
    dedup_df.to_excel(OUTPUT_XLSX_PATH, index=False)

    print(f"[SUCCESS] XLSX 생성 완료: {OUTPUT_XLSX_PATH}")
    print(f"[INFO] 최종 사용자 수: {len(dedup_df)}")


def main():
    try:
        csv_files = collect_csv_files()

        if not csv_files:
            print("[INFO] 대상 CSV 없음")
            return

        merge_filter_and_deduplicate(csv_files)

    except Exception as e:
        print(f"[FATAL ERROR] {e}")


if __name__ == "__main__":
    main()
