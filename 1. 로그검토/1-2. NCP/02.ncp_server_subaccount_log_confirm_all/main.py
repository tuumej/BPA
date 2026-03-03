import os
import json
import pandas as pd
from datetime import datetime
from config import BASE_DIR, KEYWORDS, OUTPUT_PREFIX, LOCATION_MAP


def load_json_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        return pd.DataFrame(data)
    elif isinstance(data, dict):
        return pd.DataFrame([data])
    return pd.DataFrame()


def convert_event_time_to_kst(df):
    if "eventTime" not in df.columns:
        return df

    if pd.api.types.is_numeric_dtype(df["eventTime"]):
        df["eventTime"] = pd.to_datetime(
            df["eventTime"], unit="ms", utc=True, errors="coerce"
        )
    else:
        df["eventTime"] = pd.to_datetime(
            df["eventTime"], utc=True, errors="coerce"
        )

    df["eventTime"] = (
        df["eventTime"]
        .dt.tz_convert("Asia/Seoul")
        .dt.strftime("%Y-%m-%d %H:%M:%S")
    )

    return df


def extract_acg_name(df):
    if "productData" not in df.columns:
        df["ACG"] = ""
        return df

    def get_acg_value(val):
        try:
            if isinstance(val, dict):
                return val.get("accessControlGroupName", "")
            if isinstance(val, str):
                data = json.loads(val)
                return data.get("accessControlGroupName", "")
        except Exception:
            return ""
        return ""

    df["ACG"] = df["productData"].apply(get_acg_value)
    return df


def add_review_columns(df):
    def detect_review_item(val):
        items = []
        text = ""

        try:
            if isinstance(val, dict):
                text = json.dumps(val)
            elif isinstance(val, str):
                text = val
        except Exception:
            pass

        if "0.0.0.0/0" in text:
            items.append("0.0.0.0/0")
        if "1-65535" in text:
            items.append("1-65535")

        return ",".join(items)

    df["검토 항목"] = df["productData"].apply(detect_review_item) if "productData" in df.columns else ""
    df["검토 결과"] = ""
    df["비고"] = ""
    return df


def add_subaccount_columns(df):
    def extract_username(val):
        try:
            if isinstance(val, dict):
                return val.get("userName", "")
            if isinstance(val, str):
                data = json.loads(val)
                return data.get("userName", "")
        except Exception:
            return ""
        return ""

    def map_location(row):
        ip = ""
        if "sourceIp" in row:
            ip = row.get("sourceIp")
        elif "soureIp" in row:
            ip = row.get("soureIp")
        return LOCATION_MAP.get(ip, "")

    df["userName"] = df["productData"].apply(extract_username) if "productData" in df.columns else ""
    df["위치정보"] = df.apply(map_location, axis=1)
    df["접속사유"] = ""
    df["비고"] = ""
    return df


def fill_empty_with_dash(df):
    """
    모든 컬럼에서 값이 없는 셀을 '-' 로 치환
    """
    df = df.fillna("-")
    df = df.replace("", "-")
    return df


def main():
    yyyy_mm = "202601"#datetime.now().strftime("%Y%m")

    for idx, keyword in enumerate(KEYWORDS, start=1):
        try:
            target_dir = os.path.join(BASE_DIR, f"{yyyy_mm}_{keyword}")
            print(f"[{idx}] Processing folder: {target_dir}")

            if not os.path.isdir(target_dir):
                print("  - 폴더 없음, 스킵")
                continue

            acg_df_list = []
            subaccount_df_list = []

            for file_name in os.listdir(target_dir):
                if not file_name.lower().endswith(".json"):
                    continue

                file_path = os.path.join(target_dir, file_name)
                df = load_json_file(file_path)

                if df.empty:
                    continue

                df["source_file"] = file_name
                fname_lower = file_name.lower()

                if "server" in fname_lower:
                    df = convert_event_time_to_kst(df)
                    df = extract_acg_name(df)
                    df = add_review_columns(df)
                    df = fill_empty_with_dash(df)
                    acg_df_list.append(df)

                elif "account" in fname_lower:
                    df = convert_event_time_to_kst(df)
                    df = add_subaccount_columns(df)
                    df = fill_empty_with_dash(df)
                    subaccount_df_list.append(df)

            output_file = os.path.join(
                BASE_DIR,
                f"{OUTPUT_PREFIX}_{yyyy_mm}_{keyword}.xlsx"
            )

            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                if acg_df_list:
                    pd.concat(acg_df_list, ignore_index=True).to_excel(
                        writer, sheet_name="ACG", index=False
                    )

                if subaccount_df_list:
                    pd.concat(subaccount_df_list, ignore_index=True).to_excel(
                        writer, sheet_name="SUBACCOUNT", index=False
                    )

            print(f"  - 결과 생성 완료: {output_file}")

        except Exception as e:
            print(f"[{idx}] ERROR 발생: {e}")


if __name__ == "__main__":
    main()
