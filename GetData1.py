import os
import requests
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime
from zoneinfo import ZoneInfo

env_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(env_path)

TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST")
SITE_NAME = os.getenv("SITE_NAME")
LIST_NAME = os.getenv("LIST_NAME")
GAP_LIST_NAME = os.getenv("GAP_LIST_NAME")


def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    res = requests.post(url, data=data)
    if not res.ok:
        print("Token error:", res.status_code)
        print(res.text)
    res.raise_for_status()
    return res.json()["access_token"]


def get_site_id(token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
    headers = {"Authorization": f"Bearer {token}"}

    res = requests.get(url, headers=headers)
    if not res.ok:
        print("Site error:", res.status_code)
        print(res.text)
    res.raise_for_status()
    return res.json()["id"]


def get_list_id(token, site_id, list_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    headers = {"Authorization": f"Bearer {token}"}

    res = requests.get(url, headers=headers)
    if not res.ok:
        print("List error:", res.status_code)
        print(res.text)
    res.raise_for_status()

    target = list_name.strip().lower()

    for item in res.json().get("value", []):
        display_name = item.get("displayName", "").strip()
        if display_name.lower() == target:
            return item["id"]

    raise Exception(f"List not found: {list_name}")


def get_all_items(token, site_id, list_id):
    headers = {"Authorization": f"Bearer {token}"}

    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
        f"?$expand=fields&$top=5000"
    )

    all_items = []

    while url:
        res = requests.get(url, headers=headers)
        if not res.ok:
            print("Get items error:", res.status_code)
            print(res.text)
        res.raise_for_status()

        data = res.json()
        all_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return all_items


def flatten_items(items):
    rows = []

    for item in items:
        fields = item.get("fields", {}).copy()

        row = {
            "ItemID": item.get("id"),
            "CreatedDateTime": item.get("createdDateTime"),
            "LastModifiedDateTime": item.get("lastModifiedDateTime"),
            "WebUrl": item.get("webUrl"),
        }

        for k, v in fields.items():
            row[k] = v

        rows.append(row)

    return pd.DataFrame(rows)


def get_roster_dataframe():
    token = get_token()
    site_id = get_site_id(token)
    list_id = get_list_id(token, site_id, LIST_NAME)

    items = get_all_items(token, site_id, list_id)
    return flatten_items(items)


def get_workers_for_app():
    df = get_roster_dataframe()

    print("Columns from SharePoint:")
    print(df.columns.tolist())

    required_columns = [
        "Title",
        "First_x0020_Name",
        "Last_x0020_Name",
        "OPMS",
        "Position",
        "Project",
        "Date_x0020_From",
        "WorkType",
    ]

    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise Exception(f"Missing columns in SharePoint list: {missing}")

    worker_df = df[required_columns].copy()

    worker_df = worker_df.rename(
        columns={
            "Title": "name",
            "First_x0020_Name": "first_name",
            "Last_x0020_Name": "last_name",
            "OPMS": "opms",
            "Position": "position",
            "Project": "project",
            "Date_x0020_From": "roster_date",
            "WorkType": "work_type",
        }
    )

    worker_df = worker_df.fillna("")

    for col in worker_df.columns:
        worker_df[col] = worker_df[col].astype(str).str.strip()

    worker_df = worker_df[worker_df["name"] != ""]

    worker_df = worker_df[
        ~worker_df["name"].str.contains(r"\d", regex=True, na=False)
    ]

    exclude_keywords = [
        "vehicle",
        "truck",
        "utility",
        "toyota",
        "hilux",
        "hino",
        "trailer",
        "plant",
        "equipment",
        "mlvt",
        "tphp",
        "ute",
    ]

    exclude_pattern = "|".join(exclude_keywords)

    worker_df = worker_df[
        ~worker_df["name"]
        .str.lower()
        .str.contains(exclude_pattern, regex=True, na=False)
    ]

    worker_df["site"] = worker_df["project"]

    supervisor_keywords = [
        "manager",
        "director",
        "supervisor",
        "superintendent",
    ]

    worker_df["is_supervisor_candidate"] = worker_df["position"].str.lower().apply(
        lambda x: any(keyword in x for keyword in supervisor_keywords)
    )

    worker_df = worker_df.drop_duplicates(subset=["opms"], keep="last")

    return worker_df.to_dict(orient="records")


def get_supervisors_for_app():
    workers = get_workers_for_app()

    supervisors = [
        {
            "name": w["name"],
            "opms": w["opms"],
            "position": w["position"],
            "site": w["site"],
            "project": w["project"],
        }
        for w in workers
        if w["is_supervisor_candidate"]
    ]

    return supervisors


# ==============================
# PPL-Timesheets write-back logic
# ==============================

def parse_sharepoint_datetime(value):
    if not value:
        return None

    value = str(value).strip()

    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=ZoneInfo("Australia/Perth"))
        return parsed.astimezone(ZoneInfo("Australia/Perth"))
    except Exception:
        pass

    try:
        return datetime.strptime(value, "%Y/%m/%d %H:%M:%S").replace(
            tzinfo=ZoneInfo("Australia/Perth")
        )
    except Exception:
        pass

    try:
        return datetime.strptime(value, "%Y-%m-%dT%H:%M").replace(
            tzinfo=ZoneInfo("Australia/Perth")
        )
    except Exception:
        pass

    try:
        return datetime.strptime(value, "%Y-%m-%dT%H:%M:%S").replace(
            tzinfo=ZoneInfo("Australia/Perth")
        )
    except Exception:
        pass

    return None


def get_requested_record_time(data):
    custom_time = str(data.get("custom_time", "")).strip()

    if custom_time:
        parsed_time = parse_sharepoint_datetime(custom_time)

        if parsed_time is None:
            raise Exception("Invalid custom time format.")

        return parsed_time.astimezone(ZoneInfo("Australia/Perth"))

    return datetime.now(ZoneInfo("Australia/Perth"))


def has_custom_time(data):
    return bool(str(data.get("custom_time", "")).strip())


def build_data_with_forced_time(data, forced_time):
    copied = dict(data)
    copied["custom_time"] = forced_time.isoformat()
    return copied


def get_gap_context():
    if not GAP_LIST_NAME:
        raise Exception("Missing GAP_LIST_NAME in .env")

    token = get_token()
    site_id = get_site_id(token)
    gap_list_id = get_list_id(token, site_id, GAP_LIST_NAME)

    return token, site_id, gap_list_id


def get_timesheet_items_by_opms(opms):
    token, site_id, gap_list_id = get_gap_context()

    headers = {"Authorization": f"Bearer {token}"}

    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{gap_list_id}/items"
        f"?$expand=fields&$top=5000"
    )

    all_items = []

    while url:
        res = requests.get(url, headers=headers)

        if not res.ok:
            print("Graph read error status:", res.status_code)
            print("Graph read error body:", res.text)

        res.raise_for_status()

        data = res.json()
        all_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    target_opms = str(opms).strip()

    matched_items = []

    for item in all_items:
        fields = item.get("fields", {})
        item_opms = str(fields.get("OPMS", "")).strip()

        if item_opms == target_opms:
            matched_items.append(item)

    return matched_items


def build_record_from_item(item):
    fields = item.get("fields", {})

    return {
        "item_id": item.get("id"),
        "title": fields.get("Title", ""),
        "first_name": fields.get("FirstName", ""),
        "last_name": fields.get("LastName", ""),
        "opms": fields.get("OPMS", ""),
        "position": fields.get("Position", ""),
        "date_time": fields.get("Date", ""),
        "date_time_obj": item.get("_parsed_date"),
        "project": fields.get("Project", ""),
        "site": fields.get("Site", ""),
        "unique_key": fields.get("Uniquekey", ""),
        "status": fields.get("Status", ""),
        "reason": fields.get("Reason", ""),
        "raw": item,
    }


def find_latest_today_timesheet_record(opms, reference_time=None):
    if reference_time is None:
        reference_time = datetime.now(ZoneInfo("Australia/Perth"))

    target_date = reference_time.astimezone(ZoneInfo("Australia/Perth")).strftime(
        "%Y-%m-%d"
    )

    items = get_timesheet_items_by_opms(opms)

    same_day_items = []

    for item in items:
        fields = item.get("fields", {})
        date_time_value = str(fields.get("Date", "")).strip()
        parsed_time = parse_sharepoint_datetime(date_time_value)

        if parsed_time is None:
            continue

        perth_time = parsed_time.astimezone(ZoneInfo("Australia/Perth"))

        if perth_time.strftime("%Y-%m-%d") == target_date:
            item["_parsed_date"] = perth_time
            same_day_items.append(item)

    if not same_day_items:
        return None

    same_day_items = sorted(
        same_day_items,
        key=lambda x: x.get("_parsed_date"),
        reverse=True,
    )

    return build_record_from_item(same_day_items[0])


def find_latest_timesheet_record_by_opms(opms):
    items = get_timesheet_items_by_opms(opms)

    valid_items = []

    for item in items:
        fields = item.get("fields", {})
        date_time_value = str(fields.get("Date", "")).strip()
        parsed_time = parse_sharepoint_datetime(date_time_value)

        if parsed_time is None:
            continue

        item["_parsed_date"] = parsed_time.astimezone(ZoneInfo("Australia/Perth"))
        valid_items.append(item)

    if not valid_items:
        return None

    valid_items = sorted(
        valid_items,
        key=lambda x: x.get("_parsed_date"),
        reverse=True,
    )

    return build_record_from_item(valid_items[0])


def create_timesheet_record(data, status):
    token, site_id, gap_list_id = get_gap_context()

    record_time = get_requested_record_time(data)
    date_time_value = record_time.isoformat()

    unique_key = (
        f"{str(data.get('opms', '')).strip()}-"
        f"{record_time.strftime('%Y/%m/%d %H:%M:%S')}"
    )

    # ========= 基础字段 =========
    name = data.get("name", "").strip()
    first_name = data.get("first_name", "").strip()
    last_name = data.get("last_name", "").strip()
    position = data.get("position", "").strip()
    opms = str(data.get("opms", "")).strip()
    project = (data.get("project") or data.get("site") or "").strip()
    reason = data.get("reason", "").strip()

    # ✅ 关键：Supervisor 从前端拿
    supervisor = data.get("supervisor", "").strip()

    payload = {
        "fields": {
            "Title": name,
            "FirstName": first_name,
            "LastName": last_name,
            "Position": position,
            "OPMS": opms,
            "Date": date_time_value,
            "Project": project,
            "Uniquekey": unique_key,
            "Status": status,
            "Reason": reason,
            "Supervisor": supervisor
        }
    }

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{gap_list_id}/items"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    print("Creating timesheet record:")
    print(payload)

    res = requests.post(url, headers=headers, json=payload)

    if not res.ok:
        print("Graph create error status:", res.status_code)
        print("Graph create error body:", res.text)
        print("Graph create payload:", payload)

    res.raise_for_status()
    return res.json()


def handle_sign_out(data):
    opms = data["opms"]

    current_exit_time = datetime.now(ZoneInfo("Australia/Perth"))

    # Sign Out 用全历史最新记录，不按天隔断
    latest = find_latest_timesheet_record_by_opms(opms)

    if latest is None:
        return create_timesheet_record(
            build_data_with_forced_time(data, current_exit_time),
            "Sign Out",
        )

    latest_status = str(latest.get("status", "")).strip().lower()
    latest_time = latest.get("date_time_obj")

    if latest_time is None:
        raise Exception("Latest record time is invalid. Please check the Date field.")

    latest_time_text = latest_time.strftime("%Y/%m/%d %I:%M %p")

    if latest_status == "sign in":
        if current_exit_time < latest_time:
            raise Exception(
                f"Current Sign Out time cannot be earlier than the last Sign In / Return time: {latest_time_text}."
            )

        return create_timesheet_record(
            build_data_with_forced_time(data, current_exit_time),
            "Sign Out",
        )

    if latest_status == "sign out":
        if not has_custom_time(data):
            raise Exception(
                f"The latest record is already Sign Out at {latest_time_text}. "
                f"Please enter the missing Sign In / Return time before submitting Sign Out again."
            )

        missing_return_time = get_requested_record_time(data)

        if missing_return_time < latest_time:
            raise Exception(
                f"Missing Sign In / Return time cannot be earlier than the last Sign Out time: {latest_time_text}."
            )

        if missing_return_time > current_exit_time:
            raise Exception(
                "Missing Sign In / Return time cannot be later than current Sign Out time."
            )

        create_timesheet_record(
            build_data_with_forced_time(data, missing_return_time),
            "Sign In",
        )

        return create_timesheet_record(
            build_data_with_forced_time(data, current_exit_time),
            "Sign Out",
        )

    raise Exception(f"Invalid latest status: {latest.get('status', '')}")


def handle_sign_in(data):
    opms = data["opms"]

    current_return_time = datetime.now(ZoneInfo("Australia/Perth"))

    # Sign In 可以按当天逻辑，也可以保留正常补漏逻辑
    latest = find_latest_today_timesheet_record(opms, current_return_time)

    if latest is None:
        return create_timesheet_record(
            build_data_with_forced_time(data, current_return_time),
            "Sign In",
        )

    latest_status = str(latest.get("status", "")).strip().lower()
    latest_time = latest.get("date_time_obj")

    if latest_time is None:
        raise Exception("Latest record time is invalid. Please check the Date field.")

    if latest_status == "sign out":
        if current_return_time < latest_time:
            raise Exception(
                "Sign In / Return time cannot be earlier than last Sign Out / Exit time."
            )

        return create_timesheet_record(
            build_data_with_forced_time(data, current_return_time),
            "Sign In",
        )

    if latest_status == "sign in":
        if not has_custom_time(data):
            raise Exception(
                "The latest record is already Sign In. "
                "Please enter the missing Sign Out / Exit time before submitting Sign In again."
            )

        missing_exit_time = get_requested_record_time(data)

        if missing_exit_time < latest_time:
            raise Exception(
                "Missing Sign Out / Exit time cannot be earlier than last Sign In / Return time."
            )

        if missing_exit_time > current_return_time:
            raise Exception(
                "Missing Sign Out / Exit time cannot be later than current Sign In / Return time."
            )

        create_timesheet_record(
            build_data_with_forced_time(data, missing_exit_time),
            "Sign Out",
        )

        return create_timesheet_record(
            build_data_with_forced_time(data, current_return_time),
            "Sign In",
        )

    raise Exception(f"Invalid latest status: {latest.get('status', '')}")


def print_gap_list_columns():
    token, site_id, gap_list_id = get_gap_context()

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{gap_list_id}/columns"
    headers = {"Authorization": f"Bearer {token}"}

    res = requests.get(url, headers=headers)
    res.raise_for_status()

    print("\nPPL-Timesheets columns:")
    for col in res.json().get("value", []):
        print(col.get("displayName"), "=>", col.get("name"))


if __name__ == "__main__":
    workers = get_workers_for_app()
    supervisors = get_supervisors_for_app()

    print("\nTotal workers:", len(workers))
    print(workers[:10])

    print("\nTotal supervisor candidates:", len(supervisors))
    print(supervisors[:10])

    print_gap_list_columns()