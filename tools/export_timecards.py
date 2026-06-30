import argparse
import csv
import json
import os
import sys
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path


REQUIRED_COLUMNS = [
    "Source File",
    "Employee Name",
    "Work Date",
    "Clock In",
    "Clock Out",
    "Break Minutes",
    "Reviewer Status",
]


OUTPUT_COLUMNS = [
    "Source File",
    "Employee Name",
    "Employee ID",
    "Work Date",
    "Clock In",
    "Clock Out",
    "Break Minutes",
    "Hours Worked",
    "Job / Project",
    "Cost Code",
    "Reviewer Status",
    "Review Flags",
    "Notes",
]


def parse_time(value):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%I:%M %p", "%H:%M", "%I:%M%p"):
        try:
            return datetime.strptime(value.upper().replace(" ", ""), fmt.replace(" ", ""))
        except ValueError:
            continue
    raise ValueError(f"Invalid time: {value}")


def calculate_hours(row):
    start = parse_time(row.get("Clock In"))
    end = parse_time(row.get("Clock Out"))
    if not start or not end:
        return ""
    if end < start:
        end = end + timedelta(days=1)
    break_minutes = int(float(row.get("Break Minutes") or 0))
    minutes = (end - start).total_seconds() / 60 - break_minutes
    return round(max(minutes, 0) / 60, 2)


def read_rows(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = [col for col in REQUIRED_COLUMNS if col not in reader.fieldnames]
        if missing:
            raise ValueError(f"Missing required columns: {', '.join(missing)}")
        return list(reader)


def normalize_rows(rows):
    normalized = []
    issues = []
    for index, row in enumerate(rows, start=2):
        out = {column: (row.get(column) or "").strip() for column in OUTPUT_COLUMNS}
        row_issues = []
        for column in REQUIRED_COLUMNS:
            if not out.get(column):
                row_issues.append(f"missing {column}")
        try:
            out["Hours Worked"] = calculate_hours(out)
        except Exception as exc:
            row_issues.append(str(exc))
            out["Hours Worked"] = ""
        if out.get("Reviewer Status", "").lower() not in {"approved", "needs review", "rejected"}:
            row_issues.append("Reviewer Status should be Approved, Needs Review, or Rejected")
        if row_issues:
            out["Review Flags"] = "; ".join(filter(None, [out.get("Review Flags"), *row_issues]))
            issues.append({"csv_row": index, "employee": out.get("Employee Name"), "issues": row_issues})
        normalized.append(out)
    return normalized, issues


def write_csv(path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def monday_column_values(row, mapping):
    values = {}
    for monday_column_id, source_column in mapping.get("columns", {}).items():
        if source_column in row and row[source_column] != "":
            values[monday_column_id] = row[source_column]
    return values


def upload_to_monday(rows, mapping):
    token = os.environ.get("MONDAY_API_TOKEN")
    board_id = os.environ.get("MONDAY_BOARD_ID")
    if not token or not board_id:
        raise ValueError("Set MONDAY_API_TOKEN and MONDAY_BOARD_ID before uploading.")

    group_id = os.environ.get("MONDAY_GROUP_ID") or mapping.get("group_id") or None
    endpoint = "https://api.monday.com/v2"
    headers = {"Authorization": token, "Content-Type": "application/json"}
    mutation = """
    mutation ($boardId: ID!, $groupId: String, $itemName: String!, $columnValues: JSON!) {
      create_item(board_id: $boardId, group_id: $groupId, item_name: $itemName, column_values: $columnValues) {
        id
      }
    }
    """

    created = []
    template = mapping.get("item_name_template", "{Employee Name} - {Work Date}")
    for row in rows:
        if row.get("Reviewer Status", "").lower() != "approved":
            continue
        item_name = template.format(**row)
        variables = {
            "boardId": str(board_id),
            "groupId": group_id,
            "itemName": item_name,
            "columnValues": json.dumps(monday_column_values(row, mapping)),
        }
        request = urllib.request.Request(
            endpoint,
            data=json.dumps({"query": mutation, "variables": variables}).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
        if "errors" in payload:
            raise RuntimeError(payload["errors"])
        created.append(payload["data"]["create_item"]["id"])
    return created


def main():
    parser = argparse.ArgumentParser(description="Validate reviewed timecards and export payroll-ready CSV.")
    parser.add_argument("input_csv", help="Reviewed timecard CSV exported from the Excel review workbook.")
    parser.add_argument("--output-csv", default="outputs/timecard_summary.csv")
    parser.add_argument("--monday-mapping", default="config/monday_mapping.sample.json")
    parser.add_argument("--upload-monday", action="store_true")
    args = parser.parse_args()

    rows = read_rows(args.input_csv)
    normalized, issues = normalize_rows(rows)
    write_csv(Path(args.output_csv), normalized)

    print(f"Wrote {args.output_csv} with {len(normalized)} rows.")
    if issues:
        print(f"Review warnings: {len(issues)} row(s) need attention.")
        for issue in issues[:10]:
            print(f"Row {issue['csv_row']}: {issue['employee'] or 'Unknown'} - {', '.join(issue['issues'])}")

    if args.upload_monday:
        with open(args.monday_mapping, "r", encoding="utf-8") as handle:
            mapping = json.load(handle)
        created = upload_to_monday(normalized, mapping)
        print(f"Created {len(created)} Monday.com item(s).")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
