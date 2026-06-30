from __future__ import annotations

import argparse
import json
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RULES = ROOT / "config" / "reading_rules.json"
DEFAULT_MAPPING = ROOT / "config" / "app_field_mapping.sample.json"
DEFAULT_NODE = (
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
    / "node"
    / "bin"
    / "node.exe"
)


@dataclass(frozen=True)
class TimecardEntry:
    source_page: int | str
    source_pdf: str
    date: str
    work_performed: str
    job_number: str
    employee_name: str
    start_time: str
    end_time: str
    total_hours: float | None
    include_in_summary: bool


@dataclass(frozen=True)
class TimecardSection:
    source_page: int | str
    date: str
    work_performed: str
    entries: list[TimecardEntry]


@dataclass(frozen=True)
class EmployeeSummary:
    employee_name: str
    on_site_entries: int
    total_on_site_hours: float


@dataclass(frozen=True)
class WeeklyTimecard:
    source_pdf: str
    output_base_name: str
    sections: list[TimecardSection]
    summary: list[EmployeeSummary]

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["model_version"] = "timecard.weekly_on_site.v1"
        return payload


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def parse_minutes(time_text: str) -> int | None:
    if not isinstance(time_text, str) or len(time_text) != 5 or time_text[2] != ":":
        return None
    try:
        hours = int(time_text[:2])
        minutes = int(time_text[3:])
    except ValueError:
        return None
    return hours * 60 + minutes


def hours_between(start: str, end: str) -> float | None:
    start_minutes = parse_minutes(start)
    end_minutes = parse_minutes(end)
    if start_minutes is None or end_minutes is None:
        return None
    if end_minutes < start_minutes:
        end_minutes += 24 * 60
    return round((end_minutes - start_minutes) / 60, 2)


def normalize_job_number(job_number: str, date_text: str, rules: dict[str, Any]) -> str:
    job_rules = rules.get("job_number_rules", {})
    expected_digits = int(job_rules.get("expected_digits", 7))
    if not job_rules.get("prefix_from_work_date_year", False):
        return job_number
    if not isinstance(job_number, str) or not job_number.isdigit() or len(job_number) != expected_digits:
        return job_number
    if not isinstance(date_text, str) or len(date_text) < 4 or not date_text[:4].isdigit():
        return job_number
    return f"{date_text[2:4]}{job_number[2:]}"


def normalize_employee_name(name: str, rules: dict[str, Any]) -> str:
    name = (name or "").strip()
    if not name or name == rules.get("placeholders", {}).get("unknown_lead"):
        return name
    lower_name = name.lower()
    for employee in rules.get("employee_names", []):
        candidates = [employee.get("canonical", ""), *employee.get("aliases", [])]
        if any(candidate.lower() == lower_name for candidate in candidates):
            return employee["canonical"]
    return name


def employee_full_name(name: str, rules: dict[str, Any]) -> str:
    normalized = normalize_employee_name(name, rules)
    if not normalized or normalized == rules.get("placeholders", {}).get("unknown_lead"):
        return normalized
    for employee in rules.get("employee_names", []):
        if employee.get("canonical") == normalized:
            last_name = employee.get("last_name", "").strip()
            return f"{normalized} {last_name}".strip()
    return normalized


def unique_people(people: list[str], rules: dict[str, Any]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for person in people:
        normalized = normalize_employee_name(person, rules)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        result.append(normalized)
    return result


def build_weekly_timecard(extraction_path: Path, rules_path: Path = DEFAULT_RULES) -> WeeklyTimecard:
    extraction = load_json(extraction_path)
    rules = load_json(rules_path)
    source_pdf = extraction["source_pdf"]
    output_base_name = Path(source_pdf).stem
    unknown_lead = rules.get("placeholders", {}).get("unknown_lead", "LEAD NAME NEEDED")

    sections: list[TimecardSection] = []
    totals: dict[str, EmployeeSummary] = {}

    for raw_section in extraction["sections"]:
        source_page = raw_section["page"]
        date_text = raw_section["date"]
        work_performed = raw_section["work_performed"]
        job_number = normalize_job_number(raw_section["job_number"], date_text, rules)
        total_hours = hours_between(raw_section["on_site_start"], raw_section["on_site_end"])
        entries: list[TimecardEntry] = []

        for person in unique_people(raw_section.get("people", []), rules):
            normalized_name = normalize_employee_name(person, rules)
            full_name = employee_full_name(person, rules)
            include_in_summary = total_hours is not None and normalized_name != unknown_lead
            entry = TimecardEntry(
                source_page=source_page,
                source_pdf=source_pdf,
                date=date_text,
                work_performed=work_performed,
                job_number=job_number,
                employee_name=full_name,
                start_time=raw_section["on_site_start"],
                end_time=raw_section["on_site_end"],
                total_hours=total_hours,
                include_in_summary=include_in_summary,
            )
            entries.append(entry)

            if include_in_summary:
                current = totals.get(full_name)
                if current is None:
                    totals[full_name] = EmployeeSummary(full_name, 1, total_hours)
                else:
                    totals[full_name] = EmployeeSummary(
                        full_name,
                        current.on_site_entries + 1,
                        round(current.total_on_site_hours + total_hours, 2),
                    )

        sections.append(
            TimecardSection(
                source_page=source_page,
                date=date_text,
                work_performed=work_performed,
                entries=entries,
            )
        )

    return WeeklyTimecard(
        source_pdf=source_pdf,
        output_base_name=output_base_name,
        sections=sections,
        summary=sorted(totals.values(), key=lambda item: item.employee_name),
    )


def render_template(template: str, values: dict[str, Any]) -> str:
    return template.format(**{key: "" if value is None else value for key, value in values.items()})


def map_entry_fields(entry: dict[str, Any], field_map: dict[str, str]) -> dict[str, Any]:
    return {target: entry.get(source) for target, source in field_map.items()}


def build_app_payload(model: WeeklyTimecard, mapping_path: Path = DEFAULT_MAPPING) -> dict[str, Any]:
    mapping = load_json(mapping_path)
    entries = [
        asdict(entry)
        for section in model.sections
        for entry in section.entries
        if entry.include_in_summary
    ]

    monday = mapping["monday"]
    vibe = mapping["vibe"]
    return {
        "source_pdf": model.source_pdf,
        "monday": [
            {
                "item_name": render_template(monday["item_name"], entry),
                "fields": map_entry_fields(entry, monday["fields"]),
            }
            for entry in entries
        ],
        "vibe": [
            {
                "record_name": render_template(vibe["record_name"], entry),
                "fields": map_entry_fields(entry, vibe["fields"]),
            }
            for entry in entries
        ],
    }


def build_ui_entry_queue(model: WeeklyTimecard) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for section in model.sections:
        for entry in section.entries:
            if not entry.include_in_summary:
                continue
            if not entry.job_number.isdigit():
                continue
            rows.append(
                {
                    "employee_name": entry.employee_name,
                    "project_number": entry.job_number,
                    "hours": entry.total_hours,
                    "date": entry.date,
                    "type": "Billable",
                    "description": entry.work_performed,
                    "source_page": entry.source_page,
                }
            )
    return rows


def save_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)
        handle.write("\n")


def run_excel_builder(structured_path: Path, node_path: Path = DEFAULT_NODE) -> None:
    command = [str(node_path), "tools/build_weekly_report.mjs", str(structured_path)]
    subprocess.run(command, cwd=ROOT, check=True)


def run_monday_upload_builder(workbook_path: Path) -> None:
    command = [sys.executable, "tools/build_ui_queue_from_excel.py", str(workbook_path)]
    subprocess.run(command, cwd=ROOT, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build structured timecard data and selected outputs.")
    parser.add_argument("extraction_json", type=Path)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--mapping", type=Path, default=DEFAULT_MAPPING)
    parser.add_argument(
        "--outputs",
        default="structured,excel,app-map,monday-upload",
        help="Comma-separated outputs: structured, excel, app-map, ui-queue, monday-upload",
    )
    args = parser.parse_args()

    model = build_weekly_timecard(args.extraction_json, args.rules)
    selected_outputs = {item.strip().lower() for item in args.outputs.split(",") if item.strip()}
    output_dir = ROOT / "outputs"
    structured_path = output_dir / f"{model.output_base_name}.structured.json"

    if selected_outputs & {"structured", "excel"}:
        save_json(structured_path, model.to_dict())

    if "excel" in selected_outputs:
        run_excel_builder(structured_path)

    if "monday-upload" in selected_outputs:
        workbook_path = output_dir / f"{model.output_base_name}.xlsx"
        if "excel" not in selected_outputs and not workbook_path.exists():
            save_json(structured_path, model.to_dict())
            run_excel_builder(structured_path)
        run_monday_upload_builder(workbook_path)

    if "app-map" in selected_outputs:
        app_payload = build_app_payload(model, args.mapping)
        save_json(output_dir / f"{model.output_base_name}.app_payload.json", app_payload)

    if "ui-queue" in selected_outputs:
        save_json(output_dir / f"{model.output_base_name}.ui_entry_queue.json", {"entries": build_ui_entry_queue(model)})

    print(json.dumps(model.to_dict(), indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
