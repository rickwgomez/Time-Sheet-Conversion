# Handwritten Timecard Review Workflow

This folder is set up to receive scanned or photographed handwritten timesheets, review the extracted entries, and export a clean weekly Excel report.

## Folder Layout

- `incoming_timesheets/` - put scanned timesheets here as PDF, JPG, PNG, or HEIC files.
- `outputs/` - generated Excel reports are saved here.
- `Monday Output/` - generated upload-ready JSON and CSV files for Monday/Vibe are saved here.
- `tools/` - helper scripts for building the weekly report.

## Current Workflow

1. Put timesheet scans in `incoming_timesheets/`.
2. Extract each sheet into dated sections.
3. Use only the `On Site` start and end time for the lead and assistants.
4. Generate one Excel workbook with a weekly detail sheet and a `Summary` sheet.
5. Generate one Vibe/Monday importer CSV from the workbook's `Weekly On Site` tab.

The output workbook uses the same base name as the incoming file. For example,
`incoming_timesheets/time sheets - week 1.pdf` produces
`outputs/time sheets - week 1.xlsx`.

## Review Columns

The weekly detail columns are:

- `Job#`
- `Employee Name`
- `Start Time`
- `End Time`
- `Total Hours`

Each dated section starts with `Date / Work Performed`, then lists the lead first followed by assistants.

## Reading Rules

`config/reading_rules.json` stores the rules and known names used while reading handwritten sheets. Add confirmed employee names, common misspellings, and handwriting notes there so future weekly reports can apply them automatically.

Current automatic rules include:

- 7-digit job numbers are normalized to the calendar year prefix when the sheet date is known.
- Known employee aliases are converted to their canonical names.
- Weekly totals use only `On Site` time.

Weekly extraction files live in `extractions/`. Run the complete weekly pipeline with:

`python tools/timecard_pipeline.py "extractions/Week 22.json"`

That single command is the standard batch step for weekly timesheets. It creates:

- `structured` writes `outputs/<week>.structured.json`
- `excel` writes `outputs/<week>.xlsx`
- `app-map` writes `outputs/<week>.app_payload.json` for Monday.com and Vibe-style field mapping
- `monday-upload` writes `Monday Output/<week>.upload.csv` and `Monday Output/<week>.ui_entry_queue.json`

The default run creates both export sets: the existing review/export files in `outputs/`, and the Vibe/Monday importer files in `Monday Output/`.

Use `node tools/build_weekly_report.mjs "extractions/Week 22.json"` only when intentionally rebuilding the Excel workbook by itself.

## OCR Note

Handwriting OCR quality depends heavily on the scan quality and the recognition engine. For faint scans, unclear dates, names, or job numbers should be confirmed by a person before the weekly report is finalized.
