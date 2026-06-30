import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const instructions = workbook.worksheets.add("Start Here");
const review = workbook.worksheets.add("Review Entries");
const issues = workbook.worksheets.add("Issue Summary");
const mapping = workbook.worksheets.add("Monday Mapping");

for (const sheet of [instructions, review, issues, mapping]) {
  sheet.showGridLines = false;
}

instructions.getRange("A1:F1").merge();
instructions.getRange("A1").values = [["Handwritten Timecard Review"]];
instructions.getRange("A1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF", size: 18 },
};
instructions.getRange("A3:B9").values = [
  ["Step", "Action"],
  ["1", "Put scans/photos in incoming_timesheets."],
  ["2", "Enter or paste extracted entries into Review Entries."],
  ["3", "Set Reviewer Status to Approved, Needs Review, or Rejected."],
  ["4", "Resolve rows with Review Flags before payroll export."],
  ["5", "Save Review Entries as CSV for tools/export_timecards.py."],
  ["6", "Configure Monday Mapping only if uploading to Monday.com."],
];
instructions.getRange("A3:B3").format = { fill: "#D9EAF7", font: { bold: true } };
instructions.getRange("A3:B9").format.borders = { preset: "all", style: "thin", color: "#B7C9D6" };
instructions.getRange("A:A").format.columnWidthPx = 80;
instructions.getRange("B:B").format.columnWidthPx = 520;
instructions.getRange("B4:B9").format.wrapText = true;

const headers = [
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
  "Notes",
  "Reviewer Status",
  "Review Flags",
];

review.getRange("A1:M1").values = [headers];
review.getRange("A1:M1").format = {
  fill: "#245B45",
  font: { bold: true, color: "#FFFFFF" },
};

const sampleRows = [
  ["sample_timesheet_001.jpg", "Jane Smith", "1042", new Date("2026-06-15"), "7:00 AM", "3:30 PM", 30, null, "Fence Install", "INSTALL", "", "Needs Review", "Sample row - replace"],
  ["sample_timesheet_002.jpg", "Luis Garcia", "1088", new Date("2026-06-15"), "6:45 AM", "2:45 PM", 30, null, "Warehouse", "SHOP", "", "Approved", ""],
];
review.getRange("A2:M3").values = sampleRows;
review.getRange("H2").formulas = [["=IF(OR(E2=\"\",F2=\"\"),\"\",ROUND(((F2-E2)*24)-(G2/60),2))"]];
review.getRange("H2:H200").fillDown();
review.getRange("D2:D200").format.numberFormat = "yyyy-mm-dd";
review.getRange("E2:F200").format.numberFormat = "h:mm AM/PM";
review.getRange("G2:H200").format.numberFormat = "0.00";
review.getRange("A1:M200").format.borders = { preset: "all", style: "thin", color: "#D9E2EA" };
review.getRange("K2:M200").format.wrapText = true;
review.getRange("A:A").format.columnWidthPx = 190;
review.getRange("B:B").format.columnWidthPx = 160;
review.getRange("C:C").format.columnWidthPx = 105;
review.getRange("D:F").format.columnWidthPx = 105;
review.getRange("G:H").format.columnWidthPx = 120;
review.getRange("I:J").format.columnWidthPx = 145;
review.getRange("K:K").format.columnWidthPx = 220;
review.getRange("L:L").format.columnWidthPx = 135;
review.getRange("M:M").format.columnWidthPx = 240;
review.freezePanes.freezeRows(1);
review.getRange("L2:L200").dataValidation = {
  rule: { type: "list", values: ["Approved", "Needs Review", "Rejected"] },
};
review.getRange("L2:L200").conditionalFormats.add("containsText", {
  text: "Needs Review",
  format: { fill: "#FFF2CC", font: { color: "#7A4F00" } },
});
review.getRange("L2:L200").conditionalFormats.add("containsText", {
  text: "Rejected",
  format: { fill: "#F4CCCC", font: { color: "#990000" } },
});
review.getRange("L2:L200").conditionalFormats.add("containsText", {
  text: "Approved",
  format: { fill: "#D9EAD3", font: { color: "#274E13" } },
});

issues.getRange("A1:D1").merge();
issues.getRange("A1").values = [["Review Summary"]];
issues.getRange("A1").format = {
  fill: "#5B3F8C",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};
issues.getRange("A3:B7").values = [
  ["Metric", "Value"],
  ["Total Entries", null],
  ["Approved", null],
  ["Needs Review", null],
  ["Rejected", null],
];
issues.getRange("B4").formulas = [["=COUNTA('Review Entries'!B2:B200)"]];
issues.getRange("B5").formulas = [["=COUNTIF('Review Entries'!L2:L200,\"Approved\")"]];
issues.getRange("B6").formulas = [["=COUNTIF('Review Entries'!L2:L200,\"Needs Review\")"]];
issues.getRange("B7").formulas = [["=COUNTIF('Review Entries'!L2:L200,\"Rejected\")"]];
issues.getRange("D3:E3").values = [["Status", "Count"]];
issues.getRange("D4:E6").formulas = [
  ["=\"Approved\"", "=B5"],
  ["=\"Needs Review\"", "=B6"],
  ["=\"Rejected\"", "=B7"],
];
issues.getRange("A3:B7").format.borders = { preset: "all", style: "thin", color: "#D9D2E9" };
issues.getRange("D3:E6").format.borders = { preset: "all", style: "thin", color: "#D9D2E9" };
issues.getRange("A3:B3").format = { fill: "#EADCF8", font: { bold: true } };
issues.getRange("D3:E3").format = { fill: "#EADCF8", font: { bold: true } };
issues.getRange("A:A").format.columnWidthPx = 160;
issues.getRange("B:B").format.columnWidthPx = 120;
issues.getRange("D:E").format.columnWidthPx = 140;
const statusChart = issues.charts.add("bar", issues.getRange("D3:E6"));
statusChart.title = "Entries by Review Status";
statusChart.hasLegend = false;
statusChart.setPosition("G3", "M18");

mapping.getRange("A1:C1").values = [["Monday Column ID", "Source Column", "Notes"]];
mapping.getRange("A1:C1").format = {
  fill: "#604A7B",
  font: { bold: true, color: "#FFFFFF" },
};
mapping.getRange("A2:C14").values = [
  ["employee_name", "Employee Name", "Replace with the actual Monday column ID."],
  ["employee_id", "Employee ID", ""],
  ["work_date", "Work Date", ""],
  ["clock_in", "Clock In", ""],
  ["clock_out", "Clock Out", ""],
  ["break_minutes", "Break Minutes", ""],
  ["hours_worked", "Hours Worked", ""],
  ["job_project", "Job / Project", ""],
  ["cost_code", "Cost Code", ""],
  ["reviewer_status", "Reviewer Status", ""],
  ["review_flags", "Review Flags", ""],
  ["notes", "Notes", ""],
  ["source_file", "Source File", ""],
];
mapping.getRange("A1:C14").format.borders = { preset: "all", style: "thin", color: "#D9D2E9" };
mapping.getRange("A:C").format.columnWidthPx = 200;
mapping.getRange("C:C").format.columnWidthPx = 330;
mapping.getRange("C2:C14").format.wrapText = true;

await workbook.inspect({
  kind: "table",
  range: "'Review Entries'!A1:M5",
  include: "values,formulas",
  maxChars: 2000,
});
await workbook.render({ sheetName: "Start Here", autoCrop: "all", scale: 1, format: "png" });
await workbook.render({ sheetName: "Review Entries", range: "A1:M12", scale: 1, format: "png" });
await workbook.render({ sheetName: "Issue Summary", autoCrop: "all", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/timecard_review_template.xlsx`);
