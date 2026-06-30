import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
await fs.mkdir(outputDir, { recursive: true });
const sourcePdf = "incoming_timesheets/time sheets - week 1.pdf";
const outputBaseName = sourcePdf
  .split("/")
  .at(-1)
  .replace(/\.[^.]+$/, "");
const readingRules = JSON.parse(await fs.readFile("config/reading_rules.json", "utf8"));

// One timesheet page becomes one dated section. Only the "On Site" time is
// reported for the lead and each assistant assigned to that page.
const sections = [
  {
    date: "DATE NEEDS REVIEW",
    workPerformed: "Stretch / Finish",
    jobNumber: "2600206",
    onSiteStart: "08:36",
    onSiteEnd: "16:30",
    people: [
      { name: "LEAD NAME NEEDED", role: "Lead" },
      { name: "Marcelino", role: "Assistant" },
    ],
  },
  {
    date: "2026-06-11",
    workPerformed: "Finish",
    jobNumber: "3600106",
    onSiteStart: "09:30",
    onSiteEnd: "10:30",
    people: [
      { name: "Eric", role: "Lead" },
      { name: "Marcelino", role: "Assistant" },
    ],
  },
  {
    date: "2026-06-09",
    workPerformed: "Stretch Finish",
    jobNumber: "3600227",
    onSiteStart: "09:30",
    onSiteEnd: "15:00",
    people: [
      { name: "Zach", role: "Lead" },
      { name: "Felipe", role: "Assistant" },
      { name: "Jose", role: "Assistant" },
      { name: "Frank", role: "Assistant" },
    ],
  },
];

function parseMinutes(timeText) {
  const [hours, minutes] = timeText.split(":").map(Number);
  return hours * 60 + minutes;
}

function hoursBetween(start, end) {
  let startMinutes = parseMinutes(start);
  let endMinutes = parseMinutes(end);
  if (endMinutes < startMinutes) {
    endMinutes += 24 * 60;
  }
  return Math.round(((endMinutes - startMinutes) / 60) * 100) / 100;
}

function normalizeJobNumber(jobNumber, dateText) {
  const expectedDigits = readingRules.job_number_rules?.expected_digits ?? 7;
  const shouldUseYear = readingRules.job_number_rules?.prefix_from_work_date_year ?? false;
  if (
    !shouldUseYear ||
    !new RegExp(`^\\d{${expectedDigits}}$`).test(jobNumber) ||
    !/^\d{4}-/.test(dateText)
  ) {
    return jobNumber;
  }
  const yearPrefix = dateText.slice(2, 4);
  return `${yearPrefix}${jobNumber.slice(2)}`;
}

function normalizeEmployeeName(rawName) {
  const name = (rawName || "").trim();
  if (!name || name === readingRules.placeholders?.unknown_lead) {
    return name;
  }
  const lowerName = name.toLowerCase();
  const match = readingRules.employee_names?.find((employee) => {
    const candidates = [employee.canonical, ...(employee.aliases || [])];
    return candidates.some((candidate) => candidate.toLowerCase() === lowerName);
  });
  return match?.canonical ?? name;
}

const employeeTotals = new Map();
for (const section of sections) {
  const hours = hoursBetween(section.onSiteStart, section.onSiteEnd);
  for (const person of section.people) {
    const employeeName = normalizeEmployeeName(person.name);
    const current = employeeTotals.get(employeeName) ?? { name: employeeName, entries: 0, totalHours: 0 };
    current.entries += 1;
    current.totalHours = Math.round((current.totalHours + hours) * 100) / 100;
    employeeTotals.set(employeeName, current);
  }
}

const workbook = Workbook.create();
const weekly = workbook.worksheets.add("Weekly On Site");
const summary = workbook.worksheets.add("Summary");

for (const sheet of [weekly, summary]) {
  sheet.showGridLines = false;
}

weekly.getRange("A1:E1").merge();
weekly.getRange("A1").values = [["Weekly On Site Time"]];
weekly.getRange("A1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};

let row = 3;
for (const section of sections) {
  const sectionTitle = `${section.date} / ${section.workPerformed}`;
  weekly.getRange(`A${row}:E${row}`).merge();
  weekly.getRange(`A${row}`).values = [[sectionTitle]];
  weekly.getRange(`A${row}`).format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#1F4E79", size: 13 },
  };
  row += 1;

  weekly.getRange(`A${row}:E${row}`).values = [["Job#", "Employee Name", "Start Time", "End Time", "Total Hours"]];
  weekly.getRange(`A${row}:E${row}`).format = {
    fill: "#245B45",
    font: { bold: true, color: "#FFFFFF" },
  };
  row += 1;

  const hours = hoursBetween(section.onSiteStart, section.onSiteEnd);
  const jobNumber = normalizeJobNumber(section.jobNumber, section.date);
  const rows = section.people.map((person) => [
    jobNumber,
    normalizeEmployeeName(person.name),
    section.onSiteStart,
    section.onSiteEnd,
    hours,
  ]);
  weekly.getRange(`A${row}:E${row + rows.length - 1}`).values = rows;

  // Lead is always listed first; give the first row a subtle tint so field crews
  // can see the intended order when reviewing the report.
  weekly.getRange(`A${row}:E${row}`).format.fill = "#EFF6FF";
  weekly.getRange(`A${row}:E${row + rows.length - 1}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#B7C9D6",
  };
  weekly.getRange(`E${row}:E${row + rows.length - 1}`).format.numberFormat = "0.00";
  row += rows.length + 2;
}

weekly.getRange("A:A").format.columnWidthPx = 110;
weekly.getRange("B:B").format.columnWidthPx = 190;
weekly.getRange("C:D").format.columnWidthPx = 110;
weekly.getRange("E:E").format.columnWidthPx = 115;
weekly.freezePanes.freezeRows(1);

summary.getRange("A1:C1").merge();
summary.getRange("A1").values = [["Weekly On Site Summary"]];
summary.getRange("A1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};
summary.getRange("A3:C3").values = [["Employee Name", "On Site Entries", "Total On Site Hours"]];
summary.getRange("A3:C3").format = {
  fill: "#245B45",
  font: { bold: true, color: "#FFFFFF" },
};

const summaryRows = [...employeeTotals.values()]
  .sort((a, b) => a.name.localeCompare(b.name))
  .map((item) => [item.name, item.entries, item.totalHours]);

summary.getRange(`A4:C${summaryRows.length + 3}`).values = summaryRows;
summary.getRange(`A3:C${summaryRows.length + 3}`).format.borders = {
  preset: "all",
  style: "thin",
  color: "#B7C9D6",
};
summary.getRange(`C4:C${summaryRows.length + 3}`).format.numberFormat = "0.00";
summary.getRange("A:A").format.columnWidthPx = 190;
summary.getRange("B:B").format.columnWidthPx = 120;
summary.getRange("C:C").format.columnWidthPx = 145;
summary.freezePanes.freezeRows(3);

await workbook.inspect({
  kind: "table",
  range: "'Weekly On Site'!A1:E25",
  include: "values,formulas",
  maxChars: 5000,
});
await workbook.inspect({
  kind: "table",
  range: "Summary!A1:C20",
  include: "values,formulas",
  maxChars: 3000,
});
await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
});
await workbook.render({ sheetName: "Weekly On Site", range: "A1:E25", scale: 1, format: "png" });
await workbook.render({ sheetName: "Summary", autoCrop: "all", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/${outputBaseName}.xlsx`);
