import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const extractionPath = process.argv[2] ?? "extractions/Week 22.json";
const outputDir = "outputs";
await fs.mkdir(outputDir, { recursive: true });

const readingRules = JSON.parse(await fs.readFile("config/reading_rules.json", "utf8"));
const extraction = JSON.parse(await fs.readFile(extractionPath, "utf8"));
const outputBaseName = extraction.source_pdf
  .split("/")
  .at(-1)
  .replace(/\.[^.]+$/, "");
const isStructuredModel = extraction.model_version === "timecard.weekly_on_site.v1";

function parseMinutes(timeText) {
  if (!/^\d{2}:\d{2}$/.test(timeText)) {
    return null;
  }
  const [hours, minutes] = timeText.split(":").map(Number);
  return hours * 60 + minutes;
}

function hoursBetween(start, end) {
  const startMinutes = parseMinutes(start);
  let endMinutes = parseMinutes(end);
  if (startMinutes === null || endMinutes === null) {
    return "";
  }
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

function uniquePeople(people) {
  const seen = new Set();
  const result = [];
  for (const person of people) {
    const normalized = normalizeEmployeeName(person);
    if (!normalized || seen.has(normalized)) {
      continue;
    }
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

const sections = isStructuredModel
  ? extraction.sections.map((section) => ({
      date: section.date,
      work_performed: section.work_performed,
      entries: section.entries.map((entry) => ({
        job_number: entry.job_number,
        employee_name: entry.employee_name,
        start_time: entry.start_time,
        end_time: entry.end_time,
        hours: entry.total_hours ?? "",
        include_in_summary: entry.include_in_summary,
        prevailing_wage: Boolean(entry.prevailing_wage),
      })),
    }))
  : extraction.sections.map((section) => {
      const jobNumber = normalizeJobNumber(section.job_number, section.date);
      const people = uniquePeople(section.people);
      const hours = hoursBetween(section.on_site_start, section.on_site_end);
      const prevailingWage = Boolean(section.prevailing_wage || section.pw || section.PW);
      return {
        date: section.date,
        work_performed: section.work_performed,
        entries: people.map((person) => ({
          job_number: jobNumber,
          employee_name: person,
          start_time: section.on_site_start,
          end_time: section.on_site_end,
          hours,
          include_in_summary: hours !== "" && person !== readingRules.placeholders?.unknown_lead,
          prevailing_wage: prevailingWage,
        })),
      };
    });

const employeeTotals = new Map();
for (const section of sections) {
  for (const entry of section.entries) {
    if (!entry.include_in_summary || entry.hours === "") {
      continue;
    }
    const person = entry.employee_name;
    const current = employeeTotals.get(person) ?? { name: person, entries: 0, totalHours: 0 };
    current.entries += 1;
    current.totalHours = Math.round((current.totalHours + entry.hours) * 100) / 100;
    employeeTotals.set(person, current);
  }
}

const workbook = Workbook.create();
const weekly = workbook.worksheets.add("Weekly On Site");
const summary = workbook.worksheets.add("Summary");

for (const sheet of [weekly, summary]) {
  sheet.showGridLines = false;
}

weekly.getRange("A1:E1").merge();
weekly.getRange("A1").values = [[`${outputBaseName} - Weekly On Site Time`]];
weekly.getRange("A1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};

let row = 3;
for (const section of sections) {
  const sectionTitle = `${section.date} / ${section.work_performed}`;
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

  const rows = section.entries.length
    ? section.entries.map((entry) => [
        entry.job_number,
        entry.employee_name,
        entry.start_time,
        entry.end_time,
        entry.hours,
      ])
    : [[
        "JOB# NEEDED",
        readingRules.placeholders?.unknown_lead ?? "LEAD NAME NEEDED",
        "ON SITE MISSING",
        "ON SITE MISSING",
        "",
      ]];
  weekly.getRange(`A${row}:E${row + rows.length - 1}`).values = rows;
  weekly.getRange(`A${row}:E${row}`).format.fill = "#EFF6FF";
  weekly.getRange(`A${row}:E${row + rows.length - 1}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#B7C9D6",
  };
  section.entries.forEach((entry, index) => {
    if (entry.prevailing_wage) {
      weekly.getRange(`A${row + index}`).format.font = {
        bold: true,
        color: "#C00000",
      };
    }
  });
  weekly.getRange(`E${row}:E${row + rows.length - 1}`).format.numberFormat = "0.00";
  row += rows.length + 2;
}

weekly.getRange("A:A").format.columnWidthPx = 110;
weekly.getRange("B:B").format.columnWidthPx = 190;
weekly.getRange("C:D").format.columnWidthPx = 120;
weekly.getRange("E:E").format.columnWidthPx = 115;
weekly.freezePanes.freezeRows(1);

summary.getRange("A1:C1").merge();
summary.getRange("A1").values = [[`${outputBaseName} - Weekly On Site Summary`]];
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

if (summaryRows.length) {
  summary.getRange(`A4:C${summaryRows.length + 3}`).values = summaryRows;
  summary.getRange(`A3:C${summaryRows.length + 3}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#B7C9D6",
  };
  summary.getRange(`C4:C${summaryRows.length + 3}`).format.numberFormat = "0.00";
}
summary.getRange("A:A").format.columnWidthPx = 190;
summary.getRange("B:B").format.columnWidthPx = 120;
summary.getRange("C:C").format.columnWidthPx = 145;
summary.freezePanes.freezeRows(3);

await workbook.inspect({
  kind: "table",
  range: "'Weekly On Site'!A1:E120",
  include: "values,formulas",
  maxChars: 12000,
});
await workbook.inspect({
  kind: "table",
  range: "Summary!A1:C40",
  include: "values,formulas",
  maxChars: 6000,
});
await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
});
await workbook.render({ sheetName: "Weekly On Site", range: "A1:E60", scale: 1, format: "png" });
await workbook.render({ sheetName: "Summary", autoCrop: "all", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/${outputBaseName}.xlsx`);
