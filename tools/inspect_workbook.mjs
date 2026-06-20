import path from "node:path";
import process from "node:process";

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";


if (process.argv.length !== 3) {
  throw new Error("Usage: inspect_workbook.mjs <workbook.xlsx>");
}

const workbookPath = path.resolve(process.argv[2]);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 10000,
  tableMaxRows: 4,
  tableMaxCols: 12,
  tableMaxCellChars: 100,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});

process.stdout.write(JSON.stringify({ summary: summary.ndjson, errors: errors.ndjson }, null, 2));
