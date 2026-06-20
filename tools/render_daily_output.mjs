import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";


if (process.argv.length !== 4) {
  throw new Error("Usage: render_daily_output.mjs <workbook.xlsx> <output-dir>");
}

const workbookPath = path.resolve(process.argv[2]);
const outputDir = path.resolve(process.argv[3]);
await fs.mkdir(outputDir, { recursive: true });

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const outputs = [];
for (const [index, sheet] of workbook.worksheets.items.entries()) {
  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 0.75,
    format: "png",
  });
  const safeName = sheet.name.replaceAll("/", "_");
  const outputPath = path.join(outputDir, `${String(index + 1).padStart(2, "0")}_${safeName}.png`);
  await fs.writeFile(outputPath, new Uint8Array(await preview.arrayBuffer()));
  outputs.push(outputPath);
}

process.stdout.write(JSON.stringify(outputs, null, 2));
