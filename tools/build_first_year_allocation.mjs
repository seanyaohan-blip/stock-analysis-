import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(root, "outputs", "first_year_optimization");
const outputPath = path.join(outputDir, "1st年配置表_优化版.xlsx");

function parseSimpleCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines[0].replace(/^\uFEFF/, "").split(",");
  return lines.slice(1).filter(Boolean).map((line) => {
    const values = line.split(",");
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

function parseAccountMeta(text) {
  const first = text.split(/\r?\n/, 1)[0].replace(/^#\s*/, "");
  const parts = first.split(",");
  const result = {};
  for (let i = 0; i + 1 < parts.length; i += 2) result[parts[i].trim()] = Number(parts[i + 1]);
  return result;
}

const allocationRows = parseSimpleCsv(await fs.readFile(path.join(root, "first_year_allocation.csv"), "utf8"));
const positionsText = await fs.readFile(path.join(root, "positions.csv"), "utf8");
const accountMeta = parseAccountMeta(positionsText);
const positions = parseSimpleCsv(positionsText.split(/\r?\n/).filter((line) => !line.startsWith("#")).join("\n"));
const positionByCode = new Map(positions.map((row) => [String(row.Code).padStart(6, "0"), row]));

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Summary");
const allocation = workbook.worksheets.add("Allocation");
const mapping = workbook.worksheets.add("Mapping");
const rules = workbook.worksheets.add("Rules");
const sources = workbook.worksheets.add("Sources_Audit");

const navy = "#1F4E78";
const paleBlue = "#D9EAF7";
const paleGreen = "#E2F0D9";
const paleYellow = "#FFF2CC";
const paleRed = "#FCE4D6";
const gray = "#F2F4F7";

function styleHeader(range) {
  range.format.fill = navy;
  range.format.font = { bold: true, color: "#FFFFFF" };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
}

function setWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}:${column}`).format.columnWidth = width;
}

summary.showGridLines = false;
summary.getRange("A1:H1").merge();
summary.getRange("A1").values = [["第一年度全资产配置总览"]];
summary.getRange("A1:H1").format.fill = navy;
summary.getRange("A1:H1").format.font = { bold: true, color: "#FFFFFF", size: 18 };
summary.getRange("A1:H1").format.rowHeight = 30;
summary.getRange("A3:B10").values = [
  ["指标", "数值"],
  ["当前全资产基准（元）", accountMeta.total_assets || 0],
  ["第一年目标配置比例", null],
  ["第一年动态目标金额（元）", null],
  ["当前已映射金额（元）", null],
  ["年度资金缺口（元）", null],
  ["年度完成率", null],
  ["未配置目标比例", null],
];
styleHeader(summary.getRange("A3:B3"));
summary.getRange("B4").format.font = { color: "#0000FF" };
summary.getRange("B5:B10").formulas = [
  ["=SUM(Allocation!D2:D12)"],
  ["=SUM(Allocation!G2:G12)"],
  ["=SUM(Allocation!H2:H12)"],
  ["=B6-B7"],
  ["=IF(B6=0,0,B7/B6)"],
  ["=MAX(1-B5,0)"],
];
summary.getRange("B5:B7").format.font = { color: "#008000" };
summary.getRange("B8:B10").format.font = { color: "#000000" };
summary.getRange("B4:B4").format.numberFormat = "#,##0;[Red](#,##0);-";
summary.getRange("B5:B5").format.numberFormat = "0.0%";
summary.getRange("B6:B8").format.numberFormat = "#,##0;[Red](#,##0);-";
summary.getRange("B9:B10").format.numberFormat = "0.0%";
summary.getRange("A12:H15").merge();
summary.getRange("A12").values = [["口径说明：第一年配置使用全资产口径；证券账户目标继续作为集中度硬约束。年度资金缺口只代表长期规划空间，不构成当日买入信号。目标收益为观察区间，不是承诺收益。"]];
summary.getRange("A12:H15").format.fill = paleYellow;
summary.getRange("A12:H15").format.wrapText = true;
summary.getRange("A12:H15").format.verticalAlignment = "center";
setWidths(summary, { A: 28, B: 18, C: 3, D: 14, E: 14, F: 14, G: 14, H: 14 });

const allocationHeaders = [
  "配置键", "配置项", "资产层级", "年度目标占比", "映射代码", "配置表目标金额", "动态目标金额",
  "当前金额", "配置表当前金额约", "年度资金缺口", "年度完成率", "目标收益下限", "目标收益上限",
  "配置状态", "进度状态", "执行约束", "来源",
];
allocation.getRange(`A1:Q1`).values = [allocationHeaders];
styleHeader(allocation.getRange("A1:Q1"));
const staticValues = allocationRows.map((row) => [
  row.AllocationKey, row.Label, row.AssetLayer, Number(row.TargetWeight), row.Codes || "待选",
  Number(row.SourceTargetAmount), null, null, Number(row.SourceCurrentAmount), null, null,
  Number(row.TargetReturnLow), Number(row.TargetReturnHigh), row.Status, null, row.ExecutionConstraint, row.Source,
]);
allocation.getRange(`A2:Q${allocationRows.length + 1}`).values = staticValues;
for (let row = 2; row <= allocationRows.length + 1; row += 1) {
  allocation.getRange(`G${row}`).formulas = [[`=Summary!$B$4*D${row}`]];
  allocation.getRange(`H${row}`).formulas = [[`=SUMIF(Mapping!$A$2:$A$30,A${row},Mapping!$D$2:$D$30)`]];
  allocation.getRange(`J${row}`).formulas = [[`=G${row}-H${row}`]];
  allocation.getRange(`K${row}`).formulas = [[`=IF(G${row}=0,0,H${row}/G${row})`]];
  allocation.getRange(`O${row}`).formulas = [[`=IF(E${row}="待选","待选择合格标的",IF(J${row}<=0,"达到或超过年度目标",IF(K${row}>=80%,"接近年度目标","配置中")))`]];
}
allocation.getRange(`D2:D${allocationRows.length + 1}`).format.numberFormat = "0.0%";
allocation.getRange(`F2:J${allocationRows.length + 1}`).format.numberFormat = "#,##0;[Red](#,##0);-";
allocation.getRange(`K2:M${allocationRows.length + 1}`).format.numberFormat = "0.0%";
allocation.getRange(`A2:F${allocationRows.length + 1}`).format.font = { color: "#0000FF" };
allocation.getRange(`G2:H${allocationRows.length + 1}`).format.font = { color: "#008000" };
allocation.getRange(`J2:O${allocationRows.length + 1}`).format.font = { color: "#000000" };
allocation.freezePanes.freezeRows(1);
allocation.showGridLines = false;
allocation.tables.add(`A1:Q${allocationRows.length + 1}`, true, "FirstYearAllocationTable");
setWidths(allocation, { A: 18, B: 28, C: 18, D: 13, E: 23, F: 16, G: 16, H: 14, I: 16, J: 16, K: 13, L: 13, M: 13, N: 18, O: 18, P: 46, Q: 20 });
allocation.getRange(`B2:Q${allocationRows.length + 1}`).format.wrapText = true;

const mappingRows = [];
for (const row of allocationRows) {
  const codes = (row.Codes || "").split("|").filter(Boolean);
  for (const code of codes) {
    const position = positionByCode.get(code) || {};
    mappingRows.push([row.AllocationKey, code, position.Name || "未持有", Number(position["Market Value"] || 0), row.Label]);
  }
}
mapping.getRange("A1:E1").values = [["配置键", "代码", "名称", "当前市值", "配置项"]];
styleHeader(mapping.getRange("A1:E1"));
if (mappingRows.length) mapping.getRange(`A2:E${mappingRows.length + 1}`).values = mappingRows;
mapping.getRange(`D2:D${Math.max(mappingRows.length + 1, 2)}`).format.numberFormat = "#,##0;[Red](#,##0);-";
mapping.getRange(`A2:E${Math.max(mappingRows.length + 1, 2)}`).format.font = { color: "#0000FF" };
mapping.freezePanes.freezeRows(1);
mapping.showGridLines = false;
mapping.tables.add(`A1:E${mappingRows.length + 1}`, true, "AllocationMappingTable");
setWidths(mapping, { A: 20, B: 12, C: 22, D: 16, E: 30 });

rules.getRange("A1:D1").values = [["规则模块", "规则", "定义", "执行含义"]];
styleHeader(rules.getRange("A1:D1"));
const ruleRows = [
  ["双口径", "第一年配置", "全资产目标占比与金额", "决定长期规划上限"],
  ["双口径", "证券账户目标", "单只持仓集中度目标", "达到/超过时仍可一票否决新增"],
  ["年度配置", "资金缺口", "动态目标金额减当前映射金额", "不是买入信号"],
  ["质量准入", "完整评分", "ETF/个股需完整10分评分与证据", "缺失时禁止新增"],
  ["买点过滤", "标准首批", "六项通过至少5项且无否决", "仅进入复核"],
  ["情绪纪律", "4级及以上", "火热/狂热", "暂停新增"],
  ["执行纪律", "单日上限", "一个买入方向+两个卖出/减仓方向", "系统性风险日只做风控"],
];
rules.getRange(`A2:D${ruleRows.length + 1}`).values = ruleRows;
rules.getRange(`A2:D${ruleRows.length + 1}`).format.wrapText = true;
rules.freezePanes.freezeRows(1);
rules.showGridLines = false;
rules.tables.add(`A1:D${ruleRows.length + 1}`, true, "AllocationRulesTable");
setWidths(rules, { A: 18, B: 20, C: 36, D: 44 });

sources.getRange("A1:F1").values = [["数据项", "值/范围", "单位", "截至日期", "来源", "备注"]];
styleHeader(sources.getRange("A1:F1"));
const sourceRows = [
  ["当前全资产总额", accountMeta.total_assets || 0, "元", "持仓截图日期", "positions.csv", "用于动态重算年度目标金额"],
  ["证券账户总市值", accountMeta.broker_market_value || 0, "元", "持仓截图日期", "positions.csv", "用于持仓勾稽"],
  ["第一年配置比例", "58%", "%", "来源未标注", "1st年配置表.xlsx", "其余42%需保留为未配置/现金/防御空间"],
  ["目标收益区间", "5%-35%", "%", "来源未标注", "1st年配置表.xlsx", "观察区间，不是收益承诺"],
  ["映射与约束", "11项配置", "项", "本次优化", "first_year_allocation.csv", "维护代码映射与执行约束"],
];
sources.getRange(`A2:F${sourceRows.length + 1}`).values = sourceRows;
sources.getRange(`A2:F${sourceRows.length + 1}`).format.wrapText = true;
sources.showGridLines = false;
sources.tables.add(`A1:F${sourceRows.length + 1}`, true, "AllocationSourcesTable");
setWidths(sources, { A: 24, B: 18, C: 12, D: 16, E: 28, F: 48 });

sources.getRange("H1:J12").values = [["配置项", "当前金额", "动态目标"], ...allocationRows.map((row) => [row.Label, null, null])];
for (let row = 19; row <= 29; row += 1) {
  const sourceRow = row - 17;
  const auditRow = row - 18;
  sources.getRange(`I${auditRow + 1}`).formulas = [[`=Allocation!H${sourceRow}`]];
  sources.getRange(`J${auditRow + 1}`).formulas = [[`=Allocation!G${sourceRow}`]];
}
styleHeader(sources.getRange("H1:J1"));
sources.getRange("I2:J12").format.numberFormat = "#,##0";
setWidths(sources, { H: 30, I: 16, J: 16 });
const chart = summary.charts.add("ColumnClustered", sources.getRange("H1:J12"), "Auto");
chart.title.text = "第一年配置：当前金额与动态目标";
chart.setPosition(summary.getRange("D2:H16"));
chart.width = 760;
chart.height = 430;
chart.legend.position = "bottom";

summary.getRange("A3:A10").format.fill = gray;
summary.getRange("A4:B10").format.rowHeight = 22;
await fs.mkdir(outputDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
process.stdout.write(outputPath);
