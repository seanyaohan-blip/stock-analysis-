#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
TARGET_SHEETS = [
    "Decision_Center",
    "Dashboard",
    "FirstYear_Allocation",
    "Execution_Plan",
    "Checks",
    "Framework_Rules",
    "Emotion",
    "Quality_Score",
    "Exposure",
    "Market_Data",
    "Double_Anchor",
    "Buy_Filter",
    "Positions",
    "Positions_Action",
]


def normalize_code(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0"):
        text = text[:-2]
    upper = text.upper()
    if upper.endswith((".SH", ".SZ", ".CSI")):
        return upper
    return text.zfill(6) if text.isdigit() else text


def csv_codes(path: Path) -> list[str]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        lines = [line for line in handle if not line.startswith("#")]
    return [normalize_code(row["Code"]) for row in csv.DictReader(lines)]


def headers(ws) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(ws[1], start=1)
        if cell.value is not None
    }


def blank_rows(ws, required: list[str]) -> list[dict[str, Any]]:
    mapping = headers(ws)
    issues: list[dict[str, Any]] = []
    for row in range(2, ws.max_row + 1):
        code = ws.cell(row, mapping.get("Code", 1)).value
        missing = [
            name
            for name in required
            if name not in mapping or ws.cell(row, mapping[name]).value in (None, "")
        ]
        if missing:
            issues.append({"row": row, "code": normalize_code(code), "missing": missing})
    return issues


def main(path: Path) -> int:
    result: dict[str, Any] = {"path": str(path.resolve()), "checks": {}, "issues": []}
    expected_suffix = re.compile(r"_\d{8}_\d{6}\.xlsx$")
    result["checks"]["timestamped_filename"] = bool(expected_suffix.search(path.name))

    with zipfile.ZipFile(path) as archive:
        bad_member = archive.testzip()
        result["checks"]["zip_integrity"] = bad_member is None
        result["checks"]["bad_zip_member"] = bad_member

    wb = load_workbook(path, data_only=False, read_only=False, keep_links=True)
    result["sheetnames"] = wb.sheetnames
    result["checks"]["target_sheets_present"] = all(name in wb.sheetnames for name in TARGET_SHEETS)
    result["checks"]["external_links"] = len(getattr(wb, "_external_links", []))

    watch_codes = csv_codes(ROOT / "watchlist.csv")
    position_codes = csv_codes(ROOT / "positions.csv")
    expected_market = list(dict.fromkeys(watch_codes + position_codes))
    expected_counts = {
        "FirstYear_Allocation": 11,
        "Market_Data": len(expected_market),
        "Double_Anchor": 8,
        "Buy_Filter": len(watch_codes),
        "Positions": len(position_codes),
        "Positions_Action": len(position_codes),
    }

    required = {
        "Market_Data": [
            "Code", "Name", "Latest", "PctChg", "Open", "High", "Low", "PrevClose",
            "Amount", "Avg20Amount", "Avg20AmountSource", "ETFShareChg", "Premium",
            "LeaderStatus", "PrevDayLow", "PrevDayLowSource", "QuoteTime", "DataSource",
        ],
        "Buy_Filter": [
            "Code", "Name", "Latest", "PctChg", "日内位置", "量能倍数", "分时结构",
            "量价关系", "份额变动", "折溢价", "龙头同步", "次日验证", "通过项",
            "通过明细", "未通过项", "待确认项", "一票否决", "否决原因", "建议",
        ],
        "FirstYear_Allocation": [
            "配置项", "资产层级", "映射代码", "年度目标占比", "按当前全资产目标金额",
            "最新持仓金额", "年度资金缺口", "年度完成率", "目标收益下限", "目标收益上限",
            "进度状态", "执行约束", "来源",
        ],
        "Positions": [
            "Code", "Name", "Shares", "Cost", "Latest", "Latest Source", "Market Value",
            "P/L", "P/L%", "Weight", "Full Asset Weight", "Role", "Target Weight", "V2.8.5 Action",
        ],
        "Positions_Action": [
            "Code", "Name", "Latest", "PctChg", "High", "Low", "Amount", "Avg20Amount",
            "触发提醒", "动作建议", "是否复审", "说明",
        ],
    }

    formulas = 0
    formula_errors: list[dict[str, str]] = []
    hyperlinks: list[dict[str, str]] = []
    flagged_text: list[dict[str, str]] = []
    flags = ("未取到", "接口失败", "字段不匹配", "链接失效", "陈旧", "暂无接口", "待填")

    for sheet_name in TARGET_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        sheet_result: dict[str, Any] = {
            "data_rows": max(ws.max_row - 1, 0),
            "columns": ws.max_column,
            "freeze_panes": str(ws.freeze_panes or ""),
            "auto_filter": str(ws.auto_filter.ref or ""),
        }
        if sheet_name in expected_counts:
            sheet_result["expected_rows"] = expected_counts[sheet_name]
            sheet_result["row_count_ok"] = sheet_result["data_rows"] == expected_counts[sheet_name]
        if sheet_name in required:
            sheet_result["blank_required"] = blank_rows(ws, required[sheet_name])
        result[sheet_name] = sheet_result

        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formulas += 1
                    if any(token in value for token in ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")):
                        formula_errors.append({"sheet": sheet_name, "cell": cell.coordinate, "value": value})
                if cell.hyperlink:
                    hyperlinks.append({"sheet": sheet_name, "cell": cell.coordinate, "target": str(cell.hyperlink.target)})
                if (
                    cell.row > 1
                    and isinstance(value, str)
                    and value != "未取到行情数量"
                    and any(flag in value for flag in flags)
                ):
                    flagged_text.append({"sheet": sheet_name, "cell": cell.coordinate, "value": value})

    market_ws = wb["Market_Data"]
    market_headers = headers(market_ws)
    market_codes = [normalize_code(market_ws.cell(row, market_headers["Code"]).value) for row in range(2, market_ws.max_row + 1)]
    result["checks"]["market_code_set_ok"] = set(market_codes) == set(expected_market)
    result["checks"]["market_missing_codes"] = sorted(set(expected_market) - set(market_codes))
    result["checks"]["market_extra_codes"] = sorted(set(market_codes) - set(expected_market))
    result["checks"]["market_duplicate_codes"] = sorted({code for code in market_codes if market_codes.count(code) > 1})

    quote_times: list[datetime] = []
    for row in range(2, market_ws.max_row + 1):
        raw = market_ws.cell(row, market_headers["QuoteTime"]).value
        if raw:
            try:
                quote_times.append(datetime.strptime(str(raw), "%Y-%m-%d %H:%M:%S"))
            except ValueError:
                result["issues"].append(f"Invalid QuoteTime at Market_Data row {row}: {raw}")
    result["quote_time_min"] = min(quote_times).isoformat(sep=" ") if quote_times else None
    result["quote_time_max"] = max(quote_times).isoformat(sep=" ") if quote_times else None
    result["checks"]["all_quotes_today"] = bool(quote_times) and all(item.date() == datetime.now().date() for item in quote_times)

    dashboard_ws = wb["Dashboard"]
    dashboard = {
        str(dashboard_ws.cell(row, 1).value): dashboard_ws.cell(row, 2).value
        for row in range(2, dashboard_ws.max_row + 1)
    }
    result["dashboard"] = dashboard
    result["checks"]["dashboard_missing_count"] = dashboard.get("未取到行情数量")
    result["checks"]["dashboard_missing_zero"] = dashboard.get("未取到行情数量") == 0
    result["checks"]["formula_cells"] = formulas
    result["checks"]["formula_errors"] = formula_errors
    result["checks"]["hyperlinks"] = hyperlinks
    result["checks"]["flagged_text"] = flagged_text

    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: audit_daily_output.py <workbook.xlsx>")
    raise SystemExit(main(Path(sys.argv[1])))
