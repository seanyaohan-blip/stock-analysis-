#!/usr/bin/env python3
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import main


LOCAL_MARKET_SNAPSHOT_COLUMNS = [
    "Name",
    "Latest",
    "PctChg",
    "Open",
    "High",
    "Low",
    "PrevClose",
    "Amount",
    "Avg20Amount",
    "Avg20AmountSource",
    "ETFShareChg",
    "Premium",
    "PrevDayLow",
    "PrevDayLowSource",
    "QuoteTime",
    "DataSource",
    "接口状态",
    "失败原因",
    "重试结果",
]


def load_latest_local_market_snapshot() -> pd.DataFrame:
    """Use the newest local Market_Data sheet with complete price context, without making network calls."""
    for path in sorted(main.OUTPUT_DIR.glob(f"{main.FRAMEWORK_VERSION}_每日行情输出_*.xlsx"), reverse=True):
        try:
            snapshot = pd.read_excel(path, sheet_name="Market_Data", dtype={"Code": str})
        except Exception:
            continue
        if "Code" not in snapshot.columns:
            continue
        required = ["Latest", "High", "Low", "PrevClose", "PrevDayLow", "QuoteTime"]
        if not all(column in snapshot.columns for column in required):
            continue
        complete_rows = snapshot[required].notna().all(axis=1).sum()
        if complete_rows >= 5:
            return snapshot
    return pd.DataFrame()


def merge_local_market_snapshot(market_data: pd.DataFrame, snapshot: pd.DataFrame) -> pd.DataFrame:
    if snapshot.empty:
        return market_data
    base = market_data.copy()
    base["Code"] = base["Code"].map(main.format_code)
    snap = snapshot.copy()
    snap["Code"] = snap["Code"].map(main.format_code)
    snap = snap.drop_duplicates("Code").set_index("Code")
    common_codes = [code for code in base["Code"] if code in snap.index]
    for column in LOCAL_MARKET_SNAPSHOT_COLUMNS:
        if column not in base.columns or column not in snap.columns:
            continue
        values = snap.loc[common_codes, column]
        base.loc[base["Code"].isin(common_codes), column] = base.loc[base["Code"].isin(common_codes), "Code"].map(values)
    if "Notes" in base.columns and "DataSource" in base.columns:
        refreshed = base["DataSource"].isin(main.LIVE_QUOTE_SOURCES)
        base.loc[refreshed, "Notes"] = (
            base.loc[refreshed, "Notes"]
            .astype(str)
            .str.replace(r"；未取到行情[^；]*", "", regex=True)
            .str.replace(r"未取到行情[^；]*", "", regex=True)
            .str.strip("； ")
        )
    return base


def build_offline_dashboard() -> int:
    """Generate the workbook without making external market-data requests."""
    watchlist = main.load_watchlist()
    decision_inputs = main.load_decision_inputs()
    first_year_source = main.load_first_year_allocation()
    account_meta, positions = main.load_positions()

    market_data = main.build_market_data(watchlist, {}, positions)
    market_data = merge_local_market_snapshot(market_data, load_latest_local_market_snapshot())
    positions_sheet, alerts = main.build_positions_sheet(positions, market_data, account_meta)
    first_year = main.build_first_year_allocation(first_year_source, positions_sheet, account_meta)
    first_year_summary = main.summarize_first_year_allocation(first_year, account_meta)
    double_anchor = main.build_double_anchor(market_data)
    quality_score = main.build_quality_score(watchlist, market_data)
    emotion = main.build_emotion_thermometer(market_data, decision_inputs)
    buy_filter = main.build_buy_filter(watchlist, market_data, positions, quality_score, emotion, first_year)
    positions_action = main.build_positions_action(positions, market_data)
    broker_snapshot = main.build_broker_snapshot(positions)
    exposure = main.build_exposure_summary(positions_sheet, account_meta)
    buy_point_plan_source = main.load_buy_point_plan()

    market_data = main.attach_first_year_fields(market_data, first_year)
    quality_score = main.attach_first_year_fields(quality_score, first_year)
    buy_filter = main.attach_first_year_fields(buy_filter, first_year)
    positions_sheet = main.attach_first_year_fields(positions_sheet, first_year)
    positions_action = main.attach_first_year_fields(positions_action, first_year)
    broker_snapshot = main.attach_first_year_fields(broker_snapshot, first_year)
    exposure = main.attach_first_year_fields(exposure, first_year)
    watchlist_output = main.attach_first_year_fields(watchlist, first_year)

    rules = main.build_framework_rules(decision_inputs, first_year)
    investment_profile = main.build_investment_profile(decision_inputs)
    data_sources = main.build_data_sources()
    checks = main.build_checks(watchlist, positions, positions_sheet, account_meta, market_data, quality_score, first_year)
    execution_plan = main.build_execution_plan(buy_filter, positions_action, positions_sheet, first_year)
    decision_center = main.build_decision_center(
        double_anchor,
        emotion,
        quality_score,
        buy_filter,
        positions_action,
        first_year_summary,
        decision_inputs,
    )
    dashboard = main.build_dashboard(
        account_meta,
        alerts,
        market_data,
        double_anchor,
        emotion,
        buy_filter,
        quality_score,
        positions_sheet,
        decision_inputs,
        first_year_summary,
    )
    portfolio_overview = main.build_portfolio_overview(dashboard)
    permission_row = decision_center[decision_center["层级"] == "市场权限"]
    market_permission = str(permission_row.iloc[0]["状态"]) if not permission_row.empty else "待确认"
    stage_progression = main.build_stage_progression_view(buy_filter, market_permission)
    buy_candidates = main.build_buy_candidates_view(buy_filter, market_permission)
    position_risk = main.build_position_risk_view(positions_action, positions_sheet)
    action_plan = main.build_action_plan_view(execution_plan, buy_filter, positions_sheet)
    action_dashboard = main.build_action_dashboard_view(
        decision_center,
        action_plan,
        buy_candidates,
        position_risk,
        dashboard,
    )
    long_term_tracking = main.build_long_term_tracking_view(
        watchlist_output,
        market_data,
        quality_score,
        buy_filter,
        first_year,
    )

    output_xlsx = main.make_output_xlsx_path()
    sheets = {
        "01_今日决策": action_dashboard,
        "02_今日动作": action_plan,
        "03_买入候选": buy_candidates,
        "04_阶段推进": stage_progression,
        "买点计划": main.build_buy_point_plan_view(buy_point_plan_source, positions_sheet, buy_filter),
        "05_持仓风险": position_risk,
        "06_年度配置": first_year,
        "07_组合总览": portfolio_overview,
        "Double_Anchor": double_anchor,
        "Emotion": emotion,
        "Quality_Score": quality_score,
        "Exposure": exposure,
        "Market_Data": market_data,
        "Buy_Filter": buy_filter,
        "Positions": positions_sheet,
        "Broker_Snapshot": broker_snapshot,
        "Watchlist": watchlist_output,
        "长期跟踪个股": long_term_tracking,
        "Investment_Profile": investment_profile,
        "Data_Sources": data_sources,
        "Buy_Point_Plan_Source": buy_point_plan_source,
        "Framework_Rules": rules,
        "Checks": checks,
        "使用说明": pd.DataFrame(
            [
                {"步骤": "1", "操作": "编辑 watchlist.csv", "说明": "维护观察池代码、角色、目标权重"},
                {"步骤": "2", "操作": "补充质量与六项信号", "说明": "在watchlist.csv填写完整质量分、评分证据、份额/筹码、折溢价/估值、龙头同步和次日验证"},
                {"步骤": "3", "操作": "补充量化验证口径", "说明": "按资产角色确认指数环境、ETF买点、情绪温度、正期望和风险监控；资金流入不能单独触发买入"},
                {"步骤": "4", "操作": "补充红利个股口径", "说明": "红利个股需复核分红连续性、现金流覆盖、股息率安全垫、估值中低位和高息陷阱否决"},
                {"步骤": "5", "操作": "编辑 decision_inputs.csv", "说明": "可选：填写人工情绪温度、主线周期、趋势买点类型、量化模块状态和退潮三因子状态"},
                {"步骤": "6", "操作": "编辑 first_year_allocation.csv", "说明": "维护第一年全资产目标、代码映射、收益观察区间与执行约束；年度缺口不是买入信号"},
                {"步骤": "7", "操作": "编辑 positions.csv", "说明": "维护持仓数量、成本、账户总资产和券商截图持仓快照"},
                {"步骤": "8", "操作": "运行 python main.py", "说明": "联网抓取行情并生成 Excel；如隐私或网络受限，可运行 tools/build_offline_dashboard.py 生成离线结构模板"},
                {"步骤": "9", "操作": f"打开 output 文件夹里最新的 {main.FRAMEWORK_VERSION}_每日行情输出_日期时间.xlsx", "说明": "按顺序查看 01_今日决策、02_今日动作、03_买入候选、04_阶段推进、买点计划、05_持仓风险、06_年度配置和07_组合总览"},
                {"步骤": "10", "操作": "理解数据边界", "说明": "离线模板的行情字段会标记未刷新；待刷新参考区间只作复核锚点，不作为可执行下单区间"},
            ]
        ),
    }
    main.write_excel(output_xlsx, sheets)
    print(output_xlsx)
    return 0


if __name__ == "__main__":
    raise SystemExit(build_offline_dashboard())
