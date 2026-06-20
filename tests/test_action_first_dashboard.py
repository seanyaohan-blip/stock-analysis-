from pathlib import Path
import sys
import tempfile
import unittest
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import main


class ActionFirstDashboardTests(unittest.TestCase):
    def test_calculate_buy_range_uses_pullback_band_for_etf(self):
        result = main.calculate_buy_range(
            code="159516",
            name="半导体设备ETF",
            signal="绿",
            latest=1.520,
            high=1.550,
            low=1.480,
            prev_close=1.510,
            prev_day_low=1.490,
        )

        self.assertEqual(result, (1.498, 1.515))

    def test_calculate_buy_range_uses_stricter_yellow_band_for_stock(self):
        result = main.calculate_buy_range(
            code="002594",
            name="比亚迪",
            signal="黄",
            latest=90.00,
            high=92.00,
            low=86.00,
            prev_close=89.00,
            prev_day_low=86.50,
        )

        self.assertEqual(result, (86.90, 88.10))

    def test_quote_is_recent_accepts_friday_on_saturday(self):
        self.assertTrue(
            main.is_recent_complete_quote(
                "2026-06-19 15:00:00",
                now=datetime(2026, 6, 20, 10, 0, 0),
            )
        )

    def test_quote_is_recent_rejects_missing_or_old_timestamp(self):
        now = datetime(2026, 6, 20, 10, 0, 0)

        self.assertFalse(main.is_recent_complete_quote(None, now=now))
        self.assertFalse(main.is_recent_complete_quote("2026-06-15 15:00:00", now=now))

    def test_format_buy_range_allows_full_position_observation_exception(self):
        text = main.format_buy_range_recommendation(
            code="159516",
            name="半导体设备ETF",
            technical_signal="绿",
            latest=1.520,
            high=1.550,
            low=1.480,
            prev_close=1.510,
            prev_day_low=1.490,
            quote_time="2026-06-19 15:00:00",
            data_source="东方财富补齐",
            hard_block_kind="position_full",
            hard_block_reason="当前仓位已达目标仓位",
            veto_reason="无",
            now=datetime(2026, 6, 20, 10, 0, 0),
        )

        self.assertEqual(text, "暂不建议买入；下一交易日观察区间 1.498–1.515")

    def test_format_buy_range_suppresses_other_hard_blocks(self):
        text = main.format_buy_range_recommendation(
            code="159516",
            name="半导体设备ETF",
            technical_signal="绿",
            latest=1.520,
            high=1.550,
            low=1.480,
            prev_close=1.510,
            prev_day_low=1.490,
            quote_time="2026-06-19 15:00:00",
            data_source="东方财富补齐",
            hard_block_kind="quality",
            hard_block_reason="质量评分不足",
            veto_reason="无",
            now=datetime(2026, 6, 20, 10, 0, 0),
        )

        self.assertEqual(text, "暂不建议买入（质量评分不足）")
        self.assertNotRegex(text, r"\d+\.\d+–\d+\.\d+")

    def test_market_permission_removes_numeric_range(self):
        source = pd.DataFrame(
            [
                {
                    "买点灯号": "绿",
                    "Code": "159516",
                    "Name": "半导体设备ETF",
                    "建议买入区间": "标准复核区间 1.498–1.515",
                    "建议": "复核",
                    "否决原因": "",
                    "通过项": 5,
                    "质量评分": 9.0,
                    "剩余额度": 0.02,
                    "Latest": 1.52,
                    "PctChg": 0.6,
                }
            ]
        )

        result = main.build_buy_candidates_view(source, market_permission="暂停标准新增")

        self.assertEqual(result.iloc[0]["建议买入区间"], "暂不建议买入（市场权限：暂停标准新增）")

    def test_buy_range_column_follows_name(self):
        source = pd.DataFrame(
            [
                {
                    "买点灯号": "绿",
                    "Code": "159516",
                    "Name": "半导体设备ETF",
                    "建议买入区间": "标准复核区间 1.498–1.515",
                    "建议": "复核",
                    "否决原因": "",
                    "通过项": 5,
                    "质量评分": 9.0,
                    "剩余额度": 0.02,
                    "Latest": 1.52,
                    "PctChg": 0.6,
                }
            ]
        )

        result = main.build_buy_candidates_view(source, market_permission="开放买点复核")

        self.assertEqual(result.columns[result.columns.get_loc("Name") + 1], "建议买入区间")

    def test_buy_candidates_keep_all_rows_and_sort_by_signal(self):
        buy_filter = pd.DataFrame(
            [
                {"买点灯号": "灰", "Code": "000004", "Name": "灰标的", "建议": "观察", "否决原因": "", "通过项": 1, "质量评分": None, "剩余额度": 0.04},
                {"买点灯号": "红", "Code": "000003", "Name": "红标的", "建议": "不买", "否决原因": "质量评分缺失", "通过项": 2, "质量评分": None, "剩余额度": 0.03},
                {"买点灯号": "绿", "Code": "000001", "Name": "绿标的", "建议": "复核", "否决原因": "", "通过项": 6, "质量评分": 9.0, "剩余额度": 0.02},
                {"买点灯号": "黄", "Code": "000002", "Name": "黄标的", "建议": "半额复核", "否决原因": "", "通过项": 4, "质量评分": 8.0, "剩余额度": 0.01},
            ]
        )

        result = main.build_buy_candidates_view(buy_filter, market_permission="暂停标准新增")

        self.assertEqual(result["买点灯号"].tolist(), ["绿", "黄", "红", "灰"])
        self.assertEqual(len(result), len(buy_filter))
        red_reason = result.loc[result["买点灯号"] == "红", "阻断原因"].iloc[0]
        self.assertTrue(str(red_reason).strip())

    def test_position_risk_prioritizes_triggered_and_overweight(self):
        positions_action = pd.DataFrame(
            [
                {"Code": "000001", "Name": "触发标的", "触发提醒": "是", "动作建议": "减仓复核", "说明": "进入纪律区间"},
                {"Code": "000002", "Name": "普通标的", "触发提醒": "否", "动作建议": "观察", "说明": "未触发"},
            ]
        )
        positions_sheet = pd.DataFrame(
            [
                {"Code": "000001", "Name": "触发标的", "Weight": 0.08, "Target Weight": 0.05, "Market Value": 80000, "Unrealized PnL": -10000},
                {"Code": "000002", "Name": "普通标的", "Weight": 0.03, "Target Weight": 0.05, "Market Value": 30000, "Unrealized PnL": 1000},
            ]
        )

        result = main.build_position_risk_view(positions_action, positions_sheet)

        self.assertEqual(result.iloc[0]["风险级别"], "红")
        self.assertTrue(str(result.iloc[0]["触发原因"]).strip())
        self.assertEqual(result.iloc[0]["Code"], "000001")

    def test_front_sheet_order_is_fixed(self):
        sheets = {
            "Checks": pd.DataFrame(),
            "03_买入候选": pd.DataFrame(),
            "01_今日决策": pd.DataFrame(),
            "06_组合总览": pd.DataFrame(),
            "04_持仓风险": pd.DataFrame(),
            "02_今日动作": pd.DataFrame(),
            "05_年度配置": pd.DataFrame(),
        }

        names = main.build_output_sheet_order(sheets)

        self.assertEqual(names[:6], main.FRONT_SHEET_ORDER)

    def test_action_plan_adds_position_limits(self):
        execution_plan = pd.DataFrame(
            [
                {"优先级": 1, "Code": "000001", "动作类型": "风控/减仓复核", "标的": "测试标的(000001)", "依据": "超配", "状态": "优先处理"},
            ]
        )
        buy_filter = pd.DataFrame(
            [
                {"Code": "000001", "当前仓位": 0.08, "目标仓位": 0.05, "剩余额度": 0.0},
            ]
        )
        positions_sheet = pd.DataFrame(
            [
                {"Code": "000001", "Weight": 0.08, "Target Weight": 0.05},
            ]
        )

        result = main.build_action_plan_view(execution_plan, buy_filter, positions_sheet)

        self.assertEqual(result.iloc[0]["当前仓位"], 0.08)
        self.assertEqual(result.iloc[0]["目标仓位"], 0.05)
        self.assertEqual(result.iloc[0]["剩余额度"], 0.0)

    def test_action_dashboard_contains_status_actions_blockers_and_data(self):
        decision_center = pd.DataFrame(
            [
                {"层级": "市场权限", "状态": "暂停标准新增", "证据": "双锚观察", "动作": "继续观察"},
            ]
        )
        action_plan = pd.DataFrame(
            [
                {"优先级": 1, "动作类型": "风控/减仓复核", "标的": "测试标的", "依据": "超配", "状态": "优先处理"},
            ]
        )
        buy_candidates = pd.DataFrame(
            [
                {"买点灯号": "红", "Name": "红标的", "阻断原因": "质量评分缺失"},
            ]
        )
        position_risk = pd.DataFrame([{"风险级别": "红", "Name": "测试标的"}])
        dashboard = pd.DataFrame(
            [
                {"项目": "现金安全垫状态", "内容": "达标"},
                {"项目": "行情数据时间", "内容": "未取到当日接口时间"},
            ]
        )

        result = main.build_action_dashboard_view(
            decision_center,
            action_plan,
            buy_candidates,
            position_risk,
            dashboard,
        )

        self.assertEqual(set(result["区域"]), {"核心状态", "今日动作", "主要阻断", "数据状态"})
        core_items = set(result.loc[result["区域"] == "核心状态", "项目"])
        self.assertEqual(core_items, {"市场权限", "买入候选", "减仓复核", "现金安全垫"})
        data_status = result.loc[result["区域"] == "数据状态", "状态"].iloc[0]
        self.assertIn("未刷新", data_status)

    def test_write_excel_places_front_sheets_first(self):
        dashboard_view = pd.DataFrame(
            [
                {"区域": "核心状态", "项目": "市场权限", "状态": "暂停标准新增", "证据": "双锚观察", "动作": "继续观察"},
                {"区域": "数据状态", "项目": "行情时间", "状态": "行情未刷新", "证据": "未取到当日接口时间", "动作": "禁止新增"},
            ]
        )
        sheets = {
            "Checks": pd.DataFrame([{"检查项": "测试", "状态": "OK"}]),
            "03_买入候选": pd.DataFrame([{"买点灯号": "红", "Code": "000001", "阻断原因": "测试"}]),
            "01_今日决策": dashboard_view,
            "06_组合总览": pd.DataFrame([{"项目": "测试", "内容": "测试"}]),
            "04_持仓风险": pd.DataFrame([{"风险级别": "红", "Code": "000001", "触发原因": "测试"}]),
            "02_今日动作": pd.DataFrame([{"优先级": 1, "动作类型": "观察", "标的": "测试"}]),
            "05_年度配置": pd.DataFrame([{"配置项": "测试", "年度目标占比": 0.1}]),
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "dashboard.xlsx"

            main.write_excel(path, sheets)

            workbook = load_workbook(path, read_only=False)
            self.assertEqual(workbook.sheetnames[:6], main.FRONT_SHEET_ORDER)
            self.assertIn("A1:H1", {str(item) for item in workbook["01_今日决策"].merged_cells.ranges})

    def test_portfolio_overview_removes_per_position_alert_rows(self):
        dashboard = pd.DataFrame(
            [
                {"项目": "全资产总额", "内容": 1000000},
                {"项目": "提醒-000001", "内容": "当前仓位超目标"},
                {"项目": "现金安全垫状态", "内容": "达标"},
            ]
        )

        result = main.build_portfolio_overview(dashboard)

        self.assertEqual(result["项目"].tolist(), ["全资产总额", "现金安全垫状态"])


if __name__ == "__main__":
    unittest.main()
