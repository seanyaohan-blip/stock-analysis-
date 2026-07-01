from pathlib import Path
import sys
import unittest


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tools import audit_daily_output


class AuditDailyOutputTests(unittest.TestCase):
    def test_missing_market_quote_fields_treats_dash_as_incomplete(self):
        row = {
            "Code": "688072",
            "Latest": "-",
            "PctChg": "-",
            "Open": "-",
            "High": "-",
            "Low": "-",
            "PrevClose": 832,
            "Amount": "-",
        }

        missing = audit_daily_output.missing_market_quote_fields(row)

        self.assertEqual(missing, ["最新价", "涨跌幅", "今开", "最高", "最低", "成交额"])

    def test_incomplete_market_row_requires_explicit_diagnostics(self):
        bad_row = {
            "接口状态": "当日行情已刷新",
            "失败原因": "无",
            "重试结果": "实时行情与日K补齐可用",
        }
        good_row = {
            "接口状态": "行情字段缺失",
            "失败原因": "实时行情已返回但字段不完整；缺失字段：最新价",
            "重试结果": "实时行情补齐已尝试，未取得当日完整行情",
        }

        self.assertFalse(audit_daily_output.incomplete_market_row_has_diagnostics(bad_row))
        self.assertTrue(audit_daily_output.incomplete_market_row_has_diagnostics(good_row))


if __name__ == "__main__":
    unittest.main()
