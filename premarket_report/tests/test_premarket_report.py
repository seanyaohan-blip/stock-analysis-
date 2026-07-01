from datetime import datetime
from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import main
from docx import Document
from docx.oxml.ns import qn


class FakeResponse:
    def __init__(self, text):
        self.text = text

    def raise_for_status(self):
        return None


class FakeSession:
    def __init__(self, text):
        self.text = text
        self.urls = []

    def get(self, url, timeout):
        self.urls.append((url, timeout))
        return FakeResponse(self.text)


class PremarketReportTests(unittest.TestCase):
    def test_parse_time_normalizes_utc_to_shanghai(self):
        self.assertEqual(
            main.parse_time("Thu, 18 Jun 2026 21:14:00 GMT"),
            "2026-06-19 05:14",
        )

    def test_parse_time_normalizes_chinese_date(self):
        self.assertEqual(main.parse_time("2026年6月18日"), "2026-06-18 00:00")

    def test_infer_news_time_from_dated_url(self):
        self.assertEqual(
            main.infer_news_time_from_url(
                "http://www.pbc.gov.cn/example/2026061810473541154/index.html"
            ),
            "2026-06-18",
        )
        self.assertEqual(
            main.infer_news_time_from_url(
                "https://www.ndrc.gov.cn/example/t20260618_1405995.html"
            ),
            "2026-06-18",
        )

    def test_market_timestamp_uses_shanghai_timezone(self):
        timestamp = int(datetime.fromisoformat("2026-06-19T01:00:00+00:00").timestamp())
        self.assertEqual(
            main.timestamp_to_shanghai(timestamp),
            "2026-06-19 09:00",
        )

    def test_source_attempt_summary_records_retry_and_fallback(self):
        summary = main.format_source_attempt_summary(
            "财联社",
            [
                ("https://www.cls.cn/telegraph", "HTTP成功但0条近3日有效新闻"),
                ("https://www.cls.cn/telegraph", "重试后仍为0条"),
                ("Google News site:cls.cn", "备用源成功，6条"),
            ],
        )
        self.assertIn("财联社", summary)
        self.assertIn("重试", summary)
        self.assertIn("备用源成功", summary)

    def test_market_attempt_summary_is_bounded(self):
        summary = main.format_market_attempt_summary(
            [
                ("Yahoo query1 ^GSPC", "失败 dns"),
                ("Yahoo query2 ^GSPC", "失败 timeout"),
                ("Stooq ^spx", "失败 empty"),
                ("Stooq backup", "失败 empty"),
                ("extra1", "失败"),
                ("extra2", "失败"),
            ],
            max_attempts=3,
        )
        self.assertIn("Yahoo query1", summary)
        self.assertIn("Stooq ^spx", summary)
        self.assertIn("另3次尝试失败", summary)

    def test_stooq_history_parser_uses_last_two_closes(self):
        session = FakeSession(
            "Date,Open,High,Low,Close,Volume\n"
            "2026-06-22,5000,5050,4990,5000,100\n"
            "2026-06-23,5100,5130,5070,5120,120\n"
        )
        quote = main.fetch_one_stooq_symbol(
            session,
            {"name": "标普500", "unit": "点", "kind": "index"},
            "^spx",
        )

        self.assertIsNotNone(quote)
        self.assertEqual(quote.latest, 5120.0)
        self.assertEqual(quote.previous, 5000.0)
        self.assertEqual(quote.pct_change, 2.4)
        self.assertEqual(quote.symbol, "Stooq:^spx")
        self.assertEqual(quote.data_time, "2026-06-23")

    def test_table_geometry_sets_explicit_dxa_indent(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=2)
        main.set_table_geometry(table, [3.0, 3.5])
        tbl_indent = table._tbl.tblPr.first_child_found_in("w:tblInd")
        self.assertIsNotNone(tbl_indent)
        self.assertEqual(tbl_indent.get(qn("w:type")), "dxa")
        self.assertEqual(tbl_indent.get(qn("w:w")), "120")

    def test_sheet_fallback_prefers_action_first_name(self):
        preferred = pd.DataFrame([{"项目": "新首页"}])
        legacy = pd.DataFrame([{"项目": "旧首页"}])
        sheets = {"01_今日决策": preferred, "Decision_Center": legacy}

        result = main.get_sheet_with_fallback(sheets, "01_今日决策", "Decision_Center")

        self.assertEqual(result.iloc[0]["项目"], "新首页")

    def test_sheet_fallback_reads_legacy_name_when_needed(self):
        legacy = pd.DataFrame([{"项目": "旧首页"}])
        sheets = {"Decision_Center": legacy}

        result = main.get_sheet_with_fallback(sheets, "01_今日决策", "Decision_Center")

        self.assertEqual(result.iloc[0]["项目"], "旧首页")

    def test_shifted_position_risk_sheet_name_is_supported(self):
        shifted = pd.DataFrame([{"Code": "510210", "动作建议": "提示卖出20000份"}])
        legacy = pd.DataFrame([{"Code": "legacy", "动作建议": "旧版"}])
        sheets = {"05_持仓风险": shifted, "Positions_Action": legacy}

        result = main.get_sheet_with_fallback(sheets, *main.POSITION_RISK_SHEET_NAMES)

        self.assertEqual(result.iloc[0]["Code"], "510210")


if __name__ == "__main__":
    unittest.main()
