# Suggested Buy Range Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an auditable `建议买入区间` column to the action-first dashboard template and daily market output, including the approved full-position observation-range exception.

**Architecture:** Keep price math in small pure helpers in `main.py`, then integrate the rendered text into the existing buy-filter pipeline. The calculation uses complete recent OHLC, previous close, previous-day low, signal state, and explicit blockers; `build_buy_candidates_view` applies the final market-permission gate. Workbook styling and audit rules treat the new field as a first-screen decision column while preserving the existing 17-sheet architecture.

**Tech Stack:** Python 3.12, pandas, unittest, existing openpyxl export pipeline, `@oai/artifact-tool` for workbook inspection and visual verification.

---

### Task 1: Add pure quote-freshness and buy-range helpers

**Files:**
- Modify: `tests/test_action_first_dashboard.py`
- Modify: `main.py` near `to_float` and buy-filter helpers

- [ ] **Step 1: Write failing helper tests**

Add tests that define the approved arithmetic, precision, and stale-data behavior:

```python
from datetime import datetime

def test_calculate_buy_range_uses_pullback_band_for_etf(self):
    result = main.calculate_buy_range(
        code="159516", name="半导体设备ETF", signal="绿",
        latest=1.520, high=1.550, low=1.480,
        prev_close=1.510, prev_day_low=1.490,
    )
    self.assertEqual(result, (1.498, 1.515))

def test_calculate_buy_range_uses_stricter_yellow_band_for_stock(self):
    result = main.calculate_buy_range(
        code="002594", name="比亚迪", signal="黄",
        latest=90.00, high=92.00, low=86.00,
        prev_close=89.00, prev_day_low=86.50,
    )
    self.assertEqual(result, (86.90, 88.10))

def test_quote_is_recent_accepts_friday_on_saturday(self):
    self.assertTrue(main.is_recent_complete_quote(
        "2026-06-19 15:00:00", now=datetime(2026, 6, 20, 10, 0, 0)
    ))

def test_quote_is_recent_rejects_missing_or_old_timestamp(self):
    now = datetime(2026, 6, 20, 10, 0, 0)
    self.assertFalse(main.is_recent_complete_quote(None, now=now))
    self.assertFalse(main.is_recent_complete_quote("2026-06-15 15:00:00", now=now))
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  -m unittest tests.test_action_first_dashboard.ActionFirstDashboardTests.test_calculate_buy_range_uses_pullback_band_for_etf \
  tests.test_action_first_dashboard.ActionFirstDashboardTests.test_calculate_buy_range_uses_stricter_yellow_band_for_stock \
  tests.test_action_first_dashboard.ActionFirstDashboardTests.test_quote_is_recent_accepts_friday_on_saturday \
  tests.test_action_first_dashboard.ActionFirstDashboardTests.test_quote_is_recent_rejects_missing_or_old_timestamp -v
```

Expected: FAIL because `calculate_buy_range` and `is_recent_complete_quote` do not exist.

- [ ] **Step 3: Implement the minimum pure helpers**

Add helpers with these interfaces and formulas:

```python
def is_recent_complete_quote(quote_time, now: datetime | None = None) -> bool:
    # Parse datetime/string, reject missing/future values, accept age 0-3 calendar days.

def price_tick(code: str, name: str) -> float:
    # ETF names/codes use 0.001; ordinary A-share stocks use 0.01.

def round_to_tick(value: float, tick: float) -> float:
    # Decimal half-up rounding to the instrument tick.

def calculate_buy_range(
    code: str,
    name: str,
    signal: str,
    latest,
    high,
    low,
    prev_close,
    prev_day_low,
) -> tuple[float, float] | None:
    intraday_range = high - low
    if signal == "绿":
        lower = max(prev_day_low, low + intraday_range * 0.25)
        upper = min(latest, prev_close * 1.005, low + intraday_range * 0.50)
    else:
        lower = max(prev_day_low, low + intraday_range * 0.15)
        upper = min(latest * 0.995, prev_close, low + intraday_range * 0.35)
    # Return None when required inputs are absent, the day range is invalid,
    # or the rounded upper bound is below the rounded lower bound.
```

- [ ] **Step 4: Run helper tests and verify GREEN**

Run the four tests from Step 2.

Expected: `Ran 4 tests ... OK`.

- [ ] **Step 5: Commit the helper slice**

```bash
git add main.py tests/test_action_first_dashboard.py
git commit -m "feat: calculate disciplined buy ranges"
```

### Task 2: Integrate the range into buy-filter decisions

**Files:**
- Modify: `tests/test_action_first_dashboard.py`
- Modify: `main.py` in `BUY_FILTER_COLUMNS`, `build_buy_filter`, and `build_buy_candidates_view`

- [ ] **Step 1: Write failing decision-text tests**

Add tests for visible behavior:

```python
def test_format_buy_range_allows_full_position_observation_exception(self):
    text = main.format_buy_range_recommendation(
        code="159516", name="半导体设备ETF", technical_signal="绿",
        latest=1.520, high=1.550, low=1.480, prev_close=1.510,
        prev_day_low=1.490, quote_time="2026-06-19 15:00:00",
        data_source="东方财富补齐",
        hard_block_kind="position_full", hard_block_reason="当前仓位已达目标仓位",
        veto_reason="无", now=datetime(2026, 6, 20, 10, 0, 0),
    )
    self.assertEqual(text, "暂不建议买入；下一交易日观察区间 1.498–1.515")

def test_format_buy_range_suppresses_other_hard_blocks(self):
    text = main.format_buy_range_recommendation(
        code="159516", name="半导体设备ETF", technical_signal="绿",
        latest=1.520, high=1.550, low=1.480, prev_close=1.510,
        prev_day_low=1.490, quote_time="2026-06-19 15:00:00",
        data_source="东方财富补齐",
        hard_block_kind="quality", hard_block_reason="质量评分不足",
        veto_reason="无", now=datetime(2026, 6, 20, 10, 0, 0),
    )
    self.assertNotRegex(text, r"\d+\.\d+–\d+\.\d+")

def test_market_permission_removes_numeric_range(self):
    source = pd.DataFrame([{
        "买点灯号": "绿", "Code": "159516", "Name": "半导体设备ETF",
        "建议买入区间": "标准复核区间 1.498–1.515", "建议": "复核",
        "否决原因": "", "通过项": 5, "质量评分": 9.0, "剩余额度": 0.02,
        "Latest": 1.52, "PctChg": 0.6,
    }])
    result = main.build_buy_candidates_view(source, market_permission="暂停标准新增")
    self.assertEqual(result.iloc[0]["建议买入区间"], "暂不建议买入（市场权限：暂停标准新增）")

def test_buy_range_column_follows_name(self):
    result = main.build_buy_candidates_view(source, market_permission="开放买点复核")
    self.assertEqual(result.columns[result.columns.get_loc("Name") + 1], "建议买入区间")
```

- [ ] **Step 2: Run decision tests and verify RED**

Run the four new tests with `python -m unittest ... -v`.

Expected: FAIL because the formatter and integrated column behavior are missing.

- [ ] **Step 3: Implement recommendation formatting and pipeline integration**

Add:

```python
def format_buy_range_recommendation(
    code: str,
    name: str,
    technical_signal: str,
    latest,
    high,
    low,
    prev_close,
    prev_day_low,
    quote_time,
    data_source: str,
    hard_block_kind: str,
    hard_block_reason: str,
    veto_reason: str,
    now: datetime | None = None,
) -> str:
    # 1. Reject stale/incomplete prices and non-live sources such as 券商截图.
    # 2. Reject veto and every hard block except position_full.
    # 3. Return standard/half range for green/yellow.
    # 4. Return the non-executable observation range for position_full.
```

In `build_buy_filter`:

- derive `technical_signal` from the six checks before applying position constraints;
- compute `建议买入区间` from the full quote row;
- retain the existing final `买点灯号`, `建议`, and blocker behavior;
- append exactly one new output column, `建议买入区间`.

In `build_buy_candidates_view`:

- place `建议买入区间` immediately after `Name`;
- when `market_permission != "开放买点复核"`, replace all numeric range text with `暂不建议买入（市场权限：...）`;
- retain all watchlist rows and existing signal sorting.

- [ ] **Step 4: Run decision tests and the full dashboard suite**

Run:

```bash
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  -m unittest discover -s tests -v
```

Expected: all action-first dashboard tests pass.

- [ ] **Step 5: Commit the integration slice**

```bash
git add main.py tests/test_action_first_dashboard.py
git commit -m "feat: show suggested buy ranges in candidates"
```

### Task 3: Update workbook style, framework rules, and audit contract

**Files:**
- Modify: `main.py` in `build_framework_rules` and `style_excel_worksheet`
- Modify: `tools/audit_daily_output.py`
- Modify: `tests/test_action_first_dashboard.py`

- [ ] **Step 1: Write failing workbook-contract tests**

Extend the existing workbook test:

```python
self.assertEqual(
    [cell.value for cell in workbook["03_买入候选"][1]][3],
    "建议买入区间",
)
self.assertGreaterEqual(
    workbook["03_买入候选"].column_dimensions["D"].width,
    24,
)
```

Add a framework-rules assertion that the generated rules include `建议买入区间` and identify it as a non-automatic, review-only field.

- [ ] **Step 2: Run the workbook-contract tests and verify RED**

Run the targeted workbook and framework-rules tests.

Expected: FAIL because the width and framework rule have not been added.

- [ ] **Step 3: Implement workbook presentation and audit requirements**

In `main.py`:

- add a `Framework_Rules` row named `建议买入区间` with the green/yellow/full-position/stale-data rules and V2.8.5 source;
- set `建议买入区间` width to approximately 32 and wrap its text;
- preserve existing lamp fills, alternating rows, filters, freeze panes, and front-sheet order.

In `tools/audit_daily_output.py`:

- add `建议买入区间` to required headers for `03_买入候选`;
- record rows with blank recommendation text;
- flag any numeric range when `数据状态 == "行情未刷新"`;
- flag numeric ranges on non-green/non-yellow rows unless the text starts with `暂不建议买入；` and the row blocker is exclusively the full-position constraint;
- append clear issue strings so a nonzero audit finding is visible in JSON.

- [ ] **Step 4: Run tests and syntax checks**

Run:

```bash
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m unittest discover -s tests -v
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m unittest discover -s premarket_report/tests -v
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m py_compile main.py tools/audit_daily_output.py
```

Expected: all tests pass and compilation exits 0.

- [ ] **Step 5: Commit workbook contract changes**

```bash
git add main.py tools/audit_daily_output.py tests/test_action_first_dashboard.py
git commit -m "test: audit suggested buy range output"
```

### Task 4: Regenerate template and latest daily output

**Files:**
- Modify: `V2.8.5_ETF_A股统一决策仪表盘模板.xlsx`
- Create: newest timestamped workbook under `output/` (ignored delivery artifact)

- [ ] **Step 1: Generate a fresh daily workbook**

Run the bundled Python entry point. If network quotes are unavailable, allow the existing safe offline fallback to generate the workbook; do not invent prices.

```bash
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 main.py
```

Expected: a timestamped workbook is created and every candidate has either a review range or an explicit non-buy reason.

- [ ] **Step 2: Run the workbook audit**

```bash
LATEST=$(find output -maxdepth 1 -type f -name 'V2.8.5_每日行情输出_*.xlsx' -print | sort | tail -1)
/Users/seany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  tools/audit_daily_output.py "$LATEST"
```

Expected: ZIP integrity, 17 target sheets, front-sheet order, row counts, blocker coverage, and suggested-range checks pass; stale-data warnings may remain explicit but produce no unsafe numeric ranges.

- [ ] **Step 3: Promote the audited workbook to the tracked template**

Copy the audited workbook to `V2.8.5_ETF_A股统一决策仪表盘模板.xlsx`, then rerun the same audit against the template. Preserve the user's separate optimized allocation workbook unchanged.

- [ ] **Step 4: Inspect and render with `@oai/artifact-tool`**

Import the template and latest output, inspect `03_买入候选!A1:H20`, scan for formula errors, and render all sheets. Visually check at least `01_今日决策`, `03_买入候选`, `04_持仓风险`, and `06_组合总览`; fix clipping or unreadable widths before export.

- [ ] **Step 5: Keep only the newest daily output and commit the template**

After all audits pass, remove older timestamped daily output workbooks under `output/`, retaining only the newest file, then commit only the tracked template and source/test changes:

```bash
git add V2.8.5_ETF_A股统一决策仪表盘模板.xlsx
git commit -m "chore: refresh dashboard template with buy ranges"
```

### Task 5: Final verification and integration

**Files:**
- Verify: all changed source, test, template, and latest output files

- [ ] **Step 1: Run fresh full verification**

Run both unittest suites, Python compilation, workbook audit on template and latest output, ZIP integrity, required-file checks, and a git diff/status review.

Expected: zero test failures, zero syntax errors, zero workbook audit issues, exactly 17 sheets in the required order, and no user-owned optimized allocation workbook changes included in the branch commits.

- [ ] **Step 2: Merge the feature branch locally**

Use the finishing-a-development-branch workflow. Merge the verified feature branch into `main`, copy the ignored newest daily output into the main worktree, rerun the focused tests and audit, then remove the owned temporary worktree and feature branch.

- [ ] **Step 3: Report delivery status**

Link the updated template, newest daily output, design, and implementation plan. State the exact test count and whether the delivered workbook contains live numeric ranges or safe stale-data blockers.
