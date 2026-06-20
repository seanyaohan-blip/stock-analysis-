# 行动优先型决策仪表盘实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将每日输出改造成前六张决策优先 Sheet，确保买入候选完整排序、持仓风险突出，并在验证通过后清理旧模板、历史输出和历史盘前报告。

**Architecture:** 保留现有行情、质量、买点、持仓和年度配置计算函数；新增纯 DataFrame 视图构建函数，生成六张前台决策页。Excel 写入阶段对 `01_今日决策` 应用专用卡片式布局，其余页沿用统一表格样式；盘前报告与审计工具优先读取新 Sheet 名称。

**Tech Stack:** Python 3.12、pandas、openpyxl、unittest、artifact-tool、python-docx。

---

### Task 1: 决策视图回归测试

**Files:**
- Create: `tests/test_action_first_dashboard.py`
- Modify: none
- Test: `tests/test_action_first_dashboard.py`

- [ ] **Step 1: Write the failing tests**

编写测试，构造最小 `buy_filter`、`positions_sheet`、`positions_action`、`execution_plan` 和 `dashboard` DataFrame，并断言：

```python
def test_buy_candidates_keep_all_rows_and_sort_by_signal():
    result = main.build_buy_candidates_view(buy_filter, market_permission="暂停标准新增")
    assert result["买点灯号"].tolist() == ["绿", "黄", "红", "灰"]
    assert result.loc[result["买点灯号"] == "红", "阻断原因"].iloc[0]


def test_position_risk_prioritizes_triggered_and_overweight():
    result = main.build_position_risk_view(positions_action, positions_sheet)
    assert result.iloc[0]["风险级别"] == "红"
    assert result.iloc[0]["触发原因"]


def test_front_sheet_order_is_fixed():
    names = main.build_output_sheet_order({name: pd.DataFrame() for name in main.FRONT_SHEET_ORDER})
    assert names[:6] == main.FRONT_SHEET_ORDER
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
python -m unittest tests.test_action_first_dashboard -v
```

Expected: FAIL because `build_buy_candidates_view`、`build_position_risk_view`、`build_output_sheet_order` and `FRONT_SHEET_ORDER` do not exist.

- [ ] **Step 3: Record the expected failure**

确认失败来自缺失的新接口，而不是导入或测试数据错误。

### Task 2: 实现六张前台决策页

**Files:**
- Modify: `main.py`
- Test: `tests/test_action_first_dashboard.py`

- [ ] **Step 1: Add fixed sheet constants and view builders**

在 `main.py` 增加：

```python
FRONT_SHEET_ORDER = [
    "01_今日决策", "02_今日动作", "03_买入候选",
    "04_持仓风险", "05_年度配置", "06_组合总览",
]
```

实现纯函数：

- `build_buy_candidates_view(buy_filter, market_permission)`：保留全部标的，新增市场权限、阻断原因和数据状态，按绿、黄、红、灰排序；
- `build_position_risk_view(positions_action, positions_sheet)`：合并仓位、盈亏与纪律提醒，红色风险优先；
- `build_action_plan_view(execution_plan, buy_filter, positions_sheet)`：补充当前仓位、目标仓位和剩余额度；
- `build_action_dashboard_view(...)`：生成首页所需的状态、动作、阻断和数据记录；
- `build_output_sheet_order(sheets)`：固定前六张，其余明细按设计顺序后移。

- [ ] **Step 2: Run focused tests and verify GREEN**

Run:

```bash
python -m unittest tests.test_action_first_dashboard -v
```

Expected: all tests PASS.

- [ ] **Step 3: Add the six front sheets to main orchestration**

在 `main()` 中用现有底层 DataFrame 构建：

- `01_今日决策`；
- `02_今日动作`；
- `03_买入候选`；
- `04_持仓风险`；
- `05_年度配置`；
- `06_组合总览`。

后移 `Double_Anchor`、`Emotion`、`Quality_Score`、`Exposure`、`Market_Data`、`Positions`、`Broker_Snapshot`、`Watchlist`、`Framework_Rules`、`Checks` 和 `使用说明`。不再把旧版 `Decision_Center`、`Execution_Plan`、`Buy_Filter`、`Positions_Action` 和 `Dashboard` 放在前台重复展示。

- [ ] **Step 4: Implement front-sheet formatting**

在 `write_excel` 中增加 `style_action_dashboard`：标题行、四个状态卡、一句话结论、动作队列、三条阻断原因和数据状态；使用红/黄/绿/灰语义色。其余前台页冻结标题、预排序、设置关键列宽和条件色。

- [ ] **Step 5: Run all Python tests**

Run:

```bash
python -m unittest discover -s tests -v
python -m unittest discover -s premarket_report/tests -v
```

Expected: all tests PASS.

### Task 3: 更新盘前报告与审计工具

**Files:**
- Modify: `premarket_report/main.py`
- Modify: `tools/audit_daily_output.py`
- Modify: `tools/refresh_local_sheets.py`
- Modify: `premarket_report/README.md`
- Test: `premarket_report/tests/test_premarket_report.py`

- [ ] **Step 1: Write failing compatibility tests**

增加测试，断言盘前报告优先读取 `01_今日决策`、`02_今日动作`、`03_买入候选`、`04_持仓风险`、`05_年度配置` 和 `06_组合总览`，旧名称只作为兼容回退。

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
python -m unittest discover -s premarket_report/tests -v
```

Expected: new-name compatibility test FAIL.

- [ ] **Step 3: Implement compatibility and audit rules**

更新盘前报告取表别名、审计目标 Sheet、预期行数、必需列和刷新工具调用。审计新增：

- 前六张 Sheet 顺序检查；
- 买入候选全量行数检查；
- 红色候选阻断原因非空检查；
- 首页行情状态检查。

- [ ] **Step 4: Run tests and compile checks**

Run:

```bash
python -m unittest discover -s tests -v
python -m unittest discover -s premarket_report/tests -v
python -m py_compile main.py premarket_report/main.py tools/audit_daily_output.py tools/refresh_local_sheets.py
```

Expected: tests PASS and compile exits 0.

### Task 4: 生成、渲染并更新模板

**Files:**
- Generate: `output/V2.8.5_每日行情输出_<timestamp>.xlsx`
- Replace: `V2.8.5_ETF_A股统一决策仪表盘模板.xlsx`
- Generate: `premarket_report/output/V2.8.5_A股盘前决策简报_<timestamp>.docx`

- [ ] **Step 1: Generate the workbook**

Run `python main.py`。联网失败时允许使用券商快照回退，但首页必须显示“行情未刷新”，且买入候选不得判绿。

- [ ] **Step 2: Audit the workbook**

Run:

```bash
python tools/audit_daily_output.py <latest-xlsx>
```

Expected: workbook integrity、前六张顺序、候选行数、红色阻断原因和年度配置勾稽均通过；行情缺失必须作为明确数据边界输出。

- [ ] **Step 3: Render and visually inspect**

使用 `tools/render_daily_output.mjs` 渲染全部 Sheet，逐页检查前六张页面是否无截断、无重叠、重点清晰。发现问题时修改样式并重新生成、审计和渲染。

- [ ] **Step 4: Refresh template and premarket report**

将通过验证的最新工作簿复制为 `V2.8.5_ETF_A股统一决策仪表盘模板.xlsx`；运行盘前报告离线模式，确认新 Sheet 能被读取并生成报告。

### Task 5: 清理旧文档并最终验证

**Files:**
- Delete: `V2.8.4_ETF_A股纪律仪表盘模板.xlsx`
- Delete: all but newest `output/V2.8.*_每日行情输出_*.xlsx`
- Delete: all but newest `premarket_report/output/*.docx`
- Preserve: both framework DOCX files、`1st年配置表.xlsx`、optimized allocation workbook、latest template/output/report。

- [ ] **Step 1: Print keep/delete manifests**

按修改时间列出将保留和删除的文档，确认最新版路径与设计一致。

- [ ] **Step 2: Delete approved old documents**

删除旧模板、历史每日输出和历史盘前报告，不删除分析框架、原始配置和优化版配置。

- [ ] **Step 3: Verify retained files and final tests**

Run:

```bash
python -m unittest discover -s tests -v
python -m unittest discover -s premarket_report/tests -v
python tools/audit_daily_output.py <retained-latest-xlsx>
git status --short
```

Expected: tests PASS、审计无结构问题、保留清单完整、Git 仅包含本次代码/模板/设计/计划变更。

- [ ] **Step 4: Commit and push**

Run:

```bash
git add .
git commit -m "feat: add action-first investment dashboard"
git push
```

Expected: commit created and `main`/feature branch push succeeds according to the selected integration path.
