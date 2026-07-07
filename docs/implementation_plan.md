# 费用底稿 AI Review 助手 - 优化实施计划

## 基准版本：planA-backup-20260707

本计划基于自查报告 `review_self_check_findings.md`、B方案规则代码和固定资产质检Agent设计模式制定。

---

## Step 1: P0 误报修正（4项）

### 1-1 TB/BKD 科目匹配逻辑修正

**文件**：`src/expense_review_engine.py` → `review_trial_balance()`

**问题**：`account not in bkd_accounts` 精确匹配，父级科目"销售费用"在TB中存在但BKD只有明细科目 → 被误报为"未在BKD列示"

**修正**：
- 增加 `_is_leaf_account()` 辅助函数：判断TB科目是否为叶子科目
- 叶子科目才做精确匹配，父级聚合科目不做直接比对
- 父级科目值的完整检查保留到候选列表供LLM辅助

**验收**：TB/BKD误报数大幅下降，父级科目不再直接报Medium

### 1-2 序时账关键词范围过滤

**文件**：`src/expense_review_engine.py` → `review_general_ledger()`

**问题**：`GL_KEYWORDS` 命中后不检查科目范围，银行存款、在建工程等也触发

**修正**：
- 命中关键词后，增加对 `full_account` 是否包含"销售费用"或"管理费用"的判断
- 非费用类科目跳过不生成Review Note，但保留到候选列表(供LLM二次判断)

**验收**：非费用科目不再直接出现在Review Notes中

### 1-3 TOD 样本池误判修正

**文件**：`src/expense_review_engine.py` → `review_tod()`

**问题**：样本池 sheet 也被做"测试结论线索不足"检查（第617-637行）

**修正**：
- 将文本搜索范围从 `total_pool_sheets + reduced_pool_sheets + test_sheets[:5]` 改为仅 `test_sheets[:5]`
- 样本池检查只保留：空池检查(行数<=1)、明细过少检查(行数<3)

**验收**：样本池不再出现"测试结论线索不足"

### 1-4 LLM API 失败从 Review Notes 分离

**文件**：`src/llm_assistant.py` → `llm_notes_from_candidates()`  
**关联**：`src/expense_review_engine.py` → `review_trial_balance()`, `review_general_ledger()`

**问题**：LLM调用失败时生成一条 `risk_level="Low"` 的Note混入审计结果

**修正**：
- 返回值改为 `(list[ReviewNote], dict)` 元组
- dict 为 `{"ok": bool, "status": str, "candidates_submitted": int}`
- 失败信息只存在dict中，调用方将状态信息传递给上层
- 上层在UI中展示LLM状态，而非混入Review Notes

**验收**：LLM不可用时Review Notes数量不因此增加，状态信息在UI侧边栏独立显示

---

## Step 2: B方案6条新规则移植

### 2-1 BKD 货币/单位检查

**函数**：`review_bkd_currency_unit()`  
**规则ID**：`bkd_currency_unit`

检查BKD表头"货币/单位"行是否显示为IFRS（应为CNY/RMB），与Lead C8比较。

**依赖**：`VC.00 销售费用BKD`, `VD.00 管理费用BKD`  
**空值保护**：sheet不存在 → `[]`

### 2-2 BKD 上期审定数检查

**函数**：`review_bkd_prior_year()`  
**规则ID**：`bkd_prior_year`

检测上期审定数列（J列）是否出现极小比例值(<0.1)，疑似公式错误。

**依赖**：同上  
**空值保护**：sheet不存在 → `[]`  
**注意**：自查报告指出此规则可能误判，需设置阈值(suspicious_count>=3)，且排除本期金额为0的行

### 2-3 TOD 样本总体范围检查

**函数**：`review_tod_population_scope()`  
**规则ID**：`tod_population_scope`

检查TOD样本总体是否包含制造费用/研发支出。先在主底稿所有sheet中查找包含"TOD"和"预审"/"剩余期间"的sheet，不硬编码sheet名。

**依赖**：含有"TOD"和"预审"关键词的sheet  
**空值保护**：找不到对应sheet → `[]`

### 2-4 TOD 关键项(KI)检查

**函数**：`review_tod_key_items()`  
**规则ID**：`tod_key_items`

检查关键项数量和金额是否全为0。适应不同格式的sheet布局。

**依赖**：同上  
**空值保护**：找不到对应sheet → `[]`

### 2-5 TOD 负值处理检查

**函数**：`review_tod_negative_handling()`  
**规则ID**：`tod_negative_handling`

检查负值分析表中"是否抽样"列是否为空。

**依赖**：同上  
**空值保护**：找不到对应sheet → `[]`

### 2-6 缺失标准Sheet检查

**函数**：`review_missing_sheets()`  
**规则ID**：`missing_sheets`

对照标准模板（7个Sheet）检查主底稿是否缺失。与汇总页执行状态联动：如果某个sheet缺失但汇总页标记为"是"，则为High。

**依赖**：主底稿  
**标准集**：`汇总`, `Uexp.00 Lead`, `VC.00 销售费用BKD`, `VD.00 管理费用BKD`, `VC&VD.01.2 详细测试 TOD`, `VD.01.3 复核法律费用`, `VC&VD.01.4 截止性测试`

---

## Step 3: LLM 模块增强

### 3-1 新增连接测试函数

**文件**：`src/llm_assistant.py`  
**函数**：`test_connection(config: LLMConfig) -> dict`

发送max_tokens=1的轻量请求，10秒超时，返回：
```python
{"ok": bool, "latency_ms": float, "model": str, "error": str}
```

### 3-2 API失败状态分离

配合Step 1-4，将 `llm_notes_from_candidates()` 返回值改为元组，调用方适配。

---

## Step 4: UI 全面重写

### 4-1 侧边栏三区块

```
📂 底稿上传 → 多文件上传 + 自动识别
🔧 检查规则 → 9条规则 checkbox + 全选/清空
🤖 LLM增强 → API Key输入 + 测试连接按钮 + 状态灯
```

### 4-2 主区域日志流

使用 `st.status` 逐条展示执行步骤：
```
📖 读取底稿结构... ✅
🔍 [1/9] Lead 基础信息... ✅ 3条
🔍 [2/9] 汇总页执行状态... ✅ 2条
...
📊 生成报告... ✅
```

### 4-3 结果展示

- 文件头卡片(文件名 + findings总数)
- 6指标卡片行
- 下载按钮行(标注底稿/Excel/HTML/CSV)
- Tab切换(全部/High/Medium)
- 质检点执行日志(可展开)
- 系统诊断 (LLM状态/耗时)

### 4-4 session_state 管理

```python
st.session_state.results  # (summary, notes, llm_status)
st.session_state.errors   # error messages
```

---

## Step 5: 导出增强

### 5-1 标注底稿增加执行日志Sheet

新增 `AI Review 执行日志` sheet，记录规则名称、执行状态、产出的finding数量。

### 5-2 Review agent 自检

每轮改完后用独立 sub-agent 对照此计划文档检查改动是否一致。

---

## 执行顺序

| Step | 内容 | 文件 | 回滚方式 |
|------|------|------|---------|
| 1 | P0误报修正 | engine.py, llm_assistant.py | `git checkout planA-backup-20260707 -- src/` |
| 2 | 新规则移植 | engine.py | 同上 |
| 3 | LLM增强 | llm_assistant.py | 同上 |
| 4 | UI重写 | app.py | 同上 |
| 5 | 导出增强 | export_review_package.py | 同上 |

每完成一步 commit 一次。
