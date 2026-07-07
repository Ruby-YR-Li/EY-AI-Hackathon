# 费用底稿 AI Review 助手方案

## 一句话方案

基于销售费用、管理费用标准 SOP 和当前底稿案例，自动识别底稿缺失、程序执行不一致、特殊费用风险和 Review Note 缺失，并输出可定位的 Review Notes。

## 目标用户

- Senior：快速发现底稿问题并定位到 sheet、单元格或行。
- Manager：复核程序执行是否完整，判断是否存在需要追问的审计风险。

## 输入资料

| 输入 | 用途 |
| --- | --- |
| 主底稿 `U_exp SWP VC&VD 20251231 .xlsx` | 检查 Lead、汇总页、销售费用 BKD、管理费用 BKD、截止性测试。 |
| TOD 底稿 `TOD SWP U_Exp 20251231 .xlsx` | 检查样本池、剔除后样本池和抽样工具输出。 |
| SOP Excel `FY26_SOP U_exp SWP VC&VD.xlsx` | 提供标准底稿、Review checklist 和程序要求。 |
| 审计程序要求 `审计程序要求.docx` | 映射应执行程序和 Review 规则来源。 |
| 科目余额表、序时账 | 后续用于金额勾稽、样本池和特殊费用筛选。 |

## MVP 规则覆盖

| 规则 | 检查内容 | 价值 |
| --- | --- | --- |
| Lead 基础信息 | 客户、期末、TE/SAD、会计准则、记账本位币 | 发现基础填列错误，避免底稿基础信息不可靠。 |
| 汇总页执行状态 | 程序是否执行、不执行是否有原因 | 对齐 PSP 执行完整性要求。 |
| 法律费用一致性 | 法律、诉讼、咨询、中介类费用与 VD.01.3 是否冲突 | 发现潜在审计风险和程序遗漏。 |
| BKD Notes | 异常波动或定性异常是否有 Notes | 防止已识别异常没有调查记录。 |
| 截止性测试 | 凭证、证据、交易发生日期、期间标记、EY Works | 检查截止性测试字段完整性。 |
| TOD 样本池 | 总样本池、剔除后样本池、Skywind 输出 | 检查样本总体和抽样证据是否保留。 |

## 输出

工具输出 Review Notes 表格，字段包括：

- `risk_level`
- `file`
- `sheet`
- `location`
- `issue_type`
- `review_note`
- `suggested_action`
- `source`

支持下载 CSV，方便复制到底稿 Review Notes 或提交给项目组。

## 当前案例可展示发现

1. `Uexp.00 Lead!C7` 为 `CNY`，适用会计准则疑似填成货币。
2. `Uexp.00 Lead!C8` 为 `IFRS`，记账本位币疑似填成准则。
3. 管理费用 BKD 中存在咨询费、法律服务费、诉讼费、中介服务费等特殊费用。
4. 汇总页法律费用程序被标记为不执行，理由为本期无法律费用，和底稿线索可能不一致。

## 演示流程

1. 展示输入资料清单：主底稿、TOD 底稿、SOP、程序要求。
2. 点击“开始 Review”。
3. 展示风险概览：总 Review Notes、High/Medium、涉及 sheet 和规则数。
4. 展示 Review Notes 表格，重点讲基础信息填反和法律费用程序不一致。
5. 下载 CSV，说明可复制到底稿 Review Notes。

## 2 小时内不做的范围

- 不做 OCR。
- 不做真实 LLM Agent 编排。
- 不做数据库和权限控制。
- 不回写原始底稿。
- 不替代审计判断，只辅助 Senior / Manager 快速 Review。

## 时间充裕时的扩展方案

### 1. 规则覆盖增强

- 从 SOP Excel 的 `Review checklist` 自动抽取规则字典。
- 增加 Lead 与 BKD 金额勾稽。
- 增加 BKD 与科目余额表核对。
- 增加 TOD 样本池与剔除后样本池金额核对。
- 增加序时账中特殊摘要关键词筛选。

### 2. 输出增强

- 导出 Excel 格式 Review Notes。
- 生成 HTML Review 报告。
- 生成带标注的底稿副本。
- 按 High / Medium / Low 自动排序。

### 3. AI 能力增强

- 使用 LLM 将规则结果改写成更自然的 Manager Review Note。
- 基于 SOP 文本给出引用依据。
- 对不执行理由做充分性判断。
- 对异常费用摘要生成风险解释。

### 4. 产品体验增强

- 支持上传自定义底稿。
- 支持规则开关。
- 支持按 sheet、风险等级、程序编号筛选。
- 增加“复制 Review Note”按钮。

## 技术路线

- `src/expense_review_engine.py`：规则引擎。
- `src/app.py`：Streamlit 页面。
- `requirements.txt`：依赖清单。

该结构借鉴固定资产质检 agent 的模块思路，但控制在比赛可完成范围内。
