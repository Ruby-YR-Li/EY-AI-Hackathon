# 🏆 EY AI Hackathon — AI底稿Review助手（费用科目版）

> 团队成员：John · Melody · Ruby · Dana · April

本项目参加 EY AI Hackathon 比赛。Ruby 分支基于已有的固定资产Review Agent框架，改造为**销售费用&管理费用**科目审计底稿的自动化质检工具。

基于 VCVD_SOP（销售费用&管理费用审计实操指引）和模拟底稿中发现的实际问题，设计了25条费用特定检查规则。

## 功能

- 上传Excel格式费用科目底稿（U_exp SWP + TOD SWP，支持合并上传）
- 自动执行25条费用特定检查规则（纯Python + AI增强）
- 支持项目配置网页录入、Excel模板导入和TE/SAD/A3 mapping
- 支持AI模型配置（Anthropic / OpenAI / 通义千问）
- 在问题单元格添加Excel批注（Comment）
- 生成"质检报告"汇总Sheet + "核对逻辑"Sheet

## 快速开始

```bash
pip install -r requirements.txt
source .venv/bin/activate
uvicorn web.app:app --host 0.0.0.0 --port 8000
# 访问 http://localhost:8000
```

## 检查规则（25条）

| 模块 | 规则数 | 覆盖内容 |
|------|:------:|---------|
| UEXP-PSP | 2 | PSP执行完整性、拒绝理由充分性 |
| UEXP-FMT | 4 | Lead Sheet基本信息、分析日期、底稿清洁度、交叉索引 |
| UEXP-GL | 4 | Lead与BKD核对、A3签核、TE/SAD值、TE/SAD/CRA配置核对 |
| UEXP-BKD | 7 | 预期记录、排序、波动标记、定性标记、ARP分析、波动说明、X-ref |
| UEXP-TOD | 2 | SCOT区分、抽样策略X-ref |
| UEXP-CO | 3 | 截止测试期间、策略、样本金额 |
| UEXP-AI | 3 | 异常波动AI判断、费用分类、法律费用关注性 |

## 分支说明

| 分支 | 成员 | 说明 |
|------|------|------|
| `main` | — | 主分支，最终合并成果 |
| `Ruby` | Ruby | 费用底稿Review工具 |
| `John` | John | 个人开发分支 |
| `Melody` | Melody | 个人开发分支 |
| `Dana` | Dana | 个人开发分支 |
| `April` | April | 个人开发分支 |
