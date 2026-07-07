# 🏆 EY AI Hackathon

> 团队成员：John · Melody · Ruby · Dana · April

## 项目简介

本项目参加 EY AI Hackathon 比赛。每位成员在各自分支上独立开发，最终择优合并。

## 分支说明

| 分支 | 成员 | 说明 |
|------|------|------|
| `main` | — | 主分支，最终合并成果 |
| `John` | John | 个人开发分支 |
| `Melody` | Melody | 个人开发分支 |
| `Ruby` | Ruby | 个人开发分支 |
| `Dana` | Dana | 个人开发分支 |
| `April` | April | 个人开发分支 |

## 协作流程

1. 每人从 `main` 切到自己的分支开发
2. 完成功能后提交并推送
3. 代码评审后合并到 `main`

## April 分支现场赛准备

本分支用于 April 在 CSDC AI 创新大赛现场实战赛中的个人开发与成果沉淀。比赛目标是在 2 小时内完成一个贴近真实业务场景的 AI POC 原型，并形成可演示、可说明、可提交的最小成果。

### 推荐目录

| 目录 | 用途 |
|------|------|
| `data/` | 放置样例输入文件 |
| `src/` | 放置核心代码 |
| `outputs/` | 放置生成结果 |
| `docs/` | 放置题目拆解、展示话术和说明文档 |

### 默认技术路线

优先使用 Python + Streamlit 快速完成交互式 POC。若现场题目更偏企业后台或产品系统，再考虑 React + Ant Design / shadcn-ui。

### 现场开发节奏

1. 前 15 分钟：明确题目、业务痛点、输入输出和 Demo 闭环。
2. 15-70 分钟：完成最小可用功能。
3. 70-100 分钟：优化界面、样例数据和结果展示。
4. 100-120 分钟：补充 README、展示话术、检查运行和提交准备。

### 交付标准

- 能用样例数据跑通完整流程。
- 能展示清晰的输入、处理逻辑和输出结果。
- README 中写明运行方式和 Demo 流程。
- 任何提交和推送前，必须先确认变更清单。

## 费用底稿 AI Review 助手

当前 April 分支已开始实现现场赛 MVP：面向销售费用、管理费用底稿，读取 SOP、审计程序要求和当前案例底稿，自动生成可定位的 Review Notes。

### 本地运行

```powershell
pip install -r requirements.txt
streamlit run src/app.py
```

如现场网络或依赖安装不稳定，优先使用零新增依赖版本：

```powershell
python src/run_review.py
python src/local_server.py
```

`run_review.py` 会生成可追溯 Review 包：

- `outputs/review_report.html`
- `outputs/review_notes.xlsx`
- `outputs/*_AI_review_annotated.xlsx`
- `outputs/review_notes.csv`（备用）

`local_server.py` 会启动本地上传页，默认地址为 `http://127.0.0.1:8765`。

### 当前覆盖

- Lead 基础信息检查。
- 汇总页程序执行状态检查。
- 法律、诉讼、咨询、中介类特殊费用提示。
- BKD 异常项目 Notes 检查。
- 截止性测试完整性检查。
- TOD 样本池与抽样输出检查。

详细方案见 `docs/solution_outline.md`。
