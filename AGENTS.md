# Agent 协作与文件定位说明

## 沟通要求

- 所有对话使用中文。
- 每次对话最后补充“公主请查收”。
- 完成功能、修复问题、改完一组相关文件或准备切换任务时，提供简短阶段交接。
- 每次开始任务前先理解需求，不要立即写代码。先分析、说明方案，经确认后实施。

## “下班”自动收尾流程

当用户明确说“下班”“今天下班”“收工下班”等表示结束当天工作的指令时，默认自动执行收尾流程，不再额外询问是否提交：

1. 先检查工作区变更，确认不把真实底稿、质检结果、API密钥、缓存、虚拟环境或打包中间文件纳入提交。
2. 结合本轮实际修改，更新项目文档和进度台账：
   - `docs/项目文档.md`：补充或修正功能说明、规则口径、报告展示口径等长期说明。
   - `docs/项目进度台账.md`：追加当天重要变更、验证结果、影响范围和提交信息占位。
3. 运行可用的自动化测试或最小验证；如果测试无法运行，需在提交说明和回复中注明原因。
4. 使用非交互式 Git 命令完成暂存和提交，提交信息用中文概括本次改动。
5. 提交后推送到 GitHub 当前分支；如果因为网络、权限或远端配置失败，保留本地提交并在回复中说明失败原因和下一步处理建议。
6. 最后回复阶段交接：概括改了什么、测试结果、提交哈希、推送状态，以及是否有未提交/未推送遗留项。

## 质检结果快速定位（重要）

用户询问“最新测试的质检结果”“刚生成的报告”或“上传的测试底稿”时，不要先扫描整个工作区。

默认数据根目录：

```text
%LOCALAPPDATA%\QualityAgent
```

常用位置：

```text
最新质检报告：%LOCALAPPDATA%\QualityAgent\质检结果
上传测试底稿：%LOCALAPPDATA%\QualityAgent\待质检底稿
AI配置：      %LOCALAPPDATA%\QualityAgent\config\ai_config.json
项目配置：    %LOCALAPPDATA%\QualityAgent\config\projects.json
运行日志：    %LOCALAPPDATA%\QualityAgent\logs
```

PowerShell查找最新报告：

```powershell
Get-ChildItem "$env:LOCALAPPDATA\QualityAgent\质检结果" -File |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 10 FullName, LastWriteTime, Length
```

如果设置了环境变量 `QUALITY_AGENT_DATA_DIR`，应优先使用该目录，而不是
`%LOCALAPPDATA%\QualityAgent`。

只有默认目录不存在或不可写时，开发环境才可能回退到：

```text
<项目根目录>\.local_data
```

工作区整理前的历史底稿、历史报告和旧发布版本位于：

```text
%LOCALAPPDATA%\QualityAgent\workspace_archive\2026-06-23_整理前备份
```

工作区根目录不再作为日常质检结果保存位置。

## 工作区约定

- 根目录的 `启动质检网页.bat` 必须保留，方便用户直接找到。
- 核心代码在 `src/`，网页代码在 `web/`，测试在 `tests/`。
- 项目文档在 `docs/`。
- 规则清单工具在 `tools/rules/`。
- 当前发布包在 `release/`。
- 权威规则清单是 `资料库/检查规则清单.xlsx`。
- 不要把真实底稿、质检结果、API密钥、缓存或打包中间文件提交到Git。
