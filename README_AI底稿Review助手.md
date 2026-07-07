# AI审计底稿Review助手

## 快速启动

如需启用 DeepSeek LLM 增强，请先在当前终端设置环境变量：

```powershell
$env:DEEPSEEK_API_KEY="你的 DeepSeek API Key"
```

可选配置：

```powershell
$env:DEEPSEEK_BASE_URL="https://api.deepseek.com/chat/completions"
$env:DEEPSEEK_MODEL="deepseek-chat"
```

在当前目录运行：

```powershell
python app.py
```

浏览器打开：

```text
http://127.0.0.1:8765
```

## 推荐演示路径

1. 点击“使用模拟问题底稿”，系统会自动载入 `sample/模拟问题底稿_AIReviewDemo.xlsx`。
2. 查看仪表盘中的 Review Notes 数量和 High / Medium / Low 风险分布。
3. 使用筛选器查看四类 checkpoint：
   - 基础完整性检查
   - 重点审计程序检查
   - 逻辑与引用检查
   - 专业提示与优化建议
4. 展示每条 Review Note 的 Sheet + Cell 定位、审计风险和建议修改。
5. 点击“导出 Review Notes”和“下载高亮底稿”，展示可交给 Senior 整改的输出物。

## 主要文件

- `app.py`：本地网页服务。
- `review_engine.py`：Excel 解析、规则检查、导出和高亮逻辑。
- `llm_client.py`：DeepSeek API 调用、重试和 JSON 解析。
- `llm_review.py`：规则结果与 LLM 质检结果合并。
- `rules/vcvd_checkpoints.json`：VCVD SOP checkpoint 规则库。
- `sample/模拟问题底稿_AIReviewDemo.xlsx`：演示用问题底稿。
- `outputs/`：Review Notes 和高亮底稿输出目录。
- `演示话术与项目亮点.md`：现场讲解稿。

## 设计思路

底层检查采用规则化方式，保证现场稳定；设置 DeepSeek API Key 后，系统会实时调用 LLM 对规则结果进行误判删除、漏判补充和专业措辞增强。未设置 Key 或 LLM 异常时，系统会自动回退到规则结果。

双子代理质检只用于开发阶段打磨 Prompt 和规则脚本，不作为网页功能展示。
