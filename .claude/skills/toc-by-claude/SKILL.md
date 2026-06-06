---
name: toc-by-claude
description: 用 Claude 自身的多模态能力替代项目的 DeepSeek-OCR/V3 API 来完成 Pipeline 1（识别目录→写 toc_parsed.txt→加书签）。当用户想「绕过 AI API，直接用 Claude 看图识别目录并加书签」时使用。
---

# Pipeline 1（Claude 版）：绕过 AI API，用 Claude 看图识别目录并加书签

本 skill 固定一条流程：**不调用**项目里的 DeepSeek-OCR / DeepSeek-V3（即
`ai_parser.py` / `llm_client.py` 那条 API 链路），改由**你（Claude）自己的多模态能力**
读取渲染出的目录页图片，直接产出结构化目录 `toc_parsed.txt`，再写入 PDF 书签。

渲染和写书签这两个纯本地步骤由 `claude_toc_helper.py` 完成（它**不碰任何 API**）；
中间「看图识目录」这一最费 token 的部分，**派发给一个子任务（subagent）专门完成**，
让主对话保持清爽。

> 运行前置：`$env:PYTHONUTF8=1`（中文不乱码）。所有命令用 `uv run python ...`。
> 产物与进度沿用现有约定：`books-work/{书名}/`、`books_config.xlsx`（详见 ARCHITECTURE.md）。

---

## 步骤

### 0. 确定输入（书名 / 目录页 / 偏移量）
- **书名**：`books-todo/` 里的 PDF（无扩展名）。用户没说就列出 `books-todo/*.pdf` 让其选。
- **目录页范围**：PDF 中目录所在的**实际页**（如 `2-4` 或 `2,3,4`）。
- **偏移量 offset**：`PDF页码 = 印刷页码 + offset`（=书名页/版权页/前言等正文前页数）。
- 若 `books-work/books_config.xlsx` 里该书已填 `toc_pages` / `offset`，可直接复用、不必再问。
  缺失的才向用户询问。

### 1. 渲染目录页（不调用 API）
```powershell
$env:PYTHONUTF8=1
uv run python claude_toc_helper.py render "书名" --pages 2-4
```
输出 `books-work/{书名}/pages/page_NNN.png`，并记录 `toc_pages` + `rendered`。
命令会把生成的 PNG 路径逐行打印出来——记下这些路径，传给下一步的子任务。

### 2. 【子任务】看图识目录 → 写 toc_parsed.txt
用 **Agent 工具**派发一个 `general-purpose` 子任务（这是用户要求的「建立一个子任务」），
让它**用自己的视觉能力**读取第 1 步的 PNG，转写并结构化为目录，写出
`books-work/{书名}/toc_parsed.txt`。给子任务的 prompt 必须包含：

- 要读取的 PNG 绝对路径列表（来自第 1 步输出）。
- 输出文件路径：`books-work/{书名}/toc_parsed.txt`。
- **严格的行格式**：每行 `{level}|{title}|{印刷页码}`
  - `level`：1=章, 2=节, 3=小节（最多 3 级）；按目录缩进/编号层级判断。
  - `title`：目录条目标题原文（保留章节号，如 `第一章 集合`、`1.1 命题`）。
  - `印刷页码`：条目右侧书本**印刷页码**（不是 PDF 绝对页码）。
- 约束：**逐条转写、不要遗漏也不要臆造**；同级条目页码必须**不递减**（递减说明看错，需复核）；
  忽略「目录」二字标题、纯装饰行；不要输出表头或多余文字，**只写数据行**。
- 让子任务在写完后回报：总条目数、各级条目数、页码是否单调，便于主任务核对。

> 为什么用子任务：把「逐页读图+转写」这种高 token、机械的工作隔离到子 agent，
> 主对话只拿回结论（条目数/校验结果），上下文更干净——这正是用户的诉求。

### 3. 复核 toc_parsed.txt
子任务返回后，主任务读 `books-work/{书名}/toc_parsed.txt` 快速自检：
- 行格式合法（三段、中间用 `|`）。
- 同级页码不递减；层级合理（不出现孤立的 L3 等）。
- 条目数与目录页观感大致相符。
- 有问题就让子任务修正，或直接小改该文本文件（它本就是可手工编辑的）。
- 可把目录预览给用户确认后再继续。

### 4. 写入书签（不调用 API）
```powershell
uv run python claude_toc_helper.py bookmarks "书名" --offset 18
```
读 `toc_parsed.txt` → 写 `books-done/{书名}.pdf`，并记录
`offset / ocr_done / toc_parsed / bookmarks_added / bookmark_count`。
（`ocr_done`、`toc_parsed` 在这里标记为完成——因为这两步已由 Claude 顶替 API 做掉了。）

### 5. 收尾 / 交接
告诉用户已完成，并提示后续流程（不属于本 skill）：
1. `uv run python init_work.py`（首次入库或新书时刷新 Excel 配置）
2. `uv run python split_all.py --book "书名"`（按目录拆分）
3. `uv run python onenote_create_sections.py --book "书名" --write`（建分区组+空分区）
4. 用户手动 OneNote Batch 导入 → 之后 `onenote_sync_titles.py` 收尾。

---

## 注意
- **绝不**调用 `main.py`（它带 API-Key 校验且会真去跑 DeepSeek OCR）；本 skill 全程只用
  `claude_toc_helper.py` + Claude 自己的视觉，无需任何 `*_API_KEY`。
- PowerShell 传空参数/特殊字符注意引号；书名含括号时用引号包裹。
- 若某步已完成（registry 标记为 True）想重做，把 `books_config.xlsx` 对应列改回 False，或删 `toc_parsed.txt` 后重跑。
- 全流程产物可逆、可手工编辑；`--offset` 拿不准时先渲染目录页看「某条目印刷页码 ↔ 它在 PDF 的实际页」推算。
