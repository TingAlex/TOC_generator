# PDF 自动识别目录 · 添加书签 · 按章节拆分

将 PDF 中的目录页用 AI 识别，自动生成 PDF 书签（大纲），再按目录层级把整本 PDF 拆分成独立章节文件。适用于扫描版教材、无书签的电子书等场景。

---

## 整体流程

```
books-todo/*.pdf
      │
      ▼  python main.py          【Pipeline 1：书签】
      │
      ├─[1] 渲染目录页为图片 (300 DPI)
      ├─[2] DeepSeek-OCR 逐张识别
      ├─[3] DeepSeek-V3 解析为目录结构 → books-work/{书名}/toc_parsed.txt
      └─[4] 写入 PDF 书签 (--write)  → books-done/{书名}.pdf
      │
      ▼  python init_work.py     【初始化配置 Excel】
      │
      books-work/books_config.xlsx   ← 进度监控 + 每本书配置
      books-work/split_config.xlsx   ← 全局拆分格式
      │
      ▼  python split_all.py     【Pipeline 2：拆分】
      │
      books-done/{书名}_拆分/
          01/  001-第一章.pdf
               002-第一节.pdf  ...
          02/  ...
      │
      ▼  OneNote Batch 插件      【Pipeline 3：导入（外部）】
         将 books-done/{书名}_拆分/ 导入 OneNote
```

所有进度和配置统一由 `books-work/books_config.xlsx` 管理，可直接用 Excel 查看和修改。

---

## 环境要求

- **Python 3.12+**（项目使用 3.14，建议通过 `uv` 自动管理）
- **uv**（Python 包管理器）
- **API Key**（任选其一）：
  - 硅基流动（推荐，免费注册）：https://cloud.siliconflow.cn
  - DeepSeek、Anthropic 或 OpenAI

---

## 安装

```powershell
# 1. 安装 uv（如未安装）
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# 2. 克隆项目
git clone https://github.com/TingAlex/TOC_generator.git
cd TOC_generator

# 3. 安装依赖（uv 自动管理 Python 版本和虚拟环境）
uv sync

# 4. 配置 API Key
Copy-Item .env.example .env
notepad .env   # 填入 SILICONFLOW_API_KEY=sk-你的密钥
```

---

## 使用方法

### Pipeline 1：识别目录 + 添加书签

```powershell
# 将 PDF 放入 books-todo/
# dry-run：跑识别，检查结果，不写文件
uv run python main.py

# 确认无误后，正式写入书签
uv run python main.py --write
```

运行时程序会交互式询问：
1. 目录页范围（如 `7`、`7-9`、`7,8,9`）
2. 偏移量确认（PDF 实际页码 vs 印刷页码）

> **提示**：若已知目录页和偏移量，可预先在 `books_config.xlsx` 的 `toc_pages` / `offset` 列填好，程序将跳过对应询问，实现无交互批量运行。

识别完成后，`books-work/{书名}/toc_parsed.txt` 可手工编辑修正，格式为：

```
1|第一章 标题|12
2|第一节 小节|12
3|1.1.1 细目|15
```

### 初始化配置 Excel

```powershell
# 扫描所有书，生成/更新 Excel 配置（新书入库时重跑）
uv run python init_work.py
```

生成两个 Excel 文件，用 Excel 直接打开编辑：

| 文件 | 内容 |
|------|------|
| `books-work/books_config.xlsx` | 每本书一行：进度 flag、offset、拆分层级、是否完成 |
| `books-work/split_config.xlsx` | 全局格式：文件夹大小上限、单文件页数上限、前缀位数等 |

### Pipeline 2：按章节拆分 PDF

```powershell
# 预览待处理书目（不实际运行）
uv run python split_all.py --dry-run

# 批量拆分所有未完成的书
uv run python split_all.py

# 只拆某一本
uv run python split_all.py --book "必修第四册"
```

拆分完成后，`books_config.xlsx` 中该书的"拆分完成"列自动标为 True。

#### 单本调试（命令行）

```powershell
uv run python split_pdf.py "书名" --level 3 --max-pages 100
```

| 参数 | 默认 | 说明 |
|------|------|------|
| `--level` | 3 | 拆分最大层级（1=章, 2=节, 3=小节） |
| `--max-pages` | 100 | 每个文件夹总页数上限（超出换下一个文件夹） |
| `--max-pages-per-file` | 不限 | 单个输出文件最大页数，超出则切为 `_1/_2/…` 多份 |
| `--prefix-digits` | 3 | 前缀位数（3 → `001-`） |
| `--prefix-sep` | `-` | 前缀分隔符 |
| `--folder-digits` | 2 | 文件夹编号位数（2 → `01`） |
| `--offset` | 自动 | 手动覆盖页码偏移量 |

#### 与 OneNote Batch 配合使用

拆分完成后，将 `books-done/{书名}_拆分/` 作为源目录交给 OneNote Batch 插件导入。

若需控制每次导入的页数（如 OneNote 每页限 20 页），在 `split_config.xlsx` 的 `max_pages_per_file` 列填入目标值（如 `20`），超限章节将自动切为 `001-章节_1.pdf`、`001-章节_2.pdf` 等多份，每份均不超过该限制。

---

## 进度管理（Excel）

`books-work/books_config.xlsx` 是项目的统一监控面板：

| 列名 | 说明 |
|------|------|
| `书名` | PDF 文件名（无扩展名） |
| `offset` | `PDF页码 = 印刷页码 + offset` |
| `toc_pages` | 目录所在页（逗号分隔） |
| `split_level` | 拆分目录深度（1/2/3） |
| `rendered` | 目录页是否已渲染 |
| `ocr_done` | OCR 是否完成 |
| `toc_parsed` | 目录结构是否已解析 |
| `bookmarks_added` | 书签是否已写入 PDF |
| `bookmark_count` | 书签数量 |
| `拆分完成` | PDF 是否已拆分 |

**重做某步**：在 Excel 中将对应列改为 False，下次运行自动从该步重做。

**常用场景：**
- 手工修正 `toc_parsed.txt` 后重新写入书签 → 将 `bookmarks_added` 改为 False
- 重新跑 OCR → 将 `ocr_done`、`toc_parsed`、`bookmarks_added` 都改为 False
- 重新拆分 → 将 `拆分完成` 改为 False

---

## 项目结构

```
.
├── main.py              # Pipeline 1：OCR + 书签（批量）
├── test_one.py          # Pipeline 1：单本测试
├── init_work.py         # 初始化/更新 Excel 配置
├── split_pdf.py         # Pipeline 2：单本拆分（命令行）
├── split_all.py         # Pipeline 2：批量拆分（读 Excel）
│
├── ai_parser.py         # OCR + 目录解析逻辑
├── llm_client.py        # LLM 客户端适配层
├── pdf_utils.py         # PDF 渲染与书签写入
├── registry.py          # 状态管理（读写 Excel / state.json 兜底）
│
├── pyproject.toml       # 项目依赖
├── .env.example         # API Key 配置模板
│
├── books-todo/          # 放入待处理 PDF（不入库）
├── books-done/          # 处理完成的 PDF（不入库）
│   └── {书名}_拆分/     # 拆分输出（不入库）
└── books-work/          # 中间产物与配置（不入库）
    ├── books_config.xlsx    # 统一配置与进度监控
    ├── split_config.xlsx    # 拆分格式默认值
    └── {书名}/
        ├── toc_parsed.txt   # 解析后目录（可手工编辑）
        ├── ocr_raw.txt      # OCR 原始输出
        └── pages/           # 渲染图片
```

---

## 模型说明

| 用途 | 默认模型 | Provider |
|------|----------|----------|
| OCR 识别 | `deepseek-ai/DeepSeek-OCR` | 硅基流动（限时免费） |
| 目录解析 | `deepseek-ai/DeepSeek-V3` | 硅基流动（按量计费） |

在 `.env` 中设置 `LLM_MODEL=模型名` 可覆盖解析模型。

**多 Provider 优先级：**
`SILICONFLOW_API_KEY` > `DEEPSEEK_API_KEY` > `ANTHROPIC_API_KEY` > `OPENAI_API_KEY`

使用非硅基流动 provider 时，OCR 与解析合并为单次视觉模型调用。
