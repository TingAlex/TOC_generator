# 架构文档

## 系统概览

本项目包含两条独立的处理流水线，共享一套 Excel 状态管理：

```
┌─────────────────────────────────────────────────────────┐
│                     Pipeline 1：书签                      │
│  books-todo/*.pdf → OCR → 解析 → 书签 → books-done/*.pdf  │
└───────────────────────────┬─────────────────────────────┘
                            │  共享
            ┌───────────────▼───────────────┐
            │   books-work/books_config.xlsx │  ← 单一状态源
            │   books-work/split_config.xlsx │
            └───────────────┬───────────────┘
                            │  共享
┌───────────────────────────▼─────────────────────────────┐
│                     Pipeline 2：拆分                      │
│  books-done/*.pdf + toc_parsed.txt → 按章节拆分子 PDF      │
│  （可选）单文件超限时进一步切为 _1/_2/… 多份               │
└─────────────────────────────────────────────────────────┘
```

---

## 模块说明

### 入口脚本

| 脚本 | 职责 |
|------|------|
| `main.py` | Pipeline 1 批量入口；交互式询问目录页和偏移量；逐步推进每本书 |
| `test_one.py` | Pipeline 1 单本测试；硬编码书名，方便调试 |
| `init_work.py` | 扫描 books-todo/，将 state.json 数据迁移并写入 Excel；新书入库时重跑 |
| `split_pdf.py` | Pipeline 2 单本命令行入口；同时也暴露 `run_split()` 供 split_all 调用 |
| `split_all.py` | Pipeline 2 批量入口；读 Excel 配置，顺序处理所有未完成书本 |

### 核心库

| 模块 | 职责 |
|------|------|
| `registry.py` | 统一状态读写。优先操作 `books_config.xlsx`；Excel 不存在时回退到各书的 `state.json`（兼容旧版） |
| `ai_parser.py` | 调用 OCR 模型识别图片文字；调用 LLM 将 OCR 文本解析为结构化目录 |
| `llm_client.py` | LLM 适配层；统一封装硅基流动 / DeepSeek / Anthropic / OpenAI 四种 provider |
| `pdf_utils.py` | PDF 工具函数：渲染页面为图片、写入书签 |

---

## 数据流

### Pipeline 1（书签）

```
books-todo/{书名}.pdf
    │
    │  render_pages_to_images()          [pdf_utils.py]
    ▼
books-work/{书名}/pages/page_*.png
    │
    │  ocr_pages()                       [ai_parser.py → llm_client.py]
    ▼
books-work/{书名}/ocr_raw.txt
    │
    │  parse_toc_text()                  [ai_parser.py → llm_client.py]
    ▼
books-work/{书名}/toc_parsed.txt         ← 可手工编辑
    │
    │  write_bookmarks()                 [pdf_utils.py]
    ▼
books-done/{书名}.pdf
```

每完成一步，`registry.save()` 将 flag 写入 Excel（如 `rendered=True`）。中断重启后自动跳过已完成步骤。

### Pipeline 2（拆分）

```
books-work/{书名}/toc_parsed.txt
    │  parse_toc()                       [split_pdf.py]
    │  校验页码严格不递减
    ▼
条目列表 [(level, title, 印刷页码)]
    │  + offset（来自 Excel）
    │  → 计算 PDF 绝对页码范围
    ▼
┌── 前言页（000-书名.pdf，offset>0 时）
│
└── 各条目（001-章.pdf, 002-节.pdf ...）
        ├─ 若页数 ≤ max_pages_per_file（或未设）→ 单文件输出
        └─ 若页数 > max_pages_per_file → 切为 001-章_1.pdf, 001-章_2.pdf ...
        按 max_pages 累计页数分批装入 01/, 02/ 子文件夹

    ▼
books-done/{书名}_拆分/{01,02,...}/{序号-标题[_N]}.pdf
```

完成后 `split_all.py` 将 `拆分完成=True` 写回 Excel。

---

## 状态管理

### 主文件：`books-work/books_config.xlsx`

一行代表一本书，每次 `registry.save()` 自动同步：

| 列 | 类型 | 说明 |
|----|------|------|
| `书名` | str | PDF 文件名（无扩展名），唯一键 |
| `offset` | int | `PDF页码 = 印刷页码 + offset` |
| `toc_pages` | str | 目录所在页，逗号分隔 |
| `split_level` | int | 拆分目录深度（1/2/3） |
| `rendered` | bool | 目录页已渲染为图片 |
| `ocr_done` | bool | OCR 完成 |
| `toc_parsed` | bool | 目录解析完成 |
| `bookmarks_added` | bool | 书签已写入 |
| `bookmark_count` | int | 书签数量 |
| `拆分完成` | bool | Pipeline 2 完成 |

### 全局格式：`books-work/split_config.xlsx`

单行配置，影响所有书的拆分输出格式：

| 列 | 默认 | 说明 |
|----|------|------|
| `max_pages` | 100 | 每个批次文件夹的累计页数上限（超出换下一个文件夹） |
| `max_pages_per_file` | 空（不限） | 单个输出文件最大页数；超出时该文件切为 `_1/_2/…` 多份 |
| `prefix_digits` | 3 | 文件序号前缀位数（`001-`） |
| `prefix_sep` | `-` | 前缀分隔符 |
| `folder_digits` | 2 | 文件夹编号位数（`01`） |

> `max_pages_per_file` 的典型用途：与 OneNote Batch 插件配合时，将单个 PDF 页数控制在导入工具的上限以内（如 20 页）。

### 兜底：`books-work/{书名}/state.json`

仅在 `books_config.xlsx` 不存在时使用（`registry.py` 自动检测）。`init_work.py` 运行后，state.json 数据已迁移至 Excel，不再被写入。

---

## 目录解析格式（toc_parsed.txt）

```
{level}|{title}|{printed_page}
```

- `level`：层级（1=章, 2=节, 3=小节），最多 3 级
- `title`：标题文本
- `printed_page`：书本印刷页码（非 PDF 绝对页码）
- 同级条目页码必须不递减；严格递减时 split_pdf.py 报错退出

---

## 页码转换

```
PDF绝对页码 = 印刷页码 + offset
```

`offset` 等于书名页、版权页、前言等非正文页的数量。由用户在 Pipeline 1 运行时交互确认，并保存至 Excel。

---

## 批次文件夹分配算法

```python
folder_idx = 1
cumulative = 0
for section in sections:
    if cumulative + section.pages > max_pages and cumulative > 0:
        folder_idx += 1
        cumulative = 0
    place(section, folder_idx)
    cumulative += section.pages
```

- 当前文件夹已有内容 **且** 加入下一个文件会超限时，才创建新文件夹
- 单个超限文件（如整章 > max_pages）独占一个文件夹

## 文件级切片算法

当 `max_pages_per_file` 有值，且某章节页数超出该值时：

```python
n_parts = ceil(page_count / max_pages_per_file)
for i in range(n_parts):
    part_start = pdf_start + i * max_pages_per_file
    part_end   = min(part_start + max_pages_per_file - 1, pdf_end)
    save(part_path=f"{prefix}{title}_{i+1}.pdf", pages=(part_start, part_end))
```

文件夹分配仍按原始章节的总页数计算，切片文件均落入同一文件夹。

---

## 依赖

| 包 | 用途 |
|----|------|
| `pymupdf` | PDF 渲染（页→图片）、书签写入、页面提取与切片 |
| `openai` | 调用兼容 OpenAI 接口的 provider（硅基流动、DeepSeek） |
| `anthropic` | 调用 Anthropic Claude |
| `openpyxl` | 读写 Excel 配置文件 |
| `python-dotenv` | 从 `.env` 加载 API Key |
