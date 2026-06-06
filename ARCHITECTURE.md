# 架构文档

## 系统概览

本项目包含两条共享 Excel 状态的核心流水线（书签、拆分），外加一个独立的 OneNote 收尾工具（Pipeline 4，不依赖 Excel 状态）：

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
└───────────────────────────┬─────────────────────────────┘
                            │  OneNote Batch 导入（外部，Pipeline 3）
┌───────────────────────────▼─────────────────────────────┐
│              Pipeline 4：OneNote 本地整理                  │
│  本地 COM 读分区/页 → 删占位页 · 去重 · 按文件名核对改标题  │
│  数据源是 OneNote 本地缓存 + 拆分文件夹，不读 Excel 状态    │
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
| `onenote_sync_titles.py` | Pipeline 4 入口；按文件夹核对 OneNote 页标题、删占位页、去重；默认 dry-run，`--write` 才落盘 |

### 核心库

| 模块 | 职责 |
|------|------|
| `registry.py` | 统一状态读写。优先操作 `books_config.xlsx`；Excel 不存在时回退到各书的 `state.json`（兼容旧版） |
| `ai_parser.py` | 调用 OCR 模型识别图片文字；调用 LLM 将 OCR 文本解析为结构化目录 |
| `llm_client.py` | LLM 适配层；统一封装硅基流动 / DeepSeek / Anthropic / OpenAI 四种 provider |
| `pdf_utils.py` | PDF 工具函数：渲染页面为图片、写入书签 |
| `onenote_client.py` | OneNote 桌面版 COM 接口薄封装：读层级、读/写页标题、删页（送回收站）、占位页判定 |

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

## Pipeline 4：OneNote 本地整理

把拆分文件夹用 OneNote Batch 导入后，对每个分区做收尾核对。**纯本地、离线**：通过 OneNote
桌面版 COM 接口操作本地缓存，不读 Excel 状态、不走网络（可在关闭同步时安全运行，改动后续随同步上传）。

### 数据流

```
OneNote 本地缓存                         books-done/{书名}_拆分/0N/
    │  GetHierarchy()  [onenote_client]      │  *.pdf（文件名为目标标题）
    ▼                                        ▼
笔记本 → 分区 → 页（按显示顺序）         期望标题列表（按 NNN 前缀排序）
    └──────────────┬──────────────────────────┘
                   │  分区「新分区N」⇄ 文件夹「0N」
                   ▼
        ① 删开头空白占位页（is_blank_placeholder）
        ② 去重（页数 == 2× 文件数 时删后一份）
        ③ 页数 == 文件数 → 按位置一一对齐
        ④ 当前标题 != 期望 → UpdatePageContent 改标题（幂等）
```

### 对齐规则

- **分区 ⇄ 文件夹**：`新分区N` ⇄ `0N`（解析分区名时忽略前缀后的空格，如 `新分区 7` → 7）。
- **目标标题** = PDF 文件名去扩展名（保留 `NNN-` 编号前缀与 `_N` 拆分后缀）。
- **按显示顺序对齐**：依赖 OneNote 页顺序 == 打印（文件名）顺序。页数与文件数不符的分区**中止并报警**，绝不错位改名。

### 占位页判定（`is_blank_placeholder`）

双保险，避免误删真实首页：

1. 标题属于占位标题集合（`无标题页` / `无标题` / `Untitled page` / 空 …，不区分大小写）；**且**
2. `GetPageContent` 中无 `<one:Image>`、无非空 `<one:T>`。

> 打印页必含图片，绝不会被误判为占位页。失败未改名的页则显示 OneNote 默认标题「打印输出」，会被第 ④ 步改正。

### 去重算法（`--dedupe`）

仅处理「同一批被误打印两遍」这一确定场景：

```python
if dedupe and expected and len(pages) == 2 * len(expected):
    delete(pages[len(expected):])   # 删后一份重复块（进回收站）
    pages = pages[:len(expected)]   # 保留前一份，继续对齐改标题
```

只在页数**正好是文件数 2 倍**时触发；其它数量不符仍按报警中止处理。

### COM 接入要点（`onenote_client.py`）

| 坑 / 要点 | 处理 |
|-----------|------|
| comtypes 默认晚绑定，OneNote 报 `TYPE_E_LIBNOTREGISTERED`（“库没有注册”） | 用 `GetModule(("{0EA692EE-…}",1,1))`（OneNote 15.0 类型库）强制**早绑定**，再 `CreateObject(..., interface=mod.IApplication)` |
| `DeleteHierarchy` / `UpdatePageContent` 的 DATE 参数，comtypes 默认值是 `datetime`，与 ctypes 签名（实数）冲突 | 显式传 `0.0`（= 不校验修改时间）：`DeleteHierarchy(id, 0.0, False)`、`UpdatePageContent(xml, 0.0, XS_2013, True)` |
| 删除安全性 | `deletePermanently=False` → 进 OneNote 回收站，可恢复 |
| 改标题不伤正文 | `UpdatePageContent` 只提交含 `ID` 与 `Title` 的最小 Page XML，仅替换标题，图片内容不动 |
| 中文控制台乱码 | 运行前 `$env:PYTHONUTF8=1` |

> 仅支持 OneNote 桌面版（Office16，ProgID `OneNote.Application`，CLSID `{DC67E480-…}`）；不支持 UWP「OneNote for Windows 10」。

### 安全机制

- 默认 **dry-run**，`--write` 才落盘；推荐先 `--list` / dry-run 核对对照表。
- 所有删除进**回收站**，标题改动可手动撤销。

---

## 依赖

| 包 | 用途 |
|----|------|
| `pymupdf` | PDF 渲染（页→图片）、书签写入、页面提取与切片 |
| `openai` | 调用兼容 OpenAI 接口的 provider（硅基流动、DeepSeek） |
| `anthropic` | 调用 Anthropic Claude |
| `openpyxl` | 读写 Excel 配置文件 |
| `python-dotenv` | 从 `.env` 加载 API Key |
| `comtypes` | Pipeline 4：OneNote 本地 COM 自动化（仅 Windows + OneNote 桌面版） |
