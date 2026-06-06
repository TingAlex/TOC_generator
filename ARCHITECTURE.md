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
                            │
┌───────────────────────────▼─────────────────────────────┐
│              Pipeline 2.5：OneNote 预建分区组              │
│  按 0N 文件夹数，本地 COM 建「书名分区组」+ 空分区 01…0N   │
│  （可选 --new-notebook 新建在线笔记本，同级于现有笔记本）  │
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
| `claude_toc_helper.py` | Pipeline 1（Claude 版）辅助：`render`（渲染目录页）/`bookmarks`（写书签）两个**不调用 AI API** 的子命令；配合 skill `/toc-by-claude`，由 Claude 自身多模态能力替代 DeepSeek-OCR/V3 |
| `init_work.py` | 扫描 books-todo/（与 books-done/）登记新书到 Excel，并生成 split_config.xlsx；新书入库时重跑。亦暴露 `ensure_books_config()` 供 registry 冷启动建表 |
| `split_pdf.py` | Pipeline 2 单本命令行入口；同时也暴露 `run_split()` 供 split_all 调用 |
| `split_all.py` | Pipeline 2 批量入口；读 Excel 配置，顺序处理所有未完成书本 |
| `onenote_create_sections.py` | Pipeline 2.5 入口；按 `0N` 文件夹数在指定笔记本建「书名分区组」+ 空分区 `01…0N`；分区组查重防重名；可 `--new-notebook` 新建在线笔记本；默认 dry-run，`--write` 才落盘 |
| `onenote_sync_titles.py` | Pipeline 4 入口；按文件夹核对 OneNote 页标题、删占位页、去重；默认 dry-run，`--write` 才落盘 |
| `onenote_strip_files.py` | Pipeline 4 子工具；从指定分区删除「误插入的源文件附件」（默认仅 .pdf），保留打印图片；默认 dry-run，`--write` 才删 |

### 核心库

| 模块 | 职责 |
|------|------|
| `registry.py` | 统一状态读写。以 `books_config.xlsx` 为唯一存储；Excel 不存在时 `save()` 自动经 `init_work.ensure_books_config()` 建空表骨架再写入 |
| `ai_parser.py` | 调用 OCR 模型识别图片文字；调用 LLM 将 OCR 文本解析为结构化目录 |
| `llm_client.py` | LLM 适配层；统一封装硅基流动 / DeepSeek / Anthropic / OpenAI 四种 provider |
| `pdf_utils.py` | PDF 工具函数：渲染页面为图片、写入书签 |
| `onenote_client.py` | OneNote 桌面版 COM 接口薄封装：读层级（含笔记本 `path`）、读/写页标题、删页/删层级（送回收站）、占位页判定、列出/删除页内嵌入附件；**创建**分区组/分区/在线笔记本（`OpenHierarchy`） |

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

### 冷启动：Excel 缺失时自动建表

`books-work/` 不入库（`.gitignore`），故新机器克隆后 `books_config.xlsx` 不存在。
此时 `registry.save()` 会先调用 `init_work.ensure_books_config()` 建好仅含表头的空表骨架，
再把书本状态写入——无需先手动跑 `init_work.py`，Pipeline 1 / `claude_toc_helper.py` 可直接运行。
（旧版的 `state.json` 兜底已移除：Excel 是唯一状态源。）

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

## Pipeline 2.5：OneNote 预建分区组 + 空分区

拆分完成后、Batch 导入前的准备步。**纯本地离线**（OneNote 桌面 COM），按 `books-done/{书名}_拆分/`
下的 `0N` 文件夹数量，在指定笔记本建「以书名命名的分区组」+ 同名空分区 `01…0N`。

### 数据流

```
books-done/{书名}_拆分/0N/          OneNote 本地缓存
    │  统计 0N 子文件夹               │  get_hierarchy() / get_section_groups()
    ▼                               ▼
分区名列表 [01…0N]              目标笔记本（按名查找；可 --new-notebook 新建在线）
    └───────────────┬───────────────┘
                    │  ① 分区组查重：已存在同名（书名）→ 中止报警
                    │  ② create_section_group(notebookId, 书名)
                    │  ③ 逐个 create_section(groupId, "0N")
                    ▼
        笔记本 → 分区组「书名」 → 空分区 01…0N
```

### 关键设计

- **作用域隔离**：分区放进「书名分区组」，`01…0N` 名被该组隔离，不会与其它书的同名分区冲突；
  唯一查重的是分区组名（=书名）本身 → 撞名即中止，绝不改动既有内容。
- **新建在线笔记本**（`--new-notebook`）：取一个现有**在线**笔记本（`path` 以 `https://` 开头）的
  OneDrive 路径，求父目录拼上新名，`OpenHierarchy(newUrl, "", cftNotebook)` → 落在同一云端位置，
  而非本地笔记本。用 URL 的 `/` 分隔，不可用本地 `os.path`。
- **增量、可逆**：只新增空对象；默认 dry-run，`--write` 才落盘；删除（如撤销）均进回收站。

### COM 接入要点（创建相关，补充 onenote_client.py）

| 坑 / 要点 | 处理 |
|-----------|------|
| 创建用 `OpenHierarchy(path, relativeToObjectID, [out]objectID, cftIfNotExist)`，`[out]` 在 comtypes 早绑定下转为返回值 | 调用只传 `(path, rel, cft)`，新 ID 取返回值 |
| `CreateFileType` 枚举 | `cftNotebook=1` / `cftFolder=2`（分区组）/ `cftSection=3`（分区） |
| **分区路径必须带 `.one` 扩展名**，否则 `OpenHierarchy` 抛 `COMError 0x80042004`；分区组/文件夹则**不带**扩展名 | `create_section` 自动补 `.one`；OneNote 显示时去掉 |
| 取笔记本是否在线 | `get_hierarchy` 读出 Notebook 的 `path` 属性；在线笔记本以 `https://` 开头 |

> 与 Pipeline 4 的兼容：分区名 `01…0N`，Batch 导入后
> `onenote_sync_titles.py --section-group "书名" --section-prefix ""` 即可按 `0N ⇄ 文件夹` 对齐改标题。
> Pipeline 4 已做**分区组感知**（见下）：`--section-group` 把处理范围限定在某分区组内，
> 避免多本书的同名 `0N` 分区混淆。

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

### OneNote 开发环境须知（换机器开发必读）

这些是踩过坑、单看代码不易复原的前提，换台电脑继续开发时务必照做：

- **为什么走桌面 COM 而非 Graph API**：用户**刻意关闭 OneNote 同步**做离线编辑（避免导入/改名时的同步冲突）。
  桌面 COM 操作**本地缓存**，离线可用、无需网络与鉴权；改动落本地，等重开同步再上传。Graph 则必须联网 + OAuth。
- **平台限制**：仅 **Windows + OneNote 桌面版**（Microsoft 365 / OneNote 2016，Office16）。
  **不支持**「OneNote for Windows 10」UWP 版（没有这套 COM 接口）。
- **依赖用 `comtypes`（非 pywin32）**：comtypes 是纯 Python、开箱即用；本项目 Python 3.14 下 pywin32 装轮子麻烦。
  COM 调用**必须早绑定**（见上表的类型库 GUID），否则 OneNote 报「库没有注册」。
- **装依赖**：本项目 venv 由 `uv` 管理，**venv 里没有 pip**。加包请用 `uv add <pkg>`（写进 pyproject）后 `uv sync`，
  或临时 `uv pip install <pkg>`；不要直接 `python -m pip`。
- **运行前 `$env:PYTHONUTF8=1`**：否则中文页标题/分区名/书名路径在控制台和传参时易乱码。
- **OneNote 默认标题约定**（判定逻辑依赖它们）：新建分区的空白占位页标题是 `无标题页`（**非空**）；
  改名失败的打印页标题是 OneNote 默认的 `打印输出`。
- **Batch 生成的分区名带空格**：是 `新分区 N`（如 `新分区 7`），不是 `新分区7`——按名匹配时注意。

### 子工具：删除误插入的源文件附件（`onenote_strip_files.py`）

OneNote Batch 导入时即便取消勾选「插入 PDF 源文件」，仍可能把源 PDF 作为**附件**嵌进每一页，
使笔记本体积暴涨。本子工具从**按名指定的分区**（`--sections 名1,名2`）逐页删除这些附件，
默认只删 `.pdf`（`--ext` 可改），**保留打印页图片**。

- 识别：页 XML 里附件是 `<one:InsertedFile preferredName="x.pdf">`，自身一般不带 objectID，
  objectID 在外层 `<one:OE>` 上 → 向上找最近带 objectID 的祖先 OE 作为删除目标
  （`onenote_client.list_inserted_files`）。
- **安全护栏**：若该祖先 OE 子树内含 `<one:Image>`（打印图片），则跳过，绝不连带删图片。
- 删除：`DeletePageContent(pageId, objectId, 0.0, True)`（DATE 参数同样传 `0.0`）。
- 体积回收由 OneNote 后台压缩完成，可能略有延迟。

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
| `comtypes` | Pipeline 2.5 / 4：OneNote 本地 COM 自动化（仅 Windows + OneNote 桌面版）。纯 Python、免编译，故选它而非 pywin32（Python 3.14 下更省事）；调用须早绑定，详见「OneNote 开发环境须知」 |
