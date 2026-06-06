# 架构文档

## 系统概览

本项目把 PDF 加工成带书签、可拆分、可导入 OneNote 的成果，分为几条流水线：

```
┌──────────────────────────────────────────────────────────────┐
│  Pipeline 1：书签   books-todo/*.pdf → 识别目录 → books-done/*.pdf │
│    · API 版（toc-bookmarks）：DeepSeek-OCR + DeepSeek-V3           │
│    · Claude 版（toc-claude + /toc-by-claude skill）：Claude 视觉    │
└───────────────────────────┬──────────────────────────────────┘
                            │  共享 books-work/books_config.xlsx（单一状态源）
┌───────────────────────────▼──────────────────────────────────┐
│  Pipeline 2：拆分   books-done/*.pdf + toc_parsed.txt → 章节子 PDF │
└───────────────────────────┬──────────────────────────────────┘
┌───────────────────────────▼──────────────────────────────────┐
│  Pipeline 2.5：OneNote 预建分区组 + 空分区（toc-onenote-sections） │
└───────────────────────────┬──────────────────────────────────┘
                            │  OneNote Batch 导入（外部，Pipeline 3）
┌───────────────────────────▼──────────────────────────────────┐
│  Pipeline 4：OneNote 本地整理（toc-onenote-titles / -strip）       │
│    本地 COM 读分区/页 → 删占位页 · 去重 · 按文件名核对改标题         │
└──────────────────────────────────────────────────────────────┘
```

---

## 包结构与分层

代码是一个可安装包 `src/tocgen/`，按**职责分层、依赖单向向下**（解耦 + 高内聚）：

```
cli/*            ← 薄壳：仅 argparse + 打印，不含业务逻辑
  │  调用
编排层：pipeline1（识别→书签）、split（按目录拆分）、onenote/*（建分区/改标题/删附件）
  │  调用
领域库：toc（目录模型）、pdf（渲染/书签）、llm + ai_parse（AI 识别）、onenote/client（COM）
  │  调用
基础设施：paths（路径/书名约定）、registry（每书状态）、bookconfig（Excel 模板/读写）
```

> 关键解耦点：`cli` 不写业务逻辑；所有路径常量只在 `paths` 一处定义；`level|title|page`
> 的解析/序列化/校验只在 `toc` 一处实现；三个 OneNote CLI 的共享件集中在 `onenote/common`。

### 模块职责

| 模块 | 职责 |
|------|------|
| `paths.py` | 数据目录与书名约定的**单一事实来源**（books-todo/done/work、源 PDF 解析、工作目录助手）。全部相对 CWD（命令在项目根运行） |
| `toc.py` | 目录模型：`parse_text`（AI 输出→条目，容错）、`load_file`/`dumps`/`save`、`check_nondecreasing`（页码不递减校验） |
| `pdf.py` | 无状态 PDF 工具：`parse_page_spec`、`render_pages_to_images`、`write_bookmarks`、`sanitize_filename` |
| `llm.py` | LLM 适配层：统一封装硅基流动 / DeepSeek / Anthropic / OpenAI |
| `ai_parse.py` | 调 OCR/解析模型；条目解析复用 `toc`（不再自带一份解析） |
| `registry.py` | 每书状态读写，唯一存储 `books_config.xlsx`；Excel 缺失时经 `bookconfig.ensure_books_config` 建骨架再写 |
| `bookconfig.py` | 两张 Excel 的模板、创建、读取（`books_config` + `split_config`）；`run_init` 为 `toc-init` 入口 |
| `split.py` | `run_split`：按 `toc_parsed.txt` 拆分（用 toc 校验、pdf 切页、registry 取 offset） |
| `pipeline1.py` | `process_one`：Pipeline 1 交互编排（渲染→OCR→解析→写书签），被批量/单本两个入口复用 |
| `onenote/client.py` | OneNote 桌面 COM 薄封装：读层级、建分区组/分区/在线笔记本、改/删页、列/删附件 |
| `onenote/common.py` | 三个 OneNote CLI 共享：`DEFAULT_NOTEBOOK`、`section_number`、`expected_titles`、`resolve_scope`、文件夹助手 |
| `cli/*.py` | 9 个 `toc-*` 命令入口（见 README 命令速查），各暴露 `main()` |

---

## 数据流

### Pipeline 1（书签）

```
books-todo/{书名}.pdf
    │  render_pages_to_images()          [pdf]
    ▼  books-work/{书名}/pages/page_*.png
    │  ocr_pages()                       [ai_parse → llm]   (Claude 版跳过，用视觉)
    ▼  books-work/{书名}/ocr_raw.txt
    │  parse_toc_text()                  [ai_parse → toc.parse_text]
    ▼  books-work/{书名}/toc_parsed.txt   ← 可手工编辑
    │  write_bookmarks()                 [pdf]
    ▼  books-done/{书名}.pdf
```

每完成一步 `registry.save()` 写回 Excel flag，中断重启自动跳过已完成步骤。
`process_one` 在 `pipeline1.py`，被 `cli/bookmarks.py`（批量）与 `cli/bookmarks_one.py`（单本）复用。
Claude 版（`cli/claude_toc.py`）只跑「渲染」和「写书签」，中间识图由 Claude 自身完成。

### Pipeline 2（拆分）

```
toc_parsed.txt → toc.load_file() → 按 level 过滤 → toc.check_nondecreasing()
    │  + offset（来自 Excel）→ 计算 PDF 绝对页码范围
    ▼  前言页（000-书名.pdf，offset>0 时）+ 各条目（001-章.pdf …）
       按 max_pages 累计装入 01/、02/ 子文件夹；单文件超限切为 _1/_2/… 多份
    ▼  books-done/{书名}_拆分/{01,02,…}/{序号-标题[_N]}.pdf
```

完成后 `cli/split_all.py` 把 `拆分完成=True` 写回 Excel。

---

## 状态管理

### 唯一存储：`books-work/books_config.xlsx`

一行一本书，`registry.save()` 增量同步。字段 ←→ 列见 `registry._STATE_TO_COL`，
主键统一为 `"{书名}.pdf"`。列：`书名 offset toc_pages split_level rendered ocr_done
toc_parsed bookmarks_added bookmark_count 拆分完成`。

### 冷启动：Excel 缺失时自动建表

`books-work/` 不入库（`.gitignore`），新机器克隆后 Excel 不存在。此时 `registry.save()`
会先调用 `bookconfig.ensure_books_config()` 建好仅含表头的空表骨架再写入——无需先手动跑
`toc-init`，Pipeline 1 / `toc-claude` 可直接运行。（旧版 `state.json` 兜底已移除：Excel 是唯一状态源。）

### 全局格式：`books-work/split_config.xlsx`

单行配置，影响所有书的拆分输出（`max_pages` / `max_pages_per_file` / `prefix_digits` /
`prefix_sep` / `folder_digits`）。由 `bookconfig.read_split_config()` 读取。

---

## 目录解析格式（toc_parsed.txt）

```
{level}|{title}|{printed_page}
```

- `level`：1=章, 2=节, 3=小节（最多 3 级）
- `printed_page`：书本**印刷页码**（非 PDF 绝对页码）
- 页码必须不递减；递减时 `toc.check_nondecreasing` 报错退出
- 页码转换：`PDF绝对页码 = 印刷页码 + offset`，`offset` = 正文前非正文页数

---

## 拆分算法

### 批次文件夹分配

```python
folder_idx, cumulative = 1, 0
for section in sections:
    if cumulative + section.pages > max_pages and cumulative > 0:
        folder_idx += 1; cumulative = 0
    place(section, folder_idx); cumulative += section.pages
```
当前文件夹已有内容**且**加入下一个会超限时，才新建文件夹；单个超限文件独占一个文件夹。

### 文件级切片（`max_pages_per_file` 有值且某条目超限时）

```python
n_parts = ceil(page_count / max_pages_per_file)
# 切为 001-章_1.pdf, 001-章_2.pdf …，均落入同一文件夹
```

---

## Pipeline 2.5：OneNote 预建分区组 + 空分区

拆分完成、Batch 导入前的准备步。**纯本地离线**（COM），按 `books-done/{书名}_拆分/` 下的
`0N` 文件夹数，在指定笔记本建「书名分区组」+ 同名空分区 `01…0N`。

- **作用域隔离**：分区放进「书名分区组」，`01…0N` 名被该组隔离，不与其它书的同名分区冲突；
  唯一查重的是分区组名（=书名）→ 撞名即中止，绝不改动既有内容。
- **新建在线笔记本**（`--new-notebook`）：取现有**在线**笔记本（`path` 以 `https://` 开头）的
  OneDrive 路径，求父目录拼新名，`OpenHierarchy(newUrl, "", cftNotebook)` → 落在同一云端位置。

### COM 接入要点（创建相关）

| 坑 / 要点 | 处理 |
|-----------|------|
| 创建用 `OpenHierarchy(path, relativeToObjectID, [out]objectID, cftIfNotExist)`，`[out]` 在 comtypes 早绑定下转为返回值 | 调用只传 `(path, rel, cft)`，新 ID 取返回值 |
| `CreateFileType` 枚举 | `cftNotebook=1` / `cftFolder=2`（分区组）/ `cftSection=3`（分区） |
| **分区路径必须带 `.one`**，否则 `OpenHierarchy` 抛 `COMError 0x80042004`；分区组/文件夹不带 | `create_section` 自动补 `.one`；OneNote 显示时去掉 |
| 取笔记本在线/本地 | `get_hierarchy` 读 Notebook 的 `path`；在线笔记本以 `https://` 开头 |

---

## Pipeline 4：OneNote 本地整理

把拆分文件夹用 OneNote Batch 导入后做收尾。**纯本地离线**：COM 操作本地缓存，不读 Excel、不走网络。

### 对齐规则

- **分区 ⇄ 文件夹**：去前缀后的编号 `N` ⇄ 文件夹 `0N`（`--section-prefix` 改前缀；
  配合 Pipeline 2.5 的 `01…0N` 用空前缀 `--section-prefix=`）。
- **目标标题** = PDF 文件名去扩展名（保留 `NNN-` 前缀与 `_N` 后缀）。
- **按显示顺序对齐**；页数与文件数不符的分区**中止并报警**，绝不错位改名。
- **分区组感知**：`--section-group "书名"` 把范围限定在某分区组内（`onenote/common.resolve_scope`），
  避免多本书的同名 `0N` 分区混淆。`get_hierarchy` 解析嵌套树（`Notebook.section_groups`），
  并跳过回收站（`isRecycleBin`）。

### 占位页判定（`is_blank_placeholder`）

双保险：标题属于占位集合（`无标题页` 等）**且** 页内无 `<one:Image>`、无非空 `<one:T>`。
打印页必含图片，绝不会被误判。失败未改名的页显示 OneNote 默认标题「打印输出」，由改标题步改正。

### 去重（`--dedupe`）

仅当分区页数**正好是文件数 2 倍**（误打印两遍）时，删除后一份重复块（进回收站），保留前一份继续对齐。

### 子工具：删除误插入的源文件附件（`toc-onenote-strip`）

页 XML 里附件是 `<one:InsertedFile preferredName="x.pdf">`，自身一般不带 objectID——objectID
在外层 `<one:OE>` 上，故向上找最近带 objectID 的祖先 OE 作为删除目标；若该祖先子树含
`<one:Image>`（打印图片）则跳过（绝不误删图片）。删除 `DeletePageContent(pageId, objectId, 0.0, True)`。

### COM 接入要点（读写相关）

| 坑 / 要点 | 处理 |
|-----------|------|
| comtypes 默认晚绑定，OneNote 报 `TYPE_E_LIBNOTREGISTERED`（“库没有注册”） | 用 `GetModule(("{0EA692EE-…}",1,1))`（OneNote 15.0 类型库）强制**早绑定** |
| `DeleteHierarchy` / `UpdatePageContent` / `DeletePageContent` 的 DATE 参数 | 显式传 `0.0`（= 不校验修改时间），否则 comtypes 默认 datetime 与签名冲突 |
| 删除安全性 | `deletePermanently=False` → 进 OneNote 回收站，可恢复 |
| 改标题不伤正文 | `UpdatePageContent` 只提交含 `ID` 与 `Title` 的最小 Page XML |

### OneNote 开发环境须知（换机器开发必读）

- **为什么走桌面 COM 而非 Graph**：用户刻意关闭 OneNote 同步做离线编辑；桌面 COM 操作本地缓存，
  离线免鉴权，改动随之后同步上传。Graph 则必须联网 + OAuth。
- **平台限制**：仅 Windows + OneNote 桌面版（Office16），**不支持** UWP「OneNote for Windows 10」。
- **依赖 `comtypes`（非 pywin32）**：纯 Python、免编译，Python 3.14 下更省事；COM 调用必须早绑定。
- **装依赖**：venv 由 `uv` 管理、**无 pip**，用 `uv add` / `uv sync`（或临时 `uv pip install`），别 `python -m pip`。
- **运行前 `$env:PYTHONUTF8=1`**：否则中文标题/分区名/路径乱码。
- **OneNote 默认标题**：新建分区占位页 `无标题页`（非空）；改名失败页 `打印输出`。
- **Batch 分区名带空格**：是 `新分区 N`（如 `新分区 7`），不是 `新分区7`。

---

## 依赖

| 包 | 用途 |
|----|------|
| `pymupdf` | PDF 渲染、书签写入、页面提取与切片 |
| `openai` | 兼容 OpenAI 接口的 provider（硅基流动、DeepSeek） |
| `anthropic` | Anthropic Claude |
| `openpyxl` | 读写 Excel 配置 |
| `python-dotenv` | 从 `.env` 加载 API Key |
| `comtypes` | Pipeline 2.5 / 4：OneNote 本地 COM 自动化（仅 Windows + OneNote 桌面版）。纯 Python、免编译，故选它而非 pywin32；调用须早绑定 |

构建：`hatchling`（`[build-system]`），打包 `src/tocgen`，console_scripts 见 `pyproject.toml [project.scripts]`。
