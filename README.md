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
      ▼  python onenote_create_sections.py   【Pipeline 2.5：导入前准备】
         按 0N 文件夹数，在指定笔记本建「以书名命名的分区组」+ 对应空分区 01…0N
         （可选 --new-notebook 新建在线笔记本）
      │
      ▼  OneNote Batch 插件      【Pipeline 3：导入（外部）】
         将 books-done/{书名}_拆分/ 各文件夹导入对应分区
      │
      ▼  python onenote_sync_titles.py   【Pipeline 4：OneNote 本地整理】
         核对/改正页标题 · 删除新分区占位页 · 去重（误打印两遍）
         （另：onenote_strip_files.py 删除导入时误插入的 PDF 源文件附件，瘦身笔记本）
```

所有进度和配置统一由 `books-work/books_config.xlsx` 管理，可直接用 Excel 查看和修改。

---

## 环境要求

- **Python 3.12+**（项目使用 3.14，建议通过 `uv` 自动管理）
- **uv**（Python 包管理器）
- **API Key**（任选其一）：
  - 硅基流动（推荐，免费注册）：https://cloud.siliconflow.cn
  - DeepSeek、Anthropic 或 OpenAI
- **（仅 Pipeline 4 需要）** Windows + OneNote 桌面版（Office16，非 UWP 版）；依赖 `comtypes`（已在 `pyproject.toml`，`uv sync` 自动安装）

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

### Pipeline 2.5：OneNote 预建分区组 + 空分区（导入前准备）

拆分完成后、用 OneNote Batch 导入**之前**，本工具按 `books-done/{书名}_拆分/` 下的 `0N` 文件夹数量，
在指定笔记本里建一个**以书名命名的分区组**，并在组内建好对应数量的空分区（命名与文件夹同名：`01`…`0N`）。
这样你只需做 Batch 插入，其余前后流程都交给脚本。

**纯本地离线**（OneNote 桌面 COM 接口），不走网络 Graph。把分区放进「以书名命名的分区组」里，
`01…0N` 这些分区名天然被该组隔离，不会和其它书的同名分区冲突；唯一查重的是分区组名（书名）本身。

```powershell
$env:PYTHONUTF8=1

# dry-run 预览（默认，不创建任何东西）
uv run python onenote_create_sections.py --book "更高更妙"

# 确认无误后正式创建
uv run python onenote_create_sections.py --book "更高更妙" --write

# 目标笔记本不存在 → 新建在线笔记本（同级于现有在线笔记本）
uv run python onenote_create_sections.py --book "书名" --notebook "新本子" --new-notebook --write
```

| 参数 | 说明 |
|------|------|
| `--book` | 书名，定位 `books-done/{书名}_拆分/`（支持部分匹配） |
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--new-notebook` | 笔记本不存在时**新建在线笔记本**（做成现有在线笔记本的同级；否则中止报警） |
| `--ref-notebook` | 新建时作「同级参考」的现有在线笔记本名（缺省自动取第一个在线笔记本） |
| `--root` | 拆分根目录（默认 `books-done`） |
| `--write` | 真正创建（缺省 dry-run 只预览） |

> **安全**：默认 dry-run；创建为纯增量（只新增空分区组/空分区，不删不改既有内容）；
> 若目标笔记本已存在同名（=书名）分区组，则**中止并报警**，绝不改动已有内容。
>
> **与 Pipeline 4 配合**：分区名为 `01…0N`，Batch 导入后用
> `onenote_sync_titles.py --section-prefix ""` 即可按 `0N ⇄ 文件夹` 对齐改标题。

### Pipeline 4：OneNote 本地整理（核对标题 + 删占位页 + 去重）

用 OneNote Batch 把拆分 PDF 导入后，常见两个收尾问题：

- 每个新分区开头有一个**空白占位页**（标题为「无标题页」，新建分区时自动生成）。
- 部分页面标题**没改成功**，仍是 OneNote 的默认标题「打印输出」。

本工具直接调用 **OneNote 桌面版本地 COM 接口**核对修正，**全程本地离线**，不走网络 Graph API —— 这样可以在**关闭同步**的状态下安全操作（避免导入/改名过程中的同步冲突），改动落在本地缓存，等你重新开启同步时再上传。

**前提与约定：**

- 仅 **Windows + OneNote 桌面版**（Microsoft 365 / OneNote 2016，Office16），不支持「OneNote for Windows 10」UWP 版。
- 分区 ⇄ 文件夹按 **`新分区N` ⇄ `0N`** 对应（可用 `--section-prefix` 改前缀）。
- 目标标题 = **完整文件名**（去扩展名，含 `NNN-` 编号前缀与 `_N` 拆分后缀）。
- 按**页面显示顺序**与文件夹内 PDF 一一对齐；页数与文件数不符的分区会**中止并报警**，不会瞎对齐。

```powershell
$env:PYTHONUTF8=1   # 让中文输出不乱码

# 只读探查：打印「高中数学教辅」下全部分区与每页标题
uv run python onenote_sync_titles.py --list

# dry-run 预览（默认不写）：逐分区列出待删占位页 + 标题差异
uv run python onenote_sync_titles.py --delete-placeholders --dedupe

# 确认无误后正式执行
uv run python onenote_sync_titles.py --delete-placeholders --dedupe --write

# 配合 Pipeline 2.5 的「书名分区组」+ 01…0N 分区（分区组感知，避免多本书同名 0N 混淆）
uv run python onenote_sync_titles.py --section-group "书名" --section-prefix "" --delete-placeholders --dedupe --write
```

| 参数 | 说明 |
|------|------|
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--root` | 拆分文件夹根目录（默认 `books-done/{书名}_拆分`） |
| `--section-prefix` | 分区名前缀（默认 `新分区`，自动忽略其后的空格；配合 Pipeline 2.5 的 `01…0N` 用 `""`） |
| `--section-group` | 只处理该分区组内的分区（配合 Pipeline 2.5 的「书名分区组」；缺省处理全部分区） |
| `--list` | 只读：打印分区组、分区与每页标题后退出 |
| `--delete-placeholders` | 删除每个分区开头的空白占位页 |
| `--dedupe` | 当某分区页数**正好是文件数 2 倍**（误打印两遍）时，删除后一份重复块 |
| `--write` | 真正写入（缺省为 dry-run 只预览） |

> **安全**：所有删除都进 OneNote **回收站**（可恢复）；标题改动可手动撤销。务必先看 dry-run 再加 `--write`。

### Pipeline 4 子工具：删除误插入的 PDF 源文件附件

用 OneNote Batch 导入时，**即使取消勾选「插入 PDF 源文件」，源 PDF 仍可能被作为附件嵌进每一页**，导致笔记本体积暴涨。本工具从指定分区逐页删除这些附件，**只删附件、保留打印出来的页面图片**。

```powershell
$env:PYTHONUTF8=1

# 只读探查：列出目标分区每页识别出的 PDF 附件（名字 + 来源路径），不删
uv run python onenote_strip_files.py --sections "新分区 1" --list

# dry-run 预览（默认不写）
uv run python onenote_strip_files.py --sections "新分区 1"

# 确认无误后正式删除（建议先单分区验证，再批量）
uv run python onenote_strip_files.py --sections "新分区 1,新分区 2" --write
```

| 参数 | 说明 |
|------|------|
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--sections` | 目标**分区名**，逗号分隔（精确匹配，注意 Batch 生成的分区名是 `新分区 N`，**带空格**） |
| `--section-group` | 只在该分区组内按名匹配分区（配合 Pipeline 2.5 的「书名分区组」；缺省全部分区） |
| `--ext` | 要删除的附件扩展名，逗号分隔（默认 `pdf`） |
| `--list` | 只读：列出每页识别出的附件后退出 |
| `--write` | 真正删除（缺省为 dry-run 只预览） |

> **安全机制**：每个附件嵌在独立的 `OE` 对象里，工具只删该对象；若对象子树内含打印图片则**跳过**，绝不误删页面图片。删除可在 OneNote 内 Ctrl+Z 撤销或从页面历史/回收站恢复。体积回收由 OneNote 后台压缩完成，可能略有延迟。

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
├── onenote_create_sections.py  # Pipeline 2.5：预建分区组 + 空分区（CLI）
├── onenote_sync_titles.py  # Pipeline 4：OneNote 本地整理（CLI）
├── onenote_strip_files.py  # Pipeline 4 子工具：删除误插入的 PDF 源文件附件（CLI）
│
├── ai_parser.py         # OCR + 目录解析逻辑
├── llm_client.py        # LLM 客户端适配层
├── pdf_utils.py         # PDF 渲染与书签写入
├── onenote_client.py    # OneNote 桌面版 COM 接口薄封装（Pipeline 2.5 / 4）
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
