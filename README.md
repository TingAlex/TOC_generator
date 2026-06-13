# PDF 自动识别目录 · 添加书签 · 按章节拆分 · OneNote 整理

将 PDF 中的目录页用 AI（或 Claude 视觉）识别，自动生成 PDF 书签（大纲），再按目录层级把整本 PDF 拆分成独立章节文件，并提供把成果导入 OneNote 的本地整理工具。适用于扫描版教材、无书签的电子书等场景。

> 代码组织为一个可安装的包 `src/tocgen/`，对外暴露一组 `toc-*` 命令（console_scripts）。
> **所有命令都在项目根目录下运行**（数据目录 `books-todo/`、`books-done/`、`books-work/` 相对当前目录解析）。

---

## 整体流程

```
books-todo/*.pdf
      │
      ▼  uv run toc-bookmarks            【Pipeline 1：书签（API 识别）】
      │   或  uv run toc-claude          【Pipeline 1（Claude 版，零 API）】
      │
      ├─[1] 渲染目录页为图片 (300 DPI)
      ├─[2] OCR 识别（DeepSeek-OCR / 或 Claude 视觉）
      ├─[3] 解析为目录结构 → books-work/{书名}/toc_parsed.txt
      └─[4] 写入 PDF 书签 (--write)  → books-done/{书名}.pdf
      │
      ▼  uv run toc-init                 【初始化配置 Excel】
      │
      books-work/books_config.xlsx   ← 进度监控 + 每本书配置
      books-work/split_config.xlsx   ← 全局拆分格式
      │
      ▼  uv run toc-split-all           【Pipeline 2：拆分】
      │
      books-done/{书名}_拆分/
          01/  001-第一章.pdf  002-第一节.pdf  ...
          02/  ...
      │
      ▼  uv run toc-onenote-sections    【Pipeline 2.5：导入前准备】
         按 0N 文件夹数，在指定笔记本建「书名分区组」+ 空分区 01…0N
      │
      ▼  uv run toc-onenote-import      【Pipeline 3：打印导入（需 SumatraPDF）】
         SetFilingLocation 定向分区 → SumatraPDF 静默打印 → 轮询落地
      │
      ▼  uv run toc-onenote-titles      【Pipeline 4：OneNote 本地整理】
         核对/改标题 · 删占位页 · 去重（toc-onenote-strip：删误插入附件）
```

所有进度和配置统一由 `books-work/books_config.xlsx` 管理，可直接用 Excel 查看和修改。

---

## 环境要求

- **Python 3.14+**（建议通过 `uv` 自动管理）
- **uv**（Python 包管理器）
- **API Key**（仅 Pipeline 1 的 API 版需要，任选其一）：硅基流动（推荐）、DeepSeek、Anthropic、OpenAI
- **（仅 Pipeline 2.5 / 3 / 4 需要）** Windows + OneNote 桌面版（Office16，非 UWP 版）
- **（仅 Pipeline 3 需要）** [SumatraPDF](https://www.sumatrapdfreader.org/download-free-pdf-viewer)（静默打印 PDF 到 OneNote 打印机）

---

## 安装

```powershell
# 1. 安装 uv（如未安装）
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# 2. 克隆并同步（uv 自动建虚拟环境、装依赖，并把本项目作为包安装，注册 toc-* 命令）
git clone https://github.com/TingAlex/TOC_generator.git
cd TOC_generator
uv sync

# 3. 配置 API Key（仅 Pipeline 1 API 版需要）
Copy-Item .env.example .env
notepad .env   # 填入 SILICONFLOW_API_KEY=sk-你的密钥
```

> 改动 `src/tocgen/` 后通常无需重装（editable 安装）；改了 `pyproject.toml`（如新增命令）后重跑 `uv sync`。

---

## 命令速查

| 命令 | 作用 |
|------|------|
| `toc-bookmarks` | Pipeline 1 批量：API 识别目录 + 加书签 |
| `toc-bookmarks-one` | Pipeline 1 单本调试 |
| `toc-claude` | Pipeline 1（Claude 版）辅助：`render` / `bookmarks`，不调用 API |
| `toc-init` | 扫描书目，生成/更新 Excel 配置 |
| `toc-split` | Pipeline 2 单本拆分 |
| `toc-split-all` | Pipeline 2 批量拆分（读 Excel） |
| `toc-onenote-sections` | Pipeline 2.5 预建分区组 + 空分区（支持 `--local-path` 本地笔记本） |
| `toc-onenote-clear` | 重打印前清空分区组内所有页（进回收站可恢复） |
| `toc-onenote-import` | Pipeline 3 打印 PDF 进分区（需 SumatraPDF） |
| `toc-onenote-fix` | 修复 OneNote「正在清理…」卡死（杀进程+重启，不删数据） |
| `toc-onenote-titles` | Pipeline 4 核对标题 + 删占位页 + 去重 |
| `toc-onenote-strip` | 遗留工具：删除误插入的 PDF 附件（当前打印流程不产生此问题） |

每个命令都支持 `--help`。

---

## 使用方法

### Pipeline 1：识别目录 + 添加书签

```powershell
# 将 PDF 放入 books-todo/
uv run toc-bookmarks            # dry-run：跑识别，不写文件
uv run toc-bookmarks --write    # 确认无误后正式写入书签
```

运行时交互式询问目录页范围（如 `7`、`7-9`）和偏移量（PDF 实际页码 vs 印刷页码）。
若已在 `books_config.xlsx` 的 `toc_pages` / `offset` 列预填，则跳过对应询问。

识别完成后，`books-work/{书名}/toc_parsed.txt` 可手工编辑修正，格式为 `层级|标题|印刷页码`：

```
1|第一章 标题|12
2|第一节 小节|12
3|1.1.1 细目|15
```

#### Pipeline 1（Claude 版）：绕过 AI API，用 Claude 看图识别目录

不想消耗 API 额度时，在新的 Claude Code 对话里输入 **`/toc-by-claude`** 触发该 skill，它会固定走：

1. `uv run toc-claude render "书名" --pages 2-4` 渲染目录页为 PNG（**不调用任何 API**）；
2. 派发一个**子任务**让 Claude 直接看这些 PNG，写出 `books-work/{书名}/toc_parsed.txt`；
3. `uv run toc-claude bookmarks "书名" --offset 18` 写入书签 → `books-done/{书名}.pdf`。

进度照常记入 `books_config.xlsx`。skill 定义见 `.claude/skills/toc-by-claude/SKILL.md`，
`toc-claude` 全程无需任何 `*_API_KEY`。

### 初始化配置 Excel

```powershell
uv run toc-init   # 扫描所有书，生成/更新 Excel 配置（新书入库时重跑）
```

| 文件 | 内容 |
|------|------|
| `books-work/books_config.xlsx` | 每本书一行：进度 flag、offset、拆分层级、是否完成 |
| `books-work/split_config.xlsx` | 全局格式：文件夹大小上限、单文件页数上限、前缀位数等 |

> Excel 不存在时（如新机器克隆后），Pipeline 1 / `toc-claude` 首次写状态会**自动建表**，无需先跑 `toc-init`；`toc-init` 仍用于批量登记书目和生成 `split_config.xlsx`。

### Pipeline 2：按章节拆分 PDF

```powershell
uv run toc-split-all --dry-run        # 预览待处理书目
uv run toc-split-all                  # 批量拆分所有未完成的书
uv run toc-split-all --book "必修第四册"   # 只拆某一本（部分匹配）

# 单本调试：
uv run toc-split "书名" --level 3 --max-pages 100
```

| 参数（`toc-split`） | 默认 | 说明 |
|------|------|------|
| `--level` | 3 | 拆分最大层级（1=章, 2=节, 3=小节） |
| `--max-pages` | 100 | 每个文件夹（= OneNote 分区）总页数**硬上限** |
| `--max-pages-per-file` | 不限 | 单文件最大页数，超出则切为 `_1/_2/…` 多份 |
| `--prefix-digits` | 3 | 前缀位数（3 → `001-`） |
| `--prefix-sep` | `-` | 前缀分隔符 |
| `--folder-digits` | 2 | 文件夹编号位数（2 → `01`） |
| `--offset` | 自动 | 手动覆盖页码偏移量 |

在 `split_config.xlsx` 的 `max_pages_per_file` 列填入目标值（如 `20`）控制单 PDF 页数。
文件夹按 `max_pages` **硬上限**以「文件」为单位贪心装满（只要 `max_pages_per_file ≤ max_pages` 就绝不超限），
一章若被切成多份可能落到不同文件夹——优先保证分区大小，也使每个分区打印后的 `.one` 文件尽量小（避免同步超限）。

### Pipeline 2.5：OneNote 预建分区组 + 空分区（导入前准备）

按 `books-done/{书名}_拆分/` 下的 `0N` 文件夹数量，在指定笔记本建一个**以书名命名的分区组**，
并在组内建好同名空分区 `01…0N`。**纯本地离线**（OneNote 桌面 COM），把分区放进书名分区组里使
`01…0N` 名天然隔离，不会和其它书的同名分区冲突；唯一查重的是分区组名（书名）本身。

```powershell
$env:PYTHONUTF8=1
uv run toc-onenote-sections --book "书名"            # dry-run 预览
uv run toc-onenote-sections --book "书名" --write     # 在已有笔记本里正式创建
# 推荐：新建本地笔记本（不同步，无 SharePoint 100MB 限制）
uv run toc-onenote-sections --book "书名" --local-path "C:\Users\用户\Documents\书名_本地" --write
# 新建在线笔记本（同级于现有在线笔记本）
uv run toc-onenote-sections --book "书名" --notebook "新本子" --new-notebook --write
```

| 参数 | 说明 |
|------|------|
| `--book` | 书名，定位 `books-done/{书名}_拆分/`（部分匹配） |
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--local-path` | **新建本地笔记本**的磁盘绝对路径（文件夹名即笔记本名）；不同步，无 SharePoint 限制；**推荐** |
| `--new-notebook` | 笔记本不存在时**新建在线笔记本**（做成现有在线笔记本的同级；否则中止报警） |
| `--ref-notebook` | 新建时作「同级参考」的现有在线笔记本名（缺省取第一个在线笔记本） |
| `--write` | 真正创建（默认 dry-run） |

> **安全**：默认 dry-run；创建为纯增量（只新增空分区组/空分区）；同名分区组**中止并报警**。
>
> **SharePoint 100MB 限制**：在线（OneDrive/SharePoint）笔记本的 `.one` 分区文件有 100MB 同步上限，打印大量 PDF 后极易超限报错。推荐先用 `--local-path` 建本地笔记本接收打印，打印完成后在 OneNote UI 中把整个分区组拖入目标在线笔记本（移动操作不触发该限制）。

### Pipeline 3：打印 PDF 进分区（toc-onenote-import）

把每个 `0N` 文件夹的 PDF 打进 Pipeline 2.5 建好的对应分区：COM `SetFilingLocation` 把打印输出
**定向到目标分区**，再用 **SumatraPDF** 把 PDF 静默打到 `OneNote (Desktop)` 打印机；**串行**打印——
打一份、轮询分区页数 +1 确认落地、再打下一份，保证「打印顺序 = 页显示顺序」。

```powershell
$env:PYTHONUTF8=1
# 先装 SumatraPDF（https://www.sumatrapdfreader.org/download-free-pdf-viewer）

# 重打印前先清空旧页（首次可跳过）
uv run toc-onenote-clear --notebook "书名_本地" --section-group "书名" --write

uv run toc-onenote-import --notebook "书名_本地" --section-group "书名" --section-prefix= --root books-done/书名_拆分            # dry-run 看映射
uv run toc-onenote-import --notebook "书名_本地" --section-group "书名" --section-prefix= --root books-done/书名_拆分 --write     # 正式打印
uv run toc-onenote-import --list --notebook "书名_本地" --section-group "书名"                                                   # 只读看分区/页
```

| 参数 | 说明 |
|------|------|
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--root` | 拆分文件夹根目录（如 `books-done/书名_拆分`） |
| `--section-prefix` | 分区名前缀（默认 `新分区`；配合 Pipeline 2.5 的 `01…0N` 用 `--section-prefix=`） |
| `--section-group` | 只处理该分区组内的分区（避免多本书同名 0N 混淆） |
| `--printer` | OneNote 桌面版打印机名（默认 `OneNote (Desktop)`；**勿用** UWP 版 `Send to Microsoft OneNote`） |
| `--sumatra` | SumatraPDF.exe 路径（缺省自动查 PATH 与常见安装位置） |
| `--settle` | 每份落地后停顿秒数（默认 0.5） |
| `--timeout` | 单文件等待落地超时秒数（默认 90，超时则中止本分区防错位） |
| `--fix` | 打印前先修复 OneNote「正在清理…」卡死（= 跑 `toc-onenote-fix`，见下） |
| `--write` | 真正打印（默认 dry-run） |

> **前提**：先用 `toc-onenote-sections` 建好 `01…0N` 空分区，且目标笔记本已在 OneNote 桌面版里打开。
> 打印路径**不嵌源文件附件**（无需 `toc-onenote-strip`）。打印后用 `toc-onenote-titles --delete-placeholders --write` 改标题、删占位页。
> 首跑若 OneNote 仍弹「选择打印输出位置」框，去 OneNote 选项关掉「总是询问打印输出的发送位置」。
> 打印完成后，在 OneNote UI 中把整个分区组拖入目标在线笔记本即可完成归档（不触发 SharePoint 限制）。

#### 修复 OneNote「正在清理上次打开之后的内容」卡死（toc-onenote-fix）

打印副本没处理完时，OneNote 下次启动会卡在「很抱歉，OneNote 正在清理上次打开之后的内容」。
本工具复刻 OneFix 的「Fix Relaunch」：**强杀 `ONENOTE.EXE`/`ONENOTEM.EXE` → 重启 → 等 COM 恢复响应**，
免重启电脑。**绝不删除任何缓存/数据**——本项目用户全是在线笔记本 + 关同步 + 离线编辑，未同步改动
只在本地缓存里，网上「删 `16.0` 缓存」的通用解法会丢笔记，故本工具不碰文件。

```powershell
$env:PYTHONUTF8=1
uv run toc-onenote-fix                       # 卡住时手动救
uv run toc-onenote-import ... --fix --write   # 打印前自动先修一遍
```

> 重启后旧的 COM 连接失效，本工具内部会自动等到**新实例**就绪；脚本里如需复用，请在 fix 之后重新建 `OneNoteClient`。

### Pipeline 4：OneNote 本地整理（核对标题 + 删占位页 + 去重）

Pipeline 3 打印完成后，调用 **OneNote 桌面版本地 COM 接口**核对修正，**全程本地离线**（可在关闭同步状态下安全操作）。

```powershell
$env:PYTHONUTF8=1
uv run toc-onenote-titles --list                                  # 只读探查
uv run toc-onenote-titles --root books-done/书名_拆分              # dry-run 预览
uv run toc-onenote-titles --root books-done/书名_拆分 --delete-placeholders --dedupe --write

# 配合 Pipeline 2.5 的「书名分区组」+ 01…0N 分区（分区组感知，避免多本书同名 0N 混淆）：
uv run toc-onenote-titles --section-group "书名" --section-prefix= --root books-done/书名_拆分 --delete-placeholders --dedupe --write
```

| 参数 | 说明 |
|------|------|
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--root` | 拆分文件夹根目录（如 `books-done/书名_拆分`） |
| `--section-prefix` | 分区名前缀（默认 `新分区`；配合 Pipeline 2.5 的 `01…0N` 用 `--section-prefix=`） |
| `--section-group` | 只处理该分区组内的分区（避免多本书同名 0N 混淆） |
| `--delete-placeholders` | 删除每个分区开头的空白占位页 |
| `--dedupe` | 当分区页数正好是文件数 2 倍（误打印两遍）时，删除后一份重复块 |
| `--write` | 真正写入（默认 dry-run） |

> **安全**：所有删除进 OneNote **回收站**（可恢复）；标题改动可手动撤销。务必先看 dry-run。

### 重打印前清空分区（toc-onenote-clear）

重新打印前需先清空目标分区的旧页，防止新旧内容叠加。**删除进回收站，可恢复。**

```powershell
$env:PYTHONUTF8=1
uv run toc-onenote-clear --notebook "书名_本地" --section-group "书名"          # dry-run 预览
uv run toc-onenote-clear --notebook "书名_本地" --section-group "书名" --write   # 真正删除
```

| 参数 | 说明 |
|------|------|
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--section-group` | 只清空该分区组内的分区（**强烈建议填写**，避免误删其他书） |
| `--write` | 真正删除（默认 dry-run） |

### Pipeline 4 子工具：删除误插入的 PDF 源文件附件（toc-onenote-strip，遗留工具）

> **当前打印流程（toc-onenote-import）不会产生此问题**，无需使用。此工具保留供历史数据清理。

如果笔记本里存在嵌入的 PDF 附件导致体积暴涨，可用本工具逐页删除，**只删附件、保留打印页图片**。

```powershell
$env:PYTHONUTF8=1
uv run toc-onenote-strip --sections "新分区 1" --list
uv run toc-onenote-strip --sections "新分区 1,新分区 2" --write
uv run toc-onenote-strip --section-group "书名" --sections "01,02" --write
```

| 参数 | 说明 |
|------|------|
| `--notebook` | 目标笔记本名（默认 `高中数学教辅`） |
| `--sections` | 目标**分区名**，逗号分隔（精确匹配，注意 Batch 生成的分区名是 `新分区 N`，**带空格**） |
| `--section-group` | 只在该分区组内按名匹配分区（缺省全部分区） |
| `--ext` | 要删除的附件扩展名，逗号分隔（默认 `pdf`） |
| `--list` | 只读：列出每页识别出的附件后退出 |
| `--write` | 真正删除（默认 dry-run） |

---

## 进度管理（Excel）

`books-work/books_config.xlsx` 是项目的统一监控面板（列：`书名` `offset` `toc_pages` `split_level`
`rendered` `ocr_done` `toc_parsed` `bookmarks_added` `bookmark_count` `拆分完成`）。

**重做某步**：在 Excel 中把对应列改为 `False`，下次运行自动从该步重做。例如：
- 改 `toc_parsed.txt` 后重写书签 → `bookmarks_added` 改 False
- 重新拆分 → `拆分完成` 改 False

---

## 项目结构

```
.
├── pyproject.toml           # 依赖 + console_scripts（toc-* 命令）
├── .env.example             # API Key 模板
├── .claude/skills/toc-by-claude/SKILL.md   # /toc-by-claude（Claude 版 Pipeline 1）
│
├── src/tocgen/              # 可安装包：库代码
│   ├── paths.py             # 路径与书名约定（单一事实来源）
│   ├── toc.py               # 目录模型：解析/序列化/校验
│   ├── pdf.py               # 渲染页面 / 写书签 / 文件名净化
│   ├── llm.py               # LLM 适配（硅基流动/DeepSeek/Anthropic/OpenAI）
│   ├── ai_parse.py          # OCR + 目录解析（调 llm，复用 toc）
│   ├── registry.py          # 每书状态读写（Excel）
│   ├── bookconfig.py        # Excel 模板/创建/读取（books_config + split_config）
│   ├── split.py             # 按目录拆分编排（run_split）
│   ├── pipeline1.py         # Pipeline 1 编排（process_one）
│   ├── onenote/
│   │   ├── client.py        # OneNote 桌面 COM 薄封装（含打印定向/分区页轮询）
│   │   ├── common.py        # OneNote CLI 共享：默认笔记本/编号解析/排序/范围限定
│   │   ├── printer.py       # 打印后端：SumatraPDF 静默打印到 OneNote 打印机
│   │   └── fix.py           # 修复「正在清理…」卡死：杀进程+重启+等就绪（不删数据）
│   └── cli/                 # 薄入口（argparse + 打印），对应各 toc-* 命令
│       ├── bookmarks.py  bookmarks_one.py  claude_toc.py  init.py
│       ├── split.py  split_all.py
│       └── onenote_sections.py  onenote_clear.py  onenote_import.py  onenote_fix.py  onenote_titles.py  onenote_strip.py
│
├── books-todo/   # 放入待处理 PDF（不入库）
├── books-done/   # 成品 PDF + {书名}_拆分/ 拆分输出（不入库）
└── books-work/   # 中间产物与 Excel 配置（不入库）
```

---

## 模型说明（Pipeline 1 API 版）

| 用途 | 默认模型 | Provider |
|------|----------|----------|
| OCR 识别 | `deepseek-ai/DeepSeek-OCR` | 硅基流动（限时免费） |
| 目录解析 | `deepseek-ai/DeepSeek-V3` | 硅基流动（按量计费） |

`.env` 中 `LLM_MODEL=模型名` 可覆盖解析模型。多 Provider 优先级：
`SILICONFLOW_API_KEY` > `DEEPSEEK_API_KEY` > `ANTHROPIC_API_KEY` > `OPENAI_API_KEY`。
使用非硅基流动 provider 时，OCR 与解析合并为单次视觉模型调用。
