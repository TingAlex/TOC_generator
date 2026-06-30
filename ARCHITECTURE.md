# 架构文档

## 系统概览

本项目把 PDF 加工成带书签、可拆分、可导入 OneNote 的成果，分为几条流水线：

```
┌──────────────────────────────────────────────────────────────┐
│  Pipeline 1：书签   books-todo/*.pdf → 识别目录 → books-done/*.pdf │
│    toc-claude + /toc-by-claude skill：Claude 看图，零外部 AI API   │
└───────────────────────────┬──────────────────────────────────┘
                            │  共享 books-work/books_config.xlsx（单一状态源）
┌───────────────────────────▼──────────────────────────────────┐
│  Pipeline 2：拆分   books-done/*.pdf + toc_parsed.txt → 章节子 PDF │
└───────────────────────────┬──────────────────────────────────┘
┌───────────────────────────▼──────────────────────────────────┐
│  Pipeline 2.5：OneNote 预建分区组 + 空分区（toc-onenote-sections） │
└───────────────────────────┬──────────────────────────────────┘
┌───────────────────────────▼──────────────────────────────────┐
│  Pipeline 3：打印 PDF 进分区（toc-onenote-import，需 SumatraPDF）  │
│    COM SetFilingLocation 定向分区 → SumatraPDF 静默打印 → 轮询落地 │
└───────────────────────────┬──────────────────────────────────┘
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
编排层：split（按目录拆分）、onenote/*（建分区/改标题/删附件）
  │  调用
领域库：toc（目录模型）、pdf（渲染/书签）、onenote/client（COM）
  │  调用
基础设施：paths（路径/书名约定）、registry（每书状态）、bookconfig（Excel 模板/读写）
```

> Pipeline 1（识别目录→书签）没有独立的编排模块：目录识别由 Claude 看图完成，
> `cli/claude_toc.py` 只负责「渲染目录页」「写书签」两个纯本地步骤，直接调 `pdf` + `toc`。

> 关键解耦点：`cli` 不写业务逻辑；所有路径常量只在 `paths` 一处定义；`level|title|page`
> 的读取/序列化/校验只在 `toc` 一处实现；三个 OneNote CLI 的共享件集中在 `onenote/common`。

### 模块职责

| 模块 | 职责 |
|------|------|
| `paths.py` | 数据目录与书名约定的**单一事实来源**（books-todo/done/work、源 PDF 解析、工作目录助手）。全部相对 CWD（命令在项目根运行） |
| `toc.py` | 目录模型：`load_file`（读 toc_parsed.txt）、`dumps`/`save`、`check_nondecreasing`（页码不递减校验） |
| `pdf.py` | 无状态 PDF 工具：`parse_page_spec`、`render_pages_to_images`、`write_bookmarks`、`sanitize_filename` |
| `registry.py` | 每书状态读写，唯一存储 `books_config.xlsx`；Excel 缺失时经 `bookconfig.ensure_books_config` 建骨架再写 |
| `bookconfig.py` | 两张 Excel 的模板、创建、读取（`books_config` + `split_config`）；`run_init` 为 `toc-init` 入口 |
| `split.py` | `run_split`：按 `toc_parsed.txt` 拆分（用 toc 校验、pdf 切页、registry 取 offset）；`load_boundary_overlap` 读边界重叠 sidecar |
| `boundary.py` | 拆分边界分析：`compute_boundaries`（找相邻两节的边界）+ 渲染边界页顶部裁剪 + 拼 montage（仅 pymupdf），供 Claude 看图判读 fresh/shared |
| `onenote/client.py` | OneNote 桌面 COM 薄封装：读层级、建分区组/分区/在线笔记本、改/删页、列/删附件、`set_printout_section`（打印定向）、`list_section_pages`（轮询） |
| `onenote/common.py` | 四个 OneNote CLI 共享：`DEFAULT_NOTEBOOK`、`section_number`、`sorted_pdfs`/`expected_titles`、`resolve_scope`、文件夹助手 |
| `onenote/printer.py` | 打印后端：定位 SumatraPDF + `print_pdf`（静默打印到 OneNote 桌面打印机） |
| `onenote/fix.py` | 修复 OneNote「正在清理…」卡死：`fix_relaunch`（杀进程+重启+等就绪，**不删任何文件**） |
| `cli/claude_toc.py` | `toc-claude`：Pipeline 1 的两个纯本地步骤——`render`（渲染目录页）与 `bookmarks`（由 toc_parsed.txt 写书签）；中间看图识别由 Claude 完成 |
| `cli/boundaries.py` | `toc-boundaries render`：渲染拆分边界页顶部 montage，供 Claude 判读 fresh/shared（边界重叠用） |
| `cli/*.py` | 10 个 `toc-*` 命令入口（见 README 命令速查），各暴露 `main()` |

---

## 数据流

### Pipeline 1（书签）

目录识别**完全依赖 Claude 自身的多模态能力**，项目内不含任何外部 AI / OCR API 调用。

```
books-todo/{书名}.pdf
    │  toc-claude render → render_pages_to_images()   [pdf]
    ▼  books-work/{书名}/pages/page_*.png
    │  Claude 看图识别（/toc-by-claude 派发子任务，直接写文件）
    ▼  books-work/{书名}/toc_parsed.txt   ← 可手工编辑
    │  toc-claude bookmarks → write_bookmarks()        [pdf]
    ▼  books-done/{书名}.pdf
```

`render` 与 `bookmarks` 是 `cli/claude_toc.py` 的两个子命令，各自完成后
`registry.save()` 写回 Excel flag（`rendered` / `toc_parsed` / `bookmarks_added`），
中断重启可从已记录的 `toc_pages`/`offset` 续跑。中间「看图识目录 → 写 toc_parsed.txt」
由 Claude（经 `/toc-by-claude` skill 派发的子任务）用自身视觉完成，**不调用任何 API**。

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
主键统一为 `"{书名}.pdf"`。列：`书名 offset toc_pages split_level rendered
toc_parsed bookmarks_added bookmark_count 拆分完成`。

### 冷启动：Excel 缺失时自动建表

`books-work/` 不入库（`.gitignore`），新机器克隆后 Excel 不存在。此时 `registry.save()`
会先调用 `bookconfig.ensure_books_config()` 建好仅含表头的空表骨架再写入——无需先手动跑
`toc-init`，Pipeline 1 / `toc-claude` 可直接运行。**`books_config.xlsx` 是唯一状态源。**

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

### 层级映射约定（清单 / 知识手册类书）

普通教材按「章 → 节 → 小节」自然对应 L1/L2/L3。但**清单 / 知识手册类书**（如《高中知识清单》）
原始目录常有 4 级（册 → 章 → 节 → 知识N/考点N/专题N）。识别时采用的映射：

- **册 = L1、章 = L2、节(x.x) = L3**；
- **丢弃**最细的「知识N / 考点N / 专题N」一层，**也丢弃 x.x.x 小节**——x.x.x 常与父 x.x 同页，
  按 L3 拆分会切出空段。

即这类书最多取到 `x.x` 节一级；写 `toc_parsed.txt` 时不要保留更细的层级。

---

## 拆分算法

分两步：先把章节展开成「输出文件」，再以**文件**为单位贪心装箱到文件夹（`split.py`）。

### 1. 章节 → 输出文件（每个文件 ≤ max_pages_per_file）

每个条目（level ≤ 拆分层级）算出 PDF 页面范围（`_section_ranges`）；父条目与其首个子条目同页时
只占那一页（「紧邻下一条同页 → 本条仅占一页」，避免章/节标题页与子条目内容重复整段）。再把每个
范围按 `max_pages_per_file` 切片（`_plan_files`）：

```python
if not max_pages_per_file or count <= max_pages_per_file:
    一个文件
else:
    n = ceil(count / max_pages_per_file)   # 001-章_1.pdf, 001-章_2.pdf …
```
前言页（`000-书名`，offset>0 时）也作为一个文件参与。

### 2. 输出文件 → 文件夹（硬上限 max_pages，纯贪心装满）

```python
folder_idx, cumulative = 1, 0
for f in files:                       # 每个 f.pages ≤ max_pages_per_file
    if cumulative + f.pages > max_pages and cumulative > 0:
        folder_idx += 1; cumulative = 0
    place(f, folder_idx); cumulative += f.pages
```

只要 `max_pages_per_file ≤ max_pages`（典型 20 ≤ 100），**每个文件夹必然 ≤ max_pages**——
装箱单位本身 ≤ 上限，「放不下就另起」永远成立。设计取向：**文件夹 = OneNote 分区 = 同步单元，
其大小是硬约束、优先保证**；代价是一章的多个 `_N` 文件可能落到不同文件夹（对“按大小同步、求稳”
的场景可接受且更安全）。

> `cumulative > 0` 守卫避免留下空文件夹。万一未设 `max_pages_per_file`、又有单文件 > `max_pages`
> （无法再切），它只能独占该文件夹并超限——会打印警告提示调小 `--max-pages-per-file`。

### 3. 拆分边界排查与「只重切受影响文件」

拆分后若发现**相邻两个输出文件错位**（前一个被截断、后一个头部混入上一节内容），

**先别动 `offset`**。`offset` 是全书一个的全局常量（`books_config.xlsx` 每书一行），它若错则全书皆错；
只有相邻两文件错位，几乎一定是 `toc_parsed.txt` 里**某一节的印刷页码识别错了**（Pipeline 1 看图
识别目录时的笔误），而非 offset 漂移。

**诊断**：用 `pymupdf` 渲染**源 PDF**（`books-todo/` 原始版；扫描件抽不了文字，只能渲染看图）边界附近
几页，读页脚印刷页号核对每节真实起始页。
> 实例（中学教材全解 高中数学必修第一册）：`toc_parsed.txt` 把「3.3 函数的应用」记成印刷 `181`，
> 实际是 `187`（3.2 一直延伸到 186）；3.2=173、3.4=193、本章整合=194 都正确，offset 全程=10 也没问题。
> 改这一行页码即修复。

**只重切受影响文件**（不全量重跑、不动其它已拆文件，省去 OneDrive 重新同步几十个文件）：

1. 改 `toc_parsed.txt` 对应行的页码。
2. 跑脚本：复用 `split._section_ranges` / `_plan_files` + `run_split` 的装箱循环算出 `folder_idx`，
   但**只对目标文件名前缀**（如 `029-`/`030-`）写盘。
3. **守恒前提**：边界只在同一段 PDF 跨度内移动时，受影响的两文件**合计页数不变** → 后续贪心装箱布局
   完全不变，其余文件不受影响；且只要每个目标文件 ≤ `max_pages_per_file` 不触发 `_N` 切片，文件名不变。
   （若改动会改变总页数或跨越文件夹边界，则必须全量重切。）
4. **OneDrive 占用坑**：源/目标在 OneDrive 同步盘时，`pymupdf` 原地覆盖（先删后写）可能瞬时报
   `FzErrorSystem ... Permission denied`。改为写 `*.new.pdf` 临时文件再 `os.replace` 原子替换（带重试）规避。

### 4. 边界重叠：保证 PDF1 完整（toc-boundaries + boundary_overlap.txt）

**问题**：第 N 节取 `[起页, 下一节起页-1]`，下一节起页那一**整页**归 PDF2。若下一节在该页**中间**才
开始（典型：上一节的「课后强化训练」做到一半，或下一节是「（略）」省略节、标题只占该页底部一行），
则该页顶部是第 N 节的结尾——没进 PDF1 → **PDF1 不完整**。约束：可接受 PDF2 带上一节残留，但 PDF1 必须完整。

**有界结论**：每个受影响的 PDF1 **只需 +1 页**（那张边界页）——下一节标题一定在边界页上，残留最多到这页。

**数据流**（仿 Pipeline 1：CLI 渲染 → Claude 看图 → sidecar → split 消费）：

```
toc_parsed.txt(+offset) ──toc-boundaries render──▶ books-work/{书}/boundaries/montage_*.png + 清单
   compute_boundaries 找出所有「相邻两节」边界 → 渲染各边界页(=PDF2首页)顶部~45%裁剪 → 拼 montage
                                                          │  Claude 批量看图判读每个边界
                                                          ▼      顶部是「新节标题横幅」=fresh；「上一节正文/习题残留」=shared
                          books-work/{书}/boundary_overlap.txt   （每行 `印刷页|标题`，仅列 shared 边界；可手工编辑）
                                                          │  run_split → load_boundary_overlap()
                                                          ▼
   _section_ranges：shared 边界处把**前一节** end 由 `下一节.page-1` 改为 `下一节.page`（+1 页，越界仍兜底）
```

- **键 = 边界条目 `(印刷页, 标题)`**（= 被占顶的「下一节」），与 toc 风格一致、对重新过滤稳健。
  只动 level-scan 分支，不动「紧邻同页」分支。sidecar 缺省空 → 完全向后兼容（行为同现状）。
- **整页并入**：PDF1 末页 = 边界页（含本节尾 + 下节头），与 PDF2 首页重叠一页；实现简单、无损。
- **判读偏安全**：把 shared 误判成 fresh 会让 PDF1 残缺（违约束），故**存疑就标 shared**（宁多一页不残缺）。
- **影响最小化**：overlap 只给个别节 +1 页。若该 +1 没把后续文件挤过 `max_pages` 文件夹边界
  （可在重切前用计划对比算出），则**只有那一个 PDF1 文件变**，就地重切它即可，无需全量重切/重同步。

---

## Pipeline 2.5：OneNote 预建分区组 + 空分区

Pipeline 3 打印前的准备步。**纯本地离线**（COM），按 `books-done/{书名}_拆分/` 下的
`0N` 文件夹数，在指定笔记本建「书名分区组」+ 同名空分区 `01…0N`。

- **作用域隔离**：分区放进「书名分区组」，`01…0N` 名被该组隔离，不与其它书的同名分区冲突；
  唯一查重的是分区组名（=书名）→ 撞名即中止，绝不改动既有内容。
- **新建在线笔记本**（`--new-notebook`）：取现有**在线**笔记本（`path` 以 `https://` 开头）的
  OneDrive 路径，求父目录拼新名，`OpenHierarchy(newUrl, "", cftNotebook)` → 落在同一云端位置。
- **新建本地笔记本**（`--local-path`）：`OpenHierarchy(磁盘绝对路径, "", cftNotebook)` → 纯本地
  `.one` 文件，不同步，**无 SharePoint 100MB 限制**。目标笔记本是在线笔记本时首选此方案。
  默认落在 `C:\Users\用户\Documents\OneNote 笔记本\`。注意此分支**不查重**已有同名笔记本/分区组
  （在线分支才查），动手前先自查目标目录干净。

### SharePoint 100MB 限制与本地笔记本方案

在线（OneDrive/SharePoint）笔记本有每个 `.one` 分区文件 **100MB** 的同步上限。打印进 OneNote
的每页 PDF 以高分辨率图片存储，100 页打印输出的 `.one` 文件可达 150~300MB，超限后 SharePoint
拒绝同步，OneNote 报错并中断打印流程。

**解决方案：先打印到本地笔记本，全部完成后整体移入在线笔记本。**

打印期间在本地操作，不触发 SharePoint 限制；打印完成后在 OneNote UI 中把**整个分区组**拖入
在线笔记本，OneNote 自行分批同步，不再触发 100MB 单文件报错。

```
# 1. 建本地笔记本 + 分区组 + 空分区
toc-onenote-sections --book "书名" --local-path "C:\Users\用户\Documents\书名_本地" --write

# 2. 打印
toc-onenote-import --notebook "书名_本地" --section-group "书名" --section-prefix= \
    --root books-done/书名_拆分 --write

# 3. 收尾改标题
toc-onenote-titles --notebook "书名_本地" --section-group "书名" --section-prefix= \
    --root books-done/书名_拆分 --delete-placeholders --write

# 4. 在 OneNote UI 中：把分区组整体拖入目标在线笔记本（如「薛金星教材全解-人教B」）
#    OneNote 自行处理同步，不再报 SharePoint 100MB 错误
```

### COM 接入要点（创建相关）

| 坑 / 要点 | 处理 |
|-----------|------|
| 创建用 `OpenHierarchy(path, relativeToObjectID, [out]objectID, cftIfNotExist)`，`[out]` 在 comtypes 早绑定下转为返回值 | 调用只传 `(path, rel, cft)`，新 ID 取返回值 |
| `CreateFileType` 枚举 | `cftNotebook=1` / `cftFolder=2`（分区组）/ `cftSection=3`（分区） |
| **分区路径必须带 `.one`**，否则 `OpenHierarchy` 抛 `COMError 0x80042004`；分区组/文件夹不带 | `create_section` 自动补 `.one`；OneNote 显示时去掉 |
| **本地笔记本路径必须用反斜杠 `\`**（原生 Windows 路径），正斜杠 `/` 抛 `COMError 0x80042006` | `--local-path` 传 `C:\…\本子名`；用 PowerShell 原生反斜杠路径，别用 Bash 的 `/` 风格（comtypes 把字符串原样交给 OneNote，OneNote 只认 Windows 路径） |
| 取笔记本在线/本地 | `get_hierarchy` 读 Notebook 的 `path`；在线笔记本以 `https://` 开头 |

---

## Pipeline 3：打印 PDF 进分区（toc-onenote-import）

把每个 `0N` 文件夹的 PDF 打进 Pipeline 2.5 建好的对应分区。分区 ⇄ 文件夹映射、`--section-group` 范围限定与 Pipeline 4 共用
（`section_number` / `resolve_scope` / `sorted_pdfs`）。

### 机制

```
for 每个分区(按编号排序):
    client.set_printout_section(分区.id)          # SetFilingLocation(flPrintOuts=5, fltNamedSectionNewPage=0)
    for 每个 PDF in sorted_pdfs(文件夹):
        before = len(list_section_pages(分区.id))
        printer.print_pdf(pdf, "OneNote (Desktop)", sumatra)   # SumatraPDF -print-to … -silent
        轮询 list_section_pages 直到 > before（或 --timeout 超时则中止本分区）
```

- **定向**：`SetFilingLocation(flPrintOuts=5, fltNamedSectionNewPage=0, 分区ID)` 让打印输出落到指定分区、
  每次新建一页，无需弹位置框。一个多页 PDF = OneNote 里**一页**打印输出（与 Pipeline 4「一页 ⇄ 一文件」一致）。
- **串行 + 落地确认**：必须打一份、轮询页数 +1、再打下一份，保证「打印顺序 = 页显示顺序」；超时/打印失败
  即**中止本分区**（不续打，避免迟到页与下一页错序）。事后 `toc-onenote-titles` 才能按显示顺序对齐改标题。
- **打印后端 = SumatraPDF**（`-print-to "OneNote (Desktop)" -silent`）：免费便携、静默、打完自退、不抢焦点。
  比 Adobe `/t` 稳。`onenote/printer.find_sumatra` 按 `--sumatra` → PATH → 常见安装位置定位。

### COM / 打印接入要点

| 坑 / 要点 | 处理 |
|-----------|------|
| 打印输出定向 | `SetFilingLocation(5, 0, sectionID)`：三参全 `[in]`、无 DATE 坑；持久改 OneNote 打印归档设置，无读回接口、跑完不还原（仅影响用户下次手动打印默认落点） |
| 打印机选错版本 | **必须 `OneNote (Desktop)`**（端口 `nul:`，2016 桌面栈）；避开 UWP 版 `Send to Microsoft OneNote`（端口含包名 `…8wekyb3d8bbwe…`），否则与桌面 COM 不在同一栈 |
| 打印是异步的 | SumatraPDF 退出≠页已生成；靠 `list_section_pages`（= `GetHierarchy(sectionID, hsPages)`）轮询页数 +1 确认落地 |
| 「总是询问打印输出位置」选项 | `fltNamedSectionNewPage` 理应覆盖；首跑若仍弹框，需在 OneNote 选项里关掉该询问 |
| 不嵌源文件 | 走打印路径不会插入 PDF 源文件附件，故**无需** `toc-onenote-strip` |

### 重打印前清空：`toc-onenote-clear`（`cli/onenote_clear.py`）

重新打印前需先清空目标分区的旧页（否则变成两套内容叠加）。`toc-onenote-clear` 把指定分区组内
所有分区的全部页面送入 OneNote 回收站（`DeleteHierarchy(pageId, 0.0, False)`），可恢复，默认 dry-run。

```
toc-onenote-clear --notebook "书名_本地" --section-group "书名"         # 预览
toc-onenote-clear --notebook "书名_本地" --section-group "书名" --write  # 真正删除
```

### 卡死自愈：`toc-onenote-fix`（`onenote/fix.py`）

打印副本没处理完时，OneNote 下次启动会卡在「很抱歉，OneNote 正在清理上次打开之后的内容」。
`fix_relaunch()` 复刻 OneFix 的「Fix Relaunch」：强杀 `ONENOTE.EXE`/`ONENOTEM.EXE` → 重启 → 轮询
`OneNoteClient().get_hierarchy()` 直到 COM 恢复（= 已越过清理）。`toc-onenote-import --fix` 会在打印前先跑它。

| 坑 / 要点 | 处理 |
|-----------|------|
| **绝不删缓存** | 用户全是在线笔记本 + 关同步 + 离线编辑，未同步改动只在本地 `16.0\cache`；网上「删 16.0」通用解法会丢笔记。本工具只杀进程+重启，不碰文件 |
| 重启后旧 COM 失效 | `fix_relaunch` 杀的就是当前 COM 连的进程，故重启后必须用**新** `OneNoteClient`；`wait_until_ready` 内部每轮新建 client 探活。`--import --fix` 在 fix 之后才建主 client |
| Office16 = 2016/365 | 365 也登记为 Office16，进程名/路径一致（`…\root\Office16\ONENOTE.EXE`） |

---

## Pipeline 4：OneNote 本地整理

Pipeline 3 打印完成后做收尾。**纯本地离线**：COM 操作本地缓存，不读 Excel、不走网络。

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

> **刚建分区直接打印时没有占位页可删**：分区由 Pipeline 2.5 建好后**立即**打印，第一份打印输出会
> 落在那张空占位页上、把它变成打印页，不再单独留 `无标题页`。于是每个分区正好是文件数那么多页、
> 全为「打印输出」，`--delete-placeholders` 报 `删0`，天然 1:1 干净对齐。占位页删除逻辑主要兜底
> **先建分区、隔一段时间或手动操作后才打印**而残留空占位页的情况。

### 去重（`--dedupe`）

仅当分区页数**正好是文件数 2 倍**（误打印两遍）时，删除后一份重复块（进回收站），保留前一份继续对齐。

### 子工具：删除误插入的源文件附件（`toc-onenote-strip`）

页 XML 里附件是 `<one:InsertedFile preferredName="x.pdf">`，自身一般不带 objectID——objectID
在外层 `<one:OE>` 上，故向上找最近带 objectID 的祖先 OE 作为删除目标；若该祖先子树含
`<one:Image>`（打印图片）则跳过（绝不误删图片）。删除 `DeletePageContent(pageId, objectId, 0.0, True)`。

### COM 接入要点（读写相关）

| 坑 / 要点 | 处理 |
|-----------|------|
| comtypes 默认晚绑定，OneNote 报 `TYPE_E_LIBNOTREGISTERED`（“库没有注册”） | 用 `GetModule(("{0EA692EE-BB50-4E3C-AEF0-356D91732725}",1,1))`（OneNote 15.0 类型库）强制**早绑定**，再 `CreateObject("OneNote.Application", interface=mod.IApplication)` |
| `DeleteHierarchy` / `UpdatePageContent` / `DeletePageContent` 的 DATE 参数 | 显式传 `0.0`（= 不校验修改时间），否则 comtypes 默认 datetime 与签名冲突 |
| 删除安全性 | `deletePermanently=False` → 进 OneNote 回收站，可恢复 |
| 改标题不伤正文 | `UpdatePageContent` 只提交含 `ID` 与 `Title` 的最小 Page XML |

### OneNote 开发环境须知（换机器开发必读）

- **数据目录与代码解耦**：`books-todo/ books-done/ books-work/` 默认相对**当前工作目录**解析
  （见 `paths.py`），可放在仓库外（本机即放在 OneDrive 同步盘
  `C:\Users\用户\OneDrive\高中数学教辅材料拆分\`）。换机器/换位置时不改代码，只把 **cwd 设到数据根目录**
  再跑 `toc-*`（用仓库 venv 里的可执行文件即可，`uv run` 会指向仓库内的空数据目录，故直接调
  `.venv\Scripts\toc-*.exe`）。
- **「拆分完成」flag 不可信**：`books_config.xlsx` 该列可能标 `true` 但磁盘上并无
  `books-done/{书名}_拆分/`。判断是否已拆**看磁盘目录**，别信 flag；flag=true 会让 `toc-split-all`
  跳过，需补拆时用单本 `toc-split "全名" --level N` 绕过。`--book` 是子串匹配，过滤用完整书名
  （「必修第一册」会连带命中「选择性必修第一册」）。
- **为什么走桌面 COM 而非 Graph**：用户刻意关闭 OneNote 同步做离线编辑；桌面 COM 操作本地缓存，
  离线免鉴权，改动随之后同步上传。Graph 则必须联网 + OAuth。
- **平台限制**：仅 Windows + OneNote 桌面版（Office16），**不支持** UWP「OneNote for Windows 10」。
  OneNote **365 也登记为 Office16**（exe 同为 `…\root\Office16\ONENOTE.EXE`，打印机/COM/缓存路径与 2016 一致）。
- **依赖 `comtypes`（非 pywin32）**：纯 Python、免编译，Python 3.14 下更省事；COM 调用必须早绑定。
- **装依赖**：venv 由 `uv` 管理、**无 pip**，用 `uv add` / `uv sync`（或临时 `uv pip install`），别 `python -m pip`。
- **运行前 `$env:PYTHONUTF8=1`**：否则中文标题/分区名/路径乱码。
- **OneNote 默认标题**：新建分区占位页 `无标题页`（非空）；改名失败页/打印输出页 `打印输出`。
- **Batch 分区名带空格**：是 `新分区 N`（如 `新分区 7`），不是 `新分区7`。
- **Pipeline 3 需装 [SumatraPDF](https://www.sumatrapdfreader.org/download-free-pdf-viewer)**：打印后端，
  `onenote/printer.find_sumatra` 按 `--sumatra` → PATH → 常见安装位置（含 `%LOCALAPPDATA%\SumatraPDF\`）定位。
- **打印机名随机器/语言变**：默认 `OneNote (Desktop)`；换机器先 `Get-Printer` 找到 **OneNote 桌面版**打印机
  （DriverName「Send to Microsoft OneNote 16 Driver」且端口 `nul:`），用 `--printer` 传入。**别选** UWP 那个
  （端口含包名 `…8wekyb3d8bbwe…`）。首跑若弹「选择打印输出位置」框，去 OneNote 选项关掉「总是询问…」。
- **卡死自愈**：打印没处理完会导致下次启动卡「正在清理上次内容」。用 `toc-onenote-fix` 或 `toc-onenote-import --fix`
  （杀进程+重启，**绝不删缓存**——离线编辑的未同步改动只在 `16.0\cache` 里，删了丢笔记）。

---

## 依赖

| 包 | 用途 |
|----|------|
| `pymupdf` | PDF 渲染、书签写入、页面提取与切片 |
| `openpyxl` | 读写 Excel 配置 |
| `comtypes` | Pipeline 2.5 / 3 / 4：OneNote 本地 COM 自动化（仅 Windows + OneNote 桌面版）。纯 Python、免编译，故选它而非 pywin32；调用须早绑定 |
| **SumatraPDF**（外部 exe，非 Python 包） | Pipeline 3 打印后端：`-print-to "OneNote (Desktop)" -silent` 静默打印。经 `subprocess` 调用，不入 `dependencies` |

构建：`hatchling`（`[build-system]`），打包 `src/tocgen`，console_scripts 见 `pyproject.toml [project.scripts]`。
