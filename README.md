# PDF 自动识别目录并添加书签

将 PDF 中的目录页用 AI 识别，自动生成 PDF 书签（大纲）。适用于扫描版教材、无书签的电子书等场景。

## 工作流程

```
books-todo/*.pdf
      │
      ▼ [1] 渲染目录页为图片 (300 DPI)
books-work/{书名}/pages/*.png
      │
      ▼ [2] DeepSeek-OCR 逐张识别
books-work/{书名}/ocr_raw.txt
      │
      ▼ [3] DeepSeek-V3 解析为目录结构
books-work/{书名}/toc_parsed.txt   ← 可手工编辑修正
      │
      ▼ [4] 写入 PDF 书签（需 --write）
books-done/*.pdf
```

每步进度记录在 `books-work/{书名}/state.json`，重跑时自动跳过已完成步骤。

---

## 环境要求

- **Windows 10/11**，PowerShell 5.1 或更高
- **Python 3.12+**（项目使用 3.14，建议通过 `uv` 自动管理）
- **uv**（Python 包管理器）
- **硅基流动 API Key**（免费注册：https://cloud.siliconflow.cn）

---

## 安装步骤（PowerShell）

### 1. 安装 uv

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

安装完成后重启 PowerShell，验证：

```powershell
uv --version
```

### 2. 克隆项目

```powershell
git clone https://github.com/TingAlex/TOC_generator.git
cd TOC_generator
```

### 3. 安装依赖

`uv` 会自动下载所需 Python 版本并创建虚拟环境：

```powershell
uv sync
```

主要依赖：
- `pymupdf` — PDF 渲染与书签写入
- `openai` — 调用硅基流动 API（OpenAI 兼容接口）
- `anthropic` — 可选，Anthropic Claude 支持
- `python-dotenv` — 读取 .env 配置

### 4. 配置 API Key

```powershell
Copy-Item .env.example .env
notepad .env
```

在 `.env` 中填入硅基流动的 API Key（其余留空）：

```
SILICONFLOW_API_KEY=sk-你的密钥
```

API Key 获取地址：https://cloud.siliconflow.cn/account/ak

### 5. 放入待处理 PDF

```powershell
# 如果目录不存在则创建
New-Item -ItemType Directory -Force books-todo

# 将 PDF 文件复制进去（或直接在资源管理器中拖入）
Copy-Item "C:\你的路径\书名.pdf" books-todo\
```

---

## 使用方法

### 测试单本

```powershell
# dry-run：只跑识别，不写 PDF（推荐先用这个检查识别结果）
uv run python test_one.py

# 确认无误后，写入 PDF 书签
uv run python test_one.py --write
```

### 批量处理

```powershell
# dry-run：识别所有书，不写 PDF
uv run python main.py

# 写入所有书的 PDF 书签
uv run python main.py --write
```

运行时程序会交互式询问：
1. 目录页范围（如 `7-10`）
2. 偏移量确认（PDF 实际页码 vs 印刷页码）

---

## 项目结构

```
.
├── main.py              # 主程序（批量）
├── test_one.py          # 单本测试
├── ai_parser.py         # OCR + 目录解析逻辑
├── llm_client.py        # LLM 客户端适配层
├── pdf_utils.py         # PDF 渲染与书签写入
├── registry.py          # 进度状态管理
├── pyproject.toml       # 项目依赖声明
├── .env.example         # API Key 配置模板
│
├── books-todo/          # 放入待处理 PDF（不入库）
├── books-done/          # 处理完成的 PDF 输出（不入库）
└── books-work/          # 中间产物（不入库）
    └── {书名}/
        ├── state.json       # 该书进度状态
        ├── ocr_raw.txt      # OCR 原始输出
        ├── toc_parsed.txt   # 解析后目录（可手工编辑）
        └── pages/           # 渲染图片
```

---

## 进度管理（state.json）

每本书的状态文件位于 `books-work/{书名}/state.json`：

```json
{
  "toc_pages": [7, 8, 9, 10],
  "rendered": true,
  "ocr_done": true,
  "toc_parsed": true,
  "offset": 6,
  "bookmarks_added": false,
  "bookmark_count": 33
}
```

**重做某步**：用记事本/VS Code 将对应字段改为 `false`，下次运行自动从该步重做：

| 字段 | 说明 |
|------|------|
| `toc_pages` | 目录页页码范围 |
| `rendered` | 是否已渲染为图片 |
| `ocr_done` | 是否已完成 OCR |
| `toc_parsed` | 是否已解析为目录结构 |
| `offset` | 偏移量（`PDF页码 = 印刷页码 + offset`） |
| `bookmarks_added` | 是否已写入 PDF 书签 |

**常用场景**：

- 手工修正 `toc_parsed.txt` 后重新写入 PDF：  
  将 `bookmarks_added` 改为 `false`，运行 `uv run python main.py --write`

- 重新跑 OCR（如换了更好的模型）：  
  将 `ocr_done`、`toc_parsed`、`bookmarks_added` 都改为 `false`

---

## 模型说明

| 用途 | 模型 | 备注 |
|------|------|------|
| OCR 识别 | `deepseek-ai/DeepSeek-OCR` | 硅基流动，限时免费 |
| 目录解析 | `deepseek-ai/DeepSeek-V3` | 硅基流动，按量计费 |

在 `.env` 中设置 `LLM_MODEL=模型名` 可覆盖解析模型。

## 其他 Provider 支持

优先级：`SILICONFLOW_API_KEY` > `DEEPSEEK_API_KEY` > `ANTHROPIC_API_KEY` > `OPENAI_API_KEY`

使用非硅基流动的 provider 时，OCR 与解析合并为单次调用（需该 provider 支持视觉模型）。
