"""
统一的路径与书名约定（单一事实来源）。

历史上各脚本各自写 `Path("books-todo")` / `BASE_DIR / "books-work"`，既重复又易漂移。
这里集中定义，全部相对**当前工作目录**（项目根）解析——`toc-*` 命令均在根目录运行。
"""

from pathlib import Path

BOOKS_TODO = Path("books-todo")   # 待处理 PDF（原始版）
BOOKS_DONE = Path("books-done")   # 成品 PDF（带书签）+ 拆分输出
BOOKS_WORK = Path("books-work")   # 中间产物与配置

BOOKS_CONFIG_PATH = BOOKS_WORK / "books_config.xlsx"   # 每书进度/配置
SPLIT_CONFIG_PATH = BOOKS_WORK / "split_config.xlsx"   # 全局拆分格式

SPLIT_SUFFIX = "_拆分"   # 拆分输出文件夹后缀：books-done/{书名}_拆分/


def stem(book: str) -> str:
    """书名规整：去掉可能的 .pdf 扩展名。"""
    return book[:-4] if book.lower().endswith(".pdf") else book


def book_key(book: str) -> str:
    """registry 的主键：始终带 .pdf 后缀。"""
    return f"{stem(book)}.pdf"


def work_dir(book: str) -> Path:
    return BOOKS_WORK / stem(book)


def pages_dir(book: str) -> Path:
    return work_dir(book) / "pages"


def toc_parsed_path(book: str) -> Path:
    return work_dir(book) / "toc_parsed.txt"


def ocr_raw_path(book: str) -> Path:
    return work_dir(book) / "ocr_raw.txt"


def split_root(book: str) -> Path:
    """拆分输出根目录：books-done/{书名}_拆分/"""
    return BOOKS_DONE / f"{stem(book)}{SPLIT_SUFFIX}"


def source_pdf(book: str) -> Path:
    """源 PDF：优先 books-done/（带书签版），回退 books-todo/（原始版）。"""
    s = stem(book)
    for d in (BOOKS_DONE, BOOKS_TODO):
        p = d / f"{s}.pdf"
        if p.exists():
            return p
    raise FileNotFoundError(
        f"在 books-done/ 和 books-todo/ 中均未找到：{s}.pdf")


def source_pdf_todo_first(book: str) -> Path:
    """源 PDF：优先 books-todo/（原始版），回退 books-done/。

    Pipeline 1 / Claude 渲染目录时用——希望读未加书签的原始 PDF。
    """
    s = stem(book)
    for d in (BOOKS_TODO, BOOKS_DONE):
        p = d / f"{s}.pdf"
        if p.exists():
            return p
    raise FileNotFoundError(
        f"在 books-todo/ 和 books-done/ 中均未找到：{s}.pdf")
