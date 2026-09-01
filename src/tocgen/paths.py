"""
统一的路径与书名约定（单一事实来源）。

历史上各脚本各自写 `Path("books-todo")` / `BASE_DIR / "books-work"`，既重复又易漂移。
这里集中定义，全部相对**当前工作目录**（项目根）解析——`toc-*` 命令均在根目录运行。

## 书的身份 = 相对路径

一本书的主键是它**相对 books-todo/ 的路径**（不含 `.pdf`、分隔符统一为 `/`）：

    名师大招册                      单本书，直接躺在根上
    薛金星教材全解-人教B/必修第一册      整套系列，多加一层
    高中必刷题-人教B/必修4/主书        一份教辅含多本小册子，再多一层

层级**按需建**，不强制深度。books-work / books-done 用同一条相对路径镜像：

    books-todo/薛金星教材全解-人教B/必修第一册.pdf          源 PDF
    books-work/薛金星教材全解-人教B/必修第一册/             中间产物
    books-done/薛金星教材全解-人教B/必修第一册.pdf          带书签成品
    books-done/薛金星教材全解-人教B/必修第一册_拆分/01/…    拆分输出

流水线的原子单位始终是**一个 PDF**（主书与答案册是两本独立的书），文件夹只作组织层。
"""

from pathlib import Path

BOOKS_TODO = Path("books-todo")   # 待处理 PDF（原始版）
BOOKS_DONE = Path("books-done")   # 成品 PDF（带书签）+ 拆分输出
BOOKS_WORK = Path("books-work")   # 中间产物与配置

BOOKS_CONFIG_PATH = BOOKS_WORK / "books_config.xlsx"   # 每书进度/配置
SPLIT_CONFIG_PATH = BOOKS_WORK / "split_config.xlsx"   # 全局拆分格式

SPLIT_SUFFIX = "_拆分"   # 拆分输出文件夹后缀：books-done/{书}_拆分/


def stem(book: str) -> str:
    """书名规整：分隔符统一为 `/`、去掉 `.pdf` 扩展名与首尾斜杠。

    这是书的主键形态，见模块 docstring。反斜杠（Windows 路径、Excel 里手填的）
    一律折算成正斜杠，保证同一本书在任何入口都得到同一个 key。
    """
    s = str(book).replace("\\", "/").strip("/")
    return s[:-4] if s.lower().endswith(".pdf") else s


def book_key(book: str) -> str:
    """registry 的主键：始终带 .pdf 后缀。"""
    return f"{stem(book)}.pdf"


def display_name(book: str) -> str:
    """人类可读名：路径各段用 `-` 连接。

    供 OneNote 分区组名、日志标题等展示场景用——那里不能出现 `/`。
    例：`薛金星教材全解-人教B/必修第一册` → `薛金星教材全解-人教B-必修第一册`。
    """
    return "-".join(stem(book).split("/"))


def work_dir(book: str) -> Path:
    return BOOKS_WORK / stem(book)


def pages_dir(book: str) -> Path:
    return work_dir(book) / "pages"


def toc_parsed_path(book: str) -> Path:
    return work_dir(book) / "toc_parsed.txt"


def boundary_overlap_path(book: str) -> Path:
    """边界重叠 sidecar：标记「首页被上一节占顶」的边界条目（详见 boundary.py / split.py）。"""
    return work_dir(book) / "boundary_overlap.txt"


def boundaries_dir(book: str) -> Path:
    """toc-boundaries 渲染的边界页裁剪 / montage 输出目录。"""
    return work_dir(book) / "boundaries"


def split_root(book: str) -> Path:
    """拆分输出根目录：books-done/{书}_拆分/（书含路径时同层镜像）。"""
    return BOOKS_DONE / f"{stem(book)}{SPLIT_SUFFIX}"


def done_pdf(book: str) -> Path:
    """带书签成品 PDF 的目标路径（写入前需 mkdir(parents=True) 其父目录）。"""
    return BOOKS_DONE / f"{stem(book)}.pdf"


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

    Pipeline 1（Claude 看图渲染目录）时用——希望读未加书签的原始 PDF。
    """
    s = stem(book)
    for d in (BOOKS_TODO, BOOKS_DONE):
        p = d / f"{s}.pdf"
        if p.exists():
            return p
    raise FileNotFoundError(
        f"在 books-todo/ 和 books-done/ 中均未找到：{s}.pdf")


# ── 递归发现（书可嵌套在系列 / 教辅文件夹下） ──────────────────────────────

def _is_split_output(rel: Path) -> bool:
    """rel 是否落在某个 `{书}_拆分/` 子树内——那里面全是拆分产物，不是书。"""
    return any(part.endswith(SPLIT_SUFFIX) for part in rel.parts)


def iter_book_pdfs(root: Path) -> list[str]:
    """递归列出 root 下所有源 PDF 的书 key，跳过 `_拆分/` 产物。已排序。"""
    if not root.exists():
        return []
    keys = []
    for p in root.rglob("*.pdf"):
        rel = p.relative_to(root)
        if not _is_split_output(rel):
            keys.append(rel.as_posix()[:-4])
    return sorted(keys)


def iter_work_books(root: Path | None = None) -> list[str]:
    """递归列出 books-work/ 下已有 toc_parsed.txt 的书 key。已排序。"""
    root = root or BOOKS_WORK
    if not root.exists():
        return []
    return sorted(p.parent.relative_to(root).as_posix()
                  for p in root.rglob("toc_parsed.txt"))


def iter_split_roots(root: Path | None = None) -> list[Path]:
    """递归列出 books-done/ 下所有 `{书}_拆分/` 目录。已排序。

    命中即止，不再深入其内部（拆分产物里不会再嵌套拆分目录）。
    """
    root = root or BOOKS_DONE
    if not root.exists():
        return []
    found = []

    def walk(d: Path):
        for child in sorted(d.iterdir()):
            if not child.is_dir():
                continue
            if child.name.endswith(SPLIT_SUFFIX):
                found.append(child)
            else:
                walk(child)

    walk(root)
    return found


def known_books() -> list[str]:
    """三棵树里已知的全部书 key，去重保序（todo → done → work）。"""
    keys, seen = [], set()
    for k in (iter_book_pdfs(BOOKS_TODO) + iter_book_pdfs(BOOKS_DONE)
              + iter_work_books()):
        if k not in seen:
            seen.add(k)
            keys.append(k)
    return keys


def resolve_book(name: str) -> str:
    """把用户输入解析成唯一的书 key。

    书嵌套进系列文件夹后，key 变成 `薛金星教材全解-人教B/必修第一册` 这样的长路径；
    命令行上允许只写 `必修第一册` 这类片段，唯一命中即可。精确命中优先，
    多个候选则抛 LookupError 并列出，交由调用方 sys.exit。
    """
    needle = stem(name)
    known = known_books()
    if needle in known:
        return needle
    hits = [k for k in known if needle in k]
    if len(hits) == 1:
        return hits[0]
    if not hits:
        raise LookupError(f"找不到书：{needle}")
    listed = "\n".join(f"  · {k}" for k in hits)
    raise LookupError(f"「{needle}」匹配到 {len(hits)} 本书，请写得更具体：\n{listed}")


def split_root_key(folder: Path, root: Path | None = None) -> str:
    """`{书}_拆分/` 目录 → 书 key（split_root 的逆运算）。"""
    rel = folder.relative_to(root or BOOKS_DONE).as_posix()
    return rel[:-len(SPLIT_SUFFIX)] if rel.endswith(SPLIT_SUFFIX) else rel
