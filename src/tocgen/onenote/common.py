"""
三个 OneNote CLI（建分区、改标题、删附件）共享的小工具。

此前 DEFAULT_NOTEBOOK、分区编号解析、`--section-group` 范围限定逻辑在多个脚本里
各写一份；这里收敛为一处。
"""

import re
from pathlib import Path

from .client import OneNoteClient, Notebook, Section

DEFAULT_NOTEBOOK = "高中数学教辅"

# 文件名/分区名前缀编号，如 "117-4.4.1 弦长" → 117，"新分区7" 去前缀后 → 7
_NUM_PREFIX = re.compile(r"^(\d+)")
_NUM_DIR = re.compile(r"^\d+$")


def section_number(name: str, prefix: str) -> int | None:
    """从分区名解析编号：去掉 prefix（及其后空格）后若为纯数字则返回，否则 None。

    prefix="" 时直接对整名取数字（配合 Pipeline 2.5 的 `01…0N` 分区）。
    """
    if not name.startswith(prefix):
        return None
    rest = name[len(prefix):].strip()
    return int(rest) if rest.isdigit() else None


def sorted_pdfs(folder: Path) -> list[Path]:
    """文件夹内 PDF，按编号前缀数字排序（无前缀的排末尾、再按名）。

    导入打印与改标题共用同一排序，保证「打印顺序 = 页显示顺序 = 期望标题顺序」。
    """
    def sort_key(p: Path):
        m = _NUM_PREFIX.match(p.stem)
        return (int(m.group(1)) if m else 1 << 30, p.stem)
    return sorted(folder.glob("*.pdf"), key=sort_key)


def expected_titles(folder: Path) -> list[str]:
    """文件夹内 PDF 的期望页标题 = 去扩展名的文件名，按编号前缀数字排序。"""
    return [p.stem for p in sorted_pdfs(folder)]


def section_folder_names(split_folder: Path) -> list[str]:
    """拆分文件夹下的 0N 子文件夹名（即将建的分区名），按数字排序。"""
    dirs = [d.name for d in split_folder.iterdir() if d.is_dir() and _NUM_DIR.match(d.name)]
    return sorted(dirs, key=lambda n: int(n))


def resolve_scope(client: OneNoteClient, nb: Notebook,
                  section_group: str | None) -> tuple[list[Section], str] | None:
    """
    按 `--section-group` 限定处理范围，避免多本书的同名 0N 分区混淆。

    返回 (分区列表, 范围描述)；找不到指定分区组时打印提示并返回 None。
    section_group 为空 → 用笔记本的扁平分区列表（含组内，向后兼容旧扁平布局）。
    """
    if section_group:
        grp = client.find_section_group(nb, section_group)
        if not grp:
            print(f"未找到分区组「{section_group}」。可先用 --list 查看。")
            return None
        return grp.sections, f"分区组「{grp.name}」"
    return nb.sections, "全部分区（含组内）"
