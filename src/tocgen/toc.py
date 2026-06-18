"""
目录（TOC）模型：读取、序列化、校验 —— 全项目唯一实现。

「level|title|page」的处理统一收敛在这一处：

  · 条目就是普通 dict：{"level": int, "title": str, "page": int}
  · load_file   —— 从 toc_parsed.txt 读取条目（Claude 看图后写出的文件）
  · dumps/save  —— 序列化回 `level|title|page` 文本
  · check_nondecreasing —— 校验页码不递减（拆分前置条件）
"""

from pathlib import Path


def load_file(path: Path, *, warn=print) -> list[dict]:
    """
    从 toc_parsed.txt 读取条目。坏行打印警告并跳过（不抛异常）。
    warn 可传 None 静默。
    """
    entries: list[dict] = []
    for line_no, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        parts = line.split("|")
        if len(parts) != 3:
            if warn:
                warn(f"警告：第 {line_no} 行格式不符（期望 level|title|page），已跳过：{line!r}")
            continue
        try:
            entries.append({
                "level": int(parts[0]),
                "title": parts[1].strip(),
                "page": int(parts[2]),
            })
        except ValueError:
            if warn:
                warn(f"警告：第 {line_no} 行无法解析数字，已跳过：{line!r}")
    return entries


def dumps(entries: list[dict]) -> str:
    """序列化为 `level|title|page` 多行文本（不含末尾换行）。"""
    return "\n".join(f"{e['level']}|{e['title']}|{e['page']}" for e in entries)


def save(path: Path, entries: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(dumps(entries) + "\n", encoding="utf-8")


def check_nondecreasing(entries: list[dict]) -> None:
    """校验页码不递减；发现递减抛 ValueError（指出是哪一条）。"""
    for i in range(1, len(entries)):
        prev, cur = entries[i - 1], entries[i]
        if cur["page"] < prev["page"]:
            raise ValueError(
                f"「{cur['title']}」(level {cur['level']}) 页码 {cur['page']} "
                f"小于上一条「{prev['title']}」(level {prev['level']}) 的页码 {prev['page']}，"
                f"请先修复 toc_parsed.txt"
            )
