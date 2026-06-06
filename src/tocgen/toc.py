"""
目录（TOC）模型：解析、序列化、校验 —— 全项目唯一实现。

此前「level|title|page」的解析散落在 4 处（main / ai_parser / split_pdf / claude_toc_helper），
各写各的、行为略有差异。这里收敛为一处：

  · 条目就是普通 dict：{"level": int, "title": str, "page": int}
  · parse_text  —— 从 AI/OCR 原始文字提取条目（容错：跳过坏行、去 <think>）
  · load_file   —— 从 toc_parsed.txt 读取条目
  · dumps/save  —— 序列化回 `level|title|page` 文本
  · check_nondecreasing —— 校验页码不递减（拆分前置条件）
"""

import re
from pathlib import Path

# 行格式：level|title|page
_LINE_RE = re.compile(r"^\s*(\d+)\s*\|\s*(.*?)\s*\|\s*([^\|]*?)\s*$")


def _digits(s: str) -> int:
    return int(re.sub(r"[^\d]", "", s) or "0")


def parse_text(text: str, *, strict_levels: bool = True) -> list[dict]:
    """
    从 AI/OCR 原始文字中提取目录条目。

    容错：移除 <think>…</think>；跳过非「三段|」行；页码取数字部分（无则 0）。
    strict_levels=True 时只保留 1≤level≤3 的条目。
    抛 ValueError 当一条都没提取到。
    """
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
    entries: list[dict] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split("|", 2)
        if len(parts) != 3:
            continue
        level_str, title, page_str = parts
        try:
            level = int(level_str.strip())
        except ValueError:
            continue
        page = _digits(page_str.strip())
        title = title.strip()
        if not title:
            continue
        if strict_levels and not (1 <= level <= 3):
            continue
        entries.append({"level": level, "title": title, "page": page})

    if not entries:
        raise ValueError(
            f"未能提取任何目录条目。原文前 200 字符：{text[:200]}")
    return entries


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
