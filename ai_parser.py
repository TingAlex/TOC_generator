import re
from pathlib import Path

from llm_client import get_client

SYSTEM_PROMPT = """你是PDF目录结构提取助手。用户会提供目录页的OCR识别文字，你需要提取目录条目并按格式输出。

输出要求：
- 每行一条目录项，格式严格为：层级|标题|页码
- 层级用整数：1=章/单元，2=节，3=小节，最多3层
- 标题去除多余空格，保留原始中文
- 页码为目录中印刷的数字；无明确页码则根据上下文推断，实在无法确定填0
- 只输出目录条目，不要有任何其他文字、解释或代码块

示例输出：
1|第一章 整数的运算|5
2|第一节 加法与减法|5
2|第二节 乘法与除法|12
1|第二章 分数|20"""

USER_PROMPT = '请识别以上所有目录页中的全部目录条目，按顺序合并后以"层级|标题|页码"格式每行输出一条，最多3层，不要其他文字。'


def ocr_pages(rendered: list[tuple[int, bytes]]) -> str:
    """OCR step only: send images to OCR model, return raw text."""
    images = [data for _, data in rendered]
    client = get_client()
    print("  OCR 识别中...")
    if hasattr(client, "ocr_images"):
        return client.ocr_images(images)
    # 非 OCR 专用客户端：直接用 chat_with_images 拿原始文字
    return client.chat_with_images("", "请将图片中的全部文字完整识别出来。", images)


def parse_toc_text(ocr_text: str) -> list[dict]:
    """Parse step only: convert OCR text to TOC entries."""
    client = get_client()
    print("  解析目录结构中...")
    if hasattr(client, "parse_text"):
        raw = client.parse_text(SYSTEM_PROMPT, USER_PROMPT, ocr_text)
    else:
        raw = client.chat_with_images(SYSTEM_PROMPT, USER_PROMPT, [])
    return _parse_lines(raw)


def parse_toc_from_images(rendered: list[tuple[int, bytes]]) -> list[dict]:
    """Combined OCR + parse (kept for backward compatibility)."""
    ocr_text = ocr_pages(rendered)
    return parse_toc_text(ocr_text)


def _parse_lines(text: str) -> list[dict]:
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
    entries = []
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
            page = int(re.sub(r"[^\d]", "", page_str.strip()) or "0")
        except ValueError:
            continue
        title = title.strip()
        if title and 1 <= level <= 3:
            entries.append({"level": level, "title": title, "page": page})

    if not entries:
        raise ValueError(f"未能从响应中提取任何目录条目。\n响应前 200 字符：{text[:200]}")
    return entries
