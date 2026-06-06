"""
AI 目录识别：OCR 图片 → 文字，文字 → 结构化目录条目。

只负责「调模型 + 拼 prompt」；条目的解析/校验复用 toc 模块（单一实现）。
"""

from .llm import get_client
from . import toc

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
    """仅 OCR：把图片送 OCR 模型，返回原始文字。"""
    images = [data for _, data in rendered]
    client = get_client()
    print("  OCR 识别中...")
    if hasattr(client, "ocr_images"):
        return client.ocr_images(images)
    # 非 OCR 专用客户端：直接用 chat_with_images 拿原始文字
    return client.chat_with_images("", "请将图片中的全部文字完整识别出来。", images)


def parse_toc_text(ocr_text: str) -> list[dict]:
    """仅解析：把 OCR 文字转为目录条目（解析/校验走 toc 模块）。"""
    client = get_client()
    print("  解析目录结构中...")
    if hasattr(client, "parse_text"):
        raw = client.parse_text(SYSTEM_PROMPT, USER_PROMPT, ocr_text)
    else:
        raw = client.chat_with_images(SYSTEM_PROMPT, USER_PROMPT, [])
    return toc.parse_text(raw)


def parse_toc_from_images(rendered: list[tuple[int, bytes]]) -> list[dict]:
    """OCR + 解析合并（向后兼容）。"""
    return parse_toc_text(ocr_pages(rendered))
