"""
LLM 客户端适配层。

通过环境变量选择 provider（第一个匹配的生效）：
  SILICONFLOW_API_KEY → 硅基流动，两步法：
                        Step 1: deepseek-ai/DeepSeek-OCR 逐张识别（每次只支持一张图）
                        Step 2: deepseek-ai/DeepSeek-V3 将 OCR 文本解析为目录格式
  DEEPSEEK_API_KEY    → DeepSeek（OpenAI 兼容接口，视觉模型）
  ANTHROPIC_API_KEY   → Anthropic Claude
  OPENAI_API_KEY      → OpenAI

可选覆盖：
  LLM_MODEL=<model>   覆盖 Step 2 的文本解析模型（硅基流动）或其他 provider 的主模型
                      硅基流动 Step 2 默认: deepseek-ai/DeepSeek-V3
                      DeepSeek 默认: deepseek-chat
                      Anthropic 默认: claude-sonnet-4-6
                      OpenAI 默认: gpt-4o
"""

import base64
import os
from abc import ABC, abstractmethod


class LLMClient(ABC):
    @abstractmethod
    def chat_with_images(self, system_prompt: str, user_prompt: str, images_png: list[bytes]) -> str:
        """发送文字 + PNG 图片列表，返回模型的文字回复。"""


class DeepSeekOCRClient(LLMClient):
    """
    两步法：
    1. deepseek-ai/DeepSeek-OCR 逐张图片识别 → markdown 文本（每次限传一张图）
    2. deepseek-ai/DeepSeek-V3 将 OCR 文本解析为目录行格式
    """
    OCR_MODEL = "deepseek-ai/DeepSeek-OCR"
    OCR_PROMPT = "<|grounding|>Convert the document to markdown."
    DEFAULT_PARSE_MODEL = "deepseek-ai/DeepSeek-V3"

    def __init__(self):
        from openai import OpenAI
        self._client = OpenAI(
            api_key=os.environ["SILICONFLOW_API_KEY"],
            base_url="https://api.siliconflow.cn/v1",
        )
        self._parse_model = os.environ.get("LLM_MODEL") or self.DEFAULT_PARSE_MODEL
        print(f"  [LLM] DeepSeek-OCR（识别）+ {self._parse_model}（解析）")

    def _ocr_one(self, image_bytes: bytes, idx: int, total: int) -> str:
        b64 = base64.standard_b64encode(image_bytes).decode()
        print(f"    OCR 第 {idx}/{total} 页...")
        resp = self._client.chat.completions.create(
            model=self.OCR_MODEL,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                    {"type": "text", "text": self.OCR_PROMPT},
                ],
            }],
            max_tokens=2048,
        )
        return resp.choices[0].message.content

    def ocr_images(self, images_png: list[bytes]) -> str:
        """Step 1 only: OCR all images, return combined raw text."""
        pages_text = [
            self._ocr_one(img, i, len(images_png))
            for i, img in enumerate(images_png, 1)
        ]
        return "\n\n---\n\n".join(pages_text)

    def parse_text(self, system_prompt: str, user_prompt: str, ocr_text: str) -> str:
        """Step 2 only: parse OCR text into level|title|page lines."""
        print(f"    解析目录结构（{self._parse_model}）...")
        resp = self._client.chat.completions.create(
            model=self._parse_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"以下是目录页的 OCR 识别文字：\n\n{ocr_text}\n\n{user_prompt}"},
            ],
            max_tokens=4096,
        )
        return resp.choices[0].message.content

    def chat_with_images(self, system_prompt: str, user_prompt: str, images_png: list[bytes]) -> str:
        ocr_text = self.ocr_images(images_png)
        return self.parse_text(system_prompt, user_prompt, ocr_text)


class AnthropicClient(LLMClient):
    DEFAULT_MODEL = "claude-sonnet-4-6"

    def __init__(self):
        import anthropic
        self._client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
        self._model = os.environ.get("LLM_MODEL") or self.DEFAULT_MODEL
        print(f"  [LLM] Anthropic · {self._model}")

    def chat_with_images(self, system_prompt: str, user_prompt: str, images_png: list[bytes]) -> str:
        content = []
        for i, data in enumerate(images_png, 1):
            content.append({"type": "text", "text": f"图片 {i}："})
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": base64.standard_b64encode(data).decode(),
                },
            })
        content.append({"type": "text", "text": user_prompt})

        resp = self._client.messages.create(
            model=self._model,
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": content}],
        )
        return resp.content[0].text


class OpenAICompatibleClient(LLMClient):
    """适用于所有 OpenAI 兼容接口（DeepSeek、OpenAI 等）。"""

    PROVIDERS = {
        "deepseek": {
            "base_url": "https://api.deepseek.com",
            "key_env": "DEEPSEEK_API_KEY",
            "default_model": "deepseek-chat",
        },
        "openai": {
            "base_url": None,
            "key_env": "OPENAI_API_KEY",
            "default_model": "gpt-4o",
        },
    }

    def __init__(self, provider: str):
        from openai import OpenAI
        cfg = self.PROVIDERS[provider]
        self._model = os.environ.get("LLM_MODEL") or cfg["default_model"]
        kwargs = {"api_key": os.environ[cfg["key_env"]]}
        if cfg["base_url"]:
            kwargs["base_url"] = cfg["base_url"]
        self._client = OpenAI(**kwargs)
        print(f"  [LLM] {provider.capitalize()} · {self._model}")

    def chat_with_images(self, system_prompt: str, user_prompt: str, images_png: list[bytes]) -> str:
        content = []
        for data in images_png:
            b64 = base64.standard_b64encode(data).decode()
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{b64}"},
            })
        content.append({"type": "text", "text": user_prompt})

        resp = self._client.chat.completions.create(
            model=self._model,
            max_tokens=4096,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content},
            ],
        )
        return resp.choices[0].message.content


def get_client() -> LLMClient:
    """根据环境变量自动选择 provider，找不到则报错提示。"""
    if os.environ.get("SILICONFLOW_API_KEY"):
        return DeepSeekOCRClient()
    if os.environ.get("DEEPSEEK_API_KEY"):
        return OpenAICompatibleClient("deepseek")
    if os.environ.get("ANTHROPIC_API_KEY"):
        return AnthropicClient()
    if os.environ.get("OPENAI_API_KEY"):
        return OpenAICompatibleClient("openai")

    raise EnvironmentError(
        "未找到任何 API Key。请在 .env 文件中填入以下任意一个：\n"
        "  SILICONFLOW_API_KEY=sk-...\n"
        "  DEEPSEEK_API_KEY=sk-...\n"
        "  ANTHROPIC_API_KEY=sk-ant-...\n"
        "  OPENAI_API_KEY=sk-..."
    )
