#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM backend for prajna-code-agent.

Supports:
    - DeepSeek (OpenAI-compatible)
    - OpenRouter (OpenAI-compatible)
    - OpenAI
    - Anthropic Claude

Configuration is read from environment variables:
    PRAJNA_LLM_PROVIDER   - deepseek | openrouter | openai | anthropic
    PRAJNA_LLM_API_KEY    - API key
    PRAJNA_LLM_MODEL      - model name (optional)
    PRAJNA_LLM_BASE_URL   - custom base URL (optional)

If no provider/key is configured, or the call fails, the caller should fall
back to the rule-based template engine.
"""

import json
import os
import re
from typing import Any, Dict, Optional

import requests


DEFAULT_MODELS = {
    "deepseek": "deepseek-chat",
    "openrouter": "deepseek/deepseek-chat",
    "openai": "gpt-4o-mini",
    "anthropic": "claude-3-5-sonnet-20241022",
}

DEFAULT_BASE_URLS = {
    "deepseek": "https://api.deepseek.com/v1",
    "openrouter": "https://openrouter.ai/api/v1",
    "openai": "https://api.openai.com/v1",
    "anthropic": "https://api.anthropic.com/v1",
}


SYSTEM_PROMPT = """你是一位精通 Python 与企业数据分析的工程师。你的任务是根据用户的自然语言描述，生成一段可直接运行的 Python 代码。

要求：
1. 使用 pandas 读取输入文件（路径由用户指定或默认为当前目录下的常见文件名）。
2. 完成用户描述的数据处理、统计、分析或可视化任务。
3. 将结果保存到用户指定的输出文件（.xlsx 或 .html）。
4. 代码必须完整、可运行，包含必要的 import。
5. 不要包含任何解释文字，只输出代码。
6. 不要生成 `if __name__ == "__main__":` 之外的函数调用包装。
7. 如果涉及图表，使用 matplotlib 保存为 PNG 或生成 HTML 报告；优先使用 pandas + openpyxl 输出 Excel。

输出格式：直接返回 Python 代码，代码前后不要加 ```python 或 ``` 标记。"""


def _clean_code(raw: str) -> str:
    """Strip markdown fences and trailing explanations."""
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("```", 2)[-1]
    if raw.endswith("```"):
        raw = raw.rsplit("```", 1)[0]
    raw = raw.strip()
    # Remove any trailing non-code text after the last meaningful line
    lines = raw.splitlines()
    while lines and lines[-1].strip() == "":
        lines.pop()
    return "\n".join(lines)


class LLMBackend:
    def __init__(
        self,
        provider: Optional[str] = None,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
        timeout: int = 120,
    ):
        self.provider = (provider or os.getenv("PRAJNA_LLM_PROVIDER", "")).lower()
        self.api_key = api_key or os.getenv("PRAJNA_LLM_API_KEY", "")
        self.model = model or os.getenv("PRAJNA_LLM_MODEL", DEFAULT_MODELS.get(self.provider, ""))
        self.base_url = base_url or os.getenv("PRAJNA_LLM_BASE_URL", DEFAULT_BASE_URLS.get(self.provider, ""))
        self.timeout = timeout

    def is_configured(self) -> bool:
        return bool(self.provider and self.api_key and self.model and self.base_url)

    def generate_code(self, task: str, context: Optional[Dict[str, Any]] = None) -> Optional[str]:
        if not self.is_configured():
            return None
        try:
            if self.provider == "anthropic":
                return self._call_anthropic(task)
            return self._call_openai_compatible(task)
        except Exception:
            return None

    def _call_openai_compatible(self, task: str) -> Optional[str]:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        if "openrouter" in self.base_url:
            headers["HTTP-Referer"] = "https://prajna.ai"
            headers["X-Title"] = "prajna-code-agent"

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": task},
            ],
            "temperature": 0.2,
            "max_tokens": 4096,
        }

        resp = requests.post(
            f"{self.base_url}/chat/completions",
            headers=headers,
            json=payload,
            timeout=self.timeout,
        )
        resp.raise_for_status()
        data = resp.json()
        content = data["choices"][0]["message"]["content"]
        return _clean_code(content)

    def _call_anthropic(self, task: str) -> Optional[str]:
        headers = {
            "x-api-key": self.api_key,
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01",
        }
        payload = {
            "model": self.model,
            "max_tokens": 4096,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": task}],
            "temperature": 0.2,
        }
        resp = requests.post(
            f"{self.base_url}/messages",
            headers=headers,
            json=payload,
            timeout=self.timeout,
        )
        resp.raise_for_status()
        data = resp.json()
        content = data["content"][0]["text"]
        return _clean_code(content)


if __name__ == "__main__":
    backend = LLMBackend()
    if backend.is_configured():
        code = backend.generate_code("读取 sales_data.xlsx，按区域统计销售额并生成柱状图保存为 chart.png")
        print(code)
    else:
        print("LLM not configured. Set PRAJNA_LLM_PROVIDER and PRAJNA_LLM_API_KEY.")
