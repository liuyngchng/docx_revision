"""
LLM API 客户端（OpenAI 兼容格式）
通过 httpx 直接调用，避免额外依赖 openai 包。
"""

import json
from typing import Any, Optional

import httpx

from backend.core.config import LlmConfig


class LlmClient:
    """OpenAI 兼容 API 客户端"""

    def __init__(self, config: LlmConfig):
        self.config = config
        self.base_url = config.api_base.rstrip("/")
        self.api_key = config.api_key
        self.model = config.model
        self.timeout = config.timeout

    def _headers(self) -> dict[str, str]:
        return {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }

    def chat_completion(
        self,
        messages: list[dict[str, str]],
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        response_format: Optional[dict] = None,
    ) -> dict[str, Any]:
        """
        调用 chat completions 接口。

        参数:
            messages: OpenAI 格式的消息列表
            temperature: 覆盖配置的温度
            max_tokens: 覆盖配置的最大 token
            response_format: 可选，例如 {"type": "json_object"}

        返回:
            完整的 API 响应 JSON
        """
        url = f"{self.base_url}/chat/completions"

        payload: dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "temperature": temperature if temperature is not None else self.config.temperature,
            "max_tokens": max_tokens if max_tokens is not None else self.config.max_tokens,
        }

        if response_format:
            payload["response_format"] = response_format

        try:
            with httpx.Client(timeout=self.timeout) as client:
                resp = client.post(url, headers=self._headers(), json=payload)
                resp.raise_for_status()
                return resp.json()
        except httpx.TimeoutException as e:
            raise RuntimeError(f"LLM API 请求超时（{self.timeout}秒）: {e}")
        except httpx.HTTPStatusError as e:
            raise RuntimeError(
                f"LLM API 返回错误状态 {e.response.status_code}: "
                f"{e.response.text[:500]}"
            )
        except httpx.RequestError as e:
            raise RuntimeError(f"LLM API 请求失败: {e}")

    def extract_content(self, response: dict[str, Any]) -> str:
        """从响应中提取生成内容"""
        try:
            return response["choices"][0]["message"]["content"]
        except (KeyError, IndexError) as e:
            raise RuntimeError(f"LLM API 响应格式异常: {e}，原始响应: {str(response)[:500]}")

    def chat(self, messages: list[dict[str, str]], **kwargs) -> str:
        """便捷方法：直接返回文本内容"""
        response = self.chat_completion(messages, **kwargs)
        return self.extract_content(response)

    def chat_json(self, messages: list[dict[str, str]], **kwargs) -> dict:
        """便捷方法：要求返回 JSON 并解析"""
        kwargs.setdefault("response_format", {"type": "json_object"})
        content = self.chat(messages, **kwargs)
        # 清理可能的 markdown 代码块标记
        content = content.strip()
        if content.startswith("```"):
            # 去掉 ```json 和结尾 ```
            content = content.split("\n", 1)[1] if "\n" in content else content
            content = content.rsplit("```", 1)[0]
        try:
            return json.loads(content)
        except json.JSONDecodeError as e:
            raise RuntimeError(f"LLM 返回的内容不是合法 JSON: {e}，内容: {content[:500]}")
