"""
LLM API 客户端（OpenAI 兼容格式）
通过 httpx 直接调用，避免额外依赖 openai 包。
"""

import json
import logging
from typing import Any, Optional

import httpx

from backend.core.config import LlmConfig

logger = logging.getLogger(__name__)


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

        logger.info(
            "LLM 请求: model=%s, temperature=%.2f, max_tokens=%d, "
            "url=%s, system_prompt_len=%d, user_content_len=%d",
            payload["model"],
            payload["temperature"],
            payload["max_tokens"],
            url,
            len(messages[0]["content"]) if messages else 0,
            len(messages[-1]["content"]) if messages else 0,
        )

        try:
            with httpx.Client(timeout=self.timeout) as client:
                resp = client.post(url, headers=self._headers(), json=payload)
                resp.raise_for_status()
                result = resp.json()
                # 提取并记录首条 choice 的 token 使用量
                usage = result.get("usage", {})
                logger.info(
                    "LLM 响应: prompt_tokens=%s, completion_tokens=%s, total_tokens=%s",
                    usage.get("prompt_tokens", "?"),
                    usage.get("completion_tokens", "?"),
                    usage.get("total_tokens", "?"),
                )
                return result
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
        """便捷方法：要求返回 JSON 并解析。

        对 LLM 返回内容做容错处理：
        1. 去掉 markdown 代码块围栏
        2. 修复因输出截断导致的未闭合 JSON，尽量保留已生成的 issues
        """
        kwargs.setdefault("response_format", {"type": "json_object"})
        content = self.chat(messages, **kwargs)

        logger.debug("LLM 返回内容（前 500 字符）: %s", content[:500])

        # 清理可能的 markdown 代码块标记
        content = content.strip()
        if content.startswith("```"):
            # 去掉 ```json 和结尾 ```
            content = content.split("\n", 1)[1] if "\n" in content else content
            content = content.rsplit("```", 1)[0]

        try:
            parsed = json.loads(content)
            logger.info("LLM JSON 解析成功: issues=%d", len(parsed.get("issues", [])))
            return parsed
        except json.JSONDecodeError:
            logger.warning("LLM 返回的 JSON 不完整，尝试修复...")
            repaired = self._repair_truncated_json(content)
            if repaired is None:
                raise RuntimeError(
                    f"LLM 返回的内容不是合法 JSON 且无法修复，内容: {content[:500]}"
                )
            parsed = json.loads(repaired)
            logger.info(
                "LLM JSON 修复成功: 原 issues=%d, 修复后 issues=%d",
                len(json.loads(repaired).get("issues", [])),
                len(parsed.get("issues", [])),
            )
            return parsed

    @staticmethod
    def _repair_truncated_json(text: str):
        """修复被截断的 JSON（LLM 因 max_tokens 限制常输出不完整的 JSON）。

        策略：从 `issues` 数组中逐个提取**完整**的对象（用括号配平判断），
        重建 `{"overall_assessment": ..., "issues": [...]}`，
        从而丢弃最后一个被截断的问题对象，保留其余已生成的部分。
        """
        import re

        text = text.strip()

        # 提取 overall_assessment（通常在最前面，且是完整字符串）
        assessment = ""
        m = re.search(
            r'"overall_assessment"\s*:\s*"((?:[^"\\]|\\.)*)"',
            text,
            re.DOTALL,
        )
        if m:
            # 反转义还原
            assessment = json.loads('"' + m.group(1) + '"')

        # 定位 issues 数组起始的 '['
        arr_start = text.find('"issues"')
        if arr_start == -1:
            return None
        arr_start = text.find('[', arr_start)
        if arr_start == -1:
            return None

        # 从 '[' 之后逐个提取完整对象
        issues = []
        i = arr_start + 1
        n = len(text)
        while i < n:
            # 跳过空白和逗号
            while i < n and text[i] in ' \t\r\n,':
                i += 1
            if i >= n or text[i] != '{':
                break
            # 找到这个对象的结束 '}'（配平）
            depth = 0
            in_str = False
            escape = False
            j = i
            end = -1
            while j < n:
                ch = text[j]
                if escape:
                    escape = False
                    j += 1
                    continue
                if ch == '\\' and in_str:
                    escape = True
                    j += 1
                    continue
                if ch == '"':
                    in_str = not in_str
                    j += 1
                    continue
                if in_str:
                    j += 1
                    continue
                if ch == '{':
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        end = j
                        break
                j += 1
            if end == -1:
                break  # 该对象不完整，停止
            issues.append(text[i:end + 1])
            i = end + 1

        if not issues:
            return None

        rebuilt = {
            "overall_assessment": assessment,
            "issues": [json.loads(obj) for obj in issues],
        }
        return json.dumps(rebuilt, ensure_ascii=False)
