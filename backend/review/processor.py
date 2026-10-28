"""
审核处理器：整合规则引擎和 LLM 审查，生成统一的问题清单。
"""

from pathlib import Path
from typing import Optional

import docx

from backend.core.config import AppConfig, get_config
from backend.core.models import Issue, IssueCategory, IssueSeverity, ReviewResult
from backend.rules.engine import RuleEngine
from backend.llm.client import LlmClient
from backend.llm.prompts import SYSTEM_PROMPT, USER_PROMPT_TEMPLATE


class ReviewProcessor:
    """文档审核处理器"""

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or get_config()
        self._rule_engine = None
        self._llm_client = None

    @property
    def rule_engine(self) -> RuleEngine:
        if self._rule_engine is None:
            self._rule_engine = RuleEngine(
                strictness=self.config.review.strictness
            )
        return self._rule_engine

    @property
    def llm_client(self) -> LlmClient:
        if self._llm_client is None:
            self._llm_client = LlmClient(config=self.config.llm)
        return self._llm_client

    def review(self, doc_path: str | Path) -> ReviewResult:
        """对文档进行全面审核"""
        doc_path = Path(doc_path)
        result = ReviewResult(file_name=doc_path.name)

        all_issues: list[Issue] = []

        # ---------- 阶段 1：格式规则检查 ----------
        if self.config.review.enable_rules:
            try:
                rule_issues = self.rule_engine.check_document(doc_path)
                all_issues.extend(rule_issues)
            except Exception as e:
                result.error_message = f"格式规则检查失败: {e}"

        # ---------- 阶段 2：LLM 内容审核 ----------
        if self.config.review.enable_llm:
            try:
                llm_issues = self._llm_review(doc_path)
                all_issues.extend(llm_issues)
            except Exception as e:
                if result.error_message:
                    result.error_message += f"; LLM 审核失败: {e}"
                else:
                    result.error_message = f"LLM 审核失败: {e}"

        # 按严重程度排序：error > warning > info
        severity_order = {
            IssueSeverity.ERROR: 0,
            IssueSeverity.WARNING: 1,
            IssueSeverity.INFO: 2,
        }
        all_issues.sort(key=lambda i: severity_order.get(i.severity, 99))

        result.issues = all_issues
        return result

    def _llm_review(self, doc_path: Path) -> list[Issue]:
        """调用 LLM 进行内容审核"""
        # 1. 提取文档纯文本
        doc = docx.Document(str(doc_path))
        paragraphs_text = []
        for i, para in enumerate(doc.paragraphs):
            if para.text.strip():
                paragraphs_text.append(f"[P{i}] {para.text}")

        content = "\n".join(paragraphs_text)

        if len(content) > 15000:
            content = content[:15000] + "\n\n...（文档过长，已截断）"

        # 2. 构造消息
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": USER_PROMPT_TEMPLATE.format(
                file_name=doc_path.name,
                content=content,
            )},
        ]

        # 3. 调用 LLM
        response_json = self.llm_client.chat_json(messages)

        # 4. 解析为 Issue 对象
        issues = []
        for item in response_json.get("issues", []):
            # 映射类别
            category_map = {
                "document_type": IssueCategory.DOCUMENT_TYPE,
                "title": IssueCategory.TITLE,
                "language": IssueCategory.LANGUAGE,
                "logic": IssueCategory.LOGIC,
                "formatting_norm": IssueCategory.FORMATTING_NORM,
                "heading_structure": IssueCategory.HEADING_STRUCTURE,
                "content_integrity": IssueCategory.CONTENT_INTEGRITY,
            }
            cat = category_map.get(item.get("category", ""), IssueCategory.FORMATTING_NORM)

            # 映射严重程度
            sev_map = {
                "error": IssueSeverity.ERROR,
                "warning": IssueSeverity.WARNING,
                "info": IssueSeverity.INFO,
            }
            sev = sev_map.get(item.get("severity", ""), IssueSeverity.WARNING)

            issues.append(Issue(
                category=cat,
                severity=sev,
                title=item.get("title", "内容问题"),
                description=item.get("description", ""),
                suggestion=item.get("suggestion", ""),
                original_text=item.get("original_text", ""),
                revised_text=item.get("revised_text", ""),
                source="llm",
            ))

        return issues