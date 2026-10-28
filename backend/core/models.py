"""
数据模型：问题、审核结果等结构定义。
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional
import uuid
from datetime import datetime


# ---------------------------------------------------------------------------
# 问题分类
# ---------------------------------------------------------------------------

class IssueCategory(str, Enum):
    """审核问题的分类"""
    # 格式类（规则引擎）
    PAGE_MARGIN = "page_margin"               # 页边距
    FONT = "font"                             # 字体字号
    LINE_SPACING = "line_spacing"             # 行距
    INDENT = "indent"                         # 缩进
    HEADER_FOOTER = "header_footer"           # 页眉页脚
    PAGE_NUMBER = "page_number"               # 页码
    TABLE = "table"                           # 表格格式
    SIGNATURE = "signature"                   # 印章/签名
    PARAGRAPH_ALIGNMENT = "paragraph_alignment"  # 段落对齐

    # 内容类（LLM）
    DOCUMENT_TYPE = "document_type"           # 文种选择
    TITLE = "title"                           # 标题规范
    LANGUAGE = "language"                     # 语言规范
    LOGIC = "logic"                           # 逻辑问题
    FORMATTING_NORM = "formatting_norm"       # 格式规范（LLM视角）
    HEADING_STRUCTURE = "heading_structure"   # 层次结构
    CONTENT_INTEGRITY = "content_integrity"   # 内容完整性


class IssueSeverity(str, Enum):
    """问题严重程度"""
    ERROR = "error"       # 错误——必须修改
    WARNING = "warning"   # 警告——建议修改
    INFO = "info"         # 提示——可选修改


# ---------------------------------------------------------------------------
# 具体问题
# ---------------------------------------------------------------------------

@dataclass
class Issue:
    """审核发现的单条问题"""
    id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])
    category: IssueCategory = IssueCategory.FORMATTING_NORM
    severity: IssueSeverity = IssueSeverity.WARNING
    title: str = ""                      # 简短描述
    description: str = ""                # 详细说明
    suggestion: str = ""                 # 修改建议
    # 定位信息（用于 COM 层定位修改位置）
    paragraph_index: Optional[int] = None  # 段落序号（0-based）
    run_index: Optional[int] = None        # 文字块序号（0-based）
    original_text: Optional[str] = None    # 原文
    revised_text: Optional[str] = None     # 建议修改后的文字
    # 来源
    source: str = "rule"                   # "rule" 或 "llm"


# ---------------------------------------------------------------------------
# 审核结果
# ---------------------------------------------------------------------------

@dataclass
class ReviewResult:
    """完整审核结果"""
    task_id: str = field(default_factory=lambda: str(uuid.uuid4())[:12])
    file_name: str = ""
    issues: list[Issue] = field(default_factory=list)
    status: str = "pending"               # pending / processing / done / error
    error_message: Optional[str] = None
    started_at: Optional[datetime] = None
    finished_at: Optional[datetime] = None

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == IssueSeverity.ERROR)

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == IssueSeverity.WARNING)

    @property
    def info_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == IssueSeverity.INFO)

    @property
    def total_issues(self) -> int:
        return len(self.issues)


# ---------------------------------------------------------------------------
# API 交互
# ---------------------------------------------------------------------------

@dataclass
class TaskStatus:
    """任务状态（返回给前端）"""
    task_id: str
    status: str
    file_name: str
    total_issues: int
    error_count: int
    warning_count: int
    info_count: int
    issues: list[Issue] = field(default_factory=list)
    download_url: Optional[str] = None
    error_message: Optional[str] = None