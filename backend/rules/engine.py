"""
规则引擎：使用 python-docx 读取 docx，调用 GB/T 9704 规则逐项检查。
"""

from pathlib import Path
from typing import Optional

from docx import Document
from docx.shared import Pt, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH

from backend.core.models import Issue, IssueCategory, IssueSeverity
from backend.rules import gb_t9704 as gb


# ---------------------------------------------------------------------------
# 文档信息提取
# ---------------------------------------------------------------------------

def extract_document_info(doc: Document) -> gb.DocumentInfo:
    """从 python-docx Document 对象提取页面设置和段落信息"""
    section = doc.sections[0]  # 取第一个节

    # 页面尺寸（EMU → mm）
    page_w = gb._emu_to_mm(section.page_width) if section.page_width else gb.A4_WIDTH_MM
    page_h = gb._emu_to_mm(section.page_height) if section.page_height else gb.A4_HEIGHT_MM

    # 页边距（EMU → mm）
    margin_top = gb._emu_to_mm(section.top_margin) if section.top_margin else gb.MARGIN_TOP_MM
    margin_bottom = gb._emu_to_mm(section.bottom_margin) if section.bottom_margin else gb.MARGIN_BOTTOM_MM
    margin_left = gb._emu_to_mm(section.left_margin) if section.left_margin else gb.MARGIN_LEFT_MM
    margin_right = gb._emu_to_mm(section.right_margin) if section.right_margin else gb.MARGIN_RIGHT_MM

    # 提取段落信息
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text
        runs = para.runs

        # 取第一个 run 的字体/字号作为段落代表
        font_name = None
        font_size = None
        if runs:
            first_run = runs[0]
            if first_run.font.name:
                font_name = first_run.font.name
            if first_run.font.size:
                font_size = first_run.font.size.pt

        # 段落对齐
        alignment = None
        if para.alignment is not None:
            alignment = int(para.alignment)

        # 行距
        line_spacing = None
        pf = para.paragraph_format
        if pf.line_spacing is not None:
            # line_spacing 可能是 Pt 值或倍数
            try:
                line_spacing = float(pf.line_spacing)
            except (TypeError, ValueError):
                line_spacing = None

        # 首行缩进
        first_line_indent = None
        if pf.first_line_indent is not None:
            first_line_indent = pf.first_line_indent.pt

        paragraphs.append({
            "text": text,
            "font_name": font_name,
            "font_size": font_size,
            "alignment": alignment,
            "line_spacing": line_spacing,
            "first_line_indent": first_line_indent,
        })

    return gb.DocumentInfo(
        page_width_mm=page_w,
        page_height_mm=page_h,
        margin_top_mm=margin_top,
        margin_bottom_mm=margin_bottom,
        margin_left_mm=margin_left,
        margin_right_mm=margin_right,
        paragraphs=paragraphs,
    )


# ---------------------------------------------------------------------------
# 主引擎
# ---------------------------------------------------------------------------

class RuleEngine:
    """GB/T 9704 格式规则检查引擎"""

    def __init__(self, strictness: str = "normal"):
        self.strictness = strictness

    def check_document(self, doc_path: str | Path) -> list[Issue]:
        """检查整个文档，返回所有问题列表"""
        doc_path = Path(doc_path)
        if not doc_path.exists():
            raise FileNotFoundError(f"文件不存在: {doc_path}")

        doc = Document(str(doc_path))
        doc_info = extract_document_info(doc)

        all_issues: list[Issue] = []

        # 1. 页面幅面检查
        all_issues.extend(gb.check_page_size(doc_info, self.strictness))

        # 2. 页边距检查
        all_issues.extend(gb.check_page_margins(doc_info, self.strictness))

        # 3. 段落格式检查
        total = len(doc_info.paragraphs)
        for idx, para_info in enumerate(doc_info.paragraphs):
            text = para_info.get("text", "").strip()
            if not text:
                continue

            role = gb._identify_paragraph_role(text, para_info, idx, total)

            # 段落基础格式检查
            all_issues.extend(
                gb.check_paragraph_format(para_info, idx, role, self.strictness)
            )

            # 首行缩进检查（仅正文）
            if role == "main_body":
                indent_issue = gb.check_first_line_indent(
                    para_info, idx, text, self.strictness
                )
                if indent_issue:
                    all_issues.append(indent_issue)

        return all_issues

    def check_document_object(self, doc: Document) -> list[Issue]:
        """从已打开的 Document 对象检查（避免重复打开）"""
        doc_info = extract_document_info(doc)

        all_issues: list[Issue] = []

        all_issues.extend(gb.check_page_size(doc_info, self.strictness))
        all_issues.extend(gb.check_page_margins(doc_info, self.strictness))

        total = len(doc_info.paragraphs)
        for idx, para_info in enumerate(doc_info.paragraphs):
            text = para_info.get("text", "").strip()
            if not text:
                continue

            role = gb._identify_paragraph_role(text, para_info, idx, total)
            all_issues.extend(
                gb.check_paragraph_format(para_info, idx, role, self.strictness)
            )
            if role == "main_body":
                indent_issue = gb.check_first_line_indent(
                    para_info, idx, text, self.strictness
                )
                if indent_issue:
                    all_issues.append(indent_issue)

        return all_issues


# ---------------------------------------------------------------------------
# 便捷函数
# ---------------------------------------------------------------------------

def check_document_rules(doc_path: str | Path, strictness: str = "normal") -> list[Issue]:
    """便捷函数：检查文档格式"""
    engine = RuleEngine(strictness=strictness)
    return engine.check_document(doc_path)