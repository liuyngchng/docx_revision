"""
修订应用层 —— 通过 WPS/Word COM 接口，以修订模式（Track Changes）应用修改。

这是本系统的核心模块。原理：
  1. 通过 COM 启动 WPS/Word（ProgID 在 cfg.yml 中配置）
  2. 打开待审核文档，开启 TrackRevisions
  3. 逐条应用审核问题对应的修改（文本替换、格式修正）
  4. 另存为新 docx（保留全部修订标记）
  5. 关闭应用

修订标记由 WPS/Word 自动生成，格式 100% 原生，任意 Office 软件可正常显示。

WPS 企业版兼容性说明：
  - WPS 文字组件的 COM ProgID 为 "WPS.Application"
  - WPS 兼容 Word 的大部分 COM 接口（Document、Paragraph、Range、Font 等）
  - 中文字体必须通过 Font.NameFarEast 设置，Font.Name 可能被忽略
  - SaveAs/SaveAs2 均可用，FileFormat=12 对应 wdFormatXMLDocument (.docx)
"""

from pathlib import Path
from typing import Optional

from backend.core.config import ComConfig
from backend.core.models import Issue, IssueCategory

# ---------------------------------------------------------------------------
# WPS/Word COM 常量
# ---------------------------------------------------------------------------

# 查找替换
WD_REPLACE_NONE = 0
WD_REPLACE_ONE = 1
WD_REPLACE_ALL = 2

# 保存格式 — 注意 WPS 中这些常量值可能与 Word 一致
WD_FORMAT_DOCX = 12          # wdFormatXMLDocument
WD_FORMAT_DOCX_DEFAULT = 16  # wdFormatDocumentDefault

# 段落对齐 (WdParagraphAlignment)
WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_ALIGN_RIGHT = 2
WD_ALIGN_JUSTIFY = 3

# 行距规则 (WdLineSpacing)
WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_1PT5 = 1
WD_LINE_SPACE_DOUBLE = 2
WD_LINE_SPACE_AT_LEAST = 3
WD_LINE_SPACE_EXACTLY = 4
WD_LINE_SPACE_MULTIPLE = 5

# 单位换算：1 毫米 = 72 / 25.4 磅（点）
MM_TO_PT = 72.0 / 25.4


def mm_to_pt(mm: float) -> float:
    """毫米转磅"""
    return mm * MM_TO_PT


# ---------------------------------------------------------------------------
# COM 应用器
# ---------------------------------------------------------------------------

class ComRevisionApplier:
    """
    通过 WPS/Word COM 接口应用修订。

    支持多 ProgID 自动探测（cfg.yml 中 prog_id 可填字符串或列表），
    按顺序尝试，第一个启动成功的被采用。

    在 TrackRevisions = True 下执行所有操作，WPS/Word 自动记录原生修订标记：
      - 文本修改 → w:ins / w:del
      - 格式修改 → w:rPrChange / w:pPrChange / w:sectPrChange

    使用方式：
        applier = ComRevisionApplier(com_config)
        output_path = applier.apply(input_path, issues, output_path)
    """

    def __init__(self, config: ComConfig):
        self.config = config
        self._app = None
        self._doc = None
        self._active_prog_id: Optional[str] = None  # 实际成功启动的 ProgID

    # ------------------------------------------------------------------
    # 生命周期
    # ------------------------------------------------------------------

    def _ensure_com_available(self):
        """检查 pywin32 是否可用"""
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            raise RuntimeError(
                "未安装 pywin32，无法使用 COM 接口。\n"
                "请以管理员身份运行：pip install pywin32"
            )

    def connect(self) -> str:
        """
        启动 WPS/Word 应用。
        依次尝试 cfg.yml 中配置的 ProgID 列表，返回第一个成功启动的 ProgID。
        """
        self._ensure_com_available()
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()

        prog_ids = self.config.prog_ids
        last_error = None

        for pid in prog_ids:
            try:
                self._app = win32com.client.Dispatch(pid)
                self._active_prog_id = pid
                # 验证是否真的可用（有些 ProgID 注册了但没有安装）
                _ = self._app.Name
                break
            except Exception as e:
                last_error = e
                self._app = None
                continue

        if self._app is None:
            tried = ", ".join(prog_ids)
            raise RuntimeError(
                f"无法启动 COM 应用。\n"
                f"已尝试 ProgID: {tried}\n"
                f"请确认已安装 WPS 企业版或 Microsoft Word。\n"
                f"最后一个错误: {last_error}"
            )

        self._app.Visible = self.config.visible
        if hasattr(self._app, "DisplayAlerts"):
            self._app.DisplayAlerts = 0 if not self.config.display_alerts else -1

        return self._active_prog_id

    def open_document(self, doc_path: str | Path):
        """打开文档并开启修订模式"""
        if self._app is None:
            self.connect()

        doc_path = str(Path(doc_path).resolve())
        self._doc = self._app.Documents.Open(doc_path)

        # 开启修订模式
        self._doc.TrackRevisions = True

        # 开启格式修订记录
        # WPS 企业版的属性名可能是 TrackFormatting，Word 也是
        try:
            self._doc.TrackFormatting = True
        except Exception:
            pass

        return self._doc

    def save_as(self, output_path: str | Path):
        """另存为 docx（保留修订标记）"""
        if self._doc is None:
            raise RuntimeError("没有打开的文档")

        output_path = str(Path(output_path).resolve())
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        try:
            self._doc.SaveAs2(output_path, FileFormat=WD_FORMAT_DOCX)
        except Exception:
            try:
                self._doc.SaveAs(output_path, FileFormat=WD_FORMAT_DOCX)
            except Exception:
                # 最后回退：不指定格式，依赖扩展名
                self._doc.SaveAs(output_path)

    def close(self):
        """关闭文档和应用"""
        try:
            if self._doc is not None:
                self._doc.Close(SaveChanges=0)
        except Exception:
            pass
        finally:
            self._doc = None

        try:
            if self._app is not None:
                self._app.Quit()
        except Exception:
            pass
        finally:
            self._app = None
            self._active_prog_id = None

    # ------------------------------------------------------------------
    # 修改操作（以下操作均在 TrackRevisions 开启状态下执行，
    # 自动记录为修订）
    # ------------------------------------------------------------------

    def replace_text(self, find_text: str, replace_text: str):
        """全文查找替换（记录为修订）

        在 TrackRevisions 开启时，Find.Execute 的替换会自动记录为修订：
        原文字标记删除 (w:del)，新文字标记插入 (w:ins)。
        """
        if not find_text or not replace_text:
            return

        find = self._doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Execute(
            FindText=find_text,
            ReplaceWith=replace_text,
            Replace=WD_REPLACE_ALL,
            MatchCase=False,
            MatchWholeWord=False,
            MatchWildcards=False,
            Forward=True,
            Wrap=1,  # wdFindContinue
        )

    def replace_text_one(self, find_text: str, replace_text: str) -> bool:
        """替换单处（用于精确控制），返回是否替换成功"""
        if not find_text:
            return False

        find = self._doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        return bool(find.Execute(
            FindText=find_text,
            ReplaceWith=replace_text,
            Replace=WD_REPLACE_ONE,
            Forward=True,
            Wrap=1,
        ))

    def set_page_margins(self, top_mm: float, bottom_mm: float,
                         left_mm: float, right_mm: float):
        """设置页边距（记录为修订）"""
        for section in self._doc.Sections:
            ps = section.PageSetup
            ps.TopMargin = mm_to_pt(top_mm)
            ps.BottomMargin = mm_to_pt(bottom_mm)
            ps.LeftMargin = mm_to_pt(left_mm)
            ps.RightMargin = mm_to_pt(right_mm)

    def set_page_size(self, width_mm: float, height_mm: float):
        """设置纸张大小（记录为修订）"""
        for section in self._doc.Sections:
            ps = section.PageSetup
            ps.PageWidth = mm_to_pt(width_mm)
            ps.PageHeight = mm_to_pt(height_mm)

    def set_paragraph_font(self, paragraph_index: int, font_name: Optional[str],
                           font_size_pt: Optional[float] = None,
                           alignment: Optional[int] = None,
                           line_spacing_pt: Optional[float] = None):
        """
        设置段落字体格式（全部记录为格式修订 w:rPrChange / w:pPrChange）。

        WPS 企业版注意：中文字体必须同时设置 Name（西文）和 NameFarEast（东亚），
        否则 WPS 可能不会正确识别中文字体名。
        """
        com_index = paragraph_index + 1
        try:
            para = self._doc.Paragraphs(com_index)
            rng = para.Range

            if font_name:
                rng.Font.Name = font_name
                # WPS 中文字体关键：NameFarEast
                try:
                    rng.Font.NameFarEast = font_name
                except Exception:
                    pass

            if font_size_pt is not None:
                rng.Font.Size = font_size_pt

            if alignment is not None:
                para.Alignment = alignment

            if line_spacing_pt is not None:
                para.LineSpacingRule = WD_LINE_SPACE_EXACTLY
                para.LineSpacing = line_spacing_pt

        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 字体失败: {e}")

    def set_first_line_indent(self, paragraph_index: int, indent_pt: float):
        """设置首行缩进（记录为格式修订 w:pPrChange）"""
        com_index = paragraph_index + 1
        try:
            para = self._doc.Paragraphs(com_index)
            para.FirstLineIndent = indent_pt
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 缩进失败: {e}")

    def set_paragraph_alignment(self, paragraph_index: int, alignment: int):
        """设置段落对齐方式"""
        com_index = paragraph_index + 1
        try:
            para = self._doc.Paragraphs(com_index)
            para.Alignment = alignment
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 对齐失败: {e}")

    def set_paragraph_line_spacing(self, paragraph_index: int, spacing_pt: float):
        """设置段落行距为固定值"""
        com_index = paragraph_index + 1
        try:
            para = self._doc.Paragraphs(com_index)
            para.LineSpacingRule = WD_LINE_SPACE_EXACTLY
            para.LineSpacing = spacing_pt
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 行距失败: {e}")

    # ------------------------------------------------------------------
    # 高层接口：应用问题清单
    # ------------------------------------------------------------------

    def apply_issues(self, issues: list[Issue]) -> dict:
        """
        根据问题清单应用修改。
        返回统计信息：{"applied": N, "skipped": N}
        """
        applied = 0
        skipped = 0

        for issue in issues:
            try:
                if self._apply_single_issue(issue):
                    applied += 1
                else:
                    skipped += 1
            except Exception as e:
                print(f"[WARN] 应用问题 {issue.id} ({issue.title}) 失败: {e}")
                skipped += 1

        return {"applied": applied, "skipped": skipped}

    def _apply_single_issue(self, issue: Issue):
        """应用单条问题对应的修改。所有修改都在 TrackRevisions 开启下进行。"""
        from backend.rules import gb_t9704 as gb

        applied = False

        # ---- 文本替换（LLM 内容修改） ----
        if issue.original_text and issue.revised_text:
            self.replace_text_one(issue.original_text, issue.revised_text)
            applied = True

        # ---- 页边距 ----
        if issue.category == IssueCategory.PAGE_MARGIN:
            self.set_page_margins(
                gb.MARGIN_TOP_MM, gb.MARGIN_BOTTOM_MM,
                gb.MARGIN_LEFT_MM, gb.MARGIN_RIGHT_MM,
            )
            applied = True

        # ---- 字体/字号（段落级） ----
        if issue.category == IssueCategory.FONT and issue.paragraph_index is not None:
            if "标题" in (issue.title or ""):
                target_font = gb.FONT_BIAOSONG
                target_size = gb.SIZE_ER_HAO
            elif any(kw in (issue.title or "") for kw in ("抄送", "印发", "页码")):
                target_font = gb.FONT_FANGSONG
                target_size = gb.SIZE_SI_HAO
            else:
                target_font = gb.FONT_FANGSONG
                target_size = gb.SIZE_SAN_HAO

            self.set_paragraph_font(
                issue.paragraph_index, target_font, target_size,
            )
            applied = True

        # ---- 段落对齐 ----
        if issue.category == IssueCategory.PARAGRAPH_ALIGNMENT \
                and issue.paragraph_index is not None:
            if "居中" in (issue.title or ""):
                self.set_paragraph_alignment(issue.paragraph_index, WD_ALIGN_CENTER)
            elif "两端" in (issue.title or ""):
                self.set_paragraph_alignment(issue.paragraph_index, WD_ALIGN_JUSTIFY)
            applied = True

        # ---- 首行缩进 ----
        if issue.category == IssueCategory.INDENT and issue.paragraph_index is not None:
            indent_pt = gb.SIZE_SAN_HAO * 2  # 2字符 = 2×16pt
            self.set_first_line_indent(issue.paragraph_index, indent_pt)
            applied = True

        # ---- 行距 ----
        if issue.category == IssueCategory.LINE_SPACING \
                and issue.paragraph_index is not None:
            self.set_paragraph_line_spacing(issue.paragraph_index, gb.LINE_SPACING_28)
            applied = True

        return applied

    # ------------------------------------------------------------------
    # 总入口
    # ------------------------------------------------------------------

    def apply(self, input_path: str | Path, issues: list[Issue],
              output_path: str | Path) -> str:
        """
        完整流程：打开 → 修订 → 保存 → 关闭。
        返回输出文件路径。
        """
        input_path = Path(input_path)
        output_path = Path(output_path)

        self.connect()
        self.open_document(input_path)

        try:
            stats = self.apply_issues(issues)
            self.save_as(output_path)
        finally:
            self.close()

        return str(output_path)


# ---------------------------------------------------------------------------
# 便捷函数
# ---------------------------------------------------------------------------

def apply_revisions(input_path: str | Path, issues: list[Issue],
                    output_path: str | Path, com_config: ComConfig) -> str:
    """便捷函数：通过 COM 应用修订"""
    applier = ComRevisionApplier(com_config)
    return applier.apply(input_path, issues, output_path)