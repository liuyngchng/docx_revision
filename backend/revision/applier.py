"""
修订应用层 —— 通过 WPS/Word COM 接口，以修订模式（Track Changes）应用修改。

这是本系统的核心模块。原理：
  1. 通过 COM 启动 WPS/Word（ProgID 在 cfg.yml 中配置）
  2. 打开待审核文档，开启 TrackRevisions
  3. 逐条应用审核问题对应的修改（文本替换、格式修正）
  4. 另存为新 docx（保留全部修订标记）
  5. 关闭应用

修订标记由 WPS/Word 自动生成，格式 100% 原生，任意 Office 软件可正常显示。

实例安全策略（重要）：
  - 若本机已有用户正在使用的 WPS/Word 实例，则附着使用，绝不修改其
    可见性，也绝不执行 Quit()（避免关闭用户正在编辑的文档）；
  - 仅当由本系统自行启动实例时，才隐藏窗口并在结束后 Quit()。
  - COM 是单线程模型：全局锁保证同一时刻只有一个任务在操作 COM，
    其余任务排队等待。

WPS 兼容性说明：
  - WPS 文字的 COM ProgID 为 "Kwps.Application"（个人版；注意不是 "WPS.Application"，
    "WPS.Application" 是企业版的 ProgID）
  - WPS 兼容 Word 的大部分 COM 接口（Document、Paragraph、Range、Font 等）
  - 中文字体必须通过 Font.NameFarEast 设置，Font.Name 可能被忽略
  - SaveAs/SaveAs2 均可用，FileFormat=12 对应 wdFormatXMLDocument (.docx)
"""

import threading
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

# 保存格式
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

# Range.Information 参数：是否位于表格内
WD_WITH_IN_TABLE = 12

# 单位换算：1 毫米 = 72 / 25.4 磅（点）
MM_TO_PT = 72.0 / 25.4


def mm_to_pt(mm: float) -> float:
    """毫米转磅"""
    return mm * MM_TO_PT


# COM 单线程模型：全局锁，同一时刻只允许一个任务操作 WPS/Word
_COM_LOCK = threading.Lock()


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
    """

    def __init__(self, config: ComConfig):
        self.config = config
        self._app = None
        self._doc = None
        self._active_prog_id: Optional[str] = None  # 实际成功启动的 ProgID
        self._owns_app = False                      # 实例是否由本系统启动
        self._com_initialized = False               # 本线程是否由本类初始化了 COM

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
        启动或附着 WPS/Word 应用。

        - 本机已有运行中的实例（GetActiveObject 命中）→ 附着使用，
          不修改窗口可见性，结束后不 Quit；
        - 没有运行中的实例 → DispatchEx 启动独立实例，本系统拥有它，
          结束后 Quit。

        依次尝试 cfg.yml 中配置的 ProgID 列表，返回实际采用的 ProgID。
        """
        self._ensure_com_available()
        import win32com.client
        import pythoncom

        # 初始化当前线程的 COM。若线程已被以其他模式初始化（如 MTA，
        # RPC_E_CHANGED_MODE），忽略错误继续——进程外 COM 服务器在
        # MTA 下同样可用。
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
        except Exception:
            self._com_initialized = False

        prog_ids = self.config.prog_ids
        last_error = None

        for pid in prog_ids:
            try:
                # 1. 探测是否已有用户正在使用的实例
                existed = False
                try:
                    win32com.client.GetActiveObject(pid)
                    existed = True
                except Exception:
                    existed = False

                # 2. 附着已有实例，或启动新实例
                if existed:
                    self._app = win32com.client.GetActiveObject(pid)
                    self._owns_app = False
                else:
                    self._app = win32com.client.DispatchEx(pid)
                    self._owns_app = True

                self._active_prog_id = pid
                _ = self._app.Name  # 验证实例可用

                # 3. 只对自建实例设置窗口与告警，绝不打扰用户正在使用的实例
                if self._owns_app:
                    self._app.Visible = self.config.visible
                    if hasattr(self._app, "DisplayAlerts"):
                        self._app.DisplayAlerts = 0 if not self.config.display_alerts else -1
                break
            except Exception as e:
                last_error = e
                self._app = None
                self._owns_app = False
                continue

        if self._app is None:
            tried = ", ".join(prog_ids)
            raise RuntimeError(
                f"无法启动 COM 应用。\n"
                f"已尝试 ProgID: {tried}\n"
                f"请确认已安装 WPS 企业版/专业版或 Microsoft Word。\n"
                f"WPS 文字的 ProgID 应为 Kwps.Application。\n"
                f"最后一个错误: {last_error}"
            )

        return self._active_prog_id

    def open_document(self, doc_path: str | Path):
        """打开文档并开启修订模式"""
        if self._app is None:
            self.connect()

        doc_path = str(Path(doc_path).resolve())
        self._doc = self._app.Documents.Open(doc_path)

        # 开启修订模式
        self._doc.TrackRevisions = True

        # 开启格式修订记录（Word 支持，WPS 可能不支持，失败忽略）
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
        """关闭文档；仅当实例由本系统启动时才 Quit（保护用户正在编辑的文档）"""
        try:
            if self._doc is not None:
                self._doc.Close(SaveChanges=0)
        except Exception:
            pass
        finally:
            self._doc = None

        try:
            if self._app is not None and self._owns_app:
                self._app.Quit()
        except Exception:
            pass
        finally:
            self._app = None
            self._active_prog_id = None
            self._owns_app = False

        # 与 connect() 中的 CoInitialize 对称
        if self._com_initialized:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialized = False

    # ------------------------------------------------------------------
    # 段落定位
    # ------------------------------------------------------------------

    def _body_paragraph(self, paragraph_index: int):
        """
        按正文段落序号定位（0-based，与规则引擎 python-docx 的
        doc.paragraphs 枚举一致：不含表格内段落）。

        文档含表格时，COM 的 doc.Paragraphs 会包含表格内段落导致序号
        错位；这里跳过表格内段落保证两端一致。找不到时返回 None。
        """
        count = 0
        for p in self._doc.Paragraphs:
            try:
                if p.Range.Information(WD_WITH_IN_TABLE):
                    continue
            except Exception:
                pass
            if count == paragraph_index:
                return p
            count += 1
        return None

    def _body_para_or_warn(self, paragraph_index: int):
        para = self._body_paragraph(paragraph_index)
        if para is None:
            print(
                f"[WARN] 正文段落 {paragraph_index} 未找到"
                f"（序号可能因文档结构变化而错位），跳过该条修订"
            )
        return para

    # ------------------------------------------------------------------
    # 修改操作（以下操作均在 TrackRevisions 开启状态下执行，
    # 自动记录为修订）
    # ------------------------------------------------------------------

    def replace_text(self, find_text: str, replace_text: str):
        """全文查找替换（记录为修订）"""
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
                         left_mm: float, right_mm: float) -> bool:
        """设置页边距（记录为修订）"""
        for section in self._doc.Sections:
            ps = section.PageSetup
            ps.TopMargin = mm_to_pt(top_mm)
            ps.BottomMargin = mm_to_pt(bottom_mm)
            ps.LeftMargin = mm_to_pt(left_mm)
            ps.RightMargin = mm_to_pt(right_mm)
        return True

    def set_page_size(self, width_mm: float, height_mm: float) -> bool:
        """设置纸张大小（记录为修订）"""
        for section in self._doc.Sections:
            ps = section.PageSetup
            ps.PageWidth = mm_to_pt(width_mm)
            ps.PageHeight = mm_to_pt(height_mm)
        return True

    def set_paragraph_font(self, paragraph_index: int, font_name: Optional[str],
                           font_size_pt: Optional[float] = None,
                           alignment: Optional[int] = None,
                           line_spacing_pt: Optional[float] = None) -> bool:
        """
        设置段落字体格式（全部记录为格式修订 w:rPrChange / w:pPrChange）。

        WPS 注意：中文字体必须同时设置 Name（西文）和 NameFarEast（东亚），
        否则 WPS 可能不会正确识别中文字体名。
        """
        para = self._body_para_or_warn(paragraph_index)
        if para is None:
            return False
        try:
            rng = para.Range

            if font_name:
                rng.Font.Name = font_name
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

            return True
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 字体失败: {e}")
            return False

    def set_first_line_indent(self, paragraph_index: int, indent_chars: float = 2.0) -> bool:
        """设置首行缩进（字符单位，记录为格式修订 w:pPrChange）"""
        para = self._body_para_or_warn(paragraph_index)
        if para is None:
            return False
        try:
            # 优先字符单位（中文公文标准做法），失败回退磅值
            try:
                para.CharacterUnitFirstLineIndent = indent_chars
            except Exception:
                para.FirstLineIndent = indent_chars * 16  # 按三号字 16pt 换算
            return True
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 缩进失败: {e}")
            return False

    def set_paragraph_alignment(self, paragraph_index: int, alignment: int) -> bool:
        """设置段落对齐方式"""
        para = self._body_para_or_warn(paragraph_index)
        if para is None:
            return False
        try:
            para.Alignment = alignment
            return True
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 对齐失败: {e}")
            return False

    def set_paragraph_line_spacing(self, paragraph_index: int, spacing_pt: float) -> bool:
        """设置段落行距为固定值"""
        para = self._body_para_or_warn(paragraph_index)
        if para is None:
            return False
        try:
            para.LineSpacingRule = WD_LINE_SPACE_EXACTLY
            para.LineSpacing = spacing_pt
            return True
        except Exception as e:
            print(f"[WARN] 设置段落 {paragraph_index} 行距失败: {e}")
            return False

    # ------------------------------------------------------------------
    # 高层接口：应用问题清单
    # ------------------------------------------------------------------

    def apply_issues(self, issues: list[Issue]) -> dict:
        """
        根据问题清单应用修改。
        返回统计信息：{"applied": N, "skipped": N, "warnings": [...]}
        """
        applied = 0
        skipped = 0
        warnings: list[str] = []

        for issue in issues:
            try:
                ok, warns = self._apply_single_issue(issue)
                if ok:
                    applied += 1
                else:
                    skipped += 1
                warnings.extend(warns)
            except Exception as e:
                print(f"[WARN] 应用问题 {issue.id} ({issue.title}) 失败: {e}")
                skipped += 1
                warnings.append(f"{issue.id} {issue.title}: {e}")

        return {"applied": applied, "skipped": skipped, "warnings": warnings}

    def _apply_single_issue(self, issue: Issue) -> tuple[bool, list[str]]:
        """应用单条问题对应的修改。返回 (是否应用了至少一项, 告警列表)。"""
        from backend.rules import gb_t9704 as gb

        applied = False
        warnings: list[str] = []

        # ---- 文本替换（LLM 内容修改）----
        if issue.original_text and issue.revised_text:
            if self.replace_text_one(issue.original_text, issue.revised_text):
                applied = True
            else:
                warnings.append(
                    f"文本替换失败（未找到原文片段，可能标点/空白/截断不一致）: "
                    f"{issue.original_text[:30]}..."
                )

        # ---- 纸张幅面 ----
        if issue.category == IssueCategory.PAGE_SIZE:
            self.set_page_size(gb.A4_WIDTH_MM, gb.A4_HEIGHT_MM)
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
            elif any(kw in (issue.title or "") for kw in ("抄送", "印发", "页码", "文号", "签发人")):
                target_font = gb.FONT_FANGSONG
                target_size = gb.SIZE_SI_HAO if any(kw in (issue.title or "") for kw in ("抄送", "印发", "页码")) else gb.SIZE_SAN_HAO
            else:
                target_font = gb.FONT_FANGSONG
                target_size = gb.SIZE_SAN_HAO

            if self.set_paragraph_font(issue.paragraph_index, target_font, target_size):
                applied = True

        # ---- 段落对齐 ----
        if issue.category == IssueCategory.PARAGRAPH_ALIGNMENT \
                and issue.paragraph_index is not None:
            if "居中" in (issue.title or ""):
                ok = self.set_paragraph_alignment(issue.paragraph_index, WD_ALIGN_CENTER)
            elif "两端" in (issue.title or ""):
                ok = self.set_paragraph_alignment(issue.paragraph_index, WD_ALIGN_JUSTIFY)
            else:
                ok = False
            if ok:
                applied = True

        # ---- 首行缩进 ----
        if issue.category == IssueCategory.INDENT and issue.paragraph_index is not None:
            if self.set_first_line_indent(issue.paragraph_index, 2.0):
                applied = True

        # ---- 行距 ----
        if issue.category == IssueCategory.LINE_SPACING \
                and issue.paragraph_index is not None:
            if self.set_paragraph_line_spacing(issue.paragraph_index, gb.LINE_SPACING_28):
                applied = True

        return applied, warnings

    # ------------------------------------------------------------------
    # 总入口
    # ------------------------------------------------------------------

    def apply(self, input_path: str | Path, issues: list[Issue],
              output_path: str | Path) -> dict:
        """
        完整流程：排队 → 打开 → 修订 → 保存 → 关闭。

        全局锁串行化所有 COM 任务（WPS/Word 单实例模型），
        等待超时抛异常。返回 {"path": 输出路径, "stats": 应用统计}。
        """
        input_path = Path(input_path)
        output_path = Path(output_path)

        # 排队等待 COM 资源（最多 10 分钟）
        if not _COM_LOCK.acquire(timeout=600):
            raise RuntimeError("等待 WPS/Word COM 资源超时，请稍后重试")

        try:
            self.connect()
            self.open_document(input_path)
            try:
                stats = self.apply_issues(issues)
                self.save_as(output_path)
            finally:
                self.close()
        finally:
            _COM_LOCK.release()

        return {"path": str(output_path), "stats": stats}


# ---------------------------------------------------------------------------
# 便捷函数
# ---------------------------------------------------------------------------

def apply_revisions(input_path: str | Path, issues: list[Issue],
                    output_path: str | Path, com_config: ComConfig) -> dict:
    """便捷函数：通过 COM 应用修订，返回 {"path": ..., "stats": ...}"""
    applier = ComRevisionApplier(com_config)
    return applier.apply(input_path, issues, output_path)
