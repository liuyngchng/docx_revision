"""
GB/T 9704—2012《党政机关公文格式》规则引擎

基于标准条款逐条检查，通过 python-docx 读取文档属性，
无需 COM/WPS 即可完成全部格式检查。

主要检查项（来自国标）：
  5.1  幅面：A4 (210mm×297mm)
  5.2  页边距：上37mm 下35mm 左28mm 右26mm
  5.3  版心：156mm×225mm
  7.3.1 标题字体：方正小标宋简体 二号 (22pt)
  7.3.2 主送机关：仿宋_GB2312 三号 (16pt)
  7.3.3 正文：仿宋_GB2312 三号 (16pt)，行距28磅
  7.3.4 附件说明：仿宋_GB2312 三号
  7.3.5 发文机关署名/成文日期：仿宋_GB2312 三号
  7.3.6 附注：仿宋_GB2312 三号
  7.4   抄送机关：仿宋_GB2312 四号 (14pt)
  7.5   印发机关/日期：仿宋_GB2312 四号
  7.6   页码：宋体 四号 (14pt)
  9     公文中表格
"""

from dataclasses import dataclass
from typing import Optional
import re

from backend.core.models import Issue, IssueCategory, IssueSeverity


# ---------------------------------------------------------------------------
# GB/T 9704 标准常量
# ---------------------------------------------------------------------------

# A4 纸张尺寸（EMU / 毫米）
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297

# 标准页边距（毫米）
MARGIN_TOP_MM = 37      # ±1mm
MARGIN_BOTTOM_MM = 35
MARGIN_LEFT_MM = 28     # 订口
MARGIN_RIGHT_MM = 26    # 翻口

# 版心尺寸（毫米）
CORE_WIDTH_MM = 156
CORE_HEIGHT_MM = 225

# 字体规范
FONT_BIAOSONG = "方正小标宋简体"
FONT_FANGSONG = "仿宋_GB2312"       # 或 "仿宋"
FONT_FANGSONG_ALT = "仿宋"
FONT_SONG = "宋体"
FONT_HEITI = "黑体"
FONT_KAITI = "楷体_GB2312"          # 或 "楷体"
FONT_KAITI_ALT = "楷体"

# 字号规范（磅值）
SIZE_ER_HAO = 22    # 二号 — 标题
SIZE_SAN_HAO = 16   # 三号 — 正文、主送机关等
SIZE_SI_HAO = 14    # 四号 — 抄送机关、印发机关、页码

# 行距规范（磅值 / 固定值）
LINE_SPACING_28 = 28.0  # pt — 正文标准行距

# 标题：二号，居中
# 主送机关：三号，顶格
# 正文首行缩进：2字符（约等于2个三号字 = 32pt）

# 段落间距
PARA_SPACING_BEFORE = 0   # 段前
PARA_SPACING_AFTER = 0    # 段后

# 页码：四号宋体，左右各放一条一字线（单页码居右，双页码居左）


# ---------------------------------------------------------------------------
# 字体名称模糊匹配
# ---------------------------------------------------------------------------

def _match_font(actual: Optional[str], expected_list: list[str]) -> bool:
    """检查实际字体是否匹配预期字体之一（模糊匹配）"""
    if not actual:
        return False
    actual_clean = actual.strip().lower().replace(" ", "").replace("_", "")
    for expected in expected_list:
        expected_clean = expected.strip().lower().replace(" ", "").replace("_", "")
        if actual_clean == expected_clean or expected_clean in actual_clean:
            return True
    return False


def _pt_to_half_pt(pt: float) -> int:
    """磅值转半磅（OOXML 中字号单位）"""
    return int(round(pt * 2))


# ---------------------------------------------------------------------------
# 规则检查上下文
# ---------------------------------------------------------------------------

@dataclass
class DocumentInfo:
    """从 python-docx 提取的文档元信息"""
    # 页面设置
    page_width_mm: float = A4_WIDTH_MM
    page_height_mm: float = A4_HEIGHT_MM
    margin_top_mm: float = MARGIN_TOP_MM
    margin_bottom_mm: float = MARGIN_BOTTOM_MM
    margin_left_mm: float = MARGIN_LEFT_MM
    margin_right_mm: float = MARGIN_RIGHT_MM

    # 段落列表
    paragraphs: list[dict] = None

    def __post_init__(self):
        if self.paragraphs is None:
            self.paragraphs = []


# ---------------------------------------------------------------------------
# 规则检查函数
# ---------------------------------------------------------------------------

# 每项检查返回 Optional[Issue]，无问题返回 None


def _emu_to_mm(emu: int) -> float:
    """EMU (English Metric Unit) 转毫米
    1 mm = 36000 EMU（不是 360000）"""
    return emu / 36000.0


def check_page_size(doc_info: DocumentInfo, strictness: str) -> list[Issue]:
    """检查幅面：A4 (210mm×297mm)，允许偏差 ±2mm"""
    issues = []
    tolerance = 2.0 if strictness == "strict" else 3.0

    w = doc_info.page_width_mm
    h = doc_info.page_height_mm

    if abs(w - A4_WIDTH_MM) > tolerance or abs(h - A4_HEIGHT_MM) > tolerance:
        issues.append(Issue(
            category=IssueCategory.PAGE_MARGIN,
            severity=IssueSeverity.ERROR,
            title="幅面尺寸不符合 A4 标准",
            description=f"当前页面尺寸为 {w:.1f}mm × {h:.1f}mm，"
                        f"标准要求 A4 (210mm × 297mm)",
            suggestion="请在页面设置中将纸张大小设为 A4。",
            source="rule",
        ))
    return issues


def check_page_margins(doc_info: DocumentInfo, strictness: str) -> list[Issue]:
    """检查页边距：上37mm 下35mm 左28mm 右26mm"""
    issues = []
    tolerance = 1.0 if strictness == "strict" else 2.0

    checks = [
        ("上边距", doc_info.margin_top_mm, MARGIN_TOP_MM),
        ("下边距", doc_info.margin_bottom_mm, MARGIN_BOTTOM_MM),
        ("左边距（订口）", doc_info.margin_left_mm, MARGIN_LEFT_MM),
        ("右边距（翻口）", doc_info.margin_right_mm, MARGIN_RIGHT_MM),
    ]

    for name, actual, expected in checks:
        if abs(actual - expected) > tolerance:
            issues.append(Issue(
                category=IssueCategory.PAGE_MARGIN,
                severity=IssueSeverity.ERROR,
                title=f"页{name}不符合标准",
                description=f"当前{name}为 {actual:.1f}mm，标准要求 {expected}mm",
                suggestion=f"请将{name}设置为 {expected}mm。",
                source="rule",
            ))

    return issues


def check_paragraph_format(para_info: dict, para_index: int,
                            role: str, strictness: str) -> list[Issue]:
    """
    根据段落角色检查格式。
    role 可能是：title / main_body / addressee / attachment / signature /
                  date / annotation / cc / issuer / page_number / unknown
    """
    issues = []

    # 注意：python-docx 未设置的属性会是 None，`.get(key, default)` 在 key 存在但值为
    # None 时仍返回 None，因此这里用 `or` 兜底，避免 NoneType 比较崩溃。
    font_name = para_info.get("font_name") or ""
    font_size = para_info.get("font_size") or 0   # pt
    alignment = para_info.get("alignment")  # 0=left, 1=center, 2=right, 3=justify
    line_spacing = para_info.get("line_spacing") or 0  # pt
    first_line_indent = para_info.get("first_line_indent") or 0  # pt
    text = para_info.get("text", "").strip()

    if not text:
        return issues  # 空段落跳过

    # ---- 标题 ----
    if role == "title":
        # 字体应为方正小标宋简体，二号 (22pt)
        if font_name and not _match_font(font_name, [FONT_BIAOSONG]):
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="标题字体不符合标准",
                description=f'标题"{text[:30]}..."当前字体为 "{font_name}"，'
                            f"标准要求使用 {FONT_BIAOSONG}",
                suggestion=f"请将标题字体设置为 {FONT_BIAOSONG}。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        if font_size > 0 and abs(font_size - SIZE_ER_HAO) > 1:
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="标题字号不符合标准",
                description=f'标题"{text[:30]}..."当前字号为 {font_size:.0f}pt，'
                            f"标准要求二号 ({SIZE_ER_HAO}pt)",
                suggestion=f"请将标题字号设置为二号 ({SIZE_ER_HAO}pt)。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        if alignment is not None and alignment != 1:  # 非居中
            issues.append(Issue(
                category=IssueCategory.PARAGRAPH_ALIGNMENT,
                severity=IssueSeverity.WARNING,
                title="标题应居中",
                description=f'标题"{text[:30]}..."未居中对齐',
                suggestion="请将标题设置为居中对齐。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))

    # ---- 主送机关 ----
    elif role == "addressee":
        if font_name and not _match_font(font_name, [FONT_FANGSONG, FONT_FANGSONG_ALT]):
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="主送机关字体不符合标准",
                description=f'主送机关"{text[:30]}..."当前字体为 "{font_name}"，'
                            f"标准要求使用 {FONT_FANGSONG}",
                suggestion=f"请将主送机关字体设置为 {FONT_FANGSONG}。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        if font_size > 0 and abs(font_size - SIZE_SAN_HAO) > 1:
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="主送机关字号不符合标准",
                description=f'主送机关"{text[:30]}..."当前字号为 {font_size:.0f}pt，'
                            f"标准要求三号 ({SIZE_SAN_HAO}pt)",
                suggestion=f"请将主送机关字号设置为三号 ({SIZE_SAN_HAO}pt)。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))

    # ---- 正文 ----
    elif role == "main_body":
        if font_name and not _match_font(font_name, [FONT_FANGSONG, FONT_FANGSONG_ALT]):
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="正文字体不符合标准",
                description=f'正文"{text[:30]}..."当前字体为 "{font_name}"，'
                            f"标准要求使用 {FONT_FANGSONG}",
                suggestion=f"请将正文字体设置为 {FONT_FANGSONG}。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        if font_size > 0 and abs(font_size - SIZE_SAN_HAO) > 1:
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="正文字号不符合标准",
                description=f'正文"{text[:30]}..."当前字号为 {font_size:.0f}pt，'
                            f"标准要求三号 ({SIZE_SAN_HAO}pt)",
                suggestion=f"请将正文字号设置为三号 ({SIZE_SAN_HAO}pt)。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        if line_spacing > 0 and abs(line_spacing - LINE_SPACING_28) > 2:
            issues.append(Issue(
                category=IssueCategory.LINE_SPACING,
                severity=IssueSeverity.WARNING,
                title="正文行距不符合标准",
                description=f'正文"{text[:30]}..."当前行距为 {line_spacing:.1f}pt，'
                            f"标准要求固定值 {LINE_SPACING_28:.0f} 磅",
                suggestion=f"请将行距设置为固定值 {LINE_SPACING_28:.0f} 磅。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))

    # ---- 抄送机关 / 印发机关 ----
    elif role in ("cc", "issuer"):
        expected_size = SIZE_SI_HAO
        if font_size > 0 and abs(font_size - expected_size) > 1:
            label = "抄送机关" if role == "cc" else "印发机关和印发日期"
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title=f"{label}字号不符合标准",
                description=f'"{text[:30]}..."当前字号为 {font_size:.0f}pt，'
                            f"标准要求四号 ({SIZE_SI_HAO}pt)",
                suggestion=f"请将字号设置为四号 ({SIZE_SI_HAO}pt)。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))

    # ---- 附件说明 ----
    elif role == "attachment":
        if font_name and not _match_font(font_name, [FONT_FANGSONG, FONT_FANGSONG_ALT]):
            issues.append(Issue(
                category=IssueCategory.FONT,
                severity=IssueSeverity.WARNING,
                title="附件说明字体不符合标准",
                description=f'附件说明"{text[:30]}..."当前字体为 "{font_name}"，'
                            f"标准要求使用 {FONT_FANGSONG}",
                suggestion=f"请将字体设置为 {FONT_FANGSONG}。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))

    return issues


def check_first_line_indent(para_info: dict, para_index: int,
                             text: str, strictness: str) -> Optional[Issue]:
    """检查正文首行缩进（2字符 ≈ 32pt 对于三号字）"""
    # 仅对有文字的正文段落检查
    if not text.strip():
        return None
    # 根据字号动态计算预期缩进
    font_size = para_info.get("font_size") or SIZE_SAN_HAO
    expected_indent = font_size * 2  # 2字符
    actual_indent = para_info.get("first_line_indent") or 0
    tolerance = 2 if strictness == "strict" else 5

    if actual_indent > 0 and abs(actual_indent - expected_indent) > tolerance:
        return Issue(
            category=IssueCategory.INDENT,
            severity=IssueSeverity.WARNING,
            title="首行缩进不符合标准",
            description=f'段落"{text[:30]}..."首行缩进为 {actual_indent:.0f}pt，'
                        f"预期 {expected_indent:.0f}pt（2字符）",
            suggestion="请设置首行缩进为2字符。",
            paragraph_index=para_index,
            original_text=text,
            source="rule",
        )
    elif actual_indent == 0:
        return Issue(
            category=IssueCategory.INDENT,
            severity=IssueSeverity.WARNING,
            title="正文段落缺少首行缩进",
            description=f'段落"{text[:30]}..."没有设置首行缩进，'
                        f"公文正文要求首行缩进2字符",
            suggestion="请设置首行缩进为2字符。",
            paragraph_index=para_index,
            original_text=text,
            source="rule",
        )
    return None


# ---------------------------------------------------------------------------
# 段落角色识别
# ---------------------------------------------------------------------------

def _identify_paragraph_role(
    para_text: str, para_info: dict, idx: int, total: int
) -> str:
    """根据位置和内容识别段落属于公文的哪个部分"""
    text = para_text.strip()
    if not text:
        return "empty"

    # 标题：一般在前3段，且含"关于...的"等特征
    if idx <= 2 and ("关于" in text or "报告" in text or "请示" in text
                     or "通知" in text or "函" in text or "决定" in text
                     or "意见" in text or "通报" in text or "批复" in text
                     or "议案" in text or "纪要" in text or "命令" in text):
        return "title"

    if idx <= 2 and len(text) < 80:
        return "title"

    # 主送机关：正文前，包含冒号结尾的行（如"各省、自治区、直辖市人民政府："）
    if idx <= 5 and (text.endswith("：") or text.endswith(":")) and len(text) < 200:
        return "addressee"

    # 附件说明
    if text.startswith("附件") and len(text) < 100:
        return "attachment"

    # 抄送机关
    if text.startswith("抄送") or "抄送：" in text or "抄送:" in text:
        return "cc"

    # 印发机关和日期
    if text.startswith("印发") or text.startswith("（印发") or "印发机关" in text:
        return "issuer"
    if idx >= total - 3 and "印发" in text:
        return "issuer"

    # 签名 / 成文日期（最后几段中的短段落）
    if idx >= total - 5 and len(text) < 30:
        return "signature"

    # 附注
    if text.startswith("（此件") or text.startswith("(此件"):
        return "annotation"

    # 默认正文
    return "main_body"