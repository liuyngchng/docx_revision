"""
GB/T 9704—2012《党政机关公文格式》规则引擎

基于标准条款逐条检查，通过 python-docx 读取文档属性，
无需 COM/WPS 即可完成全部格式检查。

主要检查项（来自国标）：
  5.1  幅面：A4 (210mm×297mm)
  5.2  页边距：上37mm 下35mm 左28mm 右26mm
  5.3  版心：156mm×225mm
  7.2  版头：发文字号（3号仿宋）、签发人（3号仿宋，姓名楷体）
       （机关标志/份号/密级/紧急程度按角色跳过格式检查，避免误报）
  7.3.1 标题字体：小标宋体（方正小标宋简体/GBK/华文中宋）二号 (22pt)
  7.3.2 主送机关：仿宋_GB2312 三号 (16pt)
  7.3.3 正文：仿宋_GB2312 三号 (16pt)，行距28磅，首行缩进2字符
  7.3.4 附件说明：仿宋_GB2312 三号
  7.3.5 发文机关署名/成文日期：仿宋_GB2312 三号
  7.3.6 附注：仿宋_GB2312 三号
  7.4   抄送机关：仿宋_GB2312 四号 (14pt)
  7.5   印发机关/日期：仿宋_GB2312 四号
  7.6   页码：宋体 四号 (14pt)，一字线（读 footer XML 检查）
  9     公文中表格
"""

from dataclasses import dataclass
from typing import Optional
import re
import zipfile

from lxml import etree
from docx.oxml.ns import qn

from backend.core.models import Issue, IssueCategory, IssueSeverity


# ---------------------------------------------------------------------------
# GB/T 9704 标准常量
# ---------------------------------------------------------------------------

# A4 纸张尺寸（毫米）
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
FONT_BIAOSONG_ALT = "小标宋"       # 匹配 方正小标宋_GBK / 方正小标宋GBK 等变体
FONT_HUAWEN_ZHONGSONG = "华文中宋"  # 部分机关实际使用的标题字体
FONT_FANGSONG = "仿宋_GB2312"       # 或 "仿宋"
FONT_FANGSONG_ALT = "仿宋"
FONT_SONG = "宋体"
FONT_HEITI = "黑体"
FONT_KAITI = "楷体_GB2312"          # 或 "楷体"
FONT_KAITI_ALT = "楷体"

# 标题字体允许列表（含各变体）
TITLE_FONTS = [FONT_BIAOSONG, FONT_BIAOSONG_ALT, FONT_HUAWEN_ZHONGSONG]

# 字号规范（磅值）
SIZE_ER_HAO = 22    # 二号 — 标题
SIZE_SAN_HAO = 16   # 三号 — 正文、主送机关等
SIZE_SI_HAO = 14    # 四号 — 抄送机关、印发机关、页码

# 行距规范（磅值 / 固定值）
LINE_SPACING_28 = 28.0  # pt — 正文标准行距

# 标题：二号，居中
# 主送机关：三号，顶格
# 正文首行缩进：2字符
# 页码：四号宋体，左右各放一条一字线（单页码居右，双页码居左）


# ---------------------------------------------------------------------------
# 段落角色识别辅助
# ---------------------------------------------------------------------------

# 发文字号：机关代字〔年份〕序号（六角括号年份 + 序号 + "号"）
_DOC_NUMBER_RE = re.compile(r'〔\d{4}〕\s*\d+\s*号')

# 成文日期：阿拉伯数字年月日
_DATE_RE = re.compile(r'^\d{4}年\d{1,2}月\d{1,2}日$')

# 份号：短数字行
_SERIAL_RE = re.compile(r'^\d{3,6}$')

# 15 种法定文种
_DOC_TYPES = (
    "决议", "决定", "命令", "公报", "公告", "通告", "意见",
    "通知", "通报", "报告", "请示", "批复", "议案", "函", "纪要",
)

# 机关标志（红头）特征词
_ORG_SIGN_HINTS = ("中国共产党", "人民政府", "委员会", "党委", "党组")


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
# XML 辅助
# ---------------------------------------------------------------------------

def _rpr_font(rpr) -> tuple[Optional[str], Optional[float]]:
    """从 w:rPr 元素读取 (中文字体名, 字号磅值)，读不到返回 (None, None)。"""
    if rpr is None:
        return None, None

    name = None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        name = (
            rfonts.get(qn("w:eastAsia"))
            or rfonts.get(qn("w:ascii"))
            or rfonts.get(qn("w:hAnsi"))
        )

    size = None
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        size = int(sz.get(qn("w:val"))) / 2.0

    return name, size


# ---------------------------------------------------------------------------
# 规则检查函数
# ---------------------------------------------------------------------------

# 每项检查返回 Optional[Issue] 或 list[Issue]，无问题返回空


def _emu_to_mm(emu: int) -> float:
    """EMU (English Metric Unit) 转毫米
    1 mm = 36000 EMU"""
    return emu / 36000.0


def check_page_size(doc_info: DocumentInfo, strictness: str) -> list[Issue]:
    """检查幅面：A4 (210mm×297mm)，允许偏差 ±2mm"""
    issues = []
    tolerance = 2.0 if strictness == "strict" else 3.0

    w = doc_info.page_width_mm
    h = doc_info.page_height_mm

    if abs(w - A4_WIDTH_MM) > tolerance or abs(h - A4_HEIGHT_MM) > tolerance:
        issues.append(Issue(
            category=IssueCategory.PAGE_SIZE,
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


def _check_font_size(issues: list[Issue], para_info: dict, para_index: int,
                     role_label: str, text: str, expected_fonts: list[str],
                     expected_size: float, extra_font_note: str = "") -> None:
    """通用字体/字号检查，生成 Issue（内部辅助）。"""
    font_name = para_info.get("font_name", "")
    font_size = para_info.get("font_size", 0) or 0
    preview = text[:30] + ("..." if len(text) > 30 else "")

    if font_name and not _match_font(font_name, expected_fonts):
        issues.append(Issue(
            category=IssueCategory.FONT,
            severity=IssueSeverity.WARNING,
            title=f"{role_label}字体不符合标准",
            description=f'{role_label}"{preview}"当前字体为 "{font_name}"，'
                        f"标准要求使用 {'/'.join(expected_fonts)}{extra_font_note}",
            suggestion=f"请将{role_label}字体设置为 {'/'.join(expected_fonts)}。",
            paragraph_index=para_index,
            original_text=text,
            source="rule",
        ))
    if font_size > 0 and abs(font_size - expected_size) > 1:
        issues.append(Issue(
            category=IssueCategory.FONT,
            severity=IssueSeverity.WARNING,
            title=f"{role_label}字号不符合标准",
            description=f'{role_label}"{preview}"当前字号为 {font_size:.0f}pt，'
                        f"标准要求 {expected_size:.0f}pt",
            suggestion=f"请将{role_label}字号设置为 {expected_size:.0f}pt。",
            paragraph_index=para_index,
            original_text=text,
            source="rule",
        ))


def check_paragraph_format(para_info: dict, para_index: int,
                            role: str, strictness: str) -> list[Issue]:
    """
    根据段落角色检查格式。
    role 可能是：title / main_body / addressee / attachment / signature /
                  date / annotation / cc / issuer / doc_number / signer /
                  org_sign / serial_number / secrecy_level / urgency /
                  page_number / unknown
    """
    issues = []

    font_name = para_info.get("font_name", "")
    font_size = para_info.get("font_size", 0) or 0   # pt
    alignment = para_info.get("alignment", None)     # 0=left, 1=center, 2=right, 3=justify
    line_spacing = para_info.get("line_spacing", 0) or 0  # pt
    text = para_info.get("text", "").strip()

    if not text:
        return issues  # 空段落跳过

    # ---- 版头：机关标志 ----
    # 机关标志字号由机关自行规定（GB/T 9704 未固定），跳过检查避免误报
    if role == "org_sign":
        return issues

    # ---- 份号（6位阿拉伯数字，3号黑体）----
    if role == "serial_number":
        if not _SERIAL_RE.match(text):
            issues.append(Issue(
                category=IssueCategory.HEADER_FOOTER,
                severity=IssueSeverity.WARNING,
                title="份号格式不符合标准",
                description=f'份号"{text}"应为6位阿拉伯数字',
                suggestion="请将份号改为6位阿拉伯数字（如 000001）。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        _check_font_size(issues, para_info, para_index, "份号", text,
                         [FONT_HEITI], SIZE_SAN_HAO)
        return issues

    # ---- 密级和保密期限（3号黑体）----
    if role == "secrecy_level":
        _check_font_size(issues, para_info, para_index, "密级和保密期限", text,
                         [FONT_HEITI], SIZE_SAN_HAO)
        return issues

    # ---- 紧急程度（3号黑体）----
    if role == "urgency":
        _check_font_size(issues, para_info, para_index, "紧急程度", text,
                         [FONT_HEITI], SIZE_SAN_HAO)
        return issues

    # ---- 发文字号（3号仿宋）----
    if role == "doc_number":
        _check_font_size(issues, para_info, para_index, "发文字号", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO)
        return issues

    # ---- 签发人（3号仿宋；姓名部分应为楷体，段落级按仿宋检查）----
    if role == "signer":
        _check_font_size(issues, para_info, para_index, "签发人", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO,
                         extra_font_note="（签发人姓名应用楷体，请人工确认）")
        return issues

    # ---- 标题（小标宋 二号 居中）----
    if role == "title":
        _check_font_size(issues, para_info, para_index, "标题", text,
                         TITLE_FONTS, SIZE_ER_HAO)
        if alignment is not None and alignment != 1:  # 非居中
            issues.append(Issue(
                category=IssueCategory.PARAGRAPH_ALIGNMENT,
                severity=IssueSeverity.WARNING,
                title="标题应居中",
                description=f'标题"{text[:30]}"未居中对齐',
                suggestion="请将标题设置为居中对齐。",
                paragraph_index=para_index,
                original_text=text,
                source="rule",
            ))
        return issues

    # ---- 主送机关（3号仿宋）----
    if role == "addressee":
        _check_font_size(issues, para_info, para_index, "主送机关", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO)
        return issues

    # ---- 正文（3号仿宋 + 28磅固定行距）----
    if role == "main_body":
        _check_font_size(issues, para_info, para_index, "正文", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO)
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
        return issues

    # ---- 附件说明（3号仿宋）----
    if role == "attachment":
        _check_font_size(issues, para_info, para_index, "附件说明", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO)
        return issues

    # ---- 发文机关署名（3号仿宋）----
    if role == "signature":
        _check_font_size(issues, para_info, para_index, "发文机关署名", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO)
        return issues

    # ---- 成文日期（3号仿宋；右空4字需人工确认）----
    if role == "date":
        _check_font_size(issues, para_info, para_index, "成文日期", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO,
                         extra_font_note="（成文日期应右空4字，请人工确认）")
        return issues

    # ---- 附注（3号仿宋）----
    if role == "annotation":
        _check_font_size(issues, para_info, para_index, "附注", text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SAN_HAO)
        return issues

    # ---- 抄送机关 / 印发机关（4号仿宋）----
    if role in ("cc", "issuer"):
        label = "抄送机关" if role == "cc" else "印发机关和印发日期"
        _check_font_size(issues, para_info, para_index, label, text,
                         [FONT_FANGSONG, FONT_FANGSONG_ALT], SIZE_SI_HAO)
        return issues

    return issues


def check_first_line_indent(para_info: dict, para_index: int,
                             text: str, strictness: str) -> Optional[Issue]:
    """检查正文首行缩进（2字符，按段落字号换算磅值）"""
    if not text.strip():
        return None
    font_size = para_info.get("font_size", SIZE_SAN_HAO) or SIZE_SAN_HAO
    expected_indent = font_size * 2  # 2字符
    actual_indent = para_info.get("first_line_indent", 0) or 0
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
# 页码检查（读 footer XML）
# ---------------------------------------------------------------------------

def check_page_numbers(doc_path, strictness: str) -> list[Issue]:
    """检查页码：四号(14pt)宋体 + PAGE 域 + 一字线。

    通过解包 docx 读 word/footer*.xml，检查 PAGE 字段的字体字号，
    以及页脚文本中是否包含一字线（—）。
    """
    issues = []
    try:
        with zipfile.ZipFile(str(doc_path)) as z:
            footer_names = sorted(
                n for n in z.namelist()
                if re.match(r"word/footer\d*\.xml$", n)
            )
            if not footer_names:
                issues.append(Issue(
                    category=IssueCategory.PAGE_NUMBER,
                    severity=IssueSeverity.WARNING,
                    title="未检测到页码",
                    description="文档没有页脚，缺少页码。"
                                "标准要求页码为四号宋体阿拉伯数字，两侧各放一条一字线",
                    suggestion="请添加页码（四号宋体，一字线，单页右空一字、双页左空一字）。",
                    source="rule",
                ))
                return issues

            page_field_found = False
            for name in footer_names:
                root = etree.fromstring(z.read(name))

                # 简单字段：w:fldSimple instr="PAGE"
                for fld in root.iter(qn("w:fldSimple")):
                    if (fld.get(qn("w:instr")) or "").strip().upper() == "PAGE":
                        page_field_found = True
                        _check_page_field_font(fld.find(qn("w:r") + "/" + qn("w:rPr")), issues)

                # 复杂字段：fldChar begin + instrText PAGE + fldChar end
                for instr in root.iter(qn("w:instrText")):
                    if "PAGE" in (instr.text or "").upper():
                        page_field_found = True
                        _check_page_field_font(instr.getparent().find(qn("w:rPr")), issues)

                # 一字线检查（页脚文本中应含 — 或 －）
                footer_text = "".join(t.text or "" for t in root.iter(qn("w:t")))
                if page_field_found and ("—" not in footer_text and "－" not in footer_text):
                    issues.append(Issue(
                        category=IssueCategory.PAGE_NUMBER,
                        severity=IssueSeverity.WARNING,
                        title="页码缺少一字线",
                        description="页码两侧应各放一条一字线（如 — 1 —）",
                        suggestion="请在页码数字两侧添加一字线。",
                        source="rule",
                    ))

            if not page_field_found:
                issues.append(Issue(
                    category=IssueCategory.PAGE_NUMBER,
                    severity=IssueSeverity.WARNING,
                    title="未检测到页码域",
                    description="页脚存在但未检测到 PAGE 页码域，页码可能不是自动页码",
                    suggestion="请插入自动页码（插入→页码）。",
                    source="rule",
                ))
    except Exception as e:
        issues.append(Issue(
            category=IssueCategory.PAGE_NUMBER,
            severity=IssueSeverity.INFO,
            title="页码检查失败",
            description=f"无法读取页脚内容: {e}",
            suggestion="请人工确认页码格式。",
            source="rule",
        ))
    return issues


def _check_page_field_font(rpr, issues: list[Issue]) -> None:
    """检查 PAGE 字段 run 的字体字号：宋体 四号(14pt)。"""
    name, size = _rpr_font(rpr)
    if size is not None and abs(size - SIZE_SI_HAO) > 1:
        issues.append(Issue(
            category=IssueCategory.PAGE_NUMBER,
            severity=IssueSeverity.WARNING,
            title="页码字号不符合标准",
            description=f"页码当前字号为 {size:.0f}pt，标准要求四号 ({SIZE_SI_HAO}pt)",
            suggestion=f"请将页码字号设置为四号 ({SIZE_SI_HAO}pt)。",
            source="rule",
        ))
    if name and not _match_font(name, [FONT_SONG, "SimSun"]):
        issues.append(Issue(
            category=IssueCategory.PAGE_NUMBER,
            severity=IssueSeverity.WARNING,
            title="页码字体不符合标准",
            description=f'页码当前字体为 "{name}"，标准要求宋体',
            suggestion="请将页码字体设置为宋体。",
            source="rule",
        ))


# ---------------------------------------------------------------------------
# 段落角色识别
# ---------------------------------------------------------------------------

def _identify_paragraph_role(
    para_text: str, para_info: dict, idx: int, total: int
) -> str:
    """
    根据位置和内容识别段落属于公文的哪个部分。

    版头要素（文号/签发人/份号/密级/紧急程度/机关标志）先于标题判断，
    避免红头文件的版头被误判为标题产生假警报。
    """
    text = para_text.strip()
    if not text:
        return "empty"

    # ---- 版头要素 ----
    if _DOC_NUMBER_RE.search(text):
        return "doc_number"
    if text.startswith("签发人"):
        return "signer"
    if idx <= 3 and _SERIAL_RE.match(text):
        return "serial_number"          # 份号
    if idx <= 3 and text.startswith(("秘密", "机密", "绝密")):
        return "secrecy_level"          # 密级和保密期限
    if idx <= 3 and text.startswith(("特急", "加急", "平急", "急件")):
        return "urgency"                # 紧急程度

    # 主送机关：冒号结尾的顶格行（需先于机关标志判断——
    # 白头文件的机关名称行常以冒号结尾）
    if idx <= 5 and (text.endswith("：") or text.endswith(":")) and len(text) < 200:
        return "addressee"

    # 机关标志（红头大字）：第一段、含机关特征词、短行
    if idx == 0 and len(text) < 50 and any(hint in text for hint in _ORG_SIGN_HINTS):
        return "org_sign"

    # ---- 标题 ----
    # 典型结构："关于…的通知"（含"关于"+文种词）
    if "关于" in text and any(k in text for k in _DOC_TYPES):
        return "title"
    # 兜底：前 2 段、较短、以文种词结尾（标题结构为 发文机关+事由+文种）
    if idx <= 2 and len(text) < 80 and any(text.endswith(k) for k in _DOC_TYPES):
        return "title"

    # ---- 主体 ----
    if text.startswith("附件") and len(text) < 100:
        return "attachment"

    if _DATE_RE.match(text):
        return "date"

    if text.startswith("抄送") or "抄送：" in text or "抄送:" in text:
        return "cc"

    if text.startswith("印发") or text.startswith("（印发") or "印发机关" in text:
        return "issuer"
    if idx >= total - 3 and "印发" in text:
        return "issuer"

    # 附注
    if text.startswith("（此件") or text.startswith("(此件"):
        return "annotation"

    # 签名 / 成文日期（长文档最后几段中的短段落；
    # 文档过短时（total<=5）不适用，避免把正文误判为署名）
    if total > 5 and idx >= total - 5 and len(text) < 30:
        return "signature"

    # 默认正文
    return "main_body"
