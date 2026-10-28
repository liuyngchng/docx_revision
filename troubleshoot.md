# docx_revision 问题清单

> 生成日期：2026-10-28
> 范围：全仓库代码审查（backend/rules、backend/revision、backend/review、backend/api、backend/llm、run.py、cfg.yml.template、frontend）
> 总体结论：架构方向正确（规则引擎读 XML + LLM 内容审核 + COM 修订 + Web 上传下载），
> 但存在 6 个致命问题（其中 1 个有用户数据丢失风险、3 个导致全量误报/漏报、1 个会改错段落），上线前必须修复。

---

## 一、致命问题（P0）

### 1. WPS 的 COM ProgID 配错

- **位置**：`cfg.yml.template`（`com.prog_id`）、`backend/core/config.py:21`（默认值）
- **现象**：默认配置 `"WPS.Application"` 无法启动 COM，报 "Invalid class string"。
- **根因**：WPS 文字的 COM ProgID 是 `Kwps.Application`（WPS 表格是 `Ket.Application`），`WPS.Application` 不是注册的 ProgID。
- **修复**：默认改为 `"Kwps.Application"`，模板改为列表自动探测：`["Kwps.Application", "Word.Application"]`。

### 2. `Dispatch` 劫持用户正在使用的 WPS 实例，`Quit()` 会关掉用户的 WPS

- **位置**：`backend/revision/applier.py:118`（`Dispatch`）、`applier.py:137`（`Visible=False`）、`applier.py:192`（`Quit()`）
- **现象**：用户开着 WPS 写其他文档时提交审核：
  1. `Dispatch` 附着到用户正在运行的 WPS 实例；
  2. `Visible=False` 把用户所有 WPS 窗口隐藏；
  3. `close()` 里 `Quit()` 把整个 WPS 退出——**用户未保存的文档有丢失风险**。
- **修复**：
  - 改用 `win32com.client.DispatchEx(prog_id)` 尝试启动独立实例；
  - 启动前记录已有实例（`GetActiveObject` 探测），若最终附着的是已有实例则**不执行 `Quit()`**；
  - `Quit()` 只作用于本系统自己启动的实例。

### 3. 行距提取错误，所有正确设置"固定值28磅"的文档被误报

- **位置**：`backend/rules/engine.py:56-63`（提取）、`backend/rules/gb_t9704.py:291`（检查）
- **现象**：正文行距为"固定值 28 磅"（标准做法）时，规则引擎报"正文行距不符合标准"。
- **根因**：python-docx 的 `Length` 是 int 子类（EMU），`float(pf.line_spacing)` 对 `Pt(28)` 得到 355600（EMU 值）而非 28（磅），与 28 比较必然超差。
- **修复**：
  ```python
  ls = pf.line_spacing
  if ls is not None:
      if isinstance(ls, Length):        # 固定值/最小值 → 取磅值
          line_spacing = ls.pt
      else:                              # 倍数型 → 保留倍数
          line_spacing = float(ls)
  ```
  倍数型行距（如 1.5 倍）应单独报"行距应为固定值28磅"。

### 4. 字符单位首行缩进（`firstLineChars`）读不到，正确缩进的文档全部误报

- **位置**：`backend/rules/engine.py:65-68`（提取）、`backend/rules/gb_t9704.py:363-374`（检查）
- **现象**：文档用 Word/WPS 界面设置"首行缩进 2 字符"（XML 存为 `w:ind w:firstLineChars="200"`）时，引擎报"正文段落缺少首行缩进"。
- **根因**：python-docx 的 `first_line_indent` 不解析 `w:firstLineChars` 属性，返回 None → 检查走 `actual_indent == 0` 分支 → 全量误报。
- **修复**：提取时直接读 XML：
  ```python
  ind = para._p.pPr.find(qn('w:ind')) if para._p.pPr is not None else None
  if ind is not None:
      chars = ind.get(qn('w:firstLineChars'))
      if chars:  # 百倍字符数，200 = 2字符
          first_line_indent = int(chars) / 100.0 * font_size_pt
      else:
          first_line_indent = pf.first_line_indent.pt if pf.first_line_indent else 0
  ```

### 5. COM 任务无并发串行化

- **位置**：`backend/api/routes.py:120/193/234`（`background_tasks.add_task`，直接并发执行）
- **现象**：两个用户同时提交任务 → 两个后台线程同时 `Dispatch` 同一个 WPS 单实例 → 两份文档操作互相踩踏（打开、修订、保存交叉执行，结果不可预期）。
- **根因**：README 声称"当前实现按任务排队"，但代码里没有任何锁或队列。
- **修复**：全局 `threading.Lock`（或单线程任务队列）包住 `ComRevisionApplier.apply()` 全程；COM 任务同一时刻只能执行一个，其余排队等待。建议同时给前端状态加 "queued" 态。

### 6. 段落序号在规则引擎与 COM 端对不上，有表格的文档会改错段落

- **位置**：`backend/rules/engine.py:36`（`doc.paragraphs`，不含表格内段落）vs `backend/revision/applier.py:270`（`doc.Paragraphs(idx)`，含表格内段落）
- **现象**：文档中出现表格（公文附表常见）后，所有后续段落的 `paragraph_index` 在 COM 端错位，**修订会应用到错误的段落上**（字体、对齐、缩进全部改错地方）。
- **修复**（任选其一）：
  - COM 端改为与 python-docx 相同的枚举方式（跳过表格内段落），保证两端索引语义一致；
  - 或更稳妥：格式类修订全部改为"原文片段 + 出现次数"定位（`Find`），不再依赖段落序号；序号仅作辅助。

---

## 二、高（P1）

### 7. 幅面问题永远修不掉

- **位置**：`backend/rules/gb_t9704.py:143`（幅面检查用 `IssueCategory.PAGE_MARGIN`）vs `backend/revision/applier.py:360`（PAGE_MARGIN 只调 `set_page_margins`）
- **现象**：A4 尺寸不符合标准时，审核能报出来，但应用修订只改边距、不改纸张大小。
- **修复**：新增 `IssueCategory.PAGE_SIZE` 类别（或在 applier 的 PAGE_MARGIN 分支同时调用 `set_page_size`）。

### 8. 文本替换失败被计入成功

- **位置**：`backend/revision/applier.py:355-357`
- **现象**：`replace_text_one` 返回 False（原文片段不匹配：标点、空白、截断）时，`applied` 仍然 +1，前端显示"已应用"，实际什么都没改。
- **修复**：检查返回值，失败时计入 skipped 并在返回统计里带出失败条目。

### 9. 版头被当成标题，产生假警报

- **位置**：`backend/rules/gb_t9704.py:397-398`（`idx <= 2 and len(text) < 80 → title`）
- **现象**：红头文件的机关标志（"中国共产党XX市委员会"）、发文字号（"X政发〔2026〕1号"）等版头短行被识别为标题，被要求改成"二号小标宋居中"。
- **修复**：标题识别前先排除版头要素：
  - 发文字号正则：`〔\d{4}〕\s*\d+\s*号`
  - "签发人："开头的行
  - 第 1 段且为机关全称（大字、红色、无标点结尾）
  - 标题一般位于文号行之后（可结合"关于…的…"特征和位置）。

### 10. 样式继承的格式看不见，检查被静默跳过

- **位置**：`backend/rules/engine.py:40-48`
- **现象**：字体/字号来自样式（标题样式、正文样式）而非 run 直接设置时（Word 公文模板的常态），`first_run.font.size` 为 None → 字号检查被跳过，真实违规查不出来。
- **修复**：沿样式链解析有效格式：run 显式值 → 段落样式 → 基于样式（basedOn）→ Normal → document defaults。

### 11. 小标宋字体匹配过严

- **位置**：`backend/rules/gb_t9704.py:49`（`FONT_BIAOSONG = "方正小标宋简体"`）、`gb_t9704.py:80-89`（`_match_font`）
- **现象**：实际常见的"方正小标宋_GBK""方正小标宋GBK""华文中宋"均不匹配 → 误报"标题字体不符合标准"。
- **修复**：标题字体改为匹配"小标宋"子串（含方正小标宋各变体、华文中宋）。

### 12. `/api/health` 的 `wps_com_available` 误导

- **位置**：`backend/api/routes.py:293-300`
- **现象**：只检查 `import pywin32` 是否成功，未安装 WPS/Word 的机器也返回 `wps_com_available: true`。
- **修复**：实际尝试 `DispatchEx` 各 ProgID（不打开文档），成功才返回 true。

---

## 三、中（P2，审核覆盖面缺口）

### 13. 版头要素完全未审核（GB/T 9704 第 7.2 条）

份号、密级和保密期限、紧急程度、发文机关标志（红色、居中、上边缘距版心上边缘 35mm）、发文字号格式（机关代字+六角括号年份+序号，3号仿宋）、签发人（3号仿宋，姓名 3 号楷体）、红色分隔线——均未实现检查。

### 14. signature / date / annotation 角色识别了但不检查

- **位置**：`backend/rules/gb_t9704.py:180-336`（`check_paragraph_format` 无对应分支）
- 成文日期（3号仿宋、右空4字）、发文机关署名、附注的格式均不审核。

### 15. 页码不检查

页码应为 4 号半角宋体阿拉伯数字、两侧一字线、单页右空一字、双页左空一字。需要读页眉页脚 XML（`w:hdr`/`w:ftr` 关联文件），当前未实现。

### 16. LLM 审核的长文截断与替换定位脆弱

- **位置**：`backend/review/processor.py:87`（15000 字符截断）
- 长公文后半部分不审；`original_text` 必须与原文逐字一致才能被 Find 替换，提示词未强调这一点；LLM 输出无段落序号，全部依赖文本查找。

### 17. 清理与状态持久化缺失

- `cfg.yml` 的 `cleanup_days` 未使用，uploads/output 只增不减；
- `task_store` 为纯内存字典（`routes.py:50`），服务重启后任务状态全部丢失，前端轮询 404。

---

## 四、小（P3）

### 18. 杂项

| 位置 | 问题 |
|------|------|
| `applier.py:111` | `CoInitialize()` 每次 connect 都调、从不 `CoUninitialize`，未处理 `RPC_E_CHANGED_MODE` |
| `applier.py:296-303` | `set_first_line_indent` 用磅值（32pt），改字号后会不准；建议 `CharacterUnitFirstLineIndent = 2` |
| `requirements.txt` | `jinja2` 未使用 |
| `run.py:55-60` | `load_config` 调了两次（无害） |
| `cfg.yml.template` | 注释中"WPS 专业版"与"WPS 企业版"共列，但个人版无 COM 支持，需在文档中明确要求企业版/专业版 |

---

## 修复顺序建议

1. **第一批（上线阻断）**：#1 ProgID、#2 COM 实例管理（数据安全）、#5 并发串行化
2. **第二批（误报/漏报）**：#3 行距、#4 缩进、#6 段落定位一致性、#10 样式继承
3. **第三批（修订正确性）**：#7 幅面修复、#8 替换结果统计、#9 版头识别、#11 字体匹配
4. **第四批（覆盖面）**：#13-#17 版头/页码检查、状态持久化、清理任务
5. **第五批（收尾）**：#12 health、#18 杂项

## 验证环境备忘

- 目标环境：Windows 10/11 + WPS 企业版/专业版（或 MS Word）+ Python 3.10+
- 开发环境为 Ubuntu 时：规则引擎（python-docx）部分可本地验证；COM 部分必须到 Windows 上验证（`DispatchEx("Kwps.Application")`）
- 建议用一份"标准合规公文.docx"和一份"故意错乱的公文.docx"做回归测试，确认无误报、无漏报、无错改段落
