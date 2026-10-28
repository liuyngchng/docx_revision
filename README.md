# 党政机关公文格式审核与修订系统

基于 **GB/T 9704—2012《党政机关公文格式》国家标准**，对 docx 文档进行自动审核，并通过 **WPS/Word COM 接口**以标准修订模式（Track Changes）应用修改。

## 核心特性

- **格式规则检查**：页边距、纸张幅面、字体、字号、行距、缩进、对齐等（纯 Python 读 XML，无需 Office）
- **LLM 内容审核**：文种选择、标题规范、语言表达、逻辑一致性等（OpenAI 兼容 API）
- **标准修订模式**：所有修改（含文本、字体、页边距、行距）均通过 WPS/Word COM 以 Track Changes 记录，Word/WPS 打开即可看到原生修订标记
- **Web 界面**：上传 → 审核 → 下载修订后文档，全程浏览器操作

## 环境要求

- **Windows**（WPS 企业版 / Microsoft Word 必须安装其一）
- **Python 3.10+**
- **pywin32**（已列入 requirements）

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 配置 cfg.yml（填入你们公司的 LLM API 地址、Key、模型名）
#    com.prog_id 根据安装的软件选择：WPS.Application 或 Word.Application

# 3. 启动
python run.py

# 4. 浏览器打开 http://127.0.0.1:8080
```

## 配置说明（cfg.yml）

```yaml
com:
  prog_id: ["Kwps.Application", "Word.Application"]  # WPS文字的ProgID是Kwps.Application；列表按顺序自动探测
  visible: false               # 是否显示 WPS/Word 窗口（仅对系统自建实例生效）
  display_alerts: false        # 是否弹出保存对话框

llm:
  api_base: "https://your-company-api.example.com/v1"  # OpenAI 兼容地址
  api_key: "sk-..."            # 你们公司的 API Key
  model: "your-model-name"     # 模型名
  timeout: 120
  temperature: 0.1
  max_tokens: 4096

review:
  enable_rules: true           # 启用格式规则检查
  enable_llm: true             # 启用 LLM 内容审核
  strictness: "normal"         # strict / normal / loose

server:
  host: "127.0.0.1"
  port: 8080
  max_upload_mb: 20
```

## 工作流程

```
上传 docx
   ↓
① 规则引擎（本地，读 XML）
   检查页边距/幅面/字体/字号/行距/缩进/对齐
   ↓
② LLM 审核（OpenAI 兼容 API）
   检查文种/标题/语言/逻辑/格式规范
   ↓
③ 汇总问题清单（按严重程度排序）
   ↓
④ WPS/Word COM 打开文档
   TrackRevisions = True
   逐条应用修改（文本替换 + 格式修正）
   ↓
⑤ 另存为新 docx（保留全部修订标记）
   ↓
⑥ 下载
```

## 项目结构

```
docx_revision/
├── cfg.yml                    # 配置文件（LLM API + COM 设置）
├── run.py                     # 启动入口
├── requirements.txt
├── backend/
│   ├── core/
│   │   ├── config.py          # 配置加载
│   │   └── models.py          # 数据模型（Issue、ReviewResult 等）
│   ├── rules/
│   │   ├── gb_t9704.py        # GB/T 9704 标准规则
│   │   └── engine.py          # 规则引擎
│   ├── llm/
│   │   ├── client.py          # OpenAI 兼容 API 客户端
│   │   └── prompts.py         # 审核提示词
│   ├── review/
│   │   └── processor.py       # 审核处理器（规则+LLM）
│   ├── revision/
│   │   └── applier.py         # COM 修订应用层（核心）
│   └── api/
│       └── routes.py          # FastAPI 路由
├── frontend/
│   ├── templates/index.html
│   └── static/ (css/js)
├── uploads/                   # 上传暂存
└── output/                    # 修订后输出
```

## 修订模式说明

系统通过 WPS/Word 的 COM 接口，在 `TrackRevisions = True` 状态下执行所有修改操作。WPS/Word 会自动为每处修改生成标准的 OOXML 修订标记：

| 修改类型 | 产生的修订标记 |
|---------|---------------|
| 文本替换 | `<w:ins>` 插入 / `<w:del>` 删除 |
| 字体/字号 | `<w:rPrChange>` 字符格式修订 |
| 行距/缩进 | `<w:pPrChange>` 段落格式修订 |
| 页边距/幅面 | `<w:sectPrChange>` 节格式修订 |

这些标记与手动在 Word/WPS 中开启修订模式后编辑文档产生的标记完全一致。

## 注意事项

1. **本机必须有 WPS 或 Word**：COM 路线依赖已安装的 Office 软件。`com.prog_id` 需正确配置。
2. **中文字体**：WPS 中文字体需通过 `Font.NameFarEast` 设置（代码已处理）。
3. **LLM 返回格式**：需要 OpenAI 兼容的 `/chat/completions` 端点，且支持 `response_format: {"type": "json_object"}`（用于结构化审核结果）。如果你们的 API 不支持这个参数，可在 `llm/client.py` 中关闭。
4. **并发限制**：COM 是单线程模型，同一时刻只能处理一个文档。系统通过全局锁串行化所有修订任务：任务进入 `queued` 状态排队，等待 WPS/Word 空闲后自动执行（最长等待 10 分钟）。
5. **实例安全**：若本机已有正在使用的 WPS/Word，系统会附着使用该实例且不会退出它；仅当实例由本系统启动时才在结束后关闭。
