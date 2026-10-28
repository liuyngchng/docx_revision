"""诊断脚本：查看 LLM 实际收到的提示词和返回内容"""
import sys
import json
from pathlib import Path

sys.path.insert(0, '.')

from docx import Document
from backend.core.config import load_config
from backend.llm.client import LlmClient
from backend.llm.prompts import SYSTEM_PROMPT, USER_PROMPT_TEMPLATE

config = load_config()
doc_path = r'C:\Users\rd\Downloads\电子公文智能流转系统_详细设计说明书_V1.0_20261020.docx'

# 提取文本
doc = Document(doc_path)
paragraphs_text = []
for i, para in enumerate(doc.paragraphs):
    if para.text.strip():
        paragraphs_text.append(f'[P{i}] {para.text}')

content = '\n'.join(paragraphs_text)
if len(content) > 15000:
    content = content[:15000] + '\n\n...(truncated)'

messages = [
    {'role': 'system', 'content': SYSTEM_PROMPT},
    {'role': 'user', 'content': USER_PROMPT_TEMPLATE.format(
        file_name=Path(doc_path).name,
        content=content,
    )},
]

print('=== 发送给 LLM 的消息 ===')
print(f'System prompt 长度: {len(SYSTEM_PROMPT)} 字符')
print(f'User prompt 长度: {len(messages[1]["content"])} 字符')
print()

client = LlmClient(config.llm)
print('正在调用 LLM ...')
try:
    response = client.chat_completion(messages)
    raw = client.extract_content(response)
    print('=== LLM 原始返回 ===')
    print(raw[:3000])
    if len(raw) > 3000:
        print(f'\n... (总长度 {len(raw)} 字符)')

    cleaned = raw.strip()
    if cleaned.startswith('```'):
        cleaned = cleaned.split('\n', 1)[1] if '\n' in cleaned else cleaned
        cleaned = cleaned.rsplit('```', 1)[0]
    parsed = json.loads(cleaned)
    print()
    print('=== 解析结果 ===')
    print(f'overall_assessment: {parsed.get("overall_assessment", "N/A")[:200]}')
    issues = parsed.get('issues', [])
    print(f'issues 数量: {len(issues)}')
    for i, iss in enumerate(issues):
        print(f'  [{i+1}] [{iss.get("severity")}] {iss.get("title")}')
        print(f'       desc: {iss.get("description", "")[:150]}')
except Exception as e:
    print(f'LLM 调用失败: {e}')
    import traceback
    traceback.print_exc()
