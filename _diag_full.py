"""诊断脚本：完整审核流程测试，带日志输出"""
import sys
import logging

sys.path.insert(0, '.')

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    datefmt='%H:%M:%S',
    stream=sys.stdout,
    force=True,
)

from backend.core.config import load_config
from backend.review.processor import ReviewProcessor
from backend.core.models import IssueSeverity

config = load_config()
doc_path = r'C:\Users\rd\Downloads\电子公文智能流转系统_详细设计说明书_V1.0_20261020.docx'

processor = ReviewProcessor(config)
result = processor.review(doc_path)

print()
print('=' * 70)
print('审核结果汇总')
print('=' * 70)
print(f'文件: {result.file_name}')
print(f'总问题数: {result.total_issues}')
print(f'  error:   {result.error_count}')
print(f'  warning: {result.warning_count}')
print(f'  info:    {result.info_count}')
if result.error_message:
    print(f'错误信息: {result.error_message}')
print()

# 按来源统计
rule_issues = [i for i in result.issues if i.source == 'rule']
llm_issues = [i for i in result.issues if i.source == 'llm']
print(f'规则引擎问题: {len(rule_issues)} 条')
print(f'LLM 问题:     {len(llm_issues)} 条')
print()

# 打印 LLM 问题详情
print('--- LLM 发现问题明细 ---')
for i, iss in enumerate(llm_issues):
    print(f'[{i+1}] [{iss.severity.value}] ({iss.category.value}) {iss.title}')
    print(f'     desc: {iss.description[:200]}')
    print(f'     suggestion: {iss.suggestion[:200]}')
    if iss.original_text:
        print(f'     original: {iss.original_text[:120]}')
    if iss.revised_text:
        print(f'     revised: {iss.revised_text[:120]}')
    print()

# 打印规则引擎问题（按类别统计）
print('--- 规则引擎问题按类别统计 ---')
from collections import Counter
cat_counter = Counter(i.category.value for i in rule_issues)
for cat, cnt in cat_counter.most_common():
    print(f'  {cat}: {cnt} 条')

print()
print('--- 规则引擎问题样例（前10条） ---')
for i, iss in enumerate(rule_issues[:10]):
    print(f'[{i+1}] [{iss.severity.value}] {iss.title}')
    print(f'     desc: {iss.description[:150]}')
    if iss.paragraph_index is not None:
        print(f'     para_index: {iss.paragraph_index}')
