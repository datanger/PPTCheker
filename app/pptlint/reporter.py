"""
报告生成（对应任务：实现CLI与报告生成（Markdown/HTML））
支持规则检查和LLM审查的分类显示
"""
from typing import List
from jinja2 import Template

from .model import Issue


MD_TEMPLATE = Template(
    """
## 审查报告

### 📊 问题统计
- **规则检查问题**: {{ rule_issues|length }} 个
- **LLM智能审查问题**: {{ llm_issues|length }} 个
- **总计**: {{ issues|length }} 个

### 📄 按页码分组的问题详情
{% set page_numbers = [] %}
{% for issue in issues %}
    {% set _ = page_numbers.append(issue.slide_index + 1) %}
{% endfor %}
{% set unique_pages = page_numbers|unique|sort %}

{% if unique_pages %}
{% for page_num in unique_pages %}
#### 📍 第 {{ page_num }} 页

{% set page_rule_issues = rule_issues|selectattr('slide_index', 'equalto', page_num - 1)|list %}
{% set page_llm_issues = llm_issues|selectattr('slide_index', 'equalto', page_num - 1)|list %}

{% if page_rule_issues or page_llm_issues %}
**🔍 规则检查问题:**
{% if page_rule_issues %}
{% for it in page_rule_issues %}
- **{{ rule_labels.get(it.rule_id, it.rule_id) }}** | 严重性: {{ it.severity }} | 对象: {{ it.object_ref }}
  - 描述: {{ it.message }}
  - 建议: {{ it.suggestion or '-' }}
  - 可自动修复: {{ '是' if it.can_autofix else '否' }} | 已修复: {{ '是' if it.fixed else '否' }}
{% endfor %}
{% else %}
✅ 该页无规则检查问题
{% endif %}

**🤖 LLM智能审查问题:**
{% if page_llm_issues %}
{% for it in page_llm_issues %}
- **{{ rule_labels.get(it.rule_id, it.rule_id) }}** | 严重性: {{ it.severity }} | 对象: {{ it.object_ref }}
  - 描述: {{ it.message }}
  - 建议: {{ it.suggestion or '-' }}
  - 可自动修复: {{ '是' if it.can_autofix else '否' }} | 已修复: {{ '是' if it.fixed else '否' }}
{% endfor %}
{% else %}
✅ 该页无LLM审查问题
{% endif %}

**📊 第 {{ page_num }} 页问题统计:** 共 {{ page_rule_issues|length + page_llm_issues|length }} 个问题
{% else %}
✅ 该页未发现问题
{% endif %}

---
{% endfor %}
{% else %}
✅ 未发现任何问题
{% endif %}

### 📋 问题分类统计
**规则检查分类:**
{% for rule_id in rule_issues|map(attribute='rule_id')|unique|list %}
- {{ rule_labels.get(rule_id, rule_id) }}: {{ rule_issues|selectattr('rule_id', 'equalto', rule_id)|list|length }} 个
{% endfor %}

**LLM审查分类:**
{% for rule_id in llm_issues|map(attribute='rule_id')|unique|list %}
- {{ rule_labels.get(rule_id, rule_id) }}: {{ llm_issues|selectattr('rule_id', 'equalto', rule_id)|list|length }} 个
{% endfor %}
"""
)


def render_markdown(issues: List[Issue]) -> str:
    """生成Markdown格式的审查报告"""
    # 去重：同一页中同一个问题只出现一次
    deduplicated_issues = _deduplicate_issues_by_page(issues)
    
    # 区分规则检查和LLM审查的问题
    rule_issues = [it for it in deduplicated_issues if not it.rule_id.startswith("LLM_")]
    llm_issues = [it for it in deduplicated_issues if it.rule_id.startswith("LLM_")]
    # 规则ID到中文名称映射（与 annotator 中一致）
    rule_labels = {
        # 规则检查
        "FontFamilyRule": "字体不规范",
        "FontSizeRule": "字号过小",
        "ColorCountRule": "颜色过多",
        "ThemeHarmonyRule": "色调不一致",
        # LLM智能审查
        "LLM_AcronymRule": "专业缩略语需解释",
        "LLM_ContentRule": "内容逻辑问题",
        "LLM_FormatRule": "智能格式问题",
        "LLM_FluencyRule": "表达流畅性问题",
        "LLM_TitleStructureRule": "标题结构问题",
        "LLM_ThemeHarmonyRule": "主题一致性问题",
    }

    return MD_TEMPLATE.render(
        issues=deduplicated_issues,
        rule_issues=rule_issues,
        llm_issues=llm_issues,
        rule_labels=rule_labels
    )


def _deduplicate_issues_by_page(issues: List[Issue]) -> List[Issue]:
    """去重：同一页中同一个问题只出现一次"""
    # 使用字典来跟踪每个页面中已经出现的问题
    page_issues = {}  # {page_index: {issue_key: issue}}
    
    for issue in issues:
        page_index = issue.slide_index
        # 创建问题的唯一标识：rule_id + message的前50个字符
        issue_key = f"{issue.rule_id}:{issue.message[:50]}"
        
        if page_index not in page_issues:
            page_issues[page_index] = {}
        
        # 如果该页面还没有这个问题，则添加
        if issue_key not in page_issues[page_index]:
            page_issues[page_index][issue_key] = issue
    
    # 将所有去重后的问题收集到一个列表中
    deduplicated_issues = []
    for page_issues_dict in page_issues.values():
        deduplicated_issues.extend(page_issues_dict.values())
    
    return deduplicated_issues

