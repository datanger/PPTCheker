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
- **{{ it.rule_id }}** | 严重性: {{ it.severity }} | 对象: {{ it.object_ref }}
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
- **{{ it.rule_id }}** | 严重性: {{ it.severity }} | 对象: {{ it.object_ref }}
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
- {{ rule_id }}: {{ rule_issues|selectattr('rule_id', 'equalto', rule_id)|list|length }} 个
{% endfor %}

**LLM审查分类:**
{% for rule_id in llm_issues|map(attribute='rule_id')|unique|list %}
- {{ rule_id }}: {{ llm_issues|selectattr('rule_id', 'equalto', rule_id)|list|length }} 个
{% endfor %}
"""
)


def render_markdown(issues: List[Issue]) -> str:
    """生成Markdown格式的审查报告"""
    # 区分规则检查和LLM审查的问题
    rule_issues = [it for it in issues if not it.rule_id.startswith("LLM_")]
    llm_issues = [it for it in issues if it.rule_id.startswith("LLM_")]
    
    return MD_TEMPLATE.render(
        issues=issues,
        rule_issues=rule_issues,
        llm_issues=llm_issues
    )

