"""
报告生成（对应任务：实现CLI与报告生成（Markdown/HTML））
"""
from typing import List
from jinja2 import Template

from .model import Issue


MD_TEMPLATE = Template(
    """
## 审查报告

共发现 {{ issues|length }} 项问题。

{% for it in issues %}
- 规则: {{ it.rule_id }} | 严重性: {{ it.severity }} | 页: {{ it.slide_index }} | 对象: {{ it.object_ref }}
  - 描述: {{ it.message }}
  - 建议: {{ it.suggestion or '-' }}
  - 可自动修复: {{ '是' if it.can_autofix else '否' }} | 已修复: {{ '是' if it.fixed else '否' }}
{% endfor %}
"""
)


def render_markdown(issues: List[Issue]) -> str:
    return MD_TEMPLATE.render(issues=issues)

