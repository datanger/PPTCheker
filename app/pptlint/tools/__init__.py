"""
PPT审查工具包

包含以下模块：
- workflow_tools: 工作流工具函数
- structure_parsing: PPT结构分析
- llm_review: LLM智能审查
- rules: 基础规则检查
"""

from .workflow_tools import (
    load_parsing_result,
    convert_parsing_result_to_document_model,
    run_basic_rules,
    run_llm_review,
    generate_report,
    generate_annotated_ppt,
    get_workflow_statistics,
    # 新增：PPT编辑功能
    create_ppt_context,
    run_llm_edit_analysis,
    apply_edits_to_ppt,
    save_modified_ppt
)

from .structure_parsing import (
    load_parsing_result as load_structure_parsing_result,
    infer_all_structures
)

from .llm_review import (
    LLMReviewer,
    create_llm_reviewer
)

from .rules import (
    run_basic_rules as run_rules,
    check_font_and_size,
    check_color_count,
    check_theme_harmony
)

__all__ = [
    # workflow_tools
    'load_parsing_result',
    'convert_parsing_result_to_document_model',
    'run_basic_rules',
    'run_llm_review',
    'generate_report',
    'generate_annotated_ppt',
    'get_workflow_statistics',
    
    # 新增：PPT编辑功能
    'create_ppt_context',
    'run_llm_edit_analysis',
    'apply_edits_to_ppt',
    'save_modified_ppt',
    
    # structure_parsing
    'load_structure_parsing_result',
    'infer_all_structures',
    
    # llm_review
    'LLMReviewer',
    'create_llm_reviewer',
    
    # rules
    'run_rules',
    'check_font_and_size',
    'check_color_count',
    'check_theme_harmony'
]
