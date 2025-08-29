"""
工作流编排器：根据《审查需求文档》与配置，动态串联组件。
支持规则+LLM混合审查模式，LLM不可用时自动降级为纯规则。

输入：parsing_result.json 格式的数据
输出：审查结果、报告、标记PPT等
"""
from typing import List, Optional

from .config import ToolConfig
from .model import Issue
from .llm import LLMClient
from .tools.workflow_tools import (
    load_parsing_result,
    generate_report,
    generate_annotated_ppt,
    get_workflow_statistics,
    # 新增：PPT编辑功能
    create_ppt_context,
    run_llm_edit_analysis,
    apply_edits_to_ppt,
    save_modified_ppt,
    # 新增：规则检查
    convert_parsing_result_to_document_model,
    run_basic_rules
)
from .tools.llm_review import create_llm_reviewer
from .tools.structure_parsing import analyze_from_parsing_result

class WorkflowResult:
    def __init__(self):
        self.issues: List[Issue] = []
        self.report_md: Optional[str] = None
        self.rule_issues_count: int = 0
        self.llm_issues_count: int = 0


def run_review_workflow(parsing_result_path: str, cfg: ToolConfig, output_ppt: Optional[str], llm: Optional[LLMClient], original_pptx_path: Optional[str] = None) -> WorkflowResult:
    res = WorkflowResult()
    
    # 步骤1：加载 parsing_result.json
    print("📖 加载解析结果...")
    parsing_data = load_parsing_result(parsing_result_path)
    if not parsing_data or parsing_data.get("页数", 0) == 0:
        print("❌ 加载解析结果失败或数据为空")
        return res
    
    # 步骤2：分析PPT结构
    print("🔍 分析PPT结构...")
    try:
        parsing_data = analyze_from_parsing_result(parsing_data)
        print("✅ PPT结构分析完成")
        
        # 将结构分析结果重新写入parsing_result.json
        import json
        with open(parsing_result_path, "w", encoding="utf-8") as f:
            json.dump(parsing_data, f, ensure_ascii=False, indent=2)
        print(f"💾 结构分析结果已更新到: {parsing_result_path}")
        
    except Exception as e:
        print(f"⚠️ PPT结构分析失败：{e}")
        # 即使结构分析失败，也继续后续流程
    
    # 步骤3：规则检查
    print("📋 运行规则检查...")
    rule_issues = []
    try:
        doc_model = convert_parsing_result_to_document_model(parsing_data, parsing_result_path)
        rule_issues = run_basic_rules(doc_model, cfg)
        print(f"✅ 规则检查完成，发现 {len(rule_issues)} 个问题")
    except Exception as e:
        print(f"⚠️ 规则检查失败：{e}")
    
    # 步骤4：LLM审查（抽取为公共函数）
    print("🤖 运行LLM审查...")
    llm_issues = _perform_llm_review(parsing_data, cfg, llm)
    
    # 合并所有问题
    all_issues = rule_issues + llm_issues
    res.issues = all_issues
    res.rule_issues_count = len(rule_issues)
    res.llm_issues_count = len(llm_issues)
    
    # 步骤5：生成报告
    print("📊 生成审查报告...")
    res.report_md = generate_report(all_issues, rule_issues, llm_issues)
    
    # 步骤6：输出标记PPT（如果指定）
    if output_ppt:
        if not original_pptx_path:
            print("⚠️ 无法生成标记PPT：需要提供原始PPTX文件路径")
        else:
            print("🏷️ 生成标记PPT...")
            success = generate_annotated_ppt(original_pptx_path, all_issues, output_ppt)
            if success:
                print(f"✅ 标记PPT已生成: {output_ppt}")
            else:
                print("❌ 生成标记PPT失败")
    
    # 步骤7：输出统计信息
    print(f"\n🎯 审查完成！")
    print(f"   - 规则检查：{res.rule_issues_count} 个问题")
    print(f"   - LLM审查：{res.llm_issues_count} 个问题")
    print(f"   - 总计：{len(all_issues)} 个问题")
    
    return res


def _perform_llm_review(parsing_data, cfg: ToolConfig, llm: Optional[LLMClient]) -> List[Issue]:
    """公共：基于 parsing_result.json 调用LLM进行多维度审查并返回问题列表。"""
    issues: List[Issue] = []
    
    # 检查配置是否启用LLM
    if not cfg.llm_enabled:
        print("🤖 LLM审查已禁用，跳过LLM审查步骤")
        return issues
    
    # 如果LLM客户端未提供，自动创建一个
    if not llm:
        print("🤖 自动创建LLM客户端...")
        from .llm import LLMClient
        llm = LLMClient()
    
    try:
        print("🤖 创建LLM审查器...")
        reviewer = create_llm_reviewer(llm, cfg)
        
        issues = []
        
        # 根据配置开关决定是否执行各项审查
        if cfg.review_format:
            print("🤖 开始格式标准审查...")
            fmt = reviewer.review_format_standards(parsing_data)
            if fmt:
                issues.extend(fmt)
        else:
            print("🤖 格式标准审查已禁用，跳过...")
        
        if cfg.review_logic:
            print("🤖 开始内容逻辑审查...")
            logic = reviewer.review_content_logic(parsing_data)
            if logic:
                issues.extend(logic)
        else:
            print("🤖 内容逻辑审查已禁用，跳过...")
        
        if cfg.review_acronyms:
            print("🤖 开始缩略语审查...")
            acr = reviewer.review_acronyms(parsing_data)
            if acr:
                issues.extend(acr)
        else:
            print("🤖 缩略语审查已禁用，跳过...")
        
        if cfg.review_fluency:
            print("🤖 开始标题结构审查...")
            title = reviewer.review_title_structure(parsing_data)
            if title:
                issues.extend(title)
        else:
            print("🤖 标题结构审查已禁用，跳过...")
        
    except Exception as e:
        print(f"⚠️ LLM审查失败：{e}")
    return issues

def run_edit_workflow(
    parsing_result_path: str, 
    original_pptx_path: str, 
    cfg: ToolConfig, 
    output_ppt: str, 
    llm: Optional[LLMClient] = None,
    edit_requirements: str = "请分析PPT内容，提供改进建议"
) -> WorkflowResult:
    """编辑模式：使用LLM分析并自动修改PPT"""
    res = WorkflowResult()
    
    print("✏️ 启动PPT编辑模式...")
    
    # 步骤1：加载 parsing_result.json
    print("📖 加载解析结果...")
    parsing_data = load_parsing_result(parsing_result_path)
    if not parsing_data or parsing_data.get("页数", 0) == 0:
        print("❌ 加载解析结果失败或数据为空")
        return res
    
    # 步骤2：创建PPT编辑上下文
    print("🔄 创建PPT编辑上下文...")
    ppt_context = create_ppt_context(parsing_data, original_pptx_path)
    if not ppt_context:
        print("❌ 创建PPT编辑上下文失败")
        return res
    
    # 步骤3：依赖审查结果（与审查模式共用的LLM审查逻辑）
    print("🤖 运行审查以支持编辑...")
    review_issues = _perform_llm_review(parsing_data, cfg, llm)
    res.issues = review_issues
    
    # 步骤4：使用LLM分析并生成编辑建议
    print("🤖 使用LLM分析PPT内容...")
    # 如果LLM客户端未提供，自动创建一个
    if not llm:
        print("🤖 自动创建LLM客户端...")
        from .llm import LLMClient
        llm = LLMClient()
    edit_suggestions = run_llm_edit_analysis(parsing_data, llm, edit_requirements)
    
    if edit_suggestions:
        print(f"✅ LLM生成 {len(edit_suggestions)} 个编辑建议")
        
        # 步骤5：应用编辑建议到PPT
        print("🔧 应用编辑建议...")
        edit_result = apply_edits_to_ppt(ppt_context, edit_suggestions)
        
        if edit_result.success:
            # 步骤6：保存修改后的PPT
            print("💾 保存修改后的PPT...")
            if save_modified_ppt(ppt_context, output_ppt):
                print(f"✅ 编辑完成！修改后的PPT已保存到: {output_ppt}")
                
                # 生成编辑报告
                res.report_md = generate_edit_report(edit_result, edit_suggestions)
                res.rule_issues_count = len(edit_result.failed_suggestions)
                res.llm_issues_count = len(edit_result.applied_suggestions)
                
                # 输出统计信息
                print(f"\n🎯 编辑完成！")
                print(f"   - 成功应用：{len(edit_result.applied_suggestions)} 个建议")
                print(f"   - 失败建议：{len(edit_result.failed_suggestions)} 个")
                print(f"   - 修改页面：{edit_result.modified_slides}")
                
                if edit_result.error_messages:
                    print(f"   - 错误信息：{edit_result.error_messages}")
            else:
                print("❌ 保存修改后的PPT失败")
        else:
            print("❌ 应用编辑建议失败")
    else:
        print("ℹ️ LLM未生成编辑建议")
    
    return res


def generate_edit_report(edit_result, edit_suggestions: List) -> str:
    """生成编辑报告"""
    report = "# PPT编辑报告\n\n"
    
    if edit_result.success:
        report += f"## ✅ 编辑成功\n\n"
        report += f"- 成功应用：{len(edit_result.applied_suggestions)} 个建议\n"
        report += f"- 修改页面：{edit_result.modified_slides}\n"
        report += f"- 输出文件：{edit_result.output_path}\n\n"
        
        if edit_result.applied_suggestions:
            report += "## 📝 已应用的编辑\n\n"
            for suggestion in edit_result.applied_suggestions:
                report += f"### 页面 {suggestion.page_number} - 形状 {suggestion.shape_index}\n"
                report += f"- 类型：{suggestion.type}\n"
                report += f"- 当前值：{suggestion.current_value}\n"
                report += f"- 新值：{suggestion.new_value}\n"
                report += f"- 原因：{suggestion.reason}\n"
                report += f"- 优先级：{suggestion.priority}\n\n"
    else:
        report += "## ❌ 编辑失败\n\n"
    
    if edit_result.failed_suggestions:
        report += "## ⚠️ 失败的编辑\n\n"
        for suggestion in edit_result.failed_suggestions:
            report += f"### 页面 {suggestion.page_number} - 形状 {suggestion.shape_index}\n"
            report += f"- 类型：{suggestion.type}\n"
            report += f"- 原因：{suggestion.reason}\n\n"
    
    if edit_result.error_messages:
        report += "## 🚨 错误信息\n\n"
        for error in edit_result.error_messages:
            report += f"- {error}\n\n"
    
    return report


