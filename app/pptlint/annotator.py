"""
带标记PPT输出（对应任务：实现PPT注释输出模块并集成CLI）

实现要点：
- 在每页左上角新增“问题汇总”文本框；
- 对命中的 shape，将其文本末尾追加“【标记: 规则ID】”；
- 不覆盖原文件，另存为副本。
"""
from collections import defaultdict
from typing import List
from pptx import Presentation
from pptx.util import Pt, Inches

from .model import Issue


def _contains_acronym(text: str) -> bool:
    """检查文本是否包含需要解释的缩略语"""
    # 常见的需要解释的专业缩略语
    professional_acronyms = [
        'LLM', 'MCP', 'UFO', 'GPT', 'BERT', 'CNN', 'RNN', 'LSTM', 'GRU',
        'SVM', 'KNN', 'PCA', 'SVD', 'NLP', 'CV', 'ASR', 'TTS', 'OCR',
        'GAN', 'VAE', 'DQN', 'PPO', 'A3C', 'TRPO', 'SAC', 'TD3'
    ]
    
    # 检查是否包含专业缩略语
    for acronym in professional_acronyms:
        if acronym in text:
            return True
    
    return False


def annotate_pptx(src_path: str, issues: List[Issue], output_path: str) -> None:
    prs = Presentation(src_path)

    # 按页聚合问题
    issues_by_slide = defaultdict(list)
    for it in issues:
        issues_by_slide[it.slide_index].append(it)

    # 全局问题汇总：包含所有问题类型，不过滤info级别
    from collections import Counter
    rule_to_label = {
        # 规则检查问题
        "FontFamilyRule": "字体不规范",
        "FontSizeRule": "字号过小",
        "ColorCountRule": "颜色过多",
        "ThemeHarmonyRule": "色调不一致",
        # LLM智能审查问题
        "LLM_AcronymRule": "专业缩略语需解释",
        "LLM_ContentRule": "内容逻辑问题",
        "LLM_FormatRule": "智能格式问题",
        "LLM_FluencyRule": "表达流畅性问题",
        "LLM_TitleStructureRule": "标题结构问题",
    }
    
    # 统计所有问题类型
    grouped_all = Counter((rule_to_label.get(it.rule_id, "其他问题"), it.severity) for it in issues)
    global_summary_lines = [
        f"- {label} [{sev}] x{cnt}"
        for (label, sev), cnt in grouped_all.items()
    ]

    for s_idx, slide in enumerate(prs.slides):
        page_issues = issues_by_slide.get(s_idx, [])

        # 仅在首页绘制全局汇总
        if s_idx == 0 and global_summary_lines:
            left, top, width, height = Inches(0.3), Inches(0.2), Inches(6.5), Inches(1.8)
            tf_box = slide.shapes.add_textbox(left, top, width, height)
            tf = tf_box.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = "问题汇总:\n" + "\n".join(global_summary_lines)
            if run.font is not None:
                run.font.size = Pt(12)

        # 对对象内联标记
        page = prs.slides[s_idx]
        
        # 调试信息：显示该页面的所有问题
        if page_issues:
            print(f"\n页面 {s_idx} 的问题:")
            for issue in page_issues:
                print(f"  - {issue.rule_id}: {issue.object_ref} - {issue.message}")
        
        # 先处理LLM问题：收集该页面的所有LLM问题
        llm_issues = [issue for issue in page_issues if issue.rule_id.startswith("LLM_")]
        if llm_issues:
            print(f"    页面 {s_idx} 发现 {len(llm_issues)} 个LLM问题:")
            for issue in llm_issues:
                print(f"      {issue.rule_id}: {issue.object_ref} - {issue.message}")
        
        for shp in page.shapes:
            # 更安全的属性检查
            if not hasattr(shp, "text_frame") or shp.text_frame is None:
                continue
                
            # 改进对象引用匹配：支持多种引用方式
            sid = str(getattr(shp, "shape_id", ""))
            hit_rules = []
            
            # 不再自动标记所有文本对象，而是根据具体问题类型进行精确匹配
            for issue in page_issues:
                # 匹配方式1：直接shape_id匹配
                if issue.object_ref == sid:
                    hit_rules.append(issue.rule_id)
                # 匹配方式2：text_block_X_Y格式匹配（LLM返回的精确格式）
                elif issue.object_ref.startswith("text_block_") and "_" in issue.object_ref:
                    # text_block_2_1 格式：分割后是 ["text", "block", "2", "1"]
                    # 所以页码是 parts[2]，块索引是 parts[3]
                    parts = issue.object_ref.split("_")
                    if len(parts) >= 4 and parts[2] == str(s_idx):
                        print(f"    🔍 检查text_block匹配: {issue.object_ref} -> 页面 {s_idx}")
                        # 对于text_block格式，我们检查文本内容是否包含相关缩略语
                        if issue.rule_id == "LLM_AcronymRule":
                            # 检查文本内容是否包含缩略语
                            text_content = ""
                            try:
                                for para in shp.text_frame.paragraphs:
                                    for run in para.runs:
                                        text_content += run.text + " "
                            except:
                                text_content = ""
                            
                            print(f"    📝 形状 {sid} 文本内容: {text_content[:50]}...")
                            
                            # 检查是否包含需要解释的缩略语
                            if _contains_acronym(text_content):
                                hit_rules.append(issue.rule_id)
                                print(f"    ✅ 智能匹配: 形状 {sid} 包含缩略语，标记为 {issue.rule_id}")
                            else:
                                print(f"    ❌ 形状 {sid} 不包含缩略语，跳过标记")
                        else:
                            # 对于其他LLM规则，暂时跳过
                            print(f"    ⚠️ 跳过非缩略语LLM规则: {issue.rule_id}")
                            pass
                # 匹配方式3：title_X格式匹配（页面标题）
                elif issue.object_ref.startswith("title_") and issue.object_ref.endswith(f"_{s_idx}"):
                    # 对于标题问题，我们标记该页面的标题对象
                    if shp.is_title and shp.title_level:
                        hit_rules.append(issue.rule_id)
                        print(f"    标题匹配: 形状 {sid} 是标题，标记为 {issue.rule_id}")
                    elif shp == page.shapes[0]:  # 备用方案：假设第一个形状是标题
                        hit_rules.append(issue.rule_id)
                        print(f"    标题备用匹配: 形状 {sid} 是第一个形状，标记为 {issue.rule_id}")
                # 匹配方式4：page_X格式匹配（页面级别问题）
                elif issue.object_ref.startswith("page_") and issue.object_ref.endswith(f"_{s_idx}"):
                    # 对于页面级别问题，我们标记该页面的所有文本对象
                    hit_rules.append(issue.rule_id)
                # 匹配方式5：page级别的问题（向后兼容）
                elif issue.object_ref == "page":
                    # 对于page级别问题，我们标记该页面的所有文本对象
                    hit_rules.append(issue.rule_id)
            
            if not hit_rules:
                continue
                
            # 规则到中文类别的映射
            rule_to_label = {
                # 规则检查问题
                "FontFamilyRule": "字体不规范",
                "FontSizeRule": "字号过小",
                "ColorCountRule": "颜色过多",
                "ThemeHarmonyRule": "色调不一致",
                # LLM智能审查问题
                "LLM_AcronymRule": "专业缩略语需解释",
                "LLM_ContentRule": "内容逻辑问题",
                "LLM_FormatRule": "智能格式问题",
                "LLM_FluencyRule": "表达流畅性问题",
                "LLM_TitleStructureRule": "标题结构问题",
            }
            
            # 允许多个不同类别；同类多次命中以 xN 展示
            from collections import Counter
            label_counts = Counter(rule_to_label.get(rid, "其他问题") for rid in hit_rules)
            labels = [f"{lab}x{cnt}" if cnt > 1 else lab for lab, cnt in label_counts.items()]
            
            # 调试信息：显示匹配到的规则
            if hit_rules:
                print(f"页面 {s_idx} 形状 {sid} 匹配到规则: {hit_rules}")
                if any(rid.startswith("LLM_") for rid in hit_rules):
                    print(f"    -> 包含LLM规则，将应用样式和标记")
            
            try:
                # 对现有 runs 施加样式：红色 + 斜体 + 下划线
                for para in shp.text_frame.paragraphs:
                    for r in para.runs:
                        if r.font is not None:
                            r.font.italic = True
                            # 优先设置为波浪线，不支持则退化为普通下划线
                            try:
                                from pptx.enum.text import MSO_TEXT_UNDERLINE
                                r.font.underline = MSO_TEXT_UNDERLINE.WAVY_LINE
                            except Exception:
                                r.font.underline = True
                            # 设为红色
                            try:
                                from pptx.dml.color import RGBColor
                                r.font.color.rgb = RGBColor(255, 0, 0)
                            except Exception:
                                pass
                # 同时在最后追加规则摘要（去重后的中文类别），便于溯源
                para_tail = shp.text_frame.paragraphs[-1]
                tail = para_tail.add_run()
                if labels:
                    tail.text = " 【标记: " + "、".join(labels) + "】"
                else:
                    tail.text = " 【标记: 规范问题】"
                if tail.font is not None:
                    tail.font.size = Pt(10)
            except Exception as e:
                # 不阻断流程，记录错误
                print(f"标记形状 {sid} 时出错: {e}")
                pass

    prs.save(output_path)

