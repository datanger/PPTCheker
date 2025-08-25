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


def annotate_pptx(src_path: str, issues: List[Issue], output_path: str) -> None:
    prs = Presentation(src_path)

    # 按页聚合问题
    issues_by_slide = defaultdict(list)
    for it in issues:
        issues_by_slide[it.slide_index].append(it)

    # 全局问题汇总：仅在首页生成（中文类别、合并计数，过滤 info）
    from collections import Counter
    rule_to_label = {
        "FontFamilyRule": "字体不规范",
        "FontSizeRule": "字号过小",
        "AcronymRule": "缩略语未解释",
        "ColorCountRule": "颜色过多",
        "ThemeHarmonyRule": "色调不一致",
    }
    filtered_all = [it for it in issues if (it.severity or "").lower() != "info"]
    grouped_all = Counter((rule_to_label.get(it.rule_id, "其他规范问题"), it.severity) for it in filtered_all)
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
        # 注意：只能对具有文本框架的形状进行内联追加，按 shape_id 近似匹配
        page = prs.slides[s_idx]
        for shp in page.shapes:
            if not hasattr(shp, "has_text_frame") or not shp.has_text_frame:
                continue
            sid = str(getattr(shp, "shape_id", ""))
            hit_rules = [it.rule_id for it in page_issues if it.object_ref == sid]
            if not hit_rules:
                continue
            # 规则到中文类别的映射（固定类别）
            rule_to_label = {
                "FontFamilyRule": "字体不规范",
                "FontSizeRule": "字号过小",
                "AcronymRule": "缩略语未解释",
                "ColorCountRule": "颜色过多",
                "ThemeHarmonyRule": "色调不一致",
            }
            # 允许多个不同类别；同类多次命中以 xN 展示
            from collections import Counter
            label_counts = Counter(rule_to_label.get(rid, "其他规范问题") for rid in hit_rules)
            labels = [f"{lab}x{cnt}" if cnt > 1 else lab for lab, cnt in label_counts.items()]
            try:
                if shp.text_frame is None:
                    continue
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
            except Exception:
                # 不阻断流程
                pass

    prs.save(output_path)

