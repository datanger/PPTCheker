"""
基于 parser.py 解析结果的高层信息抽取工具

功能：
- 从 parsing_result_new.json（或传入的结构化数据）中提取：
  1) 每页标题（优先使用“是否是标题占位符”为 True 的文本块）
  2) 章节名与目录页识别（结合大模型）
  3) 全局主题/主题词（结合大模型）

说明：
- 所有函数仅依赖 parser 的输出结构，不修改 parser 行为
- 需要调用大模型时，复用现有 llm.py/llm_review.py 的模型调用能力
"""

from typing import List, Dict, Any, Optional, Tuple
import json
import os
import sys

try:
    # 包内相对导入（作为包调用时生效）
    from .llm import ask_llm
except Exception:
    # 兼容脚本直跑：将项目根目录加入 sys.path 后再导入
    try:
        _CURR = os.path.dirname(os.path.abspath(__file__))
        _ROOT = os.path.abspath(os.path.join(_CURR, os.pardir, os.pardir))
        if _ROOT not in sys.path:
            sys.path.insert(0, _ROOT)
        from app.pptlint.llm import ask_llm
    except Exception:
        ask_llm = None  # 若模型不可用，后续调用时做空处理


def load_parsing_result(path: str = "parsing_result_new.json") -> List[Dict[str, Any]]:
    """加载 parser 输出的 JSON 结果。
    返回 slides_data: List[Dict[str, Any]]
    """
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_slide_title(slide_map: Dict[str, Any]) -> Optional[str]:
    """提取单页标题
    规则：
    - 优先选择 key 为“文本块X”且 `是否是标题占位符` 为 True 的文本块的“拼接字符”中正文部分
    - 若未命中，则取最上层（图层编号最小）的文本块的“拼接字符”
    - 将前缀标记如【初始的字符所有属性：...】去掉，仅保留文字
    """
    candidates: List[Tuple[int, str]] = []
    for key, payload in slide_map.items():
        if not key.startswith("文本块"):
            continue
        try:
            is_title = bool(payload.get("是否是标题占位符"))
            layer = int(payload.get("图层编号", 1 << 30))
            text = str(payload.get("拼接字符", ""))
        except Exception:
            continue
        # 去除前缀的属性串，仅提取文字部分
        clean = text
        if "】" in clean:
            # 仅去掉最前面的一个属性标记
            idx = clean.find("】")
            clean = clean[idx + 1 :]
        clean = clean.replace("【换行】", "\n").strip()
        if is_title:
            candidates.append((layer - 1000000, clean))  # 强行抬高优先级
        else:
            candidates.append((layer, clean))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    title = candidates[0][1]
    return title if title else None


def extract_all_slide_titles(slides_data: List[Dict[str, Any]]) -> List[Optional[str]]:
    """批量提取每页标题，返回与 slides 对齐的列表。"""
    titles: List[Optional[str]] = []
    for slide_map in slides_data:
        titles.append(extract_slide_title(slide_map))
    return titles


def _call_llm_system(prompt: str, temperature: float = 0.2, max_tokens: int = 1024) -> str:
    """简化的大模型调用封装。根据现有 llm.ask_llm 接口实现。"""
    try:
        if ask_llm is None:
            return ""
        return ask_llm(prompt=prompt, temperature=temperature, max_tokens=max_tokens)
    except Exception:
        return ""


def infer_all_structures(slides_data: List[Dict[str, Any]], titles: List[Optional[str]]) -> Dict[str, Any]:
    """一次性向大模型询问并返回：题目、目录页、章节划分、每页标题。
    返回：{"topic": str, "contents": [int], "sections": [{"title": str, "pages": [int]}], "titles": [str]}
    """
    # 汇总上下文：仅传标题，必要时可扩展加入每页前若干字符
    lines: List[str] = []
    for i, title in enumerate(titles):
        t = title or ""
        lines.append(f"第{i+1}页 标题: {t}")

    prompt = (
        "你是PPT结构分析专家。任务：基于提供的解析后的PPT数据进行分析，分析出PPT的题目、目录、章节名、每页标题，并只输出合法JSON。\n\n"
        "定义：\n"
        "- 题目(topic)：PPT的名称，一般在首页的标题占位符中。\n"
        "- 目录(contents)：列出全卷主要章节或内容提纲的页面；常含‘目录’、‘CONTENTS’等。\n"
        "- 章节名(sections[i].title)：PPT的一级结构标题，一般对应目录中的某一项，如‘背景与目标’、‘方案设计’。\n"
        "  章节页(sections[i].pages)：该章节涵盖的页面（1-based整数数组）。\n"
        "- 每页标题(titles[i])：第 i+1 页的主标题。\n\n"
        "请分析的要点：\n"
        "- 标题中的编号/层级线索（如：1., 第一章/第二章、Part/Section 等）；\n"
        "- 目录关键词（目录、CONTENTS、Agenda 等）与条目式结构；\n"
        "- 章节边界通常在编号跳变或结构性标题（概述/背景/目标/方案/实现/结果/展望）处。\n\n"
        "输出格式（只输出JSON对象，不要解释）：\n"
        "{\n  \"topic\": str,\n  \"contents\": [int],\n  \"sections\": [{\"title\": str, \"pages\": [int]}],\n  \"titles\": [str]\n}\n\n"
        "以下为每页标题（供参考，可在输出中给出你认为更优的 titles）：\n"
        + "\n".join(lines)
    )
    raw = _call_llm_system(prompt)
    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            topic = data.get("topic") if isinstance(data.get("topic"), str) else ""
            contents = data.get("contents") if isinstance(data.get("contents"), list) else []
            sections = data.get("sections") if isinstance(data.get("sections"), list) else []
            titles_llm = data.get("titles") if isinstance(data.get("titles"), list) else []
            return {"topic": topic, "contents": contents, "sections": sections, "titles": titles_llm}
    except Exception:
        pass
    return {"topic": "", "contents": [], "sections": [], "titles": []}


def analyze_from_parsing_result(path: str = "parsing_result_new.json") -> Dict[str, Any]:
    """一站式：加载parser结果 → 提取标题（规则法作为参考） → 调一次LLM返回题目/目录/章节/每页标题。
    返回：{"topic": str, "contents": [...], "sections": [...], "titles": [...]}。
    若 LLM 未提供 titles，则回退为规则法 titles。"""
    slides_data = load_parsing_result(path)
    titles_rule = extract_all_slide_titles(slides_data)
    llm_all = infer_all_structures(slides_data, titles_rule)
    # 回退处理：titles 为空时使用规则法结果
    titles_final = llm_all.get("titles") if llm_all.get("titles") else [t or "" for t in titles_rule]
    return {
        "topic": llm_all.get("topic", ""),
        "contents": llm_all.get("contents", []),
        "sections": llm_all.get("sections", []),
        "titles": titles_final,
    }


if __name__ == "__main__":
    data = analyze_from_parsing_result("parsing_result_new.json")
    print(data)