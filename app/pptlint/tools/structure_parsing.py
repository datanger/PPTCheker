"""
PPT结构解析工具

功能：
- 从parsing_result.json中提取PPT结构信息
- 识别页面类型、标题层级、章节信息
- 为LLM审查提供结构化数据
"""
import json
import re
from typing import Dict, Any, List, Optional
try:
    # 优先使用绝对导入（兼容PyInstaller打包）
    from pptlint.model import DocumentModel, Slide, Shape, TextRun
except ImportError:
    try:
        # 尝试相对导入（开发环境）
        from ..model import DocumentModel, Slide, Shape, TextRun
    except ImportError:
        # 兼容直接运行的情况
        import sys
        import os
        sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from model import DocumentModel, Slide, Shape, TextRun


def load_parsing_result(path: str = "parsing_result.json") -> List[Dict[str, Any]]:
    """加载 parser 输出的 JSON 结果。
    返回 slides_data: List[Dict[str, Any]]
    """
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _call_llm_system(prompt: str, temperature: float = 0.2, max_tokens: int = 1024) -> str:
    """简化的大模型调用封装。根据现有 llm.LLMClient 接口实现。"""
    try:
        # 包内相对导入（作为包调用时生效）
        from ..llm import LLMClient
        if LLMClient is None:
            return ""
        # 创建LLM客户端实例
        llm_client = LLMClient()
        # 调用complete方法
        return llm_client.complete(prompt=prompt, max_tokens=max_tokens)
    except Exception:
        return ""


def infer_all_structures(slides_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """一次性向大模型询问并返回：题目、目录页、章节划分、每页标题。
    返回：{"topic": str, "contents": [int], "sections": [{"title": str, "pages": [int]}], "titles": [str]}
    """
    # 直接传递PPT原始数据给大模型分析
    prompt = f"""你是PPT结构分析专家。任务：基于提供的PPT原始数据进行分析，分析出PPT的题目、目录、章节页、每页标题，并只输出合法JSON。

        定义：
        - 题目(topic)：PPT的名称，一般在首页的标题占位符中。
        - 目录(contents)：列出全卷主要章节或内容提纲的页面；常含'目录'、'CONTENTS'等。需要返回具体的目录内容，格式为[{{"page": int, "title": str, "level": int}}]，其中level表示层级（1为一级标题，2为二级标题等）。
        - 章节页(sections)：**重要**：章节页是指PPT中真实存在的章节分隔页面，不是总结出来的。章节页的特征：
            * 首个章节页一般出现在目录页之后，也可能没有章节页
            * 章节页的内容通常是目录中的一个条目（标题或标题+序号），但一定要注意这只是通常情况，有可能人为书写错误，实际内容可能与目录不一致，需要根据真实的的章节页内容输出，即使人为错误页不要修正
            * 章节页通常只有一个文本块， 并且文本块在页面中间，内容是章节名或章节名+序号
            * 如果该页的内容比较多，一般肯定不是章节页
            * 如果PPT中没有明显的章节分隔页，则sections为空数组
            * 格式：sections[i].title为章节页的标题，sections[i].pages为该章节页的页码（通常只有一页）
        - 每页标题(titles)：第 i+1 页的主标题， 一般位于每页的左上角， 如果该页没有标题，则titles[i]为空。

        请分析的要点：
        - **不要跨越信息边界**： 每页输出的结果只能包含该页的信息，不要跨越页数， 虽然可以参考，但不要将不同页的信息合并到一起。
        - **实事求是**：如果PPT中没有明显的章节分隔页，则不要强行创建章节结构，sections保持为空。
        - **不要修正人为错误**：如果目录或章节页中存在人为错误没有对应上，也不要修正错误，这正是我们后续需要分析的。
        - **文本块位置**： 文本块位置是文本块在页面中的位置， 单位为百分比， 相对左上角，这在分析目录、章节页时非常重要。
        - **文本块数量**： 文本块数量是该页的文本块数量， 一般为1，该页有多个文本块通常不是章节页， 但有可能左上角也会有一个文本块，但其文本内容较少或为空。
        - **段落属性**： 每个文本块包含按 run 合并后的“段落属性”数组（字体、字号、颜色、样式、字符内容）。

        输出格式（只输出JSON对象，不要解释）：
        {{
        "topic": {{"text": str, "page": int}},
        "contents": [{{"text": str, "page": int}}],
        "sections": [{{"text": str, "page": int}}],
        "titles": [{{"text": str, "page": int}}]
        }}

        以下是PPT的原始数据，请直接分析：
        {json.dumps(slides_data, ensure_ascii=False, indent=2)}"""

    raw = _call_llm_system(prompt)
    
    # 调试信息：显示大模型原始返回
    # print(f"🔍 大模型原始返回: {raw}")
    
    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            return {
                "topic": data.get("topic", ""),           # 兼容 str 或 {title,page}
                "contents": data.get("contents", []),       # 兼容含/不含 level
                "sections": data.get("sections", []),       # 兼容 {title,page} 或 {title,pages}
                "titles": data.get("titles", [])            # 兼容 [str] 或 [{title,page}]
            }
    except Exception as e:
        print(f"🔍 JSON解析错误: {e}")
        pass
    return {"topic": "", "contents": [], "sections": [], "titles": []}


def analyze_from_parsing_result(parsing_data: Dict[str, Any]) -> Dict[str, Any]:
    """一站式：加载parser结果 → 调一次LLM返回题目/目录/章节/每页标题。
    返回：{"topic": str, "contents": [...], "sections": [...], "titles": [...], "structure": str, "page_types": [...], "page_titles": [...]}。
    完全依赖大模型分析，无规则法回退。"""
    llm_all = infer_all_structures(parsing_data)
    
    # 生成PPT结构汇总字符串
    structure_lines = []
    
    # 1. 主题
    topic_obj = llm_all.get('topic', '')
    topic_title = ""
    topic_page = 1
    if isinstance(topic_obj, dict):
        topic_title = topic_obj.get('text', '') or ""
        tp = topic_obj.get('page')
        if isinstance(tp, int) and tp > 0:
            topic_page = tp
    elif isinstance(topic_obj, str):
        topic_title = topic_obj
    structure_lines.append(f"主题：{topic_title or '无'} （页码：[{topic_page}]）" if topic_title else "主题：无")
    
    # 2. 目录页
    contents = llm_all.get('contents', [])
    if contents:
        structure_lines.append("目录：")
        for item in contents:
            if isinstance(item, dict):
                title = item.get('text', '')
                page = item.get('page', None)
                line = f"      {title}" + (f" （页码：[{page}]）" if isinstance(page, int) else "")
                structure_lines.append(line)
            else:
                structure_lines.append(f"      {item}")
    else:
        structure_lines.append("目录：无")
    
    # 3. 按页码顺序显示章节和标题（实事求是，还原真实内容）
    sections = llm_all.get('sections', [])
    titles = llm_all.get('titles', [])
    # 规范化：构建 page→section_title / page→title 的映射
    section_pages: Dict[int, Dict[str, Any]] = {}
    for sec in sections:
        if isinstance(sec, dict):
            # 兼容 {title,page} 或 {title,pages:[...]}
            pg = sec.get('page')
            if isinstance(pg, int):
                section_pages[pg] = sec
            else:
                pages = sec.get('pages', [])
                if isinstance(pages, list) and pages:
                    p0 = pages[0]
                    if isinstance(p0, int):
                        section_pages[p0] = sec

    titles_map: Dict[int, str] = {}
    if isinstance(titles, list):
        if titles and isinstance(titles[0], dict):
            for t in titles:
                if isinstance(t, dict):
                    pg = t.get('page')
                    title = t.get('text', '')
                    if isinstance(pg, int):
                        titles_map[pg] = title
        else:
            # 旧格式：按索引对应页码
            for i, title in enumerate(titles):
                if isinstance(title, str):
                    titles_map[i + 1] = title
    
    # 初始化页类型和页标题数组
    page_types = []
    page_titles = []
    
    total_pages = parsing_data.get('页数') or len(parsing_data.get('contents', [])) or max([0] + list(titles_map.keys()) + list(section_pages.keys()))

    if total_pages == 0:
        structure_lines.append("标题：无")
    else:
        for page_num in range(1, total_pages + 1):
            if page_num == topic_page:
                # 跳过主题行重复打印
                continue
            if any(isinstance(c, dict) and c.get('page') == page_num for c in contents):
                # 目录页
                continue
            if page_num in section_pages:
                sec = section_pages[page_num]
                stitle = sec.get('text', '')
                structure_lines.append(f"章节：{stitle} （页码：[{page_num}]）")
            else:
                t = titles_map.get(page_num, '')
                if t:
                    structure_lines.append(f"标题：{t} （页码：[{page_num}]）")
    
    # 生成structure字符串
    structure = "\n".join(structure_lines)
    print(f"🔍 结构分析结果\n: {structure}")
    
    # 生成页类型和页标题数组
    for page_num in range(1, (total_pages or 0) + 1):
        # 页类型
        if page_num == topic_page:
            page_type = "主题页"
        elif any(isinstance(c, dict) and c.get('page') == page_num for c in contents):
            page_type = "目录页"
        elif page_num in section_pages:
            page_type = "章节页"
        else:
            page_type = "内容页"
        page_types.append(page_type)

        # 页标题
        if page_num == topic_page:
            page_titles.append(topic_title)
        elif page_num in section_pages:
            page_titles.append(section_pages[page_num].get('text', ''))
        else:
            page_titles.append(titles_map.get(page_num, ''))
    
    parsing_data["structure"] = structure
    
    for i, page in enumerate(parsing_data.get("contents", [])):
        if i < len(page_types):
            page["页类型"] = page_types[i]
        if i < len(page_titles):
            page["页标题"] = page_titles[i]
    
    return parsing_data


if __name__ == "__main__":
    # 静默运行，只更新 parsing_result.json
    parsing_data = load_parsing_result("parsing_result.json")

    parsing_data = analyze_from_parsing_result(parsing_data)
    
    print(parsing_data['structure'])
    
    # 写回文件
    with open("parsing_result.json", "w", encoding="utf-8") as f:
        json.dump(parsing_data, f, ensure_ascii=False, indent=2)
