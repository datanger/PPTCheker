"""
命令行入口（对应任务：实现CLI与报告生成）
"""
import argparse
import os
from typing import List
from rich import print

from .config import load_config, ToolConfig
from .parser import parse_pptx
from .rules import run_basic_rules
from .reporter import render_markdown


def _collect_files(input_path: str) -> List[str]:
    if os.path.isdir(input_path):
        out = []
        for root, _dirs, files in os.walk(input_path):
            for f in files:
                if f.lower().endswith(".pptx"):
                    out.append(os.path.join(root, f))
        return out
    return [input_path]


def main():
    parser = argparse.ArgumentParser(description="PPT 格式审查工具")
    parser.add_argument("--input", required=True, help="PPTX 文件或目录")
    parser.add_argument("--config", required=True, help="配置文件 YAML")
    parser.add_argument("--user-req", required=False, help="用户审查需求文档（Markdown/YAML）")
    parser.add_argument("--mode", required=False, choices=["review", "edit"], default="review", help="运行模式")
    parser.add_argument("--llm", required=False, choices=["on", "off"], default="on", help="是否启用LLM（默认on）")
    parser.add_argument("--report", required=False, help="输出报告路径（.md，可选）")
    parser.add_argument("--output-ppt", required=False, help="输出PPT路径（标记或改写版 .pptx）")
    args = parser.parse_args()

    cfg: ToolConfig = load_config(args.config)
    # 用户审查需求解析
    if args.user_req:
        from .user_req import parse_user_requirements
        cfg = parse_user_requirements(args.user_req, cfg)

    # LLM 客户端
    from .llm import LLMClient
    llm = LLMClient() if args.llm == "on" else None

    files = _collect_files(args.input)
    from .workflow import run_review_workflow, run_edit_workflow
    issues_all = []
    for fp in files:
        try:
            if args.mode == "review":
                res = run_review_workflow(fp, cfg, args.output_ppt if len(files) == 1 else None, llm)
            else:
                # 编辑模式需要输出目标
                out_ppt = args.output_ppt if len(files) == 1 else None
                res = run_edit_workflow(fp, cfg, out_ppt, llm)
            issues_all.extend(res.issues)
            print(f"[green]✓[/green] 已处理: {fp}，发现问题 {len(res.issues)}")
        except Exception as e:
            print(f"[red]✗[/red] 失败: {fp}: {e}")

    if args.report:
        content = render_markdown(issues_all)
        os.makedirs(os.path.dirname(args.report), exist_ok=True)
        with open(args.report, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"[cyan]报告已生成[/cyan]: {args.report}")


if __name__ == "__main__":
    main()

