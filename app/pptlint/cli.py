"""
命令行入口（对应任务：实现CLI与报告生成）
"""
import argparse
import os
from typing import List
from rich import print

from .config import load_config, ToolConfig
from .reporter import render_markdown


def main():
	parser = argparse.ArgumentParser(description="PPT 工具工作流（以 parsing_result.json 为输入）")
	parser.add_argument("--parsing", required=True, help="解析结果文件 parsing_result.json 路径")
	parser.add_argument("--config", required=True, help="配置文件 YAML")
	parser.add_argument("--mode", required=False, choices=["review", "edit"], default="review", help="运行模式：审查/编辑")
	parser.add_argument("--llm", required=False, choices=["on", "off"], default="on", help="是否启用LLM（默认on）")
	parser.add_argument("--report", required=False, help="输出报告路径（.md，可选）")
	parser.add_argument("--output-ppt", required=False, help="输出PPT路径（标记或改写版 .pptx）")
	# 编辑模式专用
	parser.add_argument("--original-pptx", required=False, help="编辑模式：原始PPTX文件路径")
	parser.add_argument("--edit-req", required=False, default="请分析PPT内容，提供改进建议", help="编辑模式：编辑要求提示语")
	args = parser.parse_args()

	cfg: ToolConfig = load_config(args.config)

	# LLM 客户端
	from .llm import LLMClient
	llm = LLMClient() if args.llm == "on" else None

	from .workflow import run_review_workflow, run_edit_workflow

	# 参数校验
	if not os.path.exists(args.parsing):
		print(f"[red]✗[/red] parsing_result.json 不存在: {args.parsing}")
		return

	if args.mode == "edit":
		if not args.original_pptx or not os.path.exists(args.original_pptx):
			print("[red]✗[/red] 编辑模式需要提供有效的 --original-pptx")
			return
		if not args.output_ppt:
			print("[yellow]![/yellow] 未提供 --output-ppt，默认输出为 edited_output.pptx")
			args.output_ppt = "edited_output.pptx"

	# 执行
	try:
		if args.mode == "review":
			res = run_review_workflow(args.parsing, cfg, args.output_ppt, llm)
			print(f"[green]✓[/green] 审查完成。问题数：{len(res.issues)}")
		else:
			res = run_edit_workflow(
				parsing_result_path=args.parsing,
				original_pptx_path=args.original_pptx,
				cfg=cfg,
				output_ppt=args.output_ppt,
				llm=llm,
				edit_requirements=args.edit_req if hasattr(args, 'edit_req') else args.edit_req
			)
			print("[green]✓[/green] 编辑流程完成。")

		# 输出报告
		if args.report:
			content = res.report_md if getattr(res, "report_md", None) else render_markdown(res.issues)
			os.makedirs(os.path.dirname(args.report) or ".", exist_ok=True)
			with open(args.report, "w", encoding="utf-8") as f:
				f.write(content)
			print(f"[cyan]报告已生成[/cyan]: {args.report}")
	except Exception as e:
		print(f"[red]✗[/red] 运行失败: {e}")


if __name__ == "__main__":
	main()

