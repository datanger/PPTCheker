"""
命令行入口（对应任务：实现CLI与报告生成）
"""
import argparse
import os
from typing import List
from datetime import datetime
from rich import print

from .config import load_config, ToolConfig
from .reporter import render_markdown


def generate_output_paths(ppt_path: str, mode: str, output_dir: str) -> tuple:
    """自动生成所有输出文件路径"""
    # 获取原文件名（不含扩展名）
    base_name = os.path.splitext(os.path.basename(ppt_path))[0]
    
    # 获取当前日期
    current_date = datetime.now().strftime("%Y%m%d")
    
    # 生成所有路径
    parsing_result = os.path.join(output_dir, "parsing_result.json")
    report_path = os.path.join(output_dir, f"{base_name}_{mode}_{current_date}.md")
    output_ppt_path = os.path.join(output_dir, f"{base_name}_{mode}_{current_date}.pptx")
    
    return parsing_result, report_path, output_ppt_path


def main():
    parser = argparse.ArgumentParser(description="PPT 工具工作流（自动生成所有输出文件）")
    parser.add_argument("--ppt", required=True, help="输入PPT文件路径（.pptx）")
    parser.add_argument("--mode", required=True, choices=["review", "edit"], help="运行模式：审查/编辑")
    parser.add_argument("--output-dir", required=True, help="输出文件夹路径")
    parser.add_argument("--config", required=True, help="配置文件路径（YAML格式）")
    
    # LLM控制参数
    parser.add_argument("--llm", required=False, choices=["on", "off"], help="是否启用LLM（覆盖配置文件设置）")
    
    # 编辑模式专用
    parser.add_argument("--edit-req", required=False, default="请分析PPT内容，提供改进建议", help="编辑模式：编辑要求提示语")
    
    # 高级配置参数
    parser.add_argument("--font-size", type=int, help="最小字号阈值（覆盖配置文件设置）")
    parser.add_argument("--color-threshold", type=int, help="颜色数量阈值（覆盖配置文件设置）")
    
    args = parser.parse_args()

    # 检查输入文件是否存在
    if not os.path.exists(args.ppt):
        print(f"[red]✗[/red] PPT文件不存在: {args.ppt}")
        return
    
    # 创建输出目录
    os.makedirs(args.output_dir, exist_ok=True)
    
    # 自动生成所有输出路径
    parsing_result_path, report_path, output_ppt_path = generate_output_paths(args.ppt, args.mode, args.output_dir)
    
    # 检查配置文件是否存在
    if not os.path.exists(args.config):
        print(f"[red]✗[/red] 配置文件不存在: {args.config}")
        return
    
    # 加载配置文件
    cfg: ToolConfig = load_config(args.config)
    
    # 应用命令行参数覆盖配置文件设置
    if args.font_size:
        cfg.min_font_size_pt = args.font_size
    if args.color_threshold:
        cfg.color_count_threshold = args.color_threshold
    if args.llm:
        cfg.llm_enabled = (args.llm == "on")

    # 显示配置信息
    print(f"[cyan]配置信息:[/cyan]")
    print(f"  输入PPT: {args.ppt}")
    print(f"  运行模式: {args.mode}")
    print(f"  输出目录: {args.output_dir}")
    print(f"  配置文件: {args.config}")
    print(f"  字体: {cfg.jp_font_name}, 最小字号: {cfg.min_font_size_pt}pt")
    print(f"  缩略语识别: 由LLM智能识别")
    print(f"  颜色阈值: {cfg.color_count_threshold}")
    print(f"  LLM启用: {cfg.llm_enabled}")
    if cfg.llm_enabled:
        print(f"  LLM模型: {cfg.llm_model}, 温度: {cfg.llm_temperature}")
    
    # 显示审查维度配置
    print(f"  审查维度:")
    print(f"    格式规范: {'启用' if cfg.review_format else '禁用'}")
    print(f"    内容逻辑: {'启用' if cfg.review_logic else '禁用'}")
    print(f"    缩略语: {'启用' if cfg.review_acronyms else '禁用'}")
    print(f"    表达流畅性: {'启用' if cfg.review_fluency else '禁用'}")
    
    # 显示审查规则配置
    if cfg.rules:
        print(f"  审查规则:")
        for rule_name, enabled in cfg.rules.items():
            status = "启用" if enabled else "禁用"
            print(f"    {rule_name}: {status}")

    # 显示输出路径信息
    print(f"\n[cyan]输出文件:[/cyan]")
    print(f"  解析结果: {parsing_result_path}")
    print(f"  报告文件: {report_path}")
    print(f"  PPT文件: {output_ppt_path}")

    # 步骤1：解析PPT文件
    print(f"\n[blue]步骤1: 解析PPT文件...[/blue]")
    try:
        from .parser import parse_pptx
        parsing_data = parse_pptx(args.ppt, include_images=False)
        
        # 保存初始解析结果到输出目录
        import json
        with open(parsing_result_path, "w", encoding="utf-8") as f:
            json.dump(parsing_data, f, ensure_ascii=False, indent=2)
        print(f"✅ PPT初始解析完成，结果保存到: {parsing_result_path}")
        
    except Exception as e:
        print(f"[red]✗[/red] PPT解析失败: {e}")
        return

    # LLM 客户端
    from .llm import LLMClient
    llm = LLMClient() if cfg.llm_enabled else None

    from .workflow import run_review_workflow, run_edit_workflow

    # 执行
    try:
        if args.mode == "review":
            print(f"\n[green]步骤2: 开始审查模式...[/green]")
            res = run_review_workflow(parsing_result_path, cfg, output_ppt_path, llm, args.ppt)
            print(f"[green]✓[/green] 审查完成。问题数：{len(res.issues)}")
        else:
            print(f"\n[green]步骤2: 开始编辑模式...[/green]")
            res = run_edit_workflow(
                parsing_result_path=parsing_result_path,
                original_pptx_path=args.ppt,
                cfg=cfg,
                output_ppt=output_ppt_path,
                llm=llm,
                edit_requirements=args.edit_req
            )
            print("[green]✓[/green] 编辑流程完成。")

        # 输出报告
        content = res.report_md if getattr(res, "report_md", None) else render_markdown(res.issues)
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"[cyan]报告已生成[/cyan]: {report_path}")
        
        # 显示结果统计
        if hasattr(res, 'issues') and res.issues:
            print(f"\n[cyan]问题统计:[/cyan]")
            rule_counts = {}
            for issue in res.issues:
                rule_counts[issue.rule_id] = rule_counts.get(issue.rule_id, 0) + 1
            
            for rule_id, count in rule_counts.items():
                print(f"  {rule_id}: {count} 个问题")
        
        print(f"\n[green]✓[/green] 所有文件已保存到: {args.output_dir}")
        
    except Exception as e:
        print(f"[red]✗[/red] 运行失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

