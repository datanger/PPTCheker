#!/usr/bin/env python3
"""
CLI功能测试脚本
演示如何使用CLI进行PPT审查和编辑
"""

import os
import subprocess
import sys

def run_cli_command(cmd_args, description):
    """运行CLI命令并显示结果"""
    print(f"\n{'='*60}")
    print(f"🧪 测试: {description}")
    print(f"📝 命令: python -m app.pptlint.cli {' '.join(cmd_args)}")
    print(f"{'='*60}")
    
    try:
        # 运行CLI命令
        result = subprocess.run(
            [sys.executable, "-m", "app.pptlint.cli"] + cmd_args,
            capture_output=True,
            text=True,
            cwd=os.getcwd()
        )
        
        # 显示输出
        if result.stdout:
            print("✅ 标准输出:")
            print(result.stdout)
        
        if result.stderr:
            print("⚠️ 错误输出:")
            print(result.stderr)
        
        print(f"退出码: {result.returncode}")
        
        return result.returncode == 0
        
    except Exception as e:
        print(f"❌ 执行失败: {e}")
        return False

def main():
    """主测试函数"""
    print("🚀 PPT审查工具CLI功能测试")
    print("本测试将演示CLI的各种功能")
    
    # 检查必要文件
    required_files = [
        "parsing_result.json",
        "configs/config.yaml",
        "example2.pptx"
    ]
    
    missing_files = [f for f in required_files if not os.path.exists(f)]
    if missing_files:
        print(f"❌ 缺少必要文件: {missing_files}")
        print("请确保在项目根目录运行此测试")
        return
    
    print("✅ 所有必要文件已就绪")
    
    # 测试1: 基础审查模式（生成报告和标记PPT）
    test1_success = run_cli_command([
        "--parsing", "parsing_result.json",
        "--config", "configs/config.yaml",
        "--mode", "review",
        "--report", "test_report.md",
        "--output-ppt", "test_output.pptx"
    ], "基础审查模式 - 生成报告和标记PPT")
    
    # 测试2: 仅生成报告（不生成PPT）
    test2_success = run_cli_command([
        "--parsing", "parsing_result.json",
        "--config", "configs/config.yaml",
        "--mode", "review",
        "--report", "test_report_only.md"
    ], "仅生成报告模式")
    
    # 测试3: 仅生成标记PPT（不生成报告）
    test3_success = run_cli_command([
        "--parsing", "parsing_result.json",
        "--config", "configs/config.yaml",
        "--mode", "review",
        "--output-ppt", "test_ppt_only.pptx"
    ], "仅生成标记PPT模式")
    
    # 测试4: 禁用LLM的审查模式
    test4_success = run_cli_command([
        "--parsing", "parsing_result.json",
        "--config", "configs/config.yaml",
        "--mode", "review",
        "--llm", "off",
        "--report", "test_no_llm.md",
        "--output-ppt", "test_no_llm.pptx"
    ], "禁用LLM的审查模式")
    
    # 测试5: 自定义配置参数
    test5_success = run_cli_command([
        "--parsing", "parsing_result.json",
        "--config", "configs/config.yaml",
        "--mode", "review",
        "--font-size", "14",
        "--color-threshold", "3",
        "--acronym-min-len", "3",
        "--acronym-max-len", "6",
        "--report", "test_custom_config.md",
        "--output-ppt", "test_custom_config.pptx"
    ], "自定义配置参数模式")
    
    # 测试6: 编辑模式
    test6_success = run_cli_command([
        "--parsing", "parsing_result.json",
        "--config", "configs/config.yaml",
        "--mode", "edit",
        "--original-pptx", "example2.pptx",
        "--output-ppt", "test_edited.pptx",
        "--edit-req", "请优化PPT的字体大小和颜色搭配，使其更加美观易读",
        "--report", "test_edit_report.md"
    ], "编辑模式 - 优化PPT样式")
    
    # 测试7: 帮助信息
    test7_success = run_cli_command([
        "--help"
    ], "显示帮助信息")
    
    # 测试结果汇总
    print(f"\n{'='*60}")
    print("📊 测试结果汇总")
    print(f"{'='*60}")
    
    tests = [
        ("基础审查模式", test1_success),
        ("仅生成报告", test2_success),
        ("仅生成标记PPT", test3_success),
        ("禁用LLM审查", test4_success),
        ("自定义配置参数", test5_success),
        ("编辑模式", test6_success),
        ("帮助信息", test7_success)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, success in tests:
        status = "✅ 通过" if success else "❌ 失败"
        print(f"{test_name}: {status}")
        if success:
            passed += 1
    
    print(f"\n总计: {passed}/{total} 个测试通过")
    
    if passed == total:
        print("🎉 所有测试通过！CLI功能正常")
    else:
        print("⚠️ 部分测试失败，请检查相关功能")
    
    # 清理测试文件
    print(f"\n🧹 清理测试文件...")
    test_files = [
        "test_report.md", "test_output.pptx",
        "test_report_only.md", "test_ppt_only.pptx",
        "test_no_llm.md", "test_no_llm.pptx",
        "test_custom_config.md", "test_custom_config.pptx",
        "test_edited.pptx", "test_edit_report.md"
    ]
    
    for file in test_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"  删除: {file}")
            except Exception as e:
                print(f"  删除失败 {file}: {e}")
    
    print("✅ 清理完成")

if __name__ == "__main__":
    main()
