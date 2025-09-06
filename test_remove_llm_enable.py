#!/usr/bin/env python3
"""
测试删除启用LLM审查选项
"""
import sys
import os

# 添加项目路径
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

def test_remove_llm_enable():
    """测试删除启用LLM审查选项"""
    print("🧪 测试删除启用LLM审查选项...")
    
    try:
        from gui import SimpleApp
        import tkinter as tk
        
        print("✅ 成功导入GUI模块")
        
        # 创建应用实例
        app = SimpleApp()
        print("✅ 成功创建应用实例")
        
        # 显示窗口10秒后关闭
        print("🖥️ 显示窗口10秒，请检查删除启用LLM审查选项的效果...")
        app.after(10000, app.destroy)
        app.mainloop()
        
        print("✅ 删除启用LLM审查选项测试完成")
        
    except Exception as e:
        print(f"❌ 删除启用LLM审查选项测试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_remove_llm_enable()
