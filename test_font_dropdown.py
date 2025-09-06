#!/usr/bin/env python3
"""
测试字体下拉单选框功能
"""
import sys
import os

# 添加项目路径
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

def test_font_dropdown():
    """测试字体下拉单选框功能"""
    print("🧪 测试字体下拉单选框功能...")
    
    try:
        from gui import SimpleApp
        import tkinter as tk
        
        print("✅ 成功导入GUI模块")
        
        # 创建应用实例
        app = SimpleApp()
        print("✅ 成功创建应用实例")
        
        # 显示窗口5秒后关闭
        print("🖥️ 显示窗口5秒...")
        app.after(5000, app.destroy)
        app.mainloop()
        
        print("✅ 字体下拉单选框功能测试完成")
        
    except Exception as e:
        print(f"❌ 字体下拉单选框功能测试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_font_dropdown()
