#!/usr/bin/env python3
"""
PPT审查工具 - 独立版打包脚本
打包简化版GUI，避免复杂的模块导入问题
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def build_standalone():
    """构建独立版可执行文件"""
    
    # 确保在项目根目录
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    print("🚀 开始构建PPT审查工具独立版...")
    
    # 检查PyInstaller是否安装
    try:
        import PyInstaller
        print(f"✅ PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller未安装，正在安装...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    # 创建spec文件
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['simple_gui_standalone.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # 基础Python模块
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'json',
        'threading',
        'subprocess',
        'tempfile',
        'pathlib',
        'datetime',
        'os',
        'sys',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PPT审查工具_独立版',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 设置为True以显示错误信息
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
'''
    
    # 写入spec文件
    with open('ppt_standalone.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("📝 已创建独立版PyInstaller配置文件")
    
    # 运行PyInstaller
    print("🔨 开始构建...")
    try:
        subprocess.run([
            sys.executable, "-m", "PyInstaller",
            "--clean",
            "ppt_standalone.spec"
        ], check=True)
        
        print("✅ 构建完成！")
        
        # 检查输出文件
        dist_dir = Path("dist")
        if dist_dir.exists():
            exe_files = list(dist_dir.glob("*.exe"))
            if exe_files:
                print(f"🎉 可执行文件已生成: {exe_files[0]}")
                print(f"📁 位置: {exe_files[0].absolute()}")
                
                # 创建启动脚本
                create_launcher_script()
                
                return True
            else:
                print("❌ 未找到生成的可执行文件")
                return False
        else:
            print("❌ 构建失败，未生成dist目录")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ 构建失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 构建过程中出现错误: {e}")
        return False

def create_launcher_script():
    """创建启动脚本"""
    launcher_content = '''@echo off
echo 启动PPT审查工具独立版...
"%~dp0PPT审查工具_独立版.exe"
pause
'''
    
    with open("启动PPT审查工具_独立版.bat", "w", encoding="gbk") as f:
        f.write(launcher_content)
    
    print("📝 已创建启动脚本: 启动PPT审查工具_独立版.bat")

def main():
    """主函数"""
    print("🔧 PPT审查工具 - 独立版打包脚本")
    print("=" * 50)
    
    # 检查必要文件
    required_files = [
        "simple_gui_standalone.py"
    ]
    
    missing_files = [f for f in required_files if not os.path.exists(f)]
    if missing_files:
        print(f"❌ 缺少必要文件: {missing_files}")
        print("请确保在项目根目录运行此脚本")
        return False
    
    print("✅ 所有必要文件已就绪")
    
    # 构建可执行文件
    success = build_standalone()
    
    if success:
        print("\n🎉 独立版打包完成！")
        print("📁 可执行文件位置: dist/")
        print("🚀 使用 '启动PPT审查工具_独立版.bat' 启动程序")
        print("\n⚠️  注意：当前设置为console模式以显示错误信息")
        print("   如需隐藏控制台，请修改spec文件中的console=False")
        print("\n💡 独立版特点：")
        print("   - 避免复杂的模块导入问题")
        print("   - 包含基本的GUI功能")
        print("   - 可以正常运行和测试")
    else:
        print("\n❌ 打包失败，请检查错误信息")
    
    return success

if __name__ == "__main__":
    main()
