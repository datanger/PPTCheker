#!/usr/bin/env python3
"""
PPT审查工具 - PyInstaller打包脚本
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def build_exe():
    """构建可执行文件"""
    
    # 确保在项目根目录
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    print("🚀 开始构建PPT审查工具可执行文件...")
    
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
    ['app/pptlint/simple_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app/configs/config.yaml', 'configs'),
        ('dicts', 'dicts'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'yaml',
        'json',
        'urllib.request',
        'urllib.error',
        'threading',
        'subprocess',
        'platform',
        'pathlib',
        'pptlint',
        'pptlint.config',
        'pptlint.llm',
        'pptlint.parser',
        'pptlint.workflow',
        'pptlint.tools.workflow_tools',
        'pptlint.tools.llm_review',
        'pptlint.tools.structure_parsing',
        'pptlint.annotator',
        'pptlint.reporter',
        'pptlint.model',
        'pptlint.cli',
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
    name='PPT审查工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为False以隐藏控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
'''
    
    # 写入spec文件
    with open('ppt_checker.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("📝 已创建PyInstaller配置文件")
    
    # 运行PyInstaller
    print("🔨 开始构建...")
    try:
        subprocess.run([
            sys.executable, "-m", "PyInstaller",
            "--clean",
            "ppt_checker.spec"
        ], check=True)
        
        print("✅ 构建完成！")
        
        # 检查输出文件
        dist_dir = Path("dist")
        if dist_dir.exists():
            exe_files = list(dist_dir.glob("*.exe"))
            if exe_files:
                print(f"🎉 可执行文件已生成: {exe_files[0]}")
                print(f"📁 位置: {exe_files[0].absolute()}")
                
                # 创建简单的启动脚本
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

def create_launcher_script():
    """创建启动脚本"""
    launcher_content = '''@echo off
echo 正在启动PPT审查工具...
cd /d "%~dp0"
start "" "PPT审查工具.exe"
'''
    
    with open('启动PPT审查工具.bat', 'w', encoding='gbk') as f:
        f.write(launcher_content)
    
    print("📝 已创建启动脚本: 启动PPT审查工具.bat")

def clean_build():
    """清理构建文件"""
    print("🧹 清理构建文件...")
    
    dirs_to_clean = ['build', '__pycache__']
    files_to_clean = ['ppt_checker.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"🗑️ 已删除目录: {dir_name}")
    
    for file_name in files_to_clean:
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"🗑️ 已删除文件: {file_name}")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="PPT审查工具打包脚本")
    parser.add_argument("--clean", action="store_true", help="清理构建文件")
    parser.add_argument("--build", action="store_true", help="构建可执行文件")
    
    args = parser.parse_args()
    
    if args.clean:
        clean_build()
    elif args.build:
        build_exe()
    else:
        # 默认构建
        build_exe()
