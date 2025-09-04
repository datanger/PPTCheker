#!/usr/bin/env python3
"""
PPT审查工具 - 完整打包脚本
包含所有执行所需的依赖和配置

使用方法：
python build_exe.py
"""

import os
import sys
import shutil
import subprocess
import platform
from pathlib import Path

def print_step(step, description):
    """打印步骤信息"""
    print(f"\n{'='*60}")
    print(f"步骤 {step}: {description}")
    print(f"{'='*60}")

def run_command(command, description):
    """运行命令并显示结果"""
    print(f"\n🔧 {description}")
    print(f"执行命令: {command}")
    
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ 命令执行成功")
        if result.stdout:
            print(f"输出: {result.stdout}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 命令执行失败: {e}")
        if e.stdout:
            print(f"标准输出: {e.stdout}")
        if e.stderr:
            print(f"错误输出: {e.stderr}")
        return False

def check_dependencies():
    """检查必要的依赖"""
    print_step(1, "检查依赖")
    
    # 定义包名和导入名的映射
    package_mapping = {
        'pyinstaller': 'PyInstaller',
        'pyyaml': 'yaml',
        'python-pptx': 'pptx',
        'pillow': 'PIL',
        'lxml': 'lxml',
        'tkinter': 'tkinter'
    }
    
    missing_packages = []
    
    for package, import_name in package_mapping.items():
        try:
            if package == 'tkinter':
                import tkinter
                print(f"✅ {package} - 已安装")
            else:
                __import__(import_name)
                print(f"✅ {package} - 已安装")
        except ImportError:
            print(f"❌ {package} - 未安装")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n⚠️ 缺少以下依赖包: {missing_packages}")
        print("请先安装缺失的依赖:")
        for package in missing_packages:
            if package == 'pyinstaller':
                print(f"pip install {package}")
            elif package == 'pyyaml':
                print(f"pip install {package}")
            elif package == 'python-pptx':
                print(f"pip install {package}")
            elif package == 'pillow':
                print(f"pip install {package}")
            elif package == 'lxml':
                print(f"pip install {package}")
        print(f"\n或者一次性安装所有依赖:")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    print("✅ 所有依赖检查通过")
    return True

def clean_build_dirs():
    """清理构建目录"""
    print_step(2, "清理构建目录")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"🗑️ 删除目录: {dir_name}")
            shutil.rmtree(dir_name)
        else:
            print(f"✅ 目录不存在: {dir_name}")
    
    # 清理spec文件
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    for spec_file in spec_files:
        print(f"🗑️ 删除spec文件: {spec_file}")
        os.remove(spec_file)
    
    print("✅ 清理完成")

def create_pyinstaller_command():
    """创建PyInstaller命令"""
    print_step(3, "创建PyInstaller命令")
    
    # 基础命令
    base_cmd = [
        'pyinstaller',
        '--onefile',                    # 打包成单个exe文件
        '--windowed',                   # 无控制台窗口
        '--name', 'PPT审查工具',        # 可执行文件名
        '--icon', 'NONE',              # 图标（暂时不使用）
    ]
    
    # 添加数据文件
    data_files = [
        '--add-data', 'configs;configs',           # 配置文件
        '--add-data', 'pptlint;pptlint',           # pptlint模块
    ]
    
    # 添加隐藏导入
    hidden_imports = [
        '--hidden-import', 'pptlint',
        '--hidden-import', 'pptlint.config',
        '--hidden-import', 'pptlint.workflow',
        '--hidden-import', 'pptlint.llm',
        '--hidden-import', 'pptlint.parser',
        '--hidden-import', 'pptlint.cli',
        '--hidden-import', 'pptlint.tools.llm_review',
        '--hidden-import', 'pptlint.tools.structure_parsing',
        '--hidden-import', 'pptlint.tools.workflow_tools',
        '--hidden-import', 'pptlint.tools.rules',
        '--hidden-import', 'pptlint.reporter',
        '--hidden-import', 'pptlint.annotator',
        '--hidden-import', 'pptlint.serializer',
        '--hidden-import', 'pptlint.user_req',
        '--hidden-import', 'yaml',
        '--hidden-import', 'PIL',
        '--hidden-import', 'lxml',
        '--hidden-import', 'pptx',
    ]
    
    # 添加排除模块（减少文件大小）
    excludes = [
        '--exclude-module', 'matplotlib',
        '--exclude-module', 'numpy',
        '--exclude-module', 'pandas',
        '--exclude-module', 'scipy',
        '--exclude-module', 'jupyter',
        '--exclude-module', 'IPython',
        '--exclude-module', 'notebook',
    ]
    
    # 其他选项
    other_options = [
        '--clean',                      # 清理临时文件
        '--noconfirm',                 # 不询问确认
        '--log-level', 'INFO',         # 日志级别
    ]
    
    # 组合完整命令
    full_cmd = base_cmd + data_files + hidden_imports + excludes + other_options + ['gui.py']
    
    return ' '.join(full_cmd)

def build_executable():
    """构建可执行文件"""
    print_step(4, "构建可执行文件")
    
    # 创建PyInstaller命令
    pyinstaller_cmd = create_pyinstaller_command()
    
    print("📋 PyInstaller命令:")
    print(pyinstaller_cmd)
    
    # 执行打包
    if run_command(pyinstaller_cmd, "执行PyInstaller打包"):
        print("✅ 打包完成")
        return True
    else:
        print("❌ 打包失败")
        return False

def verify_build():
    """验证构建结果"""
    print_step(5, "验证构建结果")
    
    # 检查dist目录
    dist_dir = Path('dist')
    if not dist_dir.exists():
        print("❌ dist目录不存在")
        return False
    
    # 查找exe文件
    exe_files = list(dist_dir.glob('*.exe'))
    if not exe_files:
        print("❌ 未找到exe文件")
        return False
    
    exe_file = exe_files[0]
    print(f"✅ 找到可执行文件: {exe_file}")
    
    # 检查文件大小
    file_size = exe_file.stat().st_size / (1024 * 1024)  # MB
    print(f"📏 文件大小: {file_size:.2f} MB")
    
    # 检查是否包含必要文件
    print("\n🔍 检查打包内容...")
    
    # 使用PyInstaller的--list选项检查内容（如果支持）
    try:
        list_cmd = f'pyinstaller --list "{exe_file}"'
        result = subprocess.run(list_cmd, shell=True, capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            print("📋 打包内容列表:")
            print(result.stdout)
        else:
            print("⚠️ 无法获取打包内容列表")
    except Exception as e:
        print(f"⚠️ 检查打包内容时出错: {e}")
    
    return True

def create_installer():
    """创建安装包（可选）"""
    print_step(6, "创建安装包")
    
    print("💡 安装包创建功能（可选）")
    print("可以使用以下工具创建安装包:")
    print("1. Inno Setup - 创建Windows安装程序")
    print("2. NSIS - 创建Windows安装程序")
    print("3. 手动打包 - 将exe和相关文件打包成zip")
    
    # 创建简单的zip包
    try:
        import zipfile
        
        dist_dir = Path('dist')
        exe_files = list(dist_dir.glob('*.exe'))
        
        if exe_files:
            exe_file = exe_files[0]
            zip_name = f"{exe_file.stem}_完整版.zip"
            
            print(f"\n📦 创建zip包: {zip_name}")
            
            with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # 添加exe文件
                zipf.write(exe_file, exe_file.name)
                
                # 添加配置文件
                if os.path.exists('configs'):
                    for root, dirs, files in os.walk('configs'):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, '.')
                            zipf.write(file_path, arcname)
                
                # 添加README
                if os.path.exists('README.md'):
                    zipf.write('README.md', 'README.md')
            
            print(f"✅ 创建zip包成功: {zip_name}")
        else:
            print("❌ 未找到exe文件，无法创建zip包")
            
    except Exception as e:
        print(f"⚠️ 创建zip包时出错: {e}")

def main():
    """主函数"""
    print("🚀 PPT审查工具 - 完整打包脚本")
    print(f"📁 当前工作目录: {os.getcwd()}")
    print(f"🖥️ 操作系统: {platform.system()} {platform.release()}")
    print(f"🐍 Python版本: {sys.version}")
    
    # 检查依赖
    if not check_dependencies():
        print("\n❌ 依赖检查失败，请先安装缺失的依赖")
        return False
    
    # 清理构建目录
    clean_build_dirs()
    
    # 构建可执行文件
    if not build_executable():
        print("\n❌ 构建失败")
        return False
    
    # 验证构建结果
    if not verify_build():
        print("\n❌ 验证失败")
        return False
    
    # 创建安装包
    create_installer()
    
    print("\n🎉 打包完成！")
    print("\n📋 使用说明:")
    print("1. 可执行文件位于 dist/ 目录")
    print("2. 可以直接运行exe文件")
    print("3. 确保configs目录与exe在同一目录")
    print("4. 首次运行可能需要较长时间")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if success:
            print("\n✅ 所有步骤完成")
        else:
            print("\n❌ 打包过程中出现错误")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️ 用户中断打包过程")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 打包脚本执行出错: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
