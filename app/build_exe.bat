@echo off
chcp 65001 >nul
title PPT审查工具 - 打包脚本

echo.
echo ========================================
echo    PPT审查工具 - 自动打包脚本
echo ========================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python 3.7+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查是否在正确的目录
if not exist "gui.py" (
    echo 错误：请在包含gui.py的目录中运行此脚本
    pause
    exit /b 1
)

echo Python环境检查通过
echo.

:: 询问是否继续
echo 即将开始打包过程，这可能需要几分钟时间...
echo.
set /p confirm="是否继续？(Y/N): "
if /i "%confirm%" neq "Y" (
    echo 用户取消打包
    pause
    exit /b 0
)

echo.
echo 开始打包过程...
echo.

:: 执行Python打包脚本
python build_exe.py

echo.
if errorlevel 1 (
    echo 打包失败，请检查错误信息
) else (
    echo 打包完成！
    echo.
    echo 可执行文件位置：dist\PPT审查工具.exe
    echo.
    echo 提示：
    echo 1. 可以直接运行exe文件
    echo 2. 确保configs目录与exe在同一目录
    echo 3. 首次运行可能需要较长时间
)

echo.
pause
