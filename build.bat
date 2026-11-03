@echo off
chcp 65001 >nul
title 批量替换工具 Windows 版 - 一键打包
cd /d "%~dp0"

echo ============================================
echo   批量替换工具 Windows 版 - 一键打包脚本
echo ============================================
echo.

echo [1/4] 检查 Python ...
python --version >nul 2>&1
if errorlevel 1 (
    echo   未检测到 Python，请先安装 Python 3.9+（安装时勾选 "Add python.exe to PATH"）
    echo   下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "delims=" %%v in ('python --version 2^>^&1') do echo   检测到 %%v

echo [2/4] 安装依赖（PySide6 / pywin32 / pyinstaller）...
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo   依赖安装失败，请检查网络后重试
    pause
    exit /b 1
)
echo   依赖安装完成

echo [3/4] 清理旧打包产物 ...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/4] PyInstaller 打包（单文件 exe，图标 app_icon.ico）...
python -m PyInstaller --noconfirm --clean --windowed --onefile ^
    --name "批量替换工具" --icon app_icon.ico main.py
if errorlevel 1 (
    echo   打包失败，请查看上方错误信息
    pause
    exit /b 1
)

echo.
echo ============================================
echo   打包完成！
echo   输出文件: %~dp0dist\批量替换工具.exe
echo   直接双击该 exe 即可使用（可复制到任何 Windows 电脑）
echo ============================================
echo.
pause
