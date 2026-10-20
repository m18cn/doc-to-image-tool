@echo off
chcp 65001 >nul
echo ============================================================
echo 文档转图片工具 - 打包脚本
echo ============================================================
echo.

REM 检查是否在虚拟环境中
if not defined VIRTUAL_ENV (
    echo [警告] 未检测到虚拟环境，建议在虚拟环境中运行打包
    echo.
)

REM 检查PyInstaller是否已安装
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [错误] PyInstaller未安装
    echo 正在安装PyInstaller...
    pip install PyInstaller>=6.0.0
    if errorlevel 1 (
        echo [错误] PyInstaller安装失败
        pause
        exit /b 1
    )
)

echo 开始打包...
echo.

REM 运行打包脚本
python build_exe.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败
    pause
    exit /b 1
) else (
    echo.
    echo [成功] 打包完成！
    echo exe文件位置: dist\文档转图片工具.exe
    echo.
)

pause

