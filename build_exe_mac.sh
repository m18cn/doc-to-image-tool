#!/bin/bash
# ============================================================
# 文档转图片工具 - macOS打包脚本
# 必须在 macOS 上执行，产物: dist/文档转图片工具.app
#
# 前置要求:
#   1. Python 3.10+（推荐 python.org 官方安装包，自带tkinter）
#      conda 环境需额外安装: conda install python-tk
#   2. LibreOffice（PPT/Word转换依赖）:
#      brew install --cask libreoffice
#      注意: 运行时机器上也需安装LibreOffice
#
# 使用方法:
#   chmod +x build_exe_mac.sh && ./build_exe_mac.sh
# ============================================================
set -e

echo "============================================================"
echo "文档转图片工具 - macOS打包"
echo "============================================================"

# 检查操作系统
if [ "$(uname)" != "Darwin" ]; then
    echo "[错误] 此脚本必须在 macOS 上执行（PyInstaller不支持交叉打包）"
    exit 1
fi

# 检查tkinter可用性
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "[错误] 当前Python缺少tkinter模块"
    echo "  官方Python安装包自带tkinter；conda环境请执行: conda install python-tk"
    exit 1
fi

# 创建并激活虚拟环境（已存在则复用）
if [ ! -d "venv_mac" ]; then
    echo "创建虚拟环境 venv_mac..."
    python3 -m venv venv_mac
fi
source venv_mac/bin/activate

# 安装依赖（requirements中pywin32带平台标记，mac上自动跳过）
echo "安装依赖..."
pip install --upgrade pip
pip install -r requirements.txt

# 打包（--windowed 在macOS上生成.app应用包）
echo "开始打包..."
pyinstaller --clean --noconfirm \
    --name "文档转图片工具" \
    --windowed \
    --onefile \
    --exclude-module win32com \
    --exclude-module pythoncom \
    --exclude-module pywintypes \
    --exclude-module win32api \
    --exclude-module win32con \
    --hidden-import fitz \
    --hidden-import docx \
    --hidden-import PIL \
    --hidden-import PIL.Image \
    --hidden-import PIL.ImageTk \
    --hidden-import app \
    --hidden-import app.overview \
    --hidden-import app.overview.document_to_overview \
    --hidden-import app.images \
    --hidden-import app.images.document_to_images \
    --hidden-import app.overview.multi_document_to_overview \
    --hidden-import app.pdf_render \
    --hidden-import app.office_convert \
    --hidden-import app.logger_config \
    --collect-submodules docx \
    --distpath dist \
    --workpath build_mac \
    --specpath . \
    ppt_to_overview_gui.py

echo "============================================================"
echo "打包完成！"
echo "应用位置: dist/文档转图片工具.app"
echo "============================================================"
