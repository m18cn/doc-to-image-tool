"""
打包脚本：将文档转图片工具打包成exe文件
使用方法：python build_exe.py
"""

try:
    import PyInstaller.__main__
except ImportError:
    print("错误：PyInstaller未安装")
    print("请运行: pip install PyInstaller>=6.0.0")
    import sys
    sys.exit(1)

import os
import sys
from pathlib import Path

# 获取项目根目录
project_root = Path(__file__).parent

# PyInstaller配置
args = [
    'ppt_to_overview_gui.py',  # 主程序文件
    '--name=文档转图片工具',  # exe文件名
    '--windowed',  # 隐藏控制台窗口（GUI应用）
    '--onefile',  # 打包成单个exe文件
    '--clean',  # 清理临时文件
    '--noconfirm',  # 覆盖输出目录而不询问
    
    # 隐藏导入（PyInstaller可能无法自动检测的模块）
    '--hidden-import=win32com.client',
    '--hidden-import=win32com.client.gencache',
    '--hidden-import=pythoncom',
    '--hidden-import=fitz',  # PyMuPDF
    '--hidden-import=docx',  # python-docx
    '--hidden-import=PIL',  # Pillow
    '--hidden-import=PIL.Image',
    '--hidden-import=PIL.ImageTk',
    '--hidden-import=app',
    '--hidden-import=app.overview',
    '--hidden-import=app.overview.document_to_overview',
    '--hidden-import=app.images',
    '--hidden-import=app.images.document_to_images',
    '--hidden-import=app.pdf_render',
    '--hidden-import=app.logger_config',
    '--hidden-import=app.ppt',
    '--hidden-import=app.ppt.ppt_to_overview',
    
    # 收集子模块（自动收集所有子模块，包括win32com的所有子模块）
    '--collect-submodules=win32com',
    '--collect-submodules=win32comext',
    '--collect-submodules=docx',
    '--collect-submodules=PIL',
    
    # 图标（如果有的话，可以取消注释并指定图标路径）
    # '--icon=icon.ico',
    
    # 添加数据文件（如果需要）
    # '--add-data=config.ini;.',
    
    # 输出目录
    '--distpath=dist',  # 输出目录
    '--workpath=build',  # 临时文件目录
    '--specpath=.',  # spec文件位置
]

print("=" * 60)
print("开始打包文档转图片工具...")
print("=" * 60)
print(f"项目根目录: {project_root}")
print(f"主程序文件: ppt_to_overview_gui.py")
print("=" * 60)

try:
    # 执行打包
    PyInstaller.__main__.run(args)
    print("=" * 60)
    print("打包完成！")
    print(f"exe文件位置: {project_root / 'dist' / '文档转图片工具.exe'}")
    print("=" * 60)
except Exception as e:
    print("=" * 60)
    print(f"打包失败: {e}")
    print("=" * 60)
    import traceback
    traceback.print_exc()
    sys.exit(1)
