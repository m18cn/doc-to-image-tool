"""
Office文档跨平台转换工具
功能：在非Windows平台(macOS/Linux)上通过LibreOffice headless将PPT/Word转换为PDF，
     供后续PyMuPDF渲染管线使用。Windows平台仍走Office COM，不经过此模块。
"""

import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Callable

from app.logger_config import get_logger

logger = get_logger("office_convert")

# soffice转换单文件超时时间（秒），超长文档或首次启动LibreOffice可能较慢
SOFFICE_TIMEOUT = 600


def find_soffice() -> Optional[str]:
    """
    查找LibreOffice的soffice可执行文件路径

    Returns:
        soffice可执行文件的绝对路径；未找到时返回None
    """
    # 各平台常见安装路径
    candidates = [
        # macOS（LibreOffice标准安装位置）
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Linux常见位置
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
    ]
    for path in candidates:
        if Path(path).exists():
            return path

    # 最后回退到PATH环境变量查找
    return shutil.which("soffice")


def convert_to_pdf(doc_path: Path, output_dir: Path, log_callback: Optional[Callable] = None) -> Path:
    """
    使用LibreOffice headless将PPT/Word文档转换为PDF

    Args:
        doc_path: 源文档路径（.pptx/.ppt/.docx/.doc）
        output_dir: PDF输出目录（自动创建"_soffice_pdf"子目录隔离中间文件）
        log_callback: 日志回调函数（GUI界面使用）

    Returns:
        生成的PDF文件路径

    Raises:
        FileNotFoundError: 未找到LibreOffice或源文档不存在
        RuntimeError: LibreOffice转换失败或超时
    """
    doc_path = Path(doc_path)

    if not doc_path.exists():
        raise FileNotFoundError(f"源文档不存在: {doc_path}")

    soffice = find_soffice()
    if soffice is None:
        raise FileNotFoundError(
            "未找到LibreOffice。请先安装：\n"
            "  macOS: brew install --cask libreoffice 或从官网下载\n"
            "  Linux: sudo apt install libreoffice"
        )

    # PDF输出到独立子目录，避免与源文件同名冲突（soffice以源文件stem命名输出）
    pdf_out_dir = Path(output_dir) / "_soffice_pdf"
    pdf_out_dir.mkdir(parents=True, exist_ok=True)

    def _log(msg: str):
        """本地日志辅助：同时写文件日志和回调"""
        logger.info(msg)
        if log_callback:
            log_callback(msg)

    _log(f"使用LibreOffice转换: {doc_path.name} -> PDF")

    # 使用独立的UserInstallation配置目录，避免与本机正在运行的LibreOffice实例冲突
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        profile_url = Path(profile_dir).as_uri()
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_url}",
            "--headless",           # 无界面模式
            "--norestore",          # 不恢复上次崩溃会话
            "--convert-to", "pdf",
            "--outdir", str(pdf_out_dir),
            str(doc_path.absolute()),
        ]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=SOFFICE_TIMEOUT,
                encoding="utf-8",
                errors="replace",
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError(f"LibreOffice转换超时({SOFFICE_TIMEOUT}秒): {doc_path.name}")

    expected_pdf = pdf_out_dir / f"{doc_path.stem}.pdf"
    if result.returncode != 0 or not expected_pdf.exists():
        stderr = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"LibreOffice转换失败: {stderr[:500] or '未知错误'}")

    _log(f"PDF转换完成: {expected_pdf.name}")
    return expected_pdf
