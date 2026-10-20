"""
PDF页面安全渲染工具
功能：渲染PDF页面为位图时进行内存预算控制，防止超大页面（如手机长截图导出的
长图PDF、网页另存为PDF）在高DPI下分配巨量内存导致 malloc failed 崩溃。
"""

from app.logger_config import get_logger

# 单页渲染像素上限：与PIL DecompressionBomb默认阈值(MAX_IMAGE_PIXELS)一致，
# 保证后续PIL拼接环节不会触发告警或报错
MAX_RENDER_PIXELS = 89_478_485

# 每像素字节数：alpha=False时为RGB三通道
BYTES_PER_PIXEL = 3

logger = get_logger("pdf_safe_render")


def render_page_safely(page, dpi: int = 300, log_callback=None):
    """
    安全渲染PDF页面为位图，超大页面自动降级DPI以避免内存溢出

    Args:
        page: fitz.Page 对象（PyMuPDF页面）
        dpi: 目标分辨率DPI，默认300
        log_callback: 日志回调函数（GUI界面使用），仅在触发降级时调用

    Returns:
        fitz.Pixmap: 渲染结果位图（RGB模式，无alpha通道）

    Raises:
        ImportError: 未安装PyMuPDF库时抛出
    """
    import fitz

    # 计算目标缩放系数（DPI转zoom）
    zoom = dpi / 72.0
    rect = page.rect
    # 估算全页渲染所需像素总量（宽px * 高px）
    est_pixels = rect.width * zoom * rect.height * zoom

    effective_dpi = dpi
    if est_pixels > MAX_RENDER_PIXELS:
        # 超限：按面积比例缩小zoom，将像素总量压回上限内
        scale = (MAX_RENDER_PIXELS / est_pixels) ** 0.5
        effective_dpi = max(1, int(dpi * scale))
        zoom = effective_dpi / 72.0
        est_mb = est_pixels * BYTES_PER_PIXEL / 1024 / 1024
        msg = (
            f"页面尺寸过大({rect.width:.0f}x{rect.height:.0f}pt, "
            f"原始渲染约需{est_mb:.0f}MB内存)，"
            f"自动降级分辨率: {dpi} DPI -> {effective_dpi} DPI"
        )
        logger.warning(msg)
        if log_callback:
            log_callback(f"警告: {msg}")

    mat = fitz.Matrix(zoom, zoom)
    # alpha=False使用RGB三通道，比RGBA节省约1/4内存
    return page.get_pixmap(matrix=mat, alpha=False)
