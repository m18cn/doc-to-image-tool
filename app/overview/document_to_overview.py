"""
通用文档转图片并拼接成鸟瞰图工具
功能：支持PPT、Word、PDF等多种文档格式，将每一页转换为图片并拼接成鸟瞰图
"""

import os
import sys
import time
import logging
from pathlib import Path
from PIL import Image
from typing import List, Optional, Callable
from app.logger_config import get_logger
from app.pdf_render import render_page_safely
from app.office_convert import convert_to_pdf

# Windows专属依赖：COM自动化仅在Windows可用，非Windows平台走LibreOffice路径
IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    import win32com.client
    import pythoncom

# 尝试导入PDF处理库
try:
    import fitz  # PyMuPDF
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# 尝试导入Word处理库
try:
    from docx import Document
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False


class DocumentToOverviewConverter:
    """通用文档转图片并拼接成鸟瞰图的转换器"""
    
    # 支持的文件格式
    SUPPORTED_FORMATS = {
        '.pptx': 'PPT',
        '.ppt': 'PPT',
        '.docx': 'Word',
        '.doc': 'Word',
        '.pdf': 'PDF'
    }
    
    def __init__(self, document_path: str, output_dir: str = "output", log_callback: Optional[Callable] = None):
        """
        初始化转换器
        
        Args:
            document_path: 文档文件路径（支持PPT、Word、PDF）
            output_dir: 输出目录
            log_callback: 日志回调函数，用于输出日志信息（GUI界面使用）
        """
        self.document_path = Path(document_path)
        self.output_dir = Path(output_dir)
        self.temp_images_dir = self.output_dir / "temp_images"
        self.log_callback = log_callback
        # 初始化日志系统
        self.logger = get_logger("document_converter")
        
        # 验证文件是否存在
        if not self.document_path.exists():
            raise FileNotFoundError(f"文档文件不存在: {self.document_path}")
        
        # 检查文件格式
        self.file_ext = self.document_path.suffix.lower()
        if self.file_ext not in self.SUPPORTED_FORMATS:
            supported = ', '.join(self.SUPPORTED_FORMATS.keys())
            raise ValueError(f"不支持的文件格式: {self.file_ext}\n支持的格式: {supported}")
        
        self.doc_type = self.SUPPORTED_FORMATS[self.file_ext]
        
        # 检查必要的库
        if self.doc_type == 'PDF' and not PDF_SUPPORT:
            raise ImportError("PDF支持需要安装PyMuPDF库: pip install PyMuPDF")
        
        # 创建输出目录
        self.output_dir.mkdir(exist_ok=True)
        self.temp_images_dir.mkdir(exist_ok=True)
    
    def _log(self, message: str, level: str = "INFO"):
        """
        输出日志
        
        Args:
            message: 日志消息
            level: 日志级别 (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        """
        # 输出到日志文件
        log_level = getattr(logging, level.upper(), logging.INFO)
        self.logger.log(log_level, message)
        
        # 如果有回调函数（GUI界面），也输出到回调
        if self.log_callback:
            self.log_callback(message)
    
    def _render_pdf_file(self, pdf_path: Path, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """
        将PDF文件的每一页渲染为图片（跨平台共用渲染管线）

        Args:
            pdf_path: PDF文件路径
            image_format: 图片格式 (PNG, JPG等)
            dpi: 图片分辨率，默认300

        Returns:
            生成的图片文件路径列表
        """
        if not PDF_SUPPORT:
            raise ImportError("需要PyMuPDF库: pip install PyMuPDF")

        pdf_doc = fitz.open(str(pdf_path))
        page_count = len(pdf_doc)

        self._log(f"开始渲染PDF共 {page_count} 页 (DPI: {dpi})")

        image_paths = []
        for page_num in range(page_count):
            page = pdf_doc[page_num]
            image_path = self.temp_images_dir / f"page_{page_num+1:03d}.{image_format.lower()}"

            # 渲染页面为图片（含超大页面自动降级保护）
            pix = render_page_safely(page, dpi, log_callback=self.log_callback)
            pix.save(str(image_path))

            image_paths.append(str(image_path))
            self._log(f"已导出第 {page_num+1}/{page_count} 页: {image_path.name}")

        pdf_doc.close()
        self._log(f"成功导出 {len(image_paths)} 张图片")
        return image_paths

    def _ppt_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """将PPT转换为图片（Windows走PowerPoint COM，其他平台走LibreOffice）"""
        self._log(f"正在将PPT转换为图片: {self.document_path}")

        # 非Windows平台：LibreOffice转PDF后复用渲染管线
        if not IS_WINDOWS:
            pdf_path = convert_to_pdf(self.document_path, self.output_dir, log_callback=self.log_callback)
            image_paths = self._render_pdf_file(pdf_path, image_format, dpi)
            # 清理中间PDF
            try:
                os.remove(str(pdf_path))
            except Exception as e:
                self._log(f"删除中间PDF文件失败: {e}")
            return image_paths

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        
        try:
            presentation = powerpoint.Presentations.Open(str(self.document_path.absolute()))
            image_paths = []
            slide_count = presentation.Slides.Count
            
            self._log(f"PPT共有 {slide_count} 页")
            
            slide_width = int(presentation.PageSetup.SlideWidth)
            slide_height = int(presentation.PageSetup.SlideHeight)
            
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                image_path = self.temp_images_dir / f"page_{i:03d}.{image_format.lower()}"
                
                slide.Export(
                    str(image_path.absolute()),
                    image_format,
                    slide_width,
                    slide_height
                )
                
                image_paths.append(str(image_path))
                self._log(f"已导出第 {i}/{slide_count} 页: {image_path.name}")
            
            presentation.Close()
            self._log(f"成功导出 {len(image_paths)} 张图片")
            return image_paths
            
        except Exception as e:
            self._log(f"转换过程中出错: {str(e)}")
            raise
        finally:
            powerpoint.Quit()
    
    def _word_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """将Word文档转换为图片（Windows走Word COM，其他平台走LibreOffice）"""
        self._log(f"正在将Word文档转换为图片: {self.document_path}")

        # 非Windows平台：LibreOffice转PDF后复用渲染管线
        if not IS_WINDOWS:
            pdf_path = convert_to_pdf(self.document_path, self.output_dir, log_callback=self.log_callback)
            image_paths = self._render_pdf_file(pdf_path, image_format, dpi)
            # 清理中间PDF
            try:
                os.remove(str(pdf_path))
            except Exception as e:
                self._log(f"删除中间PDF文件失败: {e}")
            return image_paths

        # 初始化COM接口
        pythoncom.CoInitialize()
        word = None
        doc = None
        
        try:
            # 尝试启动Word应用程序，添加重试机制
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    self._log(f"正在启动Word应用程序 (尝试 {attempt + 1}/{max_retries})...")
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = 0  # 不显示Word窗口
                    word.DisplayAlerts = 0  # 不显示警告
                    # 等待Word完全启动
                    time.sleep(1)
                    # 测试Word是否可用
                    _ = word.Version
                    self._log("Word应用程序启动成功")
                    break
                except Exception as e:
                    if attempt < max_retries - 1:
                        self._log(f"启动Word失败，等待后重试: {str(e)}")
                        time.sleep(2)
                    else:
                        raise Exception(f"无法启动Word应用程序: {str(e)}")
            
            if word is None:
                raise Exception("Word应用程序启动失败")
            
            # 打开Word文档
            self._log("正在打开Word文档...")
            doc_path = str(self.document_path.absolute())
            
            # 尝试打开文档，添加重试
            for attempt in range(max_retries):
                try:
                    doc = word.Documents.Open(
                        FileName=doc_path,
                        ReadOnly=True,
                        ConfirmConversions=False,
                        AddToRecentFiles=False
                    )
                    break
                except Exception as e:
                    if attempt < max_retries - 1:
                        self._log(f"打开文档失败，等待后重试: {str(e)}")
                        time.sleep(1)
                    else:
                        raise Exception(f"无法打开Word文档: {str(e)}")
            
            if doc is None:
                raise Exception("无法打开Word文档")
            
            # 先导出整个文档为PDF（更可靠的方法）
            temp_pdf = self.temp_images_dir / "temp_word_export.pdf"
            self._log("正在将Word文档导出为PDF...")
            
            try:
                # 使用ExportAsFixedFormat导出为PDF
                doc.ExportAsFixedFormat(
                    OutputFileName=str(temp_pdf.absolute()),
                    ExportFormat=17,  # wdExportFormatPDF = 17
                    OpenAfterExport=False,
                    OptimizeFor=0,  # wdExportOptimizeForPrint = 0
                    BitmapMissingFonts=True,
                    DocStructureTags=True,
                    CreateBookmarks=0,  # wdExportCreateNoBookmarks = 0
                    UseISO19005_1=False
                )
                # 等待PDF生成完成
                time.sleep(1)
                if not temp_pdf.exists():
                    raise Exception("PDF文件未生成")
            except Exception as e:
                # 如果ExportAsFixedFormat失败，尝试使用SaveAs2方法
                self._log(f"ExportAsFixedFormat失败，尝试备用方法: {str(e)}")
                try:
                    # 尝试另存为PDF
                    doc.SaveAs2(
                        FileName=str(temp_pdf.absolute()),
                        FileFormat=17  # wdFormatPDF = 17
                    )
                    time.sleep(1)
                    if not temp_pdf.exists():
                        raise Exception("PDF文件未生成")
                except Exception as e2:
                    raise Exception(f"无法将Word文档转换为PDF: {str(e2)}")
            
            # 关闭文档和Word
            try:
                if doc:
                    doc.Close(SaveChanges=False)
                    doc = None
            except:
                pass
            
            try:
                if word:
                    word.Quit(SaveChanges=False)
                    word = None
            except:
                pass
            
            # 等待Word完全关闭
            time.sleep(1)
            
            # 从PDF提取所有页面为图片
            if not PDF_SUPPORT:
                raise ImportError("需要PyMuPDF库来处理Word文档: pip install PyMuPDF")
            
            self._log("正在从PDF提取页面为图片...")
            pdf_doc = fitz.open(str(temp_pdf.absolute()))
            actual_page_count = len(pdf_doc)
            
            self._log(f"Word文档共有 {actual_page_count} 页")
            
            image_paths = []
            for page_num in range(actual_page_count):
                page = pdf_doc[page_num]
                image_path = self.temp_images_dir / f"page_{page_num+1:03d}.{image_format.lower()}"

                # 渲染页面为图片（含超大页面自动降级保护）
                pix = render_page_safely(page, dpi, log_callback=self.log_callback)
                pix.save(str(image_path))
                
                image_paths.append(str(image_path))
                self._log(f"已导出第 {page_num+1}/{actual_page_count} 页: {image_path.name}")
            
            pdf_doc.close()
            
            # 删除临时PDF文件
            try:
                os.remove(str(temp_pdf.absolute()))
            except Exception as e:
                self._log(f"删除临时PDF文件失败: {e}")
            
            self._log(f"成功导出 {len(image_paths)} 张图片")
            return image_paths
            
        except Exception as e:
            self._log(f"转换过程中出错: {str(e)}")
            raise
        finally:
            # 确保清理资源
            try:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
            except:
                pass
            
            try:
                if word:
                    try:
                        word.Quit(SaveChanges=False)
                    except:
                        pass
            except:
                pass
            
            # 释放COM接口
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _pdf_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """将PDF转换为图片"""
        if not PDF_SUPPORT:
            raise ImportError("PDF支持需要安装PyMuPDF库: pip install PyMuPDF")

        self._log(f"正在将PDF转换为图片: {self.document_path}")

        try:
            # 复用跨平台渲染管线（含超大页面自动降级保护）
            return self._render_pdf_file(self.document_path, image_format, dpi)

        except Exception as e:
            self._log(f"转换过程中出错: {str(e)}")
            raise
    
    def document_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """
        将文档转换为图片（根据文档类型自动选择方法）
        
        Args:
            image_format: 图片格式 (PNG, JPG等)
            dpi: 图片分辨率，默认300
            
        Returns:
            生成的图片文件路径列表
        """
        if self.doc_type == 'PPT':
            return self._ppt_to_images(image_format, dpi)
        elif self.doc_type == 'Word':
            return self._word_to_images(image_format, dpi)
        elif self.doc_type == 'PDF':
            return self._pdf_to_images(image_format, dpi)
        else:
            raise ValueError(f"不支持的文档类型: {self.doc_type}")
    
    def merge_images_to_overview(
        self, 
        image_paths: List[str], 
        output_filename: str = "overview.png",
        layout: str = "vertical",
        spacing: int = 10,
        background_color: tuple = (255, 255, 255)
    ) -> str:
        """
        将多张图片拼接成鸟瞰图
        
        Args:
            image_paths: 图片路径列表
            layout: 布局方式 ("vertical"垂直, "horizontal"水平, "grid"网格)
            spacing: 图片之间的间距（像素）
            background_color: 背景颜色 RGB元组
            output_filename: 输出文件名
            
        Returns:
            输出文件路径
        """
        if not image_paths:
            raise ValueError("图片列表为空")
        
        self._log(f"正在拼接 {len(image_paths)} 张图片...")
        
        # 加载所有图片
        images = []
        max_width = 0
        max_height = 0
        total_width = 0
        total_height = 0
        
        for img_path in image_paths:
            img = Image.open(img_path)
            images.append(img)
            
            if layout == "vertical":
                max_width = max(max_width, img.width)
                total_height += img.height
            elif layout == "horizontal":
                max_height = max(max_height, img.height)
                total_width += img.width
            else:  # grid
                max_width = max(max_width, img.width)
                max_height = max(max_height, img.height)
        
        # 计算网格布局
        if layout == "grid":
            import math
            cols = math.ceil(math.sqrt(len(images)))
            rows = math.ceil(len(images) / cols)
            total_width = cols * max_width + (cols - 1) * spacing
            total_height = rows * max_height + (rows - 1) * spacing
        
        # 设置最终尺寸
        if layout == "vertical":
            final_width = max_width
            final_height = total_height + (len(images) - 1) * spacing
        elif layout == "horizontal":
            final_width = total_width + (len(images) - 1) * spacing
            final_height = max_height
        else:  # grid
            final_width = total_width
            final_height = total_height
        
        # 创建画布
        overview = Image.new("RGB", (final_width, final_height), background_color)
        
        # 拼接图片
        x_offset = 0
        y_offset = 0
        
        if layout == "vertical":
            for img in images:
                x_pos = (final_width - img.width) // 2
                overview.paste(img, (x_pos, y_offset))
                y_offset += img.height + spacing
                
        elif layout == "horizontal":
            for img in images:
                y_pos = (final_height - img.height) // 2
                overview.paste(img, (x_offset, y_pos))
                x_offset += img.width + spacing
                
        else:  # grid
            for idx, img in enumerate(images):
                row = idx // cols
                col = idx % cols
                x_pos = col * (max_width + spacing)
                y_pos = row * (max_height + spacing)
                overview.paste(img, (x_pos, y_pos))
        
        # 保存鸟瞰图
        output_path = self.output_dir / output_filename
        # PNG格式不支持quality参数，使用compress_level或直接保存
        try:
            overview.save(output_path, "PNG", compress_level=6)
        except TypeError:
            # 如果compress_level不支持，直接保存
            overview.save(output_path, "PNG")
        
        self._log(f"鸟瞰图已保存: {output_path}")
        return str(output_path)
    
    def convert(
        self,
        image_format: str = "PNG",
        dpi: int = 300,
        layout: str = "vertical",
        spacing: int = 10,
        output_filename: str = None,
        cleanup_temp: bool = True
    ) -> str:
        """
        完整转换流程：文档转图片 -> 拼接成鸟瞰图
        
        Args:
            image_format: 图片格式
            dpi: 分辨率
            layout: 布局方式
            spacing: 间距
            output_filename: 输出文件名，如果为None则使用文档文件名
            cleanup_temp: 是否清理临时图片文件
            
        Returns:
            最终输出文件路径
        """
        try:
            # 如果没有指定输出文件名，使用文档文件名
            if output_filename is None:
                output_filename = f"{self.document_path.stem}.png"
            
            # 步骤1: 文档转图片
            image_paths = self.document_to_images(image_format, dpi)
            
            # 步骤2: 拼接成鸟瞰图
            overview_path = self.merge_images_to_overview(
                image_paths,
                output_filename,
                layout,
                spacing
            )
            
            # 步骤3: 清理临时文件（可选）
            if cleanup_temp:
                self._log("正在清理临时文件...")
                for img_path in image_paths:
                    try:
                        os.remove(img_path)
                    except Exception as e:
                        self._log(f"删除临时文件失败 {img_path}: {e}")
                
                try:
                    self.temp_images_dir.rmdir()
                except Exception as e:
                    self._log(f"删除临时目录失败: {e}")
            
            self._log(f"\n转换完成！鸟瞰图保存在: {overview_path}")
            return overview_path
            
        except Exception as e:
            self._log(f"转换失败: {str(e)}")
            raise


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法: python document_to_overview.py <文档文件路径> [输出目录] [布局方式]")
        print("支持的格式: .pptx, .ppt, .docx, .doc, .pdf")
        print("布局方式: vertical(垂直,默认), horizontal(水平), grid(网格)")
        print("示例: python document_to_overview.py document.pdf output vertical")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    layout = sys.argv[3] if len(sys.argv) > 3 else "vertical"
    
    if layout not in ["vertical", "horizontal", "grid"]:
        print("错误: 布局方式必须是 vertical, horizontal 或 grid")
        sys.exit(1)
    
    try:
        converter = DocumentToOverviewConverter(doc_path, output_dir)
        doc_name = Path(doc_path).stem
        converter.convert(layout=layout, output_filename=f"{doc_name}.png")
    except Exception as e:
        print(f"错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

