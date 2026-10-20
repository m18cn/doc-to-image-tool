"""
文档转图片工具
功能：将PPT、Word、PDF文档的每一页转换为独立的图片文件
例如：10页文档会生成10个独立的图片文件
"""

import os
import sys
import time
import logging
from pathlib import Path
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


class DocumentToImagesConverter:
    """文档转图片转换器（每页生成独立图片）"""
    
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
        self.log_callback = log_callback
        # 初始化日志系统
        self.logger = get_logger("document_to_images")
        
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
    
    def _render_pdf_file(self, pdf_path: Path, image_format: str = "PNG", dpi: int = 300, name_prefix: str = "page") -> List[str]:
        """
        将PDF文件的每一页渲染为独立图片（跨平台共用渲染管线）

        Args:
            pdf_path: PDF文件路径
            image_format: 图片格式 (PNG, JPG等)
            dpi: 图片分辨率，默认300
            name_prefix: 输出文件名前缀

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
            # 每页生成一个独立的图片文件
            image_path = self.output_dir / f"{name_prefix}_{page_num+1:03d}.{image_format.lower()}"

            # 渲染页面为图片（含超大页面自动降级保护）
            pix = render_page_safely(page, dpi, log_callback=self.log_callback)
            pix.save(str(image_path))

            image_paths.append(str(image_path))
            self._log(f"已导出第 {page_num+1}/{page_count} 页: {image_path.name}")

        pdf_doc.close()
        self._log(f"成功导出 {len(image_paths)} 张图片")
        return image_paths

    def _ppt_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """将PPT转换为图片（每页一个文件；Windows走PowerPoint COM，其他平台走LibreOffice）"""
        self._log(f"正在将PPT转换为图片: {self.document_path}")

        # 非Windows平台：LibreOffice转PDF后复用渲染管线
        if not IS_WINDOWS:
            pdf_path = convert_to_pdf(self.document_path, self.output_dir, log_callback=self.log_callback)
            image_paths = self._render_pdf_file(pdf_path, image_format, dpi, name_prefix=self.document_path.stem)
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
            
            # 生成输出文件名前缀
            doc_name = self.document_path.stem
            
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                # 每页生成一个独立的图片文件
                image_path = self.output_dir / f"{doc_name}_page_{i:03d}.{image_format.lower()}"
                
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
            self._log(f"转换过程中出错: {str(e)}", "ERROR")
            raise
        finally:
            powerpoint.Quit()
    
    def _word_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """将Word文档转换为图片（每页一个文件；Windows走Word COM，其他平台走LibreOffice）"""
        self._log(f"正在将Word文档转换为图片: {self.document_path}")

        # 非Windows平台：LibreOffice转PDF后复用渲染管线
        if not IS_WINDOWS:
            pdf_path = convert_to_pdf(self.document_path, self.output_dir, log_callback=self.log_callback)
            image_paths = self._render_pdf_file(pdf_path, image_format, dpi, name_prefix=self.document_path.stem)
            # 清理中间PDF
            try:
                os.remove(str(pdf_path))
            except Exception as e:
                self._log(f"删除中间PDF文件失败: {e}")
            return image_paths

        if not PDF_SUPPORT:
            raise ImportError("Word文档转换需要PyMuPDF库: pip install PyMuPDF")
        
        # 初始化COM接口
        pythoncom.CoInitialize()
        word = None
        doc = None
        
        try:
            # 尝试启动Word应用程序
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    self._log(f"正在启动Word应用程序 (尝试 {attempt + 1}/{max_retries})...")
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = 0
                    word.DisplayAlerts = 0
                    time.sleep(1)
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
            
            # 导出整个文档为PDF
            temp_pdf = self.output_dir / "temp_word_export.pdf"
            self._log("正在将Word文档导出为PDF...")
            
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(temp_pdf.absolute()),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    BitmapMissingFonts=True,
                    DocStructureTags=True,
                    CreateBookmarks=0,
                    UseISO19005_1=False
                )
                time.sleep(1)
                if not temp_pdf.exists():
                    raise Exception("PDF文件未生成")
            except Exception as e:
                self._log(f"ExportAsFixedFormat失败，尝试备用方法: {str(e)}")
                try:
                    doc.SaveAs2(
                        FileName=str(temp_pdf.absolute()),
                        FileFormat=17
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
            
            time.sleep(1)
            
            # 从PDF提取所有页面为图片
            self._log("正在从PDF提取页面为图片...")
            pdf_doc = fitz.open(str(temp_pdf.absolute()))
            page_count = len(pdf_doc)
            
            self._log(f"Word文档共有 {page_count} 页")
            
            # 生成输出文件名前缀
            doc_name = self.document_path.stem
            
            image_paths = []
            for page_num in range(page_count):
                page = pdf_doc[page_num]
                # 每页生成一个独立的图片文件
                image_path = self.output_dir / f"{doc_name}_page_{page_num+1:03d}.{image_format.lower()}"

                # 渲染页面为图片（含超大页面自动降级保护）
                pix = render_page_safely(page, dpi, log_callback=self.log_callback)
                pix.save(str(image_path))
                
                image_paths.append(str(image_path))
                self._log(f"已导出第 {page_num+1}/{page_count} 页: {image_path.name}")
            
            pdf_doc.close()
            
            # 删除临时PDF文件
            try:
                os.remove(str(temp_pdf.absolute()))
            except Exception as e:
                self._log(f"删除临时PDF文件失败: {e}")
            
            self._log(f"成功导出 {len(image_paths)} 张图片")
            return image_paths
            
        except Exception as e:
            self._log(f"转换过程中出错: {str(e)}", "ERROR")
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
            
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _pdf_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """将PDF转换为图片（每页一个文件）"""
        if not PDF_SUPPORT:
            raise ImportError("PDF支持需要安装PyMuPDF库: pip install PyMuPDF")

        self._log(f"正在将PDF转换为图片: {self.document_path}")

        try:
            # 复用跨平台渲染管线（含超大页面自动降级保护）
            name_prefix = f"{self.document_path.stem}_page"
            return self._render_pdf_file(self.document_path, image_format, dpi, name_prefix=name_prefix)

        except Exception as e:
            self._log(f"转换过程中出错: {str(e)}", "ERROR")
            raise
    
    def convert(
        self,
        image_format: str = "PNG",
        dpi: int = 300
    ) -> List[str]:
        """
        将文档转换为图片（每页生成一个独立的图片文件）
        
        Args:
            image_format: 图片格式 (PNG, JPG等)
            dpi: 图片分辨率，默认300
            
        Returns:
            生成的图片文件路径列表
        """
        try:
            self._log("=" * 50)
            self._log(f"开始转换文档: {self.document_path.name}")
            self._log(f"文档类型: {self.doc_type}")
            self._log(f"输出目录: {self.output_dir}")
            self._log(f"图片格式: {image_format}")
            self._log(f"DPI: {dpi}")
            self._log("=" * 50)
            
            # 根据文档类型选择转换方法
            if self.doc_type == 'PPT':
                image_paths = self._ppt_to_images(image_format, dpi)
            elif self.doc_type == 'Word':
                image_paths = self._word_to_images(image_format, dpi)
            elif self.doc_type == 'PDF':
                image_paths = self._pdf_to_images(image_format, dpi)
            else:
                raise ValueError(f"不支持的文档类型: {self.doc_type}")
            
            self._log("=" * 50)
            self._log(f"转换完成！共生成 {len(image_paths)} 张图片")
            self._log(f"输出目录: {self.output_dir}")
            self._log("=" * 50)
            
            return image_paths
            
        except Exception as e:
            self._log(f"转换失败: {str(e)}", "ERROR")
            raise


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法: python document_to_images.py <文档文件路径> [输出目录] [DPI]")
        print("支持的格式: .pptx, .ppt, .docx, .doc, .pdf")
        print("示例: python document_to_images.py document.pdf output 300")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 300
    
    try:
        converter = DocumentToImagesConverter(doc_path, output_dir)
        image_paths = converter.convert(dpi=dpi)
        
        print(f"\n转换完成！共生成 {len(image_paths)} 张图片")
        print(f"输出目录: {output_dir}")
        print("\n生成的图片文件:")
        for img_path in image_paths:
            print(f"  - {img_path}")
            
    except Exception as e:
        print(f"错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

