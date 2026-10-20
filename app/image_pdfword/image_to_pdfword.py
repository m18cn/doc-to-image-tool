#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片转PDF、Word转换器模块

该模块提供了将单张或多张图片转换为PDF文档和Word文档的功能，
支持常见的图片格式（JPG、PNG、BMP等），并提供了灵活的参数配置。
"""

import os
import logging
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.utils import ImageReader
from docx import Document
from docx.shared import Inches

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ImageConverter:
    """
    图片转换器类，提供图片转PDF和Word的功能
    
    该类封装了图片转换的核心逻辑，支持批量转换、页面尺寸设置、
    图片居中对齐等功能。
    """
    
    def __init__(self, output_dir="output"):
        """
        初始化图片转换器
        
        Args:
            output_dir (str): 输出文件目录，默认值为"output"
        """
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
    
    def images_to_pdf(self, image_paths, output_filename="output.pdf", pagesize=A4, margin=50):
        """
        将多张图片转换为PDF文档
        
        Args:
            image_paths (list): 图片文件路径列表
            output_filename (str): 输出PDF文件名，默认值为"output.pdf"
            pagesize (tuple): 页面尺寸，默认值为A4
            margin (int): 页面边距，默认值为50
            
        Returns:
            str: 生成的PDF文件路径
            
        Raises:
            ValueError: 图片路径为空或无效
            Exception: 转换过程中发生错误
        """
        if not image_paths:
            raise ValueError("图片路径列表不能为空")
        
        try:
            output_path = os.path.join(self.output_dir, output_filename)
            c = canvas.Canvas(output_path, pagesize=pagesize)
            page_width, page_height = pagesize
            
            for image_path in image_paths:
                if not os.path.exists(image_path):
                    logger.warning(f"图片不存在: {image_path}")
                    continue
                
                try:
                    # 打开图片并获取尺寸
                    img = Image.open(image_path)
                    img_width, img_height = img.size
                    
                    # 计算缩放比例，保持图片比例并适应页面
                    max_width = page_width - 2 * margin
                    max_height = page_height - 2 * margin
                    
                    scale_width = max_width / img_width
                    scale_height = max_height / img_height
                    scale = min(scale_width, scale_height)
                    
                    new_width = img_width * scale
                    new_height = img_height * scale
                    
                    # 计算居中位置
                    x = (page_width - new_width) / 2
                    y = (page_height - new_height) / 2
                    
                    # 绘制图片
                    c.drawImage(ImageReader(img), x, y, width=new_width, height=new_height)
                    c.showPage()  # 新建一页
                    logger.info(f"已添加图片到PDF: {image_path}")
                    
                except Exception as e:
                    logger.error(f"处理图片失败 {image_path}: {str(e)}")
                    continue
            
            c.save()
            logger.info(f"PDF文件已生成: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"转换为PDF失败: {str(e)}")
            raise
    
    def images_to_word(self, image_paths, output_filename="output.docx", image_width=Inches(6)):
        """
        将多张图片转换为Word文档
        
        Args:
            image_paths (list): 图片文件路径列表
            output_filename (str): 输出Word文件名，默认值为"output.docx"
            image_width (float): 图片宽度（英寸），默认值为6英寸
            
        Returns:
            str: 生成的Word文件路径
            
        Raises:
            ValueError: 图片路径为空或无效
            Exception: 转换过程中发生错误
        """
        if not image_paths:
            raise ValueError("图片路径列表不能为空")
        
        try:
            output_path = os.path.join(self.output_dir, output_filename)
            doc = Document()
            
            for i, image_path in enumerate(image_paths):
                if not os.path.exists(image_path):
                    logger.warning(f"图片不存在: {image_path}")
                    continue
                
                try:
                    # 添加图片到Word文档
                    doc.add_picture(image_path, width=image_width)
                    
                    # 除了最后一张图片，添加分页符
                    if i < len(image_paths) - 1:
                        doc.add_page_break()
                    
                    logger.info(f"已添加图片到Word: {image_path}")
                    
                except Exception as e:
                    logger.error(f"处理图片失败 {image_path}: {str(e)}")
                    continue
            
            doc.save(output_path)
            logger.info(f"Word文件已生成: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"转换为Word失败: {str(e)}")
            raise

if __name__ == "__main__":
    # 示例用法
    converter = ImageConverter()
    
    # 示例图片路径（请替换为实际图片路径）
    sample_images = [
        # "example1.jpg",
        # "example2.png"
    ]
    
    if sample_images:
        # 转换为PDF
        converter.images_to_pdf(sample_images, "output.pdf")
        
        # 转换为Word
        converter.images_to_word(sample_images, "output.docx")
    else:
        print("请在sample_images列表中添加实际图片路径")