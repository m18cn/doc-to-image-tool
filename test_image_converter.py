#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片转换器测试脚本

该脚本用于测试图片转PDF和Word功能的完整性和稳定性
"""

import os
import sys
from app.image_pdfword.image_to_pdfword import ImageConverter

def test_image_converter():
    """
    测试图片转换器功能
    """
    # 创建转换器实例
    converter = ImageConverter(output_dir="test_output")
    
    # 检查是否有测试图片
    test_images = []
    
    # 查找当前目录下的图片文件
    for file in os.listdir('.'):
        if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
            test_images.append(file)
    
    if not test_images:
        print("未找到测试图片，请将测试图片放在当前目录下")
        print("支持的图片格式: PNG, JPG, JPEG, BMP, GIF")
        return
    
    print(f"找到 {len(test_images)} 张测试图片")
    for img in test_images:
        print(f"  - {img}")
    
    try:
        # 测试图片转PDF
        print("\n正在测试图片转PDF...")
        pdf_path = converter.images_to_pdf(test_images, "test_output.pdf")
        print(f"PDF生成成功: {pdf_path}")
        
        # 测试图片转Word
        print("\n正在测试图片转Word...")
        word_path = converter.images_to_word(test_images, "test_output.docx")
        print(f"Word生成成功: {word_path}")
        
        print("\n测试完成！所有功能正常工作。")
        
    except Exception as e:
        print(f"\n测试失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_image_converter()