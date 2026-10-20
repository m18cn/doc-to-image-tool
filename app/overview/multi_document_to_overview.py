"""
多文件文档转图片并拼接成鸟瞰图工具
功能：支持选择多个文档文件，将每个文件的每一页转换为图片，添加数字标签，最后拼接成一张鸟瞰图
"""

import os
import sys
import time
import logging
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from typing import List, Optional, Callable, Dict, Tuple
from app.overview.document_to_overview import DocumentToOverviewConverter
from app.logger_config import get_logger


class MultiDocumentToOverviewConverter:
    """多文件文档转图片并拼接成鸟瞰图的转换器"""
    
    def __init__(self, document_paths: List[str], output_dir: str = "output", log_callback: Optional[Callable] = None, temp_dir_name: str = "temp_images"):
        """
        初始化转换器
        
        Args:
            document_paths: 文档文件路径列表（支持PPT、Word、PDF）
            output_dir: 输出目录
            log_callback: 日志回调函数，用于输出日志信息（GUI界面使用）
            temp_dir_name: 临时图片目录名称，默认为"temp_images"，可以设置为"temp_images_1"、"temp_images_2"等
        """
        self.document_paths = [Path(p) for p in document_paths]
        self.output_dir = Path(output_dir)
        self.temp_images_dir = self.output_dir / temp_dir_name
        self.log_callback = log_callback
        # 初始化日志系统
        self.logger = get_logger("multi_document_converter")
        
        # 验证所有文件是否存在
        for doc_path in self.document_paths:
            if not doc_path.exists():
                raise FileNotFoundError(f"文档文件不存在: {doc_path}")
        
        # 创建输出目录
        self.output_dir.mkdir(exist_ok=True)
        self.temp_images_dir.mkdir(exist_ok=True)
        
        self._log(f"使用临时图片目录: {self.temp_images_dir}")
        
        # 存储所有生成的图片路径和标签信息
        self.all_image_paths: List[Tuple[str, int, str]] = []  # (图片路径, 标签编号, 源文件名)
    
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
    
    def convert_all_documents_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[Tuple[str, str]]:
        """
        将所有文档转换为图片，每个文件的每一页生成一个独立图片（不添加标签）
        
        Args:
            image_format: 图片格式 (PNG, JPG等)
            dpi: 图片分辨率，默认300
            
        Returns:
            生成的图片文件路径列表，每个元素为 (图片路径, 源文件名)
        """
        self.all_image_paths = []
        
        for doc_idx, doc_path in enumerate(self.document_paths):
            self._log(f"\n处理文件 {doc_idx + 1}/{len(self.document_paths)}: {doc_path.name}")
            
            try:
                # 为每个文档创建转换器
                converter = DocumentToOverviewConverter(
                    str(doc_path),
                    str(self.temp_images_dir),
                    log_callback=self.log_callback
                )
                
                # 将文档转换为图片（不添加标签）
                image_paths = converter.document_to_images(image_format, dpi)
                
                # 保存图片信息（不添加标签），确保使用绝对路径
                for img_path in image_paths:
                    img_path_abs = str(Path(img_path).absolute())
                    self.all_image_paths.append((img_path_abs, doc_path.stem))
                
                self._log(f"文件 {doc_path.name} 处理完成，共 {len(image_paths)} 页，已添加到图片列表")
                
            except Exception as e:
                self._log(f"处理文件 {doc_path.name} 时出错: {str(e)}", "ERROR")
                raise
        
        self._log(f"\n所有文件处理完成，共生成 {len(self.all_image_paths)} 张图片")
        return self.all_image_paths
    
    def add_labels_to_selected_images(self, selected_image_paths: List[str]) -> List[Tuple[str, int, str]]:
        """
        为选中的图片按顺序添加标签
        
        Args:
            selected_image_paths: 选中的图片路径列表（按用户选择的顺序）
            
        Returns:
            带标签的图片列表，每个元素为 (图片路径, 标签编号, 源文件名)
        """
        labeled_images = []
        
        # 创建路径到源文件名的映射（使用绝对路径）
        path_to_source = {}
        for path, source in self.all_image_paths:
            # path可能已经是绝对路径，也可能不是，统一处理
            path_abs = str(Path(path).absolute())
            path_to_source[path_abs] = source
        
        self._log(f"准备为 {len(selected_image_paths)} 张选中的图片添加标签")
        self._log(f"转换器中共有 {len(self.all_image_paths)} 张图片（但只处理选中的 {len(selected_image_paths)} 张）")
        
        # 验证：确保只处理传入的选中图片
        processed_count = 0
        for label_number, img_path in enumerate(selected_image_paths, start=1):
            # 确保使用绝对路径
            img_path_abs = str(Path(img_path).absolute())
            source_name = path_to_source.get(img_path_abs, "未知")
            
            # 如果找不到源文件名，尝试在all_image_paths中查找（处理路径格式差异）
            if source_name == "未知":
                for path, source in self.all_image_paths:
                    path_abs_check = str(Path(path).absolute())
                    if path_abs_check == img_path_abs:
                        source_name = source
                        break
            
            # 检查文件是否存在
            if not os.path.exists(img_path_abs):
                self._log(f"警告: 图片文件不存在，跳过 {img_path_abs}", "WARNING")
                self._log(f"可用的图片路径示例: {list(path_to_source.keys())[:3] if path_to_source else '无'}")
                continue
            
            # 不添加标签到图片上，直接使用原图片（标签仅用于排序）
            # labeled_img_path = self._add_label_to_image(img_path_abs, label_number, source_name)
            # 直接使用原图片路径，不添加视觉标签
            labeled_img_path = img_path_abs
            labeled_images.append((labeled_img_path, label_number, source_name))
            processed_count += 1
            self._log(f"准备拼接图片 #{label_number}: {Path(img_path_abs).name} (来源: {source_name})")
        
        self._log(f"成功为 {processed_count}/{len(selected_image_paths)} 张选中的图片添加标签")
        
        if not labeled_images:
            self._log(f"错误: 没有成功添加标签的图片！", "ERROR")
            self._log(f"选中的路径: {selected_image_paths[:3]}...")
            self._log(f"转换器中的路径示例: {[str(Path(p).absolute()) for p, _ in list(self.all_image_paths)[:3]]}")
        
        return labeled_images
    
    def _add_label_to_image(self, image_path: str, label_number: int, source_filename: str) -> str:
        """
        给图片添加数字标签
        
        Args:
            image_path: 原始图片路径
            label_number: 标签编号
            source_filename: 源文件名（用于生成新文件名）
            
        Returns:
            添加标签后的图片路径
        """
        # 打开原始图片并转换为RGBA模式（支持透明度）
        img = Image.open(image_path)
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
        
        # 创建图片副本
        img_with_label = img.copy()
        
        # 尝试加载字体（如果系统有中文字体）
        try:
            # Windows系统常见字体路径
            font_paths = [
                "C:/Windows/Fonts/msyh.ttc",  # 微软雅黑
                "C:/Windows/Fonts/simhei.ttf",  # 黑体
                "C:/Windows/Fonts/simsun.ttc",  # 宋体
            ]
            font = None
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        font = ImageFont.truetype(font_path, 40)
                        break
                    except:
                        continue
            if font is None:
                font = ImageFont.load_default()
        except:
            font = ImageFont.load_default()
        
        # 创建临时绘图对象用于计算文本尺寸
        temp_draw = ImageDraw.Draw(img_with_label)
        
        # 标签文本
        label_text = str(label_number)
        
        # 计算文本尺寸
        bbox = temp_draw.textbbox((0, 0), label_text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        # 标签背景框尺寸
        padding = 10
        box_width = text_width + padding * 2
        box_height = text_height + padding * 2
        
        # 标签位置（左上角）
        box_x = 10
        box_y = 10
        
        # 创建绘图对象
        draw = ImageDraw.Draw(img_with_label)
        
        # 绘制标签背景（半透明红色背景）
        overlay = Image.new('RGBA', img_with_label.size, (0, 0, 0, 0))
        overlay_draw = ImageDraw.Draw(overlay)
        overlay_draw.rectangle(
            [box_x, box_y, box_x + box_width, box_y + box_height],
            fill=(255, 0, 0, 200)  # 红色半透明背景
        )
        img_with_label = Image.alpha_composite(img_with_label, overlay)
        
        # 重新创建绘图对象
        draw = ImageDraw.Draw(img_with_label)
        
        # 绘制标签文本（白色）
        text_x = box_x + padding
        text_y = box_y + padding
        draw.text(
            (text_x, text_y),
            label_text,
            fill=(255, 255, 255, 255),  # 白色文字，完全不透明
            font=font
        )
        
        # 转换回RGB模式（如果原始图片不是RGBA）
        if img.mode == 'RGB':
            img_with_label = img_with_label.convert('RGB')
        
        # 生成新的文件名（带标签编号）
        img_path_obj = Path(image_path)
        labeled_filename = f"labeled_{label_number:04d}_{source_filename}_{img_path_obj.name}"
        labeled_path = self.temp_images_dir / labeled_filename
        
        # 保存带标签的图片（PNG格式不支持quality参数，使用compress_level）
        img_with_label.save(labeled_path, "PNG", compress_level=6)
        
        return str(labeled_path)
    
    def merge_images_to_overview(
        self, 
        image_paths_with_labels: List[Tuple[str, int, str]], 
        output_filename: str = "multi_document_overview.png",
        layout: str = "vertical",
        spacing: int = 10,
        background_color: tuple = (255, 255, 255)
    ) -> str:
        """
        将多张图片拼接成鸟瞰图
        
        Args:
            image_paths_with_labels: 图片路径列表，每个元素为 (图片路径, 标签编号, 源文件名)
            layout: 布局方式 ("vertical"垂直, "horizontal"水平, "grid"网格)
            spacing: 图片之间的间距（像素）
            background_color: 背景颜色 RGB元组
            output_filename: 输出文件名
            
        Returns:
            输出文件路径
        """
        if not image_paths_with_labels:
            raise ValueError("图片列表为空")
        
        self._log(f"正在拼接 {len(image_paths_with_labels)} 张选中的图片（只处理勾选的图片）...")
        
        # 提取图片路径（忽略标签信息）
        image_paths = [item[0] for item in image_paths_with_labels]
        
        # 加载所有图片
        images = []
        max_width = 0
        max_height = 0
        total_width = 0
        total_height = 0
        
        for img_path in image_paths:
            # 确保使用绝对路径
            img_path_abs = Path(img_path).absolute()
            
            # 验证文件是否存在
            if not img_path_abs.exists():
                self._log(f"警告: 图片文件不存在，跳过: {img_path_abs}", "WARNING")
                continue
            
            try:
                img = Image.open(str(img_path_abs))
                # 转换为RGB模式以确保兼容性
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                images.append(img)
                self._log(f"已加载图片: {img_path_abs.name}")
            except Exception as e:
                self._log(f"加载图片失败 {img_path_abs}: {str(e)}", "ERROR")
                continue
            
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
        
        # 保存鸟瞰图到输出目录（不是临时目录）
        # 确保输出目录存在（使用绝对路径）
        output_dir_abs = Path(self.output_dir).absolute()
        output_dir_abs.mkdir(parents=True, exist_ok=True)
        
        # 构建最终输出路径（保存在输出目录根目录，不在临时子目录中）
        output_path = output_dir_abs / output_filename
        
        self._log(f"输出目录: {output_dir_abs}")
        self._log(f"最终图片将保存到: {output_path}")
        
        # 保存鸟瞰图（PNG格式不支持quality参数，使用compress_level）
        try:
            # 确保图片已正确创建
            if overview is None:
                raise ValueError("图片对象为空，无法保存")
            if overview.size[0] == 0 or overview.size[1] == 0:
                raise ValueError(f"图片尺寸无效: {overview.size}")
            
            self._log(f"准备保存图片，尺寸: {overview.size}, 路径: {output_path}")
            
            # 保存图片（尝试使用compress_level，如果不支持则使用默认）
            try:
                overview.save(str(output_path), "PNG", compress_level=6)
                self._log(f"图片已保存到: {output_path}")
            except TypeError:
                # 如果compress_level不支持，直接保存
                overview.save(str(output_path), "PNG")
                self._log(f"使用默认压缩级别保存成功: {output_path}")
            
        except Exception as e:
            self._log(f"保存图片时出错: {str(e)}", "ERROR")
            import traceback
            self._log(f"错误详情: {traceback.format_exc()}", "ERROR")
            raise
        
        # 验证文件是否成功保存
        if not output_path.exists():
            raise FileNotFoundError(f"输出文件保存失败: {output_path}")
        
        # 验证文件大小（确保不是空文件）
        file_size = output_path.stat().st_size
        if file_size == 0:
            raise ValueError(f"输出文件为空: {output_path}")
        
        self._log(f"鸟瞰图已保存: {output_path}")
        self._log(f"文件大小: {file_size / 1024 / 1024:.2f} MB")
        self._log(f"保存位置: 输出目录根目录（不在临时文件夹中）")
        return str(output_path)
    
    def convert(
        self,
        image_format: str = "PNG",
        dpi: int = 300,
        layout: str = "vertical",
        spacing: int = 10,
        output_filename: str = None,
        cleanup_temp: bool = True
    ) -> Tuple[str, List[Tuple[str, int, str]]]:
        """
        完整转换流程：多文件转图片 -> 添加标签 -> 拼接成鸟瞰图
        
        Args:
            image_format: 图片格式
            dpi: 分辨率
            layout: 布局方式
            spacing: 间距
            output_filename: 输出文件名，如果为None则使用默认名称
            cleanup_temp: 是否清理临时图片文件
            
        Returns:
            (最终输出文件路径, 所有图片信息列表)
        """
        try:
            # 如果没有指定输出文件名，使用默认名称
            if output_filename is None:
                output_filename = "multi_document_overview.png"
            
            # 步骤1: 将所有文档转换为图片并添加标签
            image_paths_with_labels = self.convert_all_documents_to_images(image_format, dpi)
            
            # 步骤2: 拼接成鸟瞰图
            overview_path = self.merge_images_to_overview(
                image_paths_with_labels,
                output_filename,
                layout,
                spacing
            )
            
            # 步骤3: 清理临时文件（可选）
            if cleanup_temp:
                self._log("正在清理临时文件...")
                for img_path, _, _ in image_paths_with_labels:
                    try:
                        os.remove(img_path)
                    except Exception as e:
                        self._log(f"删除临时文件失败 {img_path}: {e}")
                
                # 清理原始临时图片
                for img_path, _, _ in image_paths_with_labels:
                    # 查找对应的原始图片（不带labeled_前缀的）
                    original_path = img_path.replace("labeled_", "").replace("_labeled_", "")
                    if os.path.exists(original_path) and original_path != img_path:
                        try:
                            os.remove(original_path)
                        except:
                            pass
                
                try:
                    # 尝试删除临时目录（如果为空）
                    if self.temp_images_dir.exists():
                        remaining_files = list(self.temp_images_dir.iterdir())
                        if not remaining_files:
                            self.temp_images_dir.rmdir()
                except Exception as e:
                    self._log(f"删除临时目录失败: {e}")
            
            self._log(f"\n转换完成！鸟瞰图保存在: {overview_path}")
            return overview_path, image_paths_with_labels
            
        except Exception as e:
            self._log(f"转换失败: {str(e)}", "ERROR")
            raise


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法: python multi_document_to_overview.py <文档文件1> [文档文件2] ... [输出目录] [布局方式]")
        print("支持的格式: .pptx, .ppt, .docx, .doc, .pdf")
        print("布局方式: vertical(垂直,默认), horizontal(水平), grid(网格)")
        print("示例: python multi_document_to_overview.py doc1.pdf doc2.pptx output vertical")
        sys.exit(1)
    
    # 解析参数
    args = sys.argv[1:]
    doc_paths = []
    output_dir = "output"
    layout = "vertical"
    
    for arg in args:
        if arg in ["vertical", "horizontal", "grid"]:
            layout = arg
        elif os.path.isdir(arg) or (not os.path.exists(arg) and not arg.endswith(('.pptx', '.ppt', '.docx', '.doc', '.pdf'))):
            output_dir = arg
        else:
            doc_paths.append(arg)
    
    if not doc_paths:
        print("错误: 请至少指定一个文档文件")
        sys.exit(1)
    
    try:
        converter = MultiDocumentToOverviewConverter(doc_paths, output_dir)
        overview_path, image_paths = converter.convert(layout=layout, output_filename="multi_document_overview.png")
        print(f"\n转换完成！鸟瞰图保存在: {overview_path}")
        print(f"共处理 {len(doc_paths)} 个文件，生成 {len(image_paths)} 张图片")
    except Exception as e:
        print(f"错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

