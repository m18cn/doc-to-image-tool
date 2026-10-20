"""
PPT转图片并拼接成鸟瞰图工具
功能：将PPT的每一页转换为图片，并将所有图片拼接成一张鸟瞰图
"""

import os
import sys
from pathlib import Path
from PIL import Image
from typing import List, Optional

# Windows专属依赖：COM自动化仅在Windows可用，此转换器仅支持Windows
IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    import win32com.client


class PPTToOverviewConverter:
    """PPT转图片并拼接成鸟瞰图的转换器"""
    
    def __init__(self, ppt_path: str, output_dir: str = "output", log_callback=None):
        """
        初始化转换器
        
        Args:
            ppt_path: PPT文件路径
            output_dir: 输出目录
            log_callback: 日志回调函数，用于输出日志信息
        """
        self.ppt_path = Path(ppt_path)
        self.output_dir = Path(output_dir)
        self.temp_images_dir = self.output_dir / "temp_images"
        self.log_callback = log_callback
        
        # 验证PPT文件是否存在
        if not self.ppt_path.exists():
            raise FileNotFoundError(f"PPT文件不存在: {self.ppt_path}")
        
        # 创建输出目录
        self.output_dir.mkdir(exist_ok=True)
        self.temp_images_dir.mkdir(exist_ok=True)
    
    def _log(self, message: str):
        """输出日志"""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)
    
    def ppt_to_images(self, image_format: str = "PNG", dpi: int = 300) -> List[str]:
        """
        将PPT的每一页转换为图片
        
        Args:
            image_format: 图片格式 (PNG, JPG等)
            dpi: 图片分辨率，默认300
            
        Returns:
            生成的图片文件路径列表
        """
        self._log(f"正在将PPT转换为图片: {self.ppt_path}")

        # 此旧版转换器仅支持Windows（新版跨平台转换请使用DocumentToOverviewConverter）
        if not IS_WINDOWS:
            raise RuntimeError("此转换器仅支持Windows平台，请使用文档转图片工具的主转换功能")

        # 启动PowerPoint应用程序
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # 设置为可见，便于调试
        
        try:
            # 打开PPT文件
            presentation = powerpoint.Presentations.Open(str(self.ppt_path.absolute()))
            
            image_paths = []
            slide_count = presentation.Slides.Count
            
            self._log(f"PPT共有 {slide_count} 页")
            
            # 获取幻灯片尺寸
            slide_width = int(presentation.PageSetup.SlideWidth)
            slide_height = int(presentation.PageSetup.SlideHeight)
            
            # 遍历每一页并导出为图片
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                image_path = self.temp_images_dir / f"slide_{i:03d}.{image_format.lower()}"
                
                # 导出幻灯片为图片
                # Export方法参数: (FileName, FilterName, ScaleWidth, ScaleHeight)
                slide.Export(
                    str(image_path.absolute()),
                    image_format,
                    slide_width,
                    slide_height
                )
                
                image_paths.append(str(image_path))
                self._log(f"已导出第 {i}/{slide_count} 页: {image_path.name}")
            
            # 关闭演示文稿
            presentation.Close()
            
            self._log(f"成功导出 {len(image_paths)} 张图片")
            return image_paths
            
        except Exception as e:
            self._log(f"转换过程中出错: {str(e)}")
            raise
        finally:
            # 退出PowerPoint应用程序
            powerpoint.Quit()
    
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
                # 居中放置
                x_pos = (final_width - img.width) // 2
                overview.paste(img, (x_pos, y_offset))
                y_offset += img.height + spacing
                
        elif layout == "horizontal":
            for img in images:
                # 居中放置
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
        overview.save(output_path, "PNG", quality=95)
        
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
        完整转换流程：PPT转图片 -> 拼接成鸟瞰图
        
        Args:
            image_format: 图片格式
            dpi: 分辨率
            layout: 布局方式
            spacing: 间距
            output_filename: 输出文件名，如果为None则使用PPT文件名
            cleanup_temp: 是否清理临时图片文件
            
        Returns:
            最终输出文件路径
        """
        try:
            # 如果没有指定输出文件名，使用PPT文件名
            if output_filename is None:
                output_filename = f"{self.ppt_path.stem}.png"
            
            # 步骤1: PPT转图片
            image_paths = self.ppt_to_images(image_format, dpi)
            
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
        print("使用方法: python ppt_to_overview.py <PPT文件路径> [输出目录] [布局方式]")
        print("布局方式: vertical(垂直,默认), horizontal(水平), grid(网格)")
        print("示例: python ppt_to_overview.py presentation.pptx output vertical")
        sys.exit(1)
    
    ppt_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    layout = sys.argv[3] if len(sys.argv) > 3 else "vertical"
    
    if layout not in ["vertical", "horizontal", "grid"]:
        print("错误: 布局方式必须是 vertical, horizontal 或 grid")
        sys.exit(1)
    
    try:
        converter = PPTToOverviewConverter(ppt_path, output_dir)
        # 使用PPT文件名作为输出文件名
        ppt_name = Path(ppt_path).stem
        converter.convert(layout=layout, output_filename=f"{ppt_name}.png")
    except Exception as e:
        print(f"错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

