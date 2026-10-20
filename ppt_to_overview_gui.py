"""
文档转图片工具 - GUI版本
功能：提供图形界面，支持PPT、Word、PDF等多种文档格式
- 鸟瞰图模式：将所有页面拼接成一张大图
- 独立图片模式：每页生成一个独立的图片文件
- 多文件鸟瞰图模式：选择多个文件，每页生成独立图片并拼接成鸟瞰图
"""

import os
import sys
import threading
import tkinter
from pathlib import Path
from typing import List, Tuple
from tkinter import (
    Tk, Label, Button, Entry, StringVar, IntVar, 
    filedialog, messagebox, ttk, scrolledtext, Canvas, Scrollbar
)
from PIL import Image, ImageTk
from app.overview.document_to_overview import DocumentToOverviewConverter
from app.images.document_to_images import DocumentToImagesConverter
from app.overview.multi_document_to_overview import MultiDocumentToOverviewConverter
from app.logger_config import get_logger

# Windows专属依赖：COM初始化仅在Windows需要，非Windows平台跳过
IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    import pythoncom


class DocumentToOverviewGUI:
    """文档转图片工具GUI界面（支持鸟瞰图和独立图片两种模式）"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("文档转图片工具")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)
        
        # 初始化日志系统
        self.logger = get_logger("document_converter_gui")
        self.logger.info("=" * 50)
        self.logger.info("启动文档转图片工具GUI界面")
        
        # 变量
        self.document_path = StringVar()
        self.document_paths = []  # 多文件模式的文件列表（累积）
        self.document_path_1 = StringVar()  # 第一个文件输入框
        self.document_path_2 = StringVar()  # 第二个文件输入框
        self.output_dir = StringVar(value="output")
        self.mode_var = StringVar(value="overview")  # 功能模式：overview(鸟瞰图) 或 images(独立图片) 或 multi(多文件鸟瞰图)
        self.layout_var = StringVar(value="vertical")
        self.spacing_var = IntVar(value=10)
        self.dpi_var = IntVar(value=300)  # DPI设置，默认300
        self.output_filename = StringVar(value="")  # 初始为空，选择文档后自动填充
        self.is_processing = False
        self.preview_images = []  # 预览图片列表
        self.image_checkboxes = {}  # 图片复选框字典 {图片路径: (checkbox_var, 原始索引, thumb_frame, info_label)}
        self.selected_order = []  # 按勾选顺序存储的图片路径列表
        self.multi_converter = None  # 多文件转换器引用
        self.all_converted_images = []  # 所有已转换的图片列表 (路径, 源文件名)
        self.file_source_map = {}  # 文件来源映射 {文件路径: 输入框编号(1或2)}
        
        self.setup_ui()
        # 初始化时根据默认模式显示/隐藏选项
        self.on_mode_change()
    
    def setup_ui(self):
        """设置UI界面"""
        # 配置根窗口网格权重
        self.root.columnconfigure(0, weight=0)  # 左侧固定宽度
        self.root.columnconfigure(1, weight=1)   # 右侧自适应
        self.root.rowconfigure(0, weight=1)
        
        # 创建主容器（水平分割）
        main_container = ttk.Frame(self.root, padding="5")
        main_container.grid(row=0, column=0, columnspan=2, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S))
        main_container.columnconfigure(0, weight=0)  # 左侧固定
        main_container.columnconfigure(1, weight=1)  # 右侧自适应
        main_container.rowconfigure(0, weight=1)
        
        # ========== 左侧：功能模式列表 ==========
        left_panel = ttk.LabelFrame(main_container, text="功能模式", padding="10")
        left_panel.grid(row=0, column=0, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S), padx=(0, 5))
        left_panel.columnconfigure(0, weight=1)
        left_panel.rowconfigure(1, weight=1)
        
        # 左侧标题
        left_title = ttk.Label(
            left_panel,
            text="选择功能",
            font=("微软雅黑", 12, "bold")
        )
        left_title.grid(row=0, column=0, pady=(0, 5))
        
        # 创建模式按钮列表（垂直排列，紧密排列）
        self.mode_buttons = {}
        modes = [
            {
                "key": "overview",
                "text": "鸟瞰图模式",
                "desc": "将所有页面拼接成一张大图"
            },
            {
                "key": "images",
                "text": "独立图片模式",
                "desc": "每页生成一个独立的图片文件"
            },
            {
                "key": "multi",
                "text": "多文件鸟瞰图模式",
                "desc": "多选文件，每页生成独立图片并拼接成鸟瞰图"
            }
        ]
        
        # 垂直显示按钮（紧密排列）
        for idx, mode_info in enumerate(modes):
            mode_key = mode_info["key"]
            mode_text = mode_info["text"]
            mode_desc = mode_info["desc"]
            
            # 创建按钮容器（最小padding，使按钮紧密排列）
            btn_container = ttk.Frame(left_panel, padding="1")
            # 按钮之间无间距，紧紧挨着
            btn_container.grid(row=idx+1, column=0, sticky=(tkinter.W, tkinter.E), pady=0)
            btn_container.columnconfigure(0, weight=1)
            
            # 创建按钮（宽度固定，垂直排列）
            mode_btn = ttk.Button(
                btn_container,
                text=mode_text,
                command=lambda key=mode_key: self.on_mode_button_click(key),
                width=25
            )
            # 按钮内部无上下间距
            mode_btn.grid(row=0, column=0, sticky=(tkinter.W, tkinter.E), pady=0)
            
            # 创建描述标签（更紧凑，字体更小）
            desc_label = ttk.Label(
                btn_container,
                text=mode_desc,
                font=("微软雅黑", 7),
                foreground="gray",
                wraplength=180,
                justify=tkinter.LEFT
            )
            # 描述标签紧贴按钮，最小间距
            desc_label.grid(row=1, column=0, sticky=tkinter.W, pady=(1, 0))
            
            # 保存按钮引用
            self.mode_buttons[mode_key] = {
                "button": mode_btn,
                "desc": desc_label,
                "container": btn_container
            }
        
        # 初始化按钮状态
        self.update_mode_buttons_state()
        
        # ========== 右侧：功能配置区域 ==========
        right_panel = ttk.Frame(main_container, padding="10")
        right_panel.grid(row=0, column=1, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S))
        right_panel.columnconfigure(1, weight=1)
        # 不在这里设置rowconfigure，在日志区域单独设置
        
        # 右侧标题
        self.right_title = ttk.Label(
            right_panel,
            text="功能配置",
            font=("微软雅黑", 14, "bold")
        )
        self.right_title.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky=tkinter.W)
        
        row = 1
        
        # 文档文件选择（多文件模式使用两个输入框）
        self.document_label = ttk.Label(right_panel, text="文档文件:", font=("微软雅黑", 10))
        self.document_label.grid(row=row, column=0, sticky=tkinter.W, pady=5)
        self.document_entry = ttk.Entry(
            right_panel, 
            textvariable=self.document_path, 
            width=50,
            font=("微软雅黑", 9)
        )
        self.document_entry.grid(row=row, column=1, sticky=(tkinter.W, tkinter.E), padx=5, pady=5)
        self.browse_button = ttk.Button(
            right_panel, 
            text="浏览...", 
            command=self.browse_document_file
        )
        self.browse_button.grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        # 多文件模式：第二个文件输入框
        self.document_label_1 = ttk.Label(right_panel, text="文档文件1:", font=("微软雅黑", 10))
        self.document_entry_1 = ttk.Entry(
            right_panel, 
            textvariable=self.document_path_1, 
            width=50,
            font=("微软雅黑", 9),
            state="readonly"
        )
        self.browse_button_1 = ttk.Button(
            right_panel, 
            text="浏览...", 
            command=lambda: self.browse_document_file_multi(1)
        )
        
        # 多文件模式：第三个文件输入框
        self.document_label_2 = ttk.Label(right_panel, text="文档文件2:", font=("微软雅黑", 10))
        self.document_entry_2 = ttk.Entry(
            right_panel, 
            textvariable=self.document_path_2, 
            width=50,
            font=("微软雅黑", 9),
            state="readonly"
        )
        self.browse_button_2 = ttk.Button(
            right_panel, 
            text="浏览...", 
            command=lambda: self.browse_document_file_multi(2)
        )
        
        # 支持格式提示
        format_label = ttk.Label(
            right_panel, 
            text="支持格式: PPT (.pptx, .ppt) | Word (.docx, .doc) | PDF (.pdf)", 
            font=("微软雅黑", 8),
            foreground="gray"
        )
        format_label.grid(row=row, column=1, sticky=tkinter.W, padx=5, pady=(0, 5))
        row += 1
        
        # 输出目录选择
        ttk.Label(right_panel, text="输出目录:", font=("微软雅黑", 10)).grid(
            row=row, column=0, sticky=tkinter.W, pady=5
        )
        ttk.Entry(
            right_panel, 
            textvariable=self.output_dir, 
            width=50,
            font=("微软雅黑", 9)
        ).grid(row=row, column=1, sticky=(tkinter.W, tkinter.E), padx=5, pady=5)
        ttk.Button(
            right_panel, 
            text="浏览...", 
            command=self.browse_output_dir
        ).grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        # 输出文件名（仅鸟瞰图模式显示）
        self.filename_label = ttk.Label(right_panel, text="输出文件名:", font=("微软雅黑", 10))
        self.filename_label.grid(row=row, column=0, sticky=tkinter.W, pady=5)
        filename_entry = ttk.Entry(
            right_panel, 
            textvariable=self.output_filename, 
            width=50,
            font=("微软雅黑", 9)
        )
        filename_entry.grid(row=row, column=1, sticky=(tkinter.W, tkinter.E), padx=5, pady=5)
        self.filename_hint = ttk.Label(
            right_panel, 
            text="(选择文档后自动填充)", 
            font=("微软雅黑", 8),
            foreground="gray"
        )
        self.filename_hint.grid(row=row, column=2, sticky=tkinter.W, padx=5)
        row += 1
        
        # 布局方式（仅鸟瞰图模式显示）
        self.layout_frame = ttk.LabelFrame(right_panel, text="布局方式", padding="10")
        self.layout_frame.grid(row=row, column=0, columnspan=3, sticky=(tkinter.W, tkinter.E), pady=10)
        self.layout_frame.columnconfigure(0, weight=1)
        
        ttk.Radiobutton(
            self.layout_frame, 
            text="垂直排列 (Vertical)", 
            variable=self.layout_var, 
            value="vertical"
        ).grid(row=0, column=0, sticky=tkinter.W, padx=10)
        ttk.Radiobutton(
            self.layout_frame, 
            text="水平排列 (Horizontal)", 
            variable=self.layout_var, 
            value="horizontal"
        ).grid(row=0, column=1, sticky=tkinter.W, padx=10)
        ttk.Radiobutton(
            self.layout_frame, 
            text="网格排列 (Grid)", 
            variable=self.layout_var, 
            value="grid"
        ).grid(row=0, column=2, sticky=tkinter.W, padx=10)
        row += 1
        
        # 间距设置（仅鸟瞰图模式显示）
        self.spacing_label = ttk.Label(right_panel, text="图片间距 (像素):", font=("微软雅黑", 10))
        self.spacing_label.grid(row=row, column=0, sticky=tkinter.W, pady=5)
        spacing_spinbox = ttk.Spinbox(
            right_panel, 
            from_=0, 
            to=100, 
            textvariable=self.spacing_var, 
            width=20
        )
        spacing_spinbox.grid(row=row, column=1, sticky=tkinter.W, padx=5, pady=5)
        row += 1
        
        # DPI设置
        dpi_frame = ttk.LabelFrame(right_panel, text="图片质量设置 (DPI)", padding="10")
        dpi_frame.grid(row=row, column=0, columnspan=3, sticky=(tkinter.W, tkinter.E), pady=10)
        dpi_frame.columnconfigure(0, weight=1)
        
        ttk.Label(dpi_frame, text="分辨率 (DPI):", font=("微软雅黑", 9)).grid(
            row=0, column=0, sticky=tkinter.W, padx=5
        )
        dpi_spinbox = ttk.Spinbox(
            dpi_frame, 
            from_=72, 
            to=600, 
            textvariable=self.dpi_var, 
            width=15,
            increment=50
        )
        dpi_spinbox.grid(row=0, column=1, sticky=tkinter.W, padx=5)
        
        # DPI说明
        dpi_info = ttk.Label(
            dpi_frame, 
            text="DPI越高图片质量越好但文件越大、转换越慢。推荐: 150(快速预览) | 300(标准质量) | 600(高质量)", 
            font=("微软雅黑", 8),
            foreground="gray",
            wraplength=500
        )
        dpi_info.grid(row=1, column=0, columnspan=2, sticky=tkinter.W, padx=5, pady=(5, 0))
        row += 1
        
        # 分隔线
        ttk.Separator(right_panel, orient=tkinter.HORIZONTAL).grid(
            row=row, column=0, columnspan=3, sticky=(tkinter.W, tkinter.E), pady=10
        )
        row += 1
        
        # 图片预览区域（仅多文件模式显示，放在转换按钮上方）
        self.preview_frame = ttk.LabelFrame(right_panel, text="图片预览（勾选要拼接的图片，按勾选顺序打标签）", padding="5")
        self.preview_frame.grid(row=row, column=0, columnspan=3, sticky=(tkinter.W, tkinter.E), pady=10)
        self.preview_frame.columnconfigure(0, weight=1)
        # 不设置rowconfigure weight，限制预览区域高度
        
        # 创建带滚动条的Canvas用于显示图片
        preview_canvas_frame = ttk.Frame(self.preview_frame)
        preview_canvas_frame.grid(row=0, column=0, sticky=(tkinter.W, tkinter.E), pady=0)
        preview_canvas_frame.columnconfigure(0, weight=1)
        # 不设置rowconfigure weight
        
        self.preview_canvas = Canvas(
            preview_canvas_frame,
            bg="white",
            width=600,
            height=250  # 固定高度，不要自适应
        )
        preview_scrollbar = Scrollbar(
            preview_canvas_frame,
            orient="vertical",
            command=self.preview_canvas.yview
        )
        self.preview_canvas.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_canvas.grid(row=0, column=0, sticky=(tkinter.W, tkinter.E))
        preview_scrollbar.grid(row=0, column=1, sticky=(tkinter.N, tkinter.S))
        
        # 内部框架用于放置图片
        self.preview_inner_frame = ttk.Frame(self.preview_canvas)
        self.preview_canvas_window = self.preview_canvas.create_window(
            (0, 0),
            window=self.preview_inner_frame,
            anchor="nw"
        )
        
        # 绑定滚动区域更新
        self.preview_inner_frame.bind(
            "<Configure>",
            lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        )
        self.preview_canvas.bind(
            "<Configure>",
            lambda e: self.preview_canvas.itemconfig(self.preview_canvas_window, width=e.width)
        )
        
        # 绑定鼠标滚轮事件到Canvas
        def on_mousewheel(event):
            self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        # Windows和Linux的滚轮事件
        self.preview_canvas.bind_all("<MouseWheel>", on_mousewheel)
        # Linux的滚轮事件
        self.preview_canvas.bind_all("<Button-4>", lambda e: self.preview_canvas.yview_scroll(-1, "units"))
        self.preview_canvas.bind_all("<Button-5>", lambda e: self.preview_canvas.yview_scroll(1, "units"))
        
        row += 1
        
        # 转换按钮
        self.convert_button = ttk.Button(
            right_panel, 
            text="开始转换", 
            command=self.start_conversion,
            style="Accent.TButton"
        )
        self.convert_button.grid(row=row, column=0, columnspan=3, pady=15, sticky=(tkinter.W, tkinter.E))
        row += 1
        
        # 进度条
        self.progress_var = StringVar(value="等待开始...")
        self.progress_label = ttk.Label(
            right_panel, 
            textvariable=self.progress_var, 
            font=("微软雅黑", 9)
        )
        self.progress_label.grid(row=row, column=0, columnspan=3, pady=5)
        row += 1
        
        self.progress_bar = ttk.Progressbar(
            right_panel, 
            mode='indeterminate', 
            length=400
        )
        self.progress_bar.grid(row=row, column=0, columnspan=3, sticky=(tkinter.W, tkinter.E), pady=5)
        row += 1
        
        # 日志输出区域
        log_frame = ttk.LabelFrame(right_panel, text="转换日志", padding="5")
        log_frame.grid(row=row, column=0, columnspan=3, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        # 设置日志区域自适应，但限制预览区域高度
        right_panel.rowconfigure(row, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=8, 
            width=70,
            font=("Consolas", 9),
            wrap=tkinter.WORD
        )
        self.log_text.grid(row=0, column=0, sticky=(tkinter.W, tkinter.E, tkinter.N, tkinter.S))
        
        # 清空日志按钮
        ttk.Button(
            log_frame, 
            text="清空日志", 
            command=self.clear_log
        ).grid(row=1, column=0, pady=5)
    
    def browse_document_file_multi(self, box_num: int):
        """浏览选择文档文件（多文件模式，支持累积添加）"""
        filenames = filedialog.askopenfilenames(
            title=f"选择文档文件（输入框{box_num}，可多选）",
            filetypes=[
                ("所有支持的文档", "*.pptx *.ppt *.docx *.doc *.pdf"),
                ("PowerPoint文件", "*.pptx *.ppt"),
                ("Word文件", "*.docx *.doc"),
                ("PDF文件", "*.pdf"),
                ("所有文件", "*.*")
            ]
        )
        if filenames:
            # 累积添加到文件列表（不重复）
            new_files = []
            for f in filenames:
                if f not in self.document_paths:
                    self.document_paths.append(f)
                    new_files.append(f)
                # 记录文件来源（输入框1或2），即使文件已存在也更新来源
                self.file_source_map[f] = box_num
            
            # 更新显示
            if box_num == 1:
                if new_files:
                    self.document_path_1.set(f"已选择 {len(new_files)} 个文件")
                else:
                    self.document_path_1.set("已选择（重复文件已跳过）")
            else:
                if new_files:
                    self.document_path_2.set(f"已选择 {len(new_files)} 个文件")
                else:
                    self.document_path_2.set("已选择（重复文件已跳过）")
            
            if new_files:
                self.log(f"输入框{box_num}已选择 {len(new_files)} 个文档文件:")
                for filename in new_files:
                    ext = Path(filename).suffix.lower()
                    doc_type = "未知"
                    if ext in ['.pptx', '.ppt']:
                        doc_type = "PPT"
                    elif ext in ['.docx', '.doc']:
                        doc_type = "Word"
                    elif ext == '.pdf':
                        doc_type = "PDF"
                    self.log(f"  - {doc_type}: {Path(filename).name}")
                self.logger.info(f"用户选择 {len(new_files)} 个文档文件（输入框{box_num}）")
                
                # 转换新选择的文件为图片（不添加标签），追加到预览
                self.log(f"正在转换新选择的 {len(new_files)} 个文档为图片，请稍候...")
                self.convert_button.config(state="disabled", text="正在转换图片...")
                # 在新线程中执行图片转换
                thread = threading.Thread(target=lambda: self.convert_images_for_preview(new_files), daemon=True)
                thread.start()
            else:
                self.log(f"输入框{box_num}选择的文件已存在，跳过")
    
    def browse_document_file(self):
        """浏览选择文档文件（支持PPT、Word、PDF）"""
        mode = self.mode_var.get()
        
        if mode == "multi":
            # 多文件模式：使用两个输入框，这里不应该被调用
            pass
        else:
            # 单文件模式
            filename = filedialog.askopenfilename(
                title="选择文档文件",
                filetypes=[
                    ("所有支持的文档", "*.pptx *.ppt *.docx *.doc *.pdf"),
                    ("PowerPoint文件", "*.pptx *.ppt"),
                    ("Word文件", "*.docx *.doc"),
                    ("PDF文件", "*.pdf"),
                    ("所有文件", "*.*")
                ]
            )
            if filename:
                self.document_path.set(filename)
                self.document_paths = [filename]
                # 自动设置输出文件名为文档文件名
                doc_name = Path(filename).stem  # 获取不含扩展名的文件名
                self.output_filename.set(f"{doc_name}.png")
                
                # 检测文件类型
                ext = Path(filename).suffix.lower()
                doc_type = "未知"
                if ext in ['.pptx', '.ppt']:
                    doc_type = "PPT"
                elif ext in ['.docx', '.doc']:
                    doc_type = "Word"
                elif ext == '.pdf':
                    doc_type = "PDF"
                
                self.log(f"已选择{doc_type}文档: {filename}")
                # 根据当前模式设置输出文件名
                if self.mode_var.get() == "overview":
                    self.output_filename.set(f"{doc_name}.png")
                    self.log(f"输出文件名已自动设置为: {doc_name}.png")
                self.logger.info(f"用户选择{doc_type}文档: {filename}")
    
    def browse_output_dir(self):
        """浏览选择输出目录"""
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_dir.set(dirname)
            self.log(f"已选择输出目录: {dirname}")
    
    def log(self, message):
        """
        添加日志信息（同时输出到GUI和日志文件）
        
        Args:
            message: 日志消息
        """
        # 输出到GUI界面
        self.log_text.insert(tkinter.END, f"{message}\n")
        self.log_text.see(tkinter.END)
        self.root.update_idletasks()
        
        # 输出到日志文件
        self.logger.info(message)
    
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tkinter.END)
    
    def clear_preview(self):
        """清空图片预览"""
        # 清除所有预览图片
        for widget in self.preview_inner_frame.winfo_children():
            widget.destroy()
        self.preview_images = []
        self.image_checkboxes = {}
        self.selected_order = []
        self.all_converted_images = []
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
    
    def clear_all_data_after_conversion(self):
        """拼接完成后清理所有数据，为下次转换做准备"""
        self.log("=" * 50)
        self.log("开始清理数据，为下次转换做准备...")
        
        # 1. 清空预览区域
        self.clear_preview()
        self.log("已清空图片预览区域")
        
        # 2. 清空文件输入框
        self.document_path_1.set("")
        self.document_path_2.set("")
        self.log("已清空文件输入框")
        
        # 3. 清空文件列表和来源映射
        self.document_paths = []
        self.file_source_map = {}
        self.log("已清空文件列表和来源映射")
        
        # 4. 重置转换器
        self.multi_converter = None
        self.log("已重置转换器")
        
        # 5. 更新按钮状态
        self.convert_button.config(state="normal", text="开始拼接")
        self.progress_var.set("请选择文件并转换图片")
        
        self.log("数据清理完成，可以开始新的转换任务")
        self.log("=" * 50)
    
    def toggle_image_selection(self, img_path: str, var: tkinter.BooleanVar):
        """切换图片选中状态（点击图片时调用）"""
        # 确保使用绝对路径
        img_path_abs = str(Path(img_path).absolute())
        var.set(not var.get())
        self.on_checkbox_change(img_path_abs, var)
    
    def get_selected_images(self):
        """
        获取所有真正被选中的图片路径列表（按勾选顺序）
        这个方法直接检查复选框状态，并按照selected_order的顺序返回
        """
        # 使用selected_order来保持用户勾选的顺序
        # 但需要验证每个图片的复选框状态，确保只返回真正选中的
        selected = []
        for img_path in self.selected_order:
            img_path_abs = str(Path(img_path).absolute())
            # 检查这个路径是否在image_checkboxes中，并且复选框是选中的
            if img_path_abs in self.image_checkboxes:
                checkbox_info = self.image_checkboxes[img_path_abs]
                if len(checkbox_info) >= 4:
                    var = checkbox_info[0]
                    if var.get():  # 复选框被选中
                        selected.append(img_path_abs)
            else:
                # 如果不在image_checkboxes中，但仍在selected_order中，也添加（可能是新添加的）
                selected.append(img_path_abs)
        
        # 如果selected_order为空或有问题，回退到遍历所有复选框（但保持顺序）
        if not selected:
            for img_path, checkbox_info in self.image_checkboxes.items():
                if len(checkbox_info) >= 4:
                    var = checkbox_info[0]
                    if var.get():  # 复选框被选中
                        img_path_abs = str(Path(img_path).absolute())
                        selected.append(img_path_abs)
        
        return selected
    
    def on_checkbox_change(self, img_path: str, var: tkinter.BooleanVar):
        """复选框状态改变时的回调"""
        # 确保使用绝对路径
        img_path_abs = str(Path(img_path).absolute())
        
        if var.get():
            # 如果被选中，添加到选中顺序列表（如果还没有）
            # 使用路径字符串比较，确保匹配
            found = False
            for order_path in self.selected_order:
                if str(Path(order_path).absolute()) == img_path_abs:
                    found = True
                    break
            if not found:
                self.selected_order.append(img_path_abs)
        else:
            # 如果取消选中，从选中顺序列表中移除
            # 使用路径字符串比较，确保匹配
            to_remove = None
            for order_path in self.selected_order:
                if str(Path(order_path).absolute()) == img_path_abs:
                    to_remove = order_path
                    break
            if to_remove is not None:
                self.selected_order.remove(to_remove)
        
        # 更新预览显示（显示选中顺序）
        self.update_preview_display()
    
    def update_preview_display(self):
        """更新预览显示，显示选中顺序"""
        # 更新所有图片的显示信息
        for img_path, checkbox_info in self.image_checkboxes.items():
            if len(checkbox_info) == 5:
                var, original_idx, thumb_frame, info_label, img_label = checkbox_info
            else:
                var, original_idx, thumb_frame, info_label = checkbox_info
                img_label = None
            
            # 获取源文件名（优先从info_label的自定义属性获取）
            source_name = "未知"
            if hasattr(info_label, 'source_name'):
                source_name = info_label.source_name
            else:
                # 从all_converted_images中查找
                for path, name in self.all_converted_images:
                    # 统一使用绝对路径比较
                    path_abs = str(Path(path).absolute())
                    img_path_abs = str(Path(img_path).absolute())
                    if path_abs == img_path_abs:
                        source_name = name
                        break
                
                # 如果无法从all_converted_images获取，尝试从info_label文本获取
                if source_name == "未知":
                    parts = info_label.cget("text").split("\n")
                    if len(parts) >= 2:
                        source_name = parts[1]
                
                # 保存到info_label的自定义属性
                info_label.source_name = source_name
            
            # 检查是否被选中（优先使用var.get()，因为这是最准确的状态）
            # 如果var.get()为True，说明复选框被勾选
            is_selected = var.get()
            
            # img_path 已经是绝对路径（作为image_checkboxes的key存储的），统一转换为绝对路径字符串
            img_path_abs = str(Path(img_path).absolute())
            
            # 同步selected_order与复选框状态
            if is_selected:
                # 如果复选框被选中，确保在selected_order中（使用绝对路径）
                found_in_order = False
                matching_path = None
                for order_path in self.selected_order:
                    order_path_abs = str(Path(order_path).absolute())
                    if order_path_abs == img_path_abs:
                        found_in_order = True
                        matching_path = order_path
                        break
                
                if not found_in_order:
                    self.selected_order.append(img_path_abs)
            else:
                # 如果复选框未选中，从selected_order中移除
                # 需要找到并移除匹配的路径（考虑路径格式差异）
                to_remove = None
                for order_path in self.selected_order:
                    order_path_abs = str(Path(order_path).absolute())
                    if order_path_abs == img_path_abs:
                        to_remove = order_path
                        break
                if to_remove is not None:
                    self.selected_order.remove(to_remove)
            
            # 更新显示（基于复选框状态）
            if is_selected:
                # 找到在selected_order中的位置（考虑路径格式差异）
                order_num = 0
                for idx, order_path in enumerate(self.selected_order):
                    order_path_abs = str(Path(order_path).absolute())
                    if order_path_abs == img_path_abs:
                        order_num = idx + 1
                        break
                
                if order_num > 0:
                    info_label.config(text=f"#{order_num} (已选中)\n{source_name}")
                else:
                    # 如果找不到顺序号，说明同步有问题，但仍然显示已选中
                    info_label.config(text=f"已选中\n{source_name}")
                thumb_frame.config(relief="solid", borderwidth=2)
            else:
                info_label.config(text=f"未选中\n{source_name}")
                thumb_frame.config(relief="solid", borderwidth=1)
    
    def update_preview(self, image_paths_with_labels: List[Tuple[str, int, str]], append: bool = False):
        """
        更新图片预览区域（显示所有图片，带复选框）
        
        Args:
            image_paths_with_labels: 图片路径列表，每个元素为 (图片路径, 标签编号, 源文件名)
            append: 是否追加显示（True=追加，False=替换）
        """
        if not append:
            # 清空之前的预览
            for widget in self.preview_inner_frame.winfo_children():
                widget.destroy()
            self.preview_images = []
            self.image_checkboxes = {}
            self.selected_order = []
            self.all_converted_images = []
        
        if not image_paths_with_labels:
            return
        
        # 追加到已转换图片列表
        for img_path, _, source_name in image_paths_with_labels:
            img_path_abs = str(Path(img_path).absolute())
            if (img_path_abs, source_name) not in self.all_converted_images:
                self.all_converted_images.append((img_path_abs, source_name))
        
        self.log(f"正在更新预览，{'追加' if append else '显示'} {len(image_paths_with_labels)} 张图片...")
        if len(image_paths_with_labels) > 0:
            first_img = image_paths_with_labels[0]
            self.log(f"  第一张图片: {Path(first_img[0]).name} (标签: {first_img[1]}, 来源: {first_img[2]})")
            if len(image_paths_with_labels) > 1:
                last_img = image_paths_with_labels[-1]
                self.log(f"  最后一张图片: {Path(last_img[0]).name} (标签: {last_img[1]}, 来源: {last_img[2]})")
        
        # 每行显示的图片数量
        cols = 4
        max_thumb_size = 150  # 缩略图最大尺寸
        
        # 计算起始索引（追加模式）
        start_idx = len(self.image_checkboxes) if append else 0
        
        # 记录成功加载的图片数量
        loaded_count = 0
        skipped_count = 0
        
        for idx, (img_path, label_num, source_name) in enumerate(image_paths_with_labels):
            try:
                # 确保使用绝对路径
                img_path_abs = str(Path(img_path).absolute())
                
                # 检查文件是否存在
                if not os.path.exists(img_path_abs):
                    self.log(f"警告: 图片文件不存在，跳过 [{idx + 1}/{len(image_paths_with_labels)}] {img_path_abs}", "WARNING")
                    skipped_count += 1
                    continue
                
                # 检查是否已经存在于image_checkboxes中（避免重复添加）
                if img_path_abs in self.image_checkboxes:
                    self.log(f"警告: 图片已存在于预览中，跳过 [{idx + 1}/{len(image_paths_with_labels)}] {Path(img_path_abs).name}", "WARNING")
                    skipped_count += 1
                    continue
                
                # 打开图片
                try:
                    img = Image.open(img_path_abs)
                    
                    # 确保图片已加载
                    img.load()
                    
                    # 计算缩略图尺寸（保持比例）
                    img.thumbnail((max_thumb_size, max_thumb_size), Image.Resampling.LANCZOS)
                    
                    # 转换为PhotoImage
                    photo = ImageTk.PhotoImage(img)
                    self.preview_images.append(photo)  # 保持引用，防止被垃圾回收
                    loaded_count += 1
                except Exception as img_error:
                    self.log(f"警告: 加载图片失败 [{idx + 1}/{len(image_paths_with_labels)}] {img_path_abs}: {str(img_error)}", "WARNING")
                    skipped_count += 1
                    continue
                
                # 计算位置（考虑已存在的图片）
                display_idx = start_idx + idx
                row = display_idx // cols
                col = display_idx % cols
                
                # 创建容器框架
                thumb_frame = ttk.Frame(self.preview_inner_frame, relief="solid", borderwidth=1)
                thumb_frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                
                # 创建复选框变量
                checkbox_var = tkinter.BooleanVar(value=False)
                
                # 创建复选框
                checkbox = ttk.Checkbutton(
                    thumb_frame,
                    variable=checkbox_var,
                    command=lambda p=img_path_abs, v=checkbox_var: self.on_checkbox_change(p, v)
                )
                checkbox.grid(row=0, column=0, padx=2, pady=2, sticky="nw")
                
                # 显示图片（添加点击事件）
                img_label = ttk.Label(thumb_frame, image=photo, cursor="hand2")
                img_label.grid(row=1, column=0, padx=2, pady=2)
                # 绑定点击事件：点击图片也能勾选
                img_label.bind("<Button-1>", lambda e, p=img_path_abs, v=checkbox_var: self.toggle_image_selection(p, v))
                
                # 显示信息（保存源文件名用于后续更新）
                info_text = f"未选中\n{source_name}"
                info_label = ttk.Label(
                    thumb_frame,
                    text=info_text,
                    font=("微软雅黑", 8),
                    justify="center"
                )
                info_label.grid(row=2, column=0, padx=2, pady=2)
                # 在info_label中保存源文件名（使用自定义属性）
                info_label.source_name = source_name
                # 绑定点击事件：点击信息标签也能勾选
                info_label.bind("<Button-1>", lambda e, p=img_path_abs, v=checkbox_var: self.toggle_image_selection(p, v))
                
                # 整个框架也可以点击
                thumb_frame.bind("<Button-1>", lambda e, p=img_path_abs, v=checkbox_var: self.toggle_image_selection(p, v))
                thumb_frame.config(cursor="hand2")
                
                # 保存复选框信息（使用绝对路径作为key）
                self.image_checkboxes[img_path_abs] = (checkbox_var, display_idx, thumb_frame, info_label, img_label)
                
            except Exception as e:
                import traceback
                error_detail = traceback.format_exc()
                self.log(f"加载预览图片失败 [{idx + 1}/{len(image_paths_with_labels)}] {img_path}: {str(e)}", "WARNING")
                self.log(f"错误详情: {error_detail}", "WARNING")
                skipped_count += 1
        
        # 更新滚动区域
        self.preview_inner_frame.update_idletasks()
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        
        # 更新所有图片的显示状态（包括之前已存在的图片）
        self.update_preview_display()
        
        # 记录预览更新结果
        total_images = len(self.image_checkboxes)
        if skipped_count > 0:
            self.log(f"预览更新完成: 成功加载 {loaded_count} 张，跳过 {skipped_count} 张，预览区域共有 {total_images} 张图片")
        else:
            self.log(f"预览更新完成: 共 {total_images} 张图片")
        
        # 验证第一张图片是否成功添加（仅在出错时记录）
        if len(image_paths_with_labels) > 0:
            first_path_abs = str(Path(image_paths_with_labels[0][0]).absolute())
            if first_path_abs not in self.image_checkboxes:
                self.log(f"错误: 第一张图片未添加到预览! 路径: {first_path_abs}", "ERROR")
    
    def on_mode_button_click(self, mode_key: str):
        """功能模式按钮点击事件"""
        # 更新模式变量
        self.mode_var.set(mode_key)
        # 更新按钮状态
        self.update_mode_buttons_state()
        # 更新界面显示
        self.on_mode_change()
    
    def update_mode_buttons_state(self):
        """更新模式按钮的选中状态"""
        current_mode = self.mode_var.get()
        for mode_key, btn_info in self.mode_buttons.items():
            if mode_key == current_mode:
                # 选中状态：使用强调样式和不同背景色
                try:
                    btn_info["button"].config(style="Accent.TButton")
                except:
                    # 如果Accent样式不可用，使用自定义样式
                    btn_info["button"].config(relief="sunken")
                # 更新描述文字颜色
                btn_info["desc"].config(foreground="blue", font=("微软雅黑", 8, "bold"))
            else:
                # 未选中状态：使用默认样式
                try:
                    btn_info["button"].config(style="TButton")
                except:
                    btn_info["button"].config(relief="raised")
                # 更新描述文字颜色
                btn_info["desc"].config(foreground="gray", font=("微软雅黑", 8))
    
    def on_mode_change(self):
        """功能模式切换时的回调"""
        mode = self.mode_var.get()
        if mode == "overview":
            # 鸟瞰图模式：显示布局、间距和输出文件名选项
            self.layout_frame.grid()
            self.spacing_label.grid()
            self.filename_label.grid()
            self.filename_hint.grid()
            self.document_label.config(text="文档文件:")
            self.document_entry.config(state="normal")
            # 隐藏多文件模式的输入框
            self.document_label_1.grid_remove()
            self.document_entry_1.grid_remove()
            self.browse_button_1.grid_remove()
            self.document_label_2.grid_remove()
            self.document_entry_2.grid_remove()
            self.browse_button_2.grid_remove()
            # 更新右侧标题
            self.right_title.config(text="功能配置 - 鸟瞰图模式")
            # 更新输出文件名提示
            if self.document_path.get() and not self.document_path.get().startswith("已选择"):
                doc_name = Path(self.document_path.get()).stem
                self.output_filename.set(f"{doc_name}.png")
            # 隐藏预览区域
            self.preview_frame.grid_remove()
        elif mode == "images":
            # 独立图片模式：隐藏布局、间距和输出文件名选项
            self.layout_frame.grid_remove()
            self.spacing_label.grid_remove()
            self.filename_label.grid_remove()
            self.filename_hint.grid_remove()
            self.document_label.config(text="文档文件:")
            self.document_entry.config(state="normal")
            # 隐藏多文件模式的输入框
            self.document_label_1.grid_remove()
            self.document_entry_1.grid_remove()
            self.browse_button_1.grid_remove()
            self.document_label_2.grid_remove()
            self.document_entry_2.grid_remove()
            self.browse_button_2.grid_remove()
            # 更新右侧标题
            self.right_title.config(text="功能配置 - 独立图片模式")
            # 独立图片模式不需要输出文件名（每页一个文件）
            self.output_filename.set("")
            # 隐藏预览区域
            self.preview_frame.grid_remove()
        else:  # multi
            # 多文件鸟瞰图模式：显示布局、间距和输出文件名选项
            self.layout_frame.grid()
            self.spacing_label.grid()
            self.filename_label.grid()
            self.filename_hint.grid()
            self.document_label.config(text="文档文件:")
            self.document_entry.config(state="readonly")
            self.document_entry.grid_remove()  # 隐藏单文件输入框
            self.browse_button.grid_remove()
            # 显示多文件模式的输入框
            self.document_label_1.grid(row=1, column=0, sticky=tkinter.W, pady=5)
            self.document_entry_1.grid(row=1, column=1, sticky=(tkinter.W, tkinter.E), padx=5, pady=5)
            self.browse_button_1.grid(row=1, column=2, padx=5, pady=5)
            self.document_label_2.grid(row=2, column=0, sticky=tkinter.W, pady=5)
            self.document_entry_2.grid(row=2, column=1, sticky=(tkinter.W, tkinter.E), padx=5, pady=5)
            self.browse_button_2.grid(row=2, column=2, padx=5, pady=5)
            # 更新右侧标题
            self.right_title.config(text="功能配置 - 多文件鸟瞰图模式")
            # 显示预览区域
            self.preview_frame.grid()
            # 确保按钮可见
            self.convert_button.grid()
            # 不清空之前的预览（累积显示）
            self.convert_button.config(text="开始拼接")
    
    def validate_inputs(self):
        """验证输入"""
        mode = self.mode_var.get()
        
        if mode == "multi":
            # 多文件模式验证：检查是否有文件被选择
            if not self.document_paths:
                self.logger.warning("验证失败: 未选择文档文件")
                messagebox.showerror("错误", "请至少在一个输入框中选择文档文件！")
                return False
            
            # 验证所有文件
            for doc_path in self.document_paths:
                doc_path_obj = Path(doc_path)
                if not doc_path_obj.exists():
                    self.logger.warning(f"验证失败: 文档文件不存在 - {doc_path_obj}")
                    messagebox.showerror("错误", f"文档文件不存在：{doc_path_obj}")
                    return False
                
                # 检查文件格式
                ext = doc_path_obj.suffix.lower()
                supported_formats = ['.pptx', '.ppt', '.docx', '.doc', '.pdf']
                if ext not in supported_formats:
                    self.logger.warning(f"验证失败: 不支持的文件格式 - {ext}")
                    messagebox.showerror(
                        "错误", 
                        f"不支持的文件格式: {ext}\n文件: {doc_path_obj.name}\n支持的格式: .pptx, .ppt, .docx, .doc, .pdf"
                    )
                    return False
        else:
            # 单文件模式验证
            if not self.document_path.get() or self.document_path.get().startswith("已选择"):
                if not self.document_paths:
                    self.logger.warning("验证失败: 未选择文档文件")
                    messagebox.showerror("错误", "请选择文档文件！")
                    return False
                doc_path = Path(self.document_paths[0])
            else:
                doc_path = Path(self.document_path.get())
            
            if not doc_path.exists():
                self.logger.warning(f"验证失败: 文档文件不存在 - {doc_path}")
                messagebox.showerror("错误", "文档文件不存在！")
                return False
            
            # 检查文件格式
            ext = doc_path.suffix.lower()
            supported_formats = ['.pptx', '.ppt', '.docx', '.doc', '.pdf']
            if ext not in supported_formats:
                self.logger.warning(f"验证失败: 不支持的文件格式 - {ext}")
                messagebox.showerror(
                    "错误", 
                    f"不支持的文件格式: {ext}\n支持的格式: .pptx, .ppt, .docx, .doc, .pdf"
                )
                return False
        
        if not self.output_dir.get():
            self.logger.warning("验证失败: 未选择输出目录")
            messagebox.showerror("错误", "请选择输出目录！")
            return False
        
        # 根据功能模式验证输出文件名
        if mode in ["overview", "multi"]:
            # 鸟瞰图模式需要输出文件名
            if not self.output_filename.get():
                if mode == "multi":
                    self.output_filename.set("multi_document_overview.png")
                else:
                    doc_name = doc_path.stem
                    self.output_filename.set(f"{doc_name}.png")
        else:
            # 独立图片模式不需要输出文件名（每页一个文件）
            pass
        
        self.logger.info("输入验证通过")
        return True
    
    def start_conversion(self):
        """开始转换（在新线程中执行）"""
        if self.is_processing:
            messagebox.showwarning("警告", "转换正在进行中，请稍候...")
            return
        
        if not self.validate_inputs():
            return
        
        # 禁用按钮
        self.is_processing = True
        self.convert_button.config(state="disabled")
        self.progress_bar.start(10)
        
        # 在新线程中执行转换
        thread = threading.Thread(target=self.convert_thread, daemon=True)
        thread.start()
    
    def convert_thread(self):
        """转换线程"""
        # 在多线程环境中初始化COM接口（仅Windows需要）
        if IS_WINDOWS:
            pythoncom.CoInitialize()
        
        try:
            mode = self.mode_var.get()
            self.logger.info("=" * 50)
            self.logger.info(f"开始转换任务 - 模式: {mode}")
            self.log("=" * 50)
            
            if mode == "multi":
                mode_name = "多文件鸟瞰图模式"
            elif mode == "overview":
                mode_name = "鸟瞰图模式"
            else:
                mode_name = "独立图片模式"
            
            self.log(f"开始转换 - {mode_name}...")
            
            if mode == "multi":
                self.log(f"文档文件数: {len(self.document_paths)}")
                for idx, doc_path in enumerate(self.document_paths):
                    self.log(f"  文件 {idx + 1}: {Path(doc_path).name}")
            else:
                self.log(f"文档文件: {self.document_path.get()}")
            
            self.log(f"输出目录: {self.output_dir.get()}")
            self.log(f"DPI设置: {self.dpi_var.get()} (分辨率)")
            
            if mode == "multi":
                # 多文件鸟瞰图模式
                # 检查转换器是否存在
                if not hasattr(self, 'multi_converter') or self.multi_converter is None:
                    self.log("错误: 图片尚未转换，请先选择文件并等待转换完成！")
                    self.root.after(0, lambda: messagebox.showerror("错误", "图片尚未转换，请先选择文件并等待转换完成！"))
                    return
                
                # 确保转换器包含所有文件
                if not hasattr(self.multi_converter, 'all_image_paths') or not self.multi_converter.all_image_paths:
                    self.log("错误: 没有可用的图片，请先选择文件并等待转换完成！")
                    self.root.after(0, lambda: messagebox.showerror("错误", "没有可用的图片，请先选择文件并等待转换完成！"))
                    return
                
                # 先同步selected_order，确保它包含所有当前选中的图片（按勾选顺序）
                # 这样可以保持用户勾选的顺序
                self.update_preview_display()  # 这会同步selected_order与复选框状态
                
                # 直接检查复选框状态，获取真正选中的图片（按勾选顺序）
                verified_selected = self.get_selected_images()
                
                if not verified_selected:
                    self.log("错误: 请至少勾选一张图片进行拼接！")
                    self.root.after(0, lambda: messagebox.showerror("错误", "请至少勾选一张图片进行拼接！"))
                    return
                
                # 更新selected_order以保持同步
                self.selected_order = verified_selected
                
                self.log(f"已选中 {len(verified_selected)} 张图片，开始拼接...")
                
                converter = self.multi_converter
                
                # 步骤1: 为选中的图片按顺序添加标签（只处理勾选的图片）
                self.log(f"步骤1: 为选中的 {len(verified_selected)} 张图片添加标签...")
                labeled_images = converter.add_labels_to_selected_images(verified_selected)
                
                # 步骤2: 拼接成鸟瞰图
                self.log("步骤2: 拼接成鸟瞰图...")
                overview_path = converter.merge_images_to_overview(
                    labeled_images,
                    self.output_filename.get() or "multi_document_overview.png",
                    self.layout_var.get(),
                    self.spacing_var.get()
                )
                
                # 统计信息
                all_count = len(converter.all_image_paths) if hasattr(converter, 'all_image_paths') else 0
                
                self.log("=" * 50)
                self.log("拼接完成！")
                self.log(f"输出文件: {overview_path}")
                self.log(f"共处理 {len(self.document_paths)} 个文件")
                self.log(f"共生成 {all_count} 张图片")
                self.log(f"选中并拼接 {len(labeled_images)} 张图片")
                self.log("=" * 50)
                
                # 记录到日志文件
                self.logger.info("拼接任务完成（多文件鸟瞰图模式）")
                self.logger.info(f"输出文件: {overview_path}")
                self.logger.info(f"共生成 {all_count} 张图片，拼接 {len(labeled_images)} 张")
                
                # 显示成功消息
                self.root.after(0, lambda: messagebox.showinfo(
                    "成功", 
                    f"拼接完成！\n\n模式: 多文件鸟瞰图\n共处理 {len(self.document_paths)} 个文件\n生成 {all_count} 张图片\n选中并拼接 {len(labeled_images)} 张图片\n输出文件:\n{overview_path}"
                ))
                
                # 拼接完成后清理数据，为下次转换做准备
                self.root.after(0, self.clear_all_data_after_conversion)
                
            elif mode == "overview":
                # 鸟瞰图模式
                self.log(f"布局方式: {self.layout_var.get()}")
                self.log(f"图片间距: {self.spacing_var.get()} 像素")
                
                # 获取文档路径
                if self.document_path.get().startswith("已选择") and self.document_paths:
                    doc_path = self.document_paths[0]
                else:
                    doc_path = self.document_path.get()
                
                # 创建鸟瞰图转换器
                converter = DocumentToOverviewConverter(
                    doc_path,
                    self.output_dir.get(),
                    log_callback=self.log
                )
                
                # 执行转换
                output_path = converter.convert(
                    layout=self.layout_var.get(),
                    spacing=self.spacing_var.get(),
                    dpi=self.dpi_var.get(),
                    output_filename=self.output_filename.get() or None,
                    cleanup_temp=True
                )
                
                self.log("=" * 50)
                self.log("转换完成！")
                self.log(f"输出文件: {output_path}")
                self.log("=" * 50)
                
                # 记录到日志文件
                self.logger.info("转换任务完成（鸟瞰图模式）")
                self.logger.info(f"输出文件: {output_path}")
                
                # 显示成功消息
                self.root.after(0, lambda: messagebox.showinfo(
                    "成功", 
                    f"转换完成！\n\n模式: 鸟瞰图\n输出文件:\n{output_path}"
                ))
            else:
                # 独立图片模式
                # 获取文档路径
                if self.document_path.get().startswith("已选择") and self.document_paths:
                    doc_path = self.document_paths[0]
                else:
                    doc_path = self.document_path.get()
                
                # 创建独立图片转换器
                converter = DocumentToImagesConverter(
                    doc_path,
                    self.output_dir.get(),
                    log_callback=self.log
                )
                
                # 执行转换
                image_paths = converter.convert(
                    dpi=self.dpi_var.get()
                )
                
                self.log("=" * 50)
                self.log("转换完成！")
                self.log(f"共生成 {len(image_paths)} 张独立图片")
                self.log(f"输出目录: {self.output_dir.get()}")
                self.log("=" * 50)
                
                # 记录到日志文件
                self.logger.info("转换任务完成（独立图片模式）")
                self.logger.info(f"共生成 {len(image_paths)} 张图片")
                
                # 显示成功消息
                self.root.after(0, lambda: messagebox.showinfo(
                    "成功", 
                    f"转换完成！\n\n模式: 独立图片\n共生成 {len(image_paths)} 张图片\n输出目录:\n{self.output_dir.get()}"
                ))
            
        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            self.log("=" * 50)
            self.log(f"错误: {error_msg}")
            self.log("=" * 50)
            
            # 记录错误到日志文件
            self.logger.error("转换任务失败", exc_info=True)
            self.logger.error(f"错误信息: {error_msg}")
            
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
        
        finally:
            # 释放COM接口（仅Windows需要）
            if IS_WINDOWS:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
            # 恢复按钮状态
            self.root.after(0, self.conversion_finished)
    
    def conversion_finished(self):
        """转换完成后的处理"""
        self.is_processing = False
        self.convert_button.config(state="normal", text="开始转换")
        self.progress_bar.stop()
        self.progress_var.set("转换完成")
    
    def convert_images_for_preview(self, files_to_convert: List[str] = None):
        """转换图片用于预览（不添加标签），支持追加模式"""
        # 初始化COM接口（仅Windows需要）
        if IS_WINDOWS:
            pythoncom.CoInitialize()
        
        try:
            # 如果指定了要转换的文件列表，只转换这些文件；否则转换所有文件
            if files_to_convert is None:
                files_to_convert = self.document_paths
            
            if not files_to_convert:
                self.log("警告: 没有文件需要转换")
                return
            
            # 确定这些文件来自哪个输入框，使用对应的临时目录
            # 如果文件来自不同输入框，使用第一个文件的来源
            source_box = None
            for f in files_to_convert:
                if f in self.file_source_map:
                    source_box = self.file_source_map[f]
                    break
            
            # 如果没有找到来源，默认使用输入框1
            if source_box is None:
                source_box = 1
            
            # 根据输入框编号使用不同的临时目录
            temp_dir_name = f"temp_images_{source_box}"
            self.log(f"使用临时目录: {temp_dir_name} (输入框{source_box})")
            
            # 创建多文件转换器（使用要转换的文件列表和对应的临时目录）
            converter = MultiDocumentToOverviewConverter(
                files_to_convert,
                self.output_dir.get(),
                log_callback=self.log,
                temp_dir_name=temp_dir_name
            )
            
            # 转换文档为图片（不添加标签）
            all_image_paths = converter.convert_all_documents_to_images(
                image_format="PNG",
                dpi=self.dpi_var.get()
            )
            
            # 转换为预览格式 (图片路径, 临时标签, 源文件名)
            # 确保路径是绝对路径，并验证路径正确性
            preview_data = []
            for idx, (path, source) in enumerate(all_image_paths):
                path_abs = str(Path(path).absolute())
                # 验证路径是否存在（仅在不存在时记录警告）
                if not os.path.exists(path_abs):
                    self.log(f"警告: 转换后的图片路径不存在: {path_abs}", "WARNING")
                preview_data.append((path_abs, len(self.all_converted_images) + idx + 1, source))
            
            # 记录预览数据信息
            self.log(f"准备显示预览，共 {len(preview_data)} 张图片")
            
            # 更新预览（追加模式）
            is_append = len(self.all_converted_images) > 0
            self.root.after(0, lambda data=preview_data, append=is_append: self.update_preview(data, append=append))
            self.root.after(0, lambda: self.convert_button.config(state="normal", text="开始拼接"))
            self.root.after(0, lambda: self.progress_var.set(f"图片转换完成，共 {len(self.all_converted_images) + len(preview_data)} 张图片，请勾选要拼接的图片"))
            
            # 保存转换器引用（如果还没有，或者合并）
            # 注意：由于不同输入框使用不同的临时目录，我们需要合并图片路径
            # 但保留各自的转换器引用，或者创建一个统一的转换器来管理所有图片
            if self.multi_converter is None:
                self.multi_converter = converter
                self.log(f"创建新的转换器，使用临时目录: {temp_dir_name}")
            else:
                # 合并新转换的图片到现有转换器
                # 重要：需要合并所有图片路径，确保包含所有文件
                # 注意：不同输入框的图片保存在不同的临时目录中，但可以合并到同一个转换器管理
                existing_paths = set(path for path, _ in self.multi_converter.all_image_paths)
                for path, source in all_image_paths:
                    if path not in existing_paths:
                        self.multi_converter.all_image_paths.append((path, source))
                
                # 同时需要更新转换器的document_paths，确保包含所有文件
                existing_docs = set(str(Path(p).absolute()) for p in self.multi_converter.document_paths)
                for doc_path in files_to_convert:
                    doc_path_abs = str(Path(doc_path).absolute())
                    if doc_path_abs not in existing_docs:
                        self.multi_converter.document_paths.append(Path(doc_path))
                
                self.log(f"已合并转换器，当前共有 {len(self.multi_converter.all_image_paths)} 张图片")
            
        except Exception as e:
            error_msg = f"转换图片失败: {str(e)}"
            self.log(f"错误: {error_msg}")
            import traceback
            self.log(f"错误详情: {traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            self.root.after(0, lambda: self.convert_button.config(state="normal", text="开始拼接"))
        finally:
            # 释放COM接口（仅Windows需要）
            if IS_WINDOWS:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass


def main():
    """主函数"""
    root = Tk()
    app = DocumentToOverviewGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

