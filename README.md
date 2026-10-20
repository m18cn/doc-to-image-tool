# PPT转图片并拼接成鸟瞰图工具

## 项目简介

本项目实现将PowerPoint演示文稿的每一页转换为图片，并将所有图片拼接成一张鸟瞰图的功能。适用于需要快速预览整个PPT内容或生成PPT缩略图的场景。

## 功能特性

- ✅ 将PPT的每一页转换为高质量图片（PNG格式）
- ✅ 支持三种布局方式：垂直、水平、网格
- ✅ 可自定义图片间距和背景颜色
- ✅ 自动清理临时文件
- ✅ 支持命令行调用、代码调用和图形界面三种方式
- ✅ 提供Windows GUI界面，操作简单直观

## 开发步骤

### 步骤1: 环境准备

#### 1.1 创建虚拟环境（用户自行创建）
```bash
# 使用venv创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate
```

#### 1.2 安装依赖包
```bash
pip install -r requirements.txt
```

**依赖说明：**
- `Pillow`: 用于图片处理和拼接
- `pywin32`: 用于调用Windows PowerPoint COM接口

**注意：** 本项目依赖Windows系统和已安装的Microsoft PowerPoint软件。

### 步骤2: 项目结构

```
project/
├── ppt_to_overview_gui.py     # GUI图形界面（启动文件，根目录）
├── requirements.txt           # 依赖包列表
├── README.md                  # 项目文档
├── logs/                      # 日志目录（自动创建）
│   └── converter_YYYYMMDD.log # 按日期命名的日志文件
├── output/                    # 输出目录（自动创建）
│   └── overview.png           # 生成的鸟瞰图
└── app/                       # 应用功能模块目录
    ├── __init__.py
    ├── logger_config.py       # 日志系统配置模块
    ├── overview/              # 鸟瞰图功能模块
    │   ├── __init__.py
    │   └── document_to_overview.py  # 文档转鸟瞰图转换器
    ├── images/                # 独立图片功能模块
    │   ├── __init__.py
    │   └── document_to_images.py    # 文档转独立图片转换器
    └── ppt/                   # PPT专用功能模块
    |   ├── __init__.py
    |   └── ppt_to_overview.py        # PPT转鸟瞰图转换器
    |── image_pdfword/                # 图片转pdf、word
    │   ├── __init__.py
    │   └── image_to_pdfword.py    # 图片转pdf、word转换器
    
```

### 步骤3: 核心功能实现

#### 3.1 PPT转图片功能

使用`win32com.client`调用PowerPoint COM接口：

1. **启动PowerPoint应用程序**
   ```python
   powerpoint = win32com.client.Dispatch("PowerPoint.Application")
   ```

2. **打开PPT文件**
   ```python
   presentation = powerpoint.Presentations.Open(ppt_path)
   ```

3. **遍历每一页并导出**
   ```python
   for i in range(1, slide_count + 1):
       slide = presentation.Slides(i)
       slide.Export(image_path, "PNG", width, height)
   ```

4. **关闭并退出**
   ```python
   presentation.Close()
   powerpoint.Quit()
   ```

#### 3.2 图片拼接功能

使用`PIL.Image`进行图片拼接：

1. **加载所有图片**
   ```python
   images = [Image.open(path) for path in image_paths]
   ```

2. **计算画布尺寸**
   - 垂直布局：宽度=最大图片宽度，高度=所有图片高度之和+间距
   - 水平布局：宽度=所有图片宽度之和+间距，高度=最大图片高度
   - 网格布局：根据图片数量计算行列数

3. **创建画布并拼接**
   ```python
   overview = Image.new("RGB", (width, height), background_color)
   overview.paste(img, (x_offset, y_offset))
   ```

4. **保存最终图片**
   ```python
   overview.save(output_path, "PNG", quality=95)
   ```

### 步骤4: 使用方法

#### 4.1 图形界面使用（推荐）

启动GUI界面（根目录启动文件）：
```bash
python ppt_to_overview_gui.py
```

**注意**：GUI文件位于根目录，作为项目的主启动入口。

GUI界面功能：
- **文档文件选择**：点击"浏览..."按钮选择要转换的文档（支持PPT、Word、PDF）
- **输出目录选择**：点击"浏览..."按钮选择输出目录
- **输出文件名**：自定义输出文件名（选择文档后自动填充）
- **布局方式**：选择垂直、水平或网格布局
- **图片间距**：设置图片之间的间距（像素）
- **DPI设置**：设置图片分辨率，控制图片质量和文件大小
- **实时日志**：查看转换过程的详细日志信息
- **进度显示**：显示转换进度状态

操作步骤：
1. 点击"文档文件"旁的"浏览..."按钮，选择文档文件（PPT/Word/PDF）
2. 点击"输出目录"旁的"浏览..."按钮，选择保存位置
3. 设置输出文件名（可选，默认使用文档文件名）
4. 选择布局方式和间距
5. 设置DPI参数（推荐：150-快速预览，300-标准质量，600-高质量）
6. 点击"开始转换"按钮
7. 等待转换完成，查看日志输出

#### 4.2 命令行使用

**鸟瞰图功能**（使用app模块）：
```bash
# 使用通用文档转换器（支持PPT/Word/PDF）
python -m app.overview.document_to_overview document.pdf output vertical
```

**独立图片功能**（使用app模块）：
```bash
# 将文档每页转为独立图片（10页文档生成10个图片文件）
python -m app.images.document_to_images document.pdf output 300
```

**PPT专用功能**（使用app模块）：
```bash
# 使用PPT专用转换器
python -m app.ppt.ppt_to_overview presentation.pptx output vertical
```

#### 4.3 代码调用

**鸟瞰图功能**：
```python
from app.overview.document_to_overview import DocumentToOverviewConverter

# 创建转换器（支持PPT/Word/PDF）
converter = DocumentToOverviewConverter("document.pdf", "output")

# 执行转换
converter.convert(
    layout="vertical",      # 布局方式: vertical/horizontal/grid
    spacing=10,             # 图片间距（像素）
    output_filename="overview.png",
    cleanup_temp=True       # 是否清理临时文件
)
```

**独立图片功能**：
```python
from app.images.document_to_images import DocumentToImagesConverter

# 创建转换器（每页生成独立图片）
converter = DocumentToImagesConverter("document.pdf", "output")
image_paths = converter.convert(dpi=300)
# 返回所有生成的图片路径列表
```

**PPT专用功能**：
```python
from app.ppt.ppt_to_overview import PPTToOverviewConverter

# 创建转换器（仅支持PPT）
converter = PPTToOverviewConverter("presentation.pptx", "output")
```

#### 4.4 高级用法

```python
# 分步执行（鸟瞰图功能）
from app.overview.document_to_overview import DocumentToOverviewConverter
converter = DocumentToOverviewConverter("document.pdf")

# 步骤1: 转换为图片
image_paths = converter.ppt_to_images(image_format="PNG", dpi=300)

# 步骤2: 拼接成鸟瞰图
overview_path = converter.merge_images_to_overview(
    image_paths,
    layout="grid",
    spacing=20,
    background_color=(240, 240, 240)
)
```

### 步骤5: 文件说明

- **`ppt_to_overview.py`**: 核心转换模块，包含转换器类和命令行接口
- **`ppt_to_overview_gui.py`**: GUI图形界面版本，提供可视化操作界面
- **`requirements.txt`**: 项目依赖包列表

### 步骤6: 参数说明

#### PPTToOverviewConverter类参数

- `ppt_path`: PPT文件路径（必需）
- `output_dir`: 输出目录（默认："output"）

#### convert()方法参数

- `image_format`: 图片格式，默认"PNG"
- `dpi`: 图片分辨率（DPI），默认300
  - **推荐值**：
    - `150`: 快速预览，文件小，转换快
    - `300`: 标准质量（默认），平衡质量和速度
    - `600`: 高质量，文件大，转换慢
  - **范围**：72-600
- `layout`: 布局方式
  - `"vertical"`: 垂直排列（默认）
  - `"horizontal"`: 水平排列
  - `"grid"`: 网格排列
- `spacing`: 图片间距（像素），默认10
- `output_filename`: 输出文件名，如果为None则使用文档文件名
- `cleanup_temp`: 是否清理临时文件，默认True

### 步骤7: 日志系统

项目集成了完整的日志系统，所有操作和错误信息都会记录到日志文件中：

- **日志目录**: `logs/`（自动创建）
- **日志文件命名**: `converter_YYYYMMDD.log`（按日期命名）
- **日志内容**:
  - 程序启动和关闭
  - 文档选择和处理过程
  - 转换进度和状态
  - 错误和异常信息
  - 用户操作记录

**日志文件管理**:
- 单个日志文件最大10MB
- 自动轮转，保留最近5个备份文件
- 日志格式：`时间戳 - 模块名 - 级别 - 消息`

**查看日志**:
- GUI界面：实时显示在"转换日志"区域
- 日志文件：查看`logs/converter_YYYYMMDD.log`文件
- 控制台：命令行运行时也会输出到控制台

### 步骤8: 输出说明

转换完成后，会在输出目录生成：

- `overview.png`: 最终的鸟瞰图
- `temp_images/`: 临时图片目录（如果`cleanup_temp=False`则保留）

### 步骤9: 注意事项

1. **系统要求**
   - Windows操作系统
   - 已安装Microsoft PowerPoint（2010或更高版本）
   - Python 3.7+

2. **权限问题**
   - 确保有PPT文件的读取权限
   - 确保输出目录有写入权限

3. **性能优化与DPI设置**
   - 大型文档文件转换可能需要较长时间
   - 建议在转换过程中不要操作Office应用程序窗口
   - **DPI参数说明**：
     - **72-150 DPI**：快速预览，文件小，转换快，适合快速浏览
     - **200-300 DPI**：标准质量（推荐），平衡质量和速度，适合大多数场景
     - **400-600 DPI**：高质量，文件大，转换慢，适合打印或精细查看
   - **转换时间参考**（以10页文档为例）：
     - 150 DPI：约10-20秒
     - 300 DPI：约20-40秒
     - 600 DPI：约40-80秒
   - 可以通过调整DPI参数平衡图片质量、文件大小和转换速度

4. **错误处理**
   - 如果PPT文件损坏或格式不支持，程序会抛出异常
   - 确保PowerPoint可以正常打开目标文件

### 步骤10: 常见问题

**Q: 提示"PowerPoint.Application"无法创建？**
A: 确保已安装Microsoft PowerPoint，并且系统支持COM接口。

**Q: 生成的图片质量不高？**
A: 可以增加DPI参数值（在GUI中设置或代码中传入`dpi`参数）。注意：
- 150 DPI适合快速预览
- 300 DPI是标准质量（推荐）
- 600 DPI是高质量，但会增加处理时间和文件大小

**Q: 转换速度太慢？**
A: 可以降低DPI参数值：
- 从600降到300可以显著提升速度
- 从300降到150可以更快，但质量会降低
- 大型文档建议使用150-200 DPI进行快速预览

**Q: 生成的图片文件太大？**
A: 降低DPI参数值可以减少文件大小：
- 300 DPI：标准质量，文件大小适中
- 150 DPI：文件小，适合网络传输
- 72 DPI：文件最小，但质量较低

**Q: 如何调整图片顺序？**
A: 图片按照PPT中的页面顺序排列，如需调整顺序，请先修改PPT文件。

**Q: 支持哪些文档格式？**
A: 支持以下格式：
- PowerPoint: `.pptx`, `.ppt`
- Word: `.docx`, `.doc`
- PDF: `.pdf`

**Q: 日志文件在哪里？**
A: 所有日志文件保存在`logs/`目录下，按日期命名（如`converter_20240101.log`）。日志文件会自动轮转，单个文件最大10MB，保留最近5个备份。

**Q: 如何查看详细的转换日志？**
A: 可以通过以下方式查看：
- GUI界面：在"转换日志"区域实时查看
- 日志文件：打开`logs/converter_YYYYMMDD.log`文件查看完整日志
- 控制台：命令行运行时直接输出到控制台

### 步骤11: 打包成exe文件

本项目提供了打包脚本，可以将程序打包成独立的exe文件，方便在没有Python环境的Windows系统上运行。

#### 11.1 安装打包工具

确保已安装PyInstaller（已包含在requirements.txt中）：

```bash
pip install PyInstaller>=6.0.0
```

#### 11.2 使用打包脚本

**方法1：使用批处理文件（推荐）**

直接运行批处理文件：

```bash
build_exe.bat
```

**方法2：使用Python脚本**

```bash
python build_exe.py
```

#### 11.3 打包说明

- **打包方式**：单文件打包（`--onefile`），生成一个独立的exe文件
- **输出位置**：打包完成后，exe文件位于`dist/文档转图片工具.exe`
- **文件大小**：由于包含所有依赖库，exe文件较大（约100-200MB），这是正常现象
- **运行要求**：
  - Windows系统
  - 已安装Microsoft Office（PowerPoint和Word）
  - 首次运行可能需要几秒钟解压时间

#### 11.4 打包选项说明

打包脚本已配置以下选项：

- `--windowed`：隐藏控制台窗口（GUI应用）
- `--onefile`：打包成单个exe文件
- `--clean`：清理临时文件
- `--hidden-import`：包含所有必要的隐藏导入模块（win32com、PyMuPDF等）

#### 11.5 自定义打包

如果需要自定义打包选项，可以编辑`build_exe.py`文件，修改PyInstaller参数。

例如，添加图标：

```python
# 在args列表中添加
'--icon=icon.ico',  # 指定图标文件
```

#### 11.6 打包后测试

打包完成后，建议：

1. 在打包机器上测试exe文件是否正常运行
2. 在没有Python环境的Windows机器上测试
3. 测试各种文档格式（PPT、Word、PDF）的转换功能
4. 检查日志文件是否正常生成

#### 11.7 常见打包问题

**Q: 打包后的exe文件很大？**
A: 这是正常的，因为包含了Python解释器和所有依赖库。单文件打包方式会生成较大的exe文件。

**Q: exe文件运行报错"找不到模块"？**
A: 检查`build_exe.py`中的`--hidden-import`参数，确保包含了所有必要的模块。

**Q: 打包时提示缺少某些模块？**
A: 确保已安装所有依赖包：`pip install -r requirements.txt`

**Q: exe文件运行很慢？**
A: 单文件打包方式在首次运行时会解压临时文件，需要几秒钟时间。这是正常现象。

**Q: 如何减小exe文件大小？**
A: 可以尝试使用`--onedir`模式（目录模式）而不是`--onefile`，但需要将整个目录一起分发。

## 技术栈

- **Python 3.7+**
- **Pillow**: 图片处理库
- **pywin32**: Windows COM接口调用
- **Microsoft PowerPoint**: 通过COM接口调用

## 许可证

本项目仅供学习和研究使用。

