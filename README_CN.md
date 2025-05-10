# any2md - 将各种文档转换为Markdown

<div align="right">
  <a href="README.md">English</a> | <a href="README_CN.md">中文文档</a>
</div>

`any2md` 是一个功能强大的Python工具，可以将各种文档格式（PDF、PPT/PPTX、PNG）转换为高质量的Markdown。它旨在保留原始文档的结构、样式和内容，同时将其转换为便于在Obsidian中使用的格式。

## 功能特点

- **多格式支持**：处理PDF文件（包括纯图像和富文本）、PowerPoint演示文稿（PPT/PPTX）和PNG图像系列
- **智能内容提取**：
  - **PyMuPDF模式**：最适合含有可提取内容的富文本PDF
  - **Qwen-VL模式**：AI驱动的图像分析，用于纯图像PDF和演示文稿中的内容定位
- **智能处理**：
  - 使用适当的Markdown层级结构保留文档结构
  - 尽可能将表格转换为Markdown格式
  - 准确处理LaTeX数学表达式
  - 维护带有适当描述的图像引用
- **性能优化**：
  - 多页并发处理，加速转换
  - API自动重试机制，提高稳定性
  - 可配置的工作线程数
- **增强工作流**：
  - 自动图像裁剪和提取
  - 图像占位符自动替换
  - 可选的调试可视化
  - 使用DeepSeek V2.5优化Markdown内容（可选）

## 系统要求

### 核心依赖

```bash
# 使用 uv（推荐）
uv pip install -r requirements.txt

# 或使用 pip
pip install requests fitz pdf2image Pillow tqdm python-dotenv tenacity
```

### 可选依赖

- **LibreOffice**：处理PPT/PPTX文件所必需
  - Windows：从[libreoffice.org](https://www.libreoffice.org/download/download/)下载
  - Linux：`sudo apt install libreoffice`
  - macOS：`brew install libreoffice`

- **额外的Python包**（Windows特有）：
  ```bash
  # 使用 uv
  uv pip install comtypes pyautogui
  
  # 或使用 pip
  pip install comtypes pyautogui
  ```

## 安装

### 安装 uv（可选但推荐）

```bash
# 使用 pipx 安装 uv（推荐）
pipx install uv

# 或使用 pip 直接安装
pip install uv
```

### 安装 any2md

```bash
# 克隆仓库
git clone https://github.com/TsekaLuk/any2md.git
cd any2md

# 使用 uv 安装依赖
uv pip install -r requirements.txt

# 或使用 pip（如果您更喜欢）
pip install -r requirements.txt
```

## API配置

此工具使用AI服务进行文档分析。在同一目录下创建一个`.env`文件，内容如下：

```
SF_API_URL=https://api.siliconflow.cn/v1/chat/completions
SF_API_KEY=your_api_key_here
SF_MODEL=Qwen/Qwen2.5-VL-72B-Instruct
SF_MAX_WORKERS=10
MAX_CONCURRENT_API_CALLS=3

# 可选：用于Markdown优化
REFINEMENT_MODEL=deepseek-ai/DeepSeek-V2.5
REFINEMENT_API_URL=https://api.siliconflow.cn/v1/chat/completions
REFINEMENT_API_KEY=your_api_key_here
```

**参数说明：**
- `SF_MAX_WORKERS`：控制所有任务（图像处理、文件处理等）的并行工作线程总数
- `MAX_CONCURRENT_API_CALLS`：限制同时向AI服务发送的API请求数量，以避免触发速率限制
  
这两个参数的区别很重要 - `SF_MAX_WORKERS`决定程序的整体并行度，而`MAX_CONCURRENT_API_CALLS`专门防止同时发出过多API调用。如果遇到API速率限制错误，请尝试减小`MAX_CONCURRENT_API_CALLS`的值，同时保持较高的`SF_MAX_WORKERS`值以保证本地处理效率。

## 使用方法

### 基本用法

```bash
python any2md.py input_file.pdf -o ./output --enable-refinement
```

### 处理整个目录

```bash
python any2md.py ./documents -o ./converted --enable-refinement
```

### 高级选项

```bash
# 使用PyMuPDF提取PDF中的图像
python any2md.py document.pdf --image-extraction-method pymupdf

# 设置自定义DPI进行图像转换
python any2md.py presentation.pptx --dpi 400

# 启用检测区域的可视化
python any2md.py document.pdf --visualize-localization

# 启用DeepSeek优化Markdown
python any2md.py document.pdf --enable-refinement

# 设置自定义工作线程数
python any2md.py document.pdf -w 5
```

### 命令行参数

| 参数 | 描述 |
|----------|-------------|
| `input_path` | PDF/PPT/PPTX文件、PNG文件或包含这些文件的目录 |
| `-o, --output` | 输出目录（默认：./output） |
| `--dpi` | PDF/演示文稿转图像的DPI（默认：300） |
| `--image-extraction-method` | 图像提取方法：qwen_vl（智能定位）或pymupdf（嵌入图像） |
| `--visualize-localization` | 可视化Qwen-VL定位结果 |
| `--enable-refinement` | 启用DeepSeek进行Markdown二次优化 |
| `-w, --workers` | 并行工作线程数 |
| `--no-clean-temp` | 禁用自动清理临时文件 |

## 工作原理

1. **文档处理**：
   - PDF逐页转换为图像
   - PPT/PPTX文件使用LibreOffice转换为PDF，然后转为图像
   - PNG图像直接处理或作为系列处理

2. **内容分析**：
   - 在`qwen_vl`模式下：AI识别文本、图像、表格和其他内容区域
   - 在`pymupdf`模式下：直接提取嵌入的图像

3. **Markdown生成**：
   - 文本内容转换为适当的Markdown，包括标题、列表等
   - LaTeX数学表达式在$或$$分隔符之间保留
   - 表格尽可能转换为Markdown格式
   - 图像被提取、保存并正确引用

4. **优化（可选）**：
   - 生成的Markdown经过分析和改进，提高可读性
   - 修复重复的标题、不一致的格式和其他问题

## 输出结构

对于每个文档，工具创建：
- 与输入文件同名的Markdown文件
- 包含提取图像的资源目录
- 启用时的可选调试可视化文件

## 局限性

- 复杂表格可能作为图像提取，而非Markdown表格
- 处理质量取决于API服务及其模型
- PPT/PPTX处理需要安装LibreOffice

## 许可证

本项目采用MIT许可证 - 详情请参阅[LICENSE](LICENSE)文件。

```
MIT License

Copyright (c) 2023 TsekaLuk (https://github.com/TsekaLuk)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 致谢

此工具利用了多个强大的开源库和AI模型：
- PyMuPDF用于PDF处理
- Qwen2.5-VL用于视觉内容分析
- DeepSeek V2.5用于Markdown优化
- LibreOffice用于演示文稿转换 