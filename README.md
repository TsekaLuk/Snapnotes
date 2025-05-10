# any2md - Convert Any Document to Markdown

<div align="right">
  <a href="README.md">English</a> | <a href="README_CN.md">中文文档</a>
</div>

`any2md` is a powerful Python utility that converts various document formats (PDF, PPT/PPTX, PNG) into high-quality Markdown. It's designed to preserve the original document structure, styling, and content while transforming it into a format that's easy to use with Obsidian.

## Features

- **Multi-format Support**: Process PDFs (both image-only and rich text), PowerPoint presentations (PPT/PPTX), and PNG image series
- **Smart Content Extraction**:
  - **PyMuPDF Mode**: Best for rich text PDFs with extractable content
  - **Qwen-VL Mode**: AI-powered image analysis for content localization in image-only PDFs and presentations
- **Intelligent Processing**:
  - Preserves document structure with proper Markdown hierarchy
  - Converts tables to Markdown format when possible
  - Accurately handles LaTeX mathematical expressions
  - Maintains image references with proper descriptions
- **Performance Optimizations**:
  - Multi-page concurrency for faster processing
  - API auto-retry mechanisms for stability
  - Configurable worker threads
- **Enhanced Workflow**:
  - Automatic image cropping and extraction
  - Image placeholder replacement
  - Optional visualization for debugging
  - Markdown refinement with DeepSeek V2.5 (optional)

## Requirements

### Core Dependencies

```bash
# Using uv (recommended)
uv pip install -r requirements.txt

# Or using pip
pip install requests fitz pdf2image Pillow tqdm python-dotenv tenacity
```

### Optional Dependencies

- **LibreOffice**: Required for PPT/PPTX processing
  - Windows: Download from [libreoffice.org](https://www.libreoffice.org/download/download/)
  - Linux: `sudo apt install libreoffice`
  - macOS: `brew install libreoffice`

- **Additional Python Packages** (Windows-specific):
  ```bash
  # Using uv
  uv pip install comtypes pyautogui
  
  # Or using pip
  pip install comtypes pyautogui
  ```

## Installation

### Install uv (Optional but Recommended)

```bash
# Install uv using pipx (recommended)
pipx install uv

# Or install directly with pip
pip install uv
```

### Install any2md

```bash
# Clone the repository
git clone https://github.com/TsekaLuk/any2md.git
cd any2md

# Install dependencies with uv
uv pip install -r requirements.txt

# Or use pip if you prefer
pip install -r requirements.txt
```

## API Configuration

This tool uses AI services for document analysis. While it's compatible with any API service that implements a compatible interface, we recommend using SiliconFlow's API service as all testing and development were done using this platform.

Create a `.env` file in the same directory with:

```
# SiliconFlow API configuration (recommended)
SF_API_URL=https://api.siliconflow.cn/v1/chat/completions
SF_API_KEY=your_api_key_here
SF_MODEL=Qwen/Qwen2.5-VL-72B-Instruct
SF_MAX_WORKERS=10
MAX_CONCURRENT_API_CALLS=3

# Optional: For markdown refinement
REFINEMENT_MODEL=deepseek-ai/DeepSeek-V2.5
REFINEMENT_API_URL=https://api.siliconflow.cn/v1/chat/completions
REFINEMENT_API_KEY=your_api_key_here
```

**Parameters Explained:**
- `SF_MAX_WORKERS`: Controls the total number of parallel worker threads for all tasks (image processing, file handling, etc.)
- `MAX_CONCURRENT_API_CALLS`: Limits the number of simultaneous API requests to the AI service to avoid rate limiting
  
The difference is important - while `SF_MAX_WORKERS` determines overall parallelization of the program, `MAX_CONCURRENT_API_CALLS` specifically prevents too many API calls at once. If you experience API rate limiting errors, try reducing `MAX_CONCURRENT_API_CALLS` while keeping `SF_MAX_WORKERS` higher for efficient local processing.

## Usage

### Basic Usage

```bash
python any2md.py input_file.pdf -o ./output --enable-refinement
```

### Process an Entire Directory

```bash
python any2md.py ./documents -o ./converted --enable-refinement
```

### Advanced Options

```bash
# Use PyMuPDF extraction method for PDFs
python any2md.py document.pdf --image-extraction-method pymupdf

# Set custom DPI for image conversion
python any2md.py presentation.pptx --dpi 400

# Enable visualization of detected regions
python any2md.py document.pdf --visualize-localization

# Enable markdown refinement with DeepSeek
python any2md.py document.pdf --enable-refinement

# Set custom number of worker threads
python any2md.py document.pdf -w 5
```

### Command Line Arguments

| Argument | Description |
|----------|-------------|
| `input_path` | PDF/PPT/PPTX file, PNG file, or directory containing these files |
| `-o, --output` | Output directory (default: ./output) |
| `--dpi` | DPI for PDF/presentation to image conversion (default: 300) |
| `--image-extraction-method` | Image extraction method: qwen_vl (smart localization) or pymupdf (embedded images) |
| `--visualize-localization` | Visualize Qwen-VL localization results |
| `--enable-refinement` | Enable second-stage Markdown refinement using DeepSeek |
| `-w, --workers` | Number of parallel worker threads |
| `--no-clean-temp` | Disable automatic cleanup of temporary files |

## How It Works

1. **Document Processing**:
   - PDFs are converted to images page by page
   - PPT/PPTX files are converted to PDF using LibreOffice, then to images
   - PNG images are processed directly or as series

2. **Content Analysis**:
   - In `qwen_vl` mode: AI identifies text, images, tables, and other content regions
   - In `pymupdf` mode: Embedded images are extracted directly

3. **Markdown Generation**:
   - Text content is converted to proper Markdown with headings, lists, etc.
   - LaTeX math expressions are preserved between $ or $$ delimiters
   - Tables are converted to Markdown format when possible
   - Images are extracted, saved, and properly referenced

4. **Refinement (Optional)**:
   - The generated Markdown is analyzed and improved for readability
   - Duplicate headers, inconsistent formatting, and other issues are fixed

## Output Structure

For each document, the tool creates:
- A Markdown file with the same name as the input file
- An assets directory containing extracted images
- Optional debug visualization files when enabled

## Limitations

- Complex tables may be extracted as images rather than Markdown tables
- Processing quality depends on the API services and their models
- PPT/PPTX processing requires LibreOffice installation

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

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

## Acknowledgments

This tool leverages several powerful open-source libraries and AI models:
- PyMuPDF for PDF processing
- Qwen2.5-VL for visual content analysis
- DeepSeek V2.5 for Markdown refinement
- LibreOffice for presentation conversion
