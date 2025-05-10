# Snapnotes

任何PDF、PPT/PPTX或PNG系列转Obsidian Markdown，支持Qwen-VL智能图片定位和可选的DeepSeek精炼。

## 功能特点

- 支持任何PDF（纯图片/富文本）、PPT/PPTX文件和PNG系列图片 → 高质量Obsidian Markdown自动转换
- 集成两种图片提取模式：PyMuPDF（适合富文本PDF）、Qwen2.5-VL智能定位（适合纯图片PDF和PPT）
- 支持多页并发、API自动重试、图片裁剪、图片占位符自动替换
- 支持可选的图像定位可视化调试

## 安装

### 使用uv安装（推荐）

[uv](https://github.com/astral-sh/uv) 是一个快速的Python包管理器和虚拟环境工具，比pip快得多。

1. 安装uv:

```bash
curl -sSf https://astral.sh/uv/install.sh | sh
```

或者对于Windows PowerShell:

```pwsh
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
```

2. 创建虚拟环境并安装依赖:

```bash
uv venv
uv pip install -e .
```

3. 对于Windows用户，添加可选依赖:

```bash
uv pip install -e ".[windows]"
```

### 使用传统pip安装

如果您更习惯于使用pip:

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python any2md.py 输入文件.pdf -o ./输出目录
```

或者使用更多选项:

```bash
python any2md.py 输入文件.pdf --dpi 300 --image-extraction-method qwen_vl --visualize-localization --enable-refinement
```

## 环境变量

在项目根目录创建`.env`文件并配置以下变量:

```
SF_API_KEY=你的API密钥
SF_API_URL=https://api.siliconflow.cn/v1/chat/completions
SF_MODEL=Qwen/Qwen2.5-VL-72B-Instruct
```

## 许可证

MIT
