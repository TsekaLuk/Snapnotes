#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
any2md.py

功能：
- 支持任何PDF（纯图片/富文本）、PPT/PPTX文件和PNG系列图片 → 高质量Obsidian Markdown自动转换
- 集成两种图片提取模式：PyMuPDF（适合富文本PDF）、Qwen2.5-VL智能定位（适合纯图片PDF和PPT）
- 支持多页并发、API自动重试、图片裁剪、图片占位符自动替换
- 支持可选的图像定位可视化调试

依赖：requests, fitz, pdf2image, Pillow, tqdm, dotenv, tenacity
"""

import os
import io
import re
import json
import base64
import argparse
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional, Union
import logging
import hashlib
import subprocess 
import shutil
import tempfile
import time
import requests
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path
from dotenv import load_dotenv
import tenacity
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import concurrent.futures
from collections import defaultdict
import logging.handlers
from threading import Semaphore # Import Semaphore

# 检测Windows系统
import platform
IS_WINDOWS = platform.system() == "Windows"

# 在Windows上尝试导入comtypes
if IS_WINDOWS:
    try:
        import comtypes.client
        COMTYPES_SUPPORT = True
    except ImportError:
        COMTYPES_SUPPORT = False
        # 将logger调用移到logger初始化之后

# 尝试导入pyautogui，用于替代方案
try:
    import pyautogui
    PYAUTOGUI_SUPPORT = True
except ImportError:
    PYAUTOGUI_SUPPORT = False

# 日志配置
import logging.handlers
# 创建日志目录
os.makedirs("logs", exist_ok=True)
# 设置日志文件路径
log_file = os.path.join("logs", f"any2md_{time.strftime('%Y%m%d_%H%M%S')}.log")

# 配置日志记录器
logger = logging.getLogger("any2md")
logger.setLevel(logging.DEBUG)

# 控制台处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_format = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
console_handler.setFormatter(console_format)

# 文件处理器 - 更详细的日志
file_handler = logging.handlers.RotatingFileHandler(
    log_file, maxBytes=10*1024*1024, backupCount=5, encoding="utf-8"
)
file_handler.setLevel(logging.DEBUG)  # 文件记录DEBUG及以上级别
file_format = logging.Formatter("%(asctime)s - %(levelname)s - %(module)s:%(lineno)d - %(message)s")
file_handler.setFormatter(file_format)

# 添加处理器到记录器
logger.addHandler(console_handler)
logger.addHandler(file_handler)

logger.info(f"日志文件已创建：{log_file}")

# 在Windows上尝试导入comtypes后给出警告
if IS_WINDOWS and not COMTYPES_SUPPORT:
    logger.warning("comtypes未安装。在Windows上使用PowerPoint导出PPTX将不可用。请使用 pip install comtypes 安装。")

# 加载环境变量
load_dotenv()

# API配置
API_URL = os.getenv("SF_API_URL", "https://api.siliconflow.cn/v1/chat/completions")
API_KEY = os.getenv("SF_API_KEY")
MODEL = os.getenv("SF_MODEL", "Qwen/Qwen2.5-VL-72B-Instruct")
SF_MAX_WORKERS = int(os.getenv("SF_MAX_WORKERS", "10")) # This will be used by ThreadPoolExecutor
MAX_CONCURRENT_API_CALLS = int(os.getenv("MAX_CONCURRENT_API_CALLS", "3")) # Max concurrent calls to the actual API
API_SEMAPHORE = Semaphore(MAX_CONCURRENT_API_CALLS) # Initialize semaphore for API calls

# 新增：用于第二阶段精炼的配置
REFINEMENT_MODEL_NAME = os.getenv("REFINEMENT_MODEL", "deepseek-ai/DeepSeek-V2.5")
REFINEMENT_API_URL = os.getenv("REFINEMENT_API_URL", API_URL) # 默认与主API相同
REFINEMENT_API_KEY = os.getenv("REFINEMENT_API_KEY", API_KEY) # 默认与主API相同

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}
REFINEMENT_HEADERS = {
    "Authorization": f"Bearer {REFINEMENT_API_KEY}",
    "Content-Type": "application/json"
}

# Qwen-VL 图像定位提示词
LOCALIZATION_SYSTEM_PROMPT = """
你是一个专门帮助用户分析文档页面图像并定位其中图示的视觉分析助手。
你的任务是：
1. 确定页面图像中所有包含非文本内容的区域，具体包括：
   a) 非结构化图像内容：如tikz几何图、Python/MATLAB/R/Julia/LaTeX等绘制的高级可视化图表、三维网格曲面图、柱状图、曲线图、示意图、插图、流程图等
   b) 复杂表格：如含有合并单元格、嵌套结构、或超过10*10规格的表格
   c) 对于PPT幻灯片中的图表、SmartArt、图形等非文本元素

2. 对于简单表格（10*10规格以内且不包含复杂嵌套逻辑或合并单元格的表格），请作为纯文本区域直接忽略，**不要**进行图像定位处理。简单表格将由后续主要的文档转写过程直接处理为Markdown表格。

3. 对于复杂表格，**必须**明确将类型标记为"table"，而不是"图示"或其他类型。这对于后续处理非常重要。

4. 对于找到的每个区域，提供其精确的像素坐标边界框（bounding box），边界框要尽可能覆盖当前非结构化图像内容，宁可多框多余的边缘空白区域，也不要遗漏可能处于可视化图像边缘的caption、legend、title、caption、arrow、label等重要图像注释信息；禁止返回与原图几乎同样大小的裁剪区域，保证能够提供模块化的、符合业务实际的图像区域。

5. 对于LaTeX数学文本内容（尤其是数学符号、数学表达式、矩阵表达式等），请作为纯文本区域直接忽略，**不要**进行图像定位处理。LaTeX数学文本内容将由后续主要的文档转写过程直接处理为Markdown数学表达式。

6. 对于水印以及其他非实质性的商业化标识或图案，默认直接忽略，**不要**进行图像定位处理。

返回格式要求：
1. 必须以JSON格式返回结果，包含一个"image_regions"数组。
2. 每个区域对象包含：id, type, bbox([x1,y1,x2,y2]), description, confidence。
3. 对于表格区域，type字段必须包含"table"或"表格"关键词。
4. 对于description字段，描述以简写的自然语言为主，请尽可能在客观简洁与区分度之间取得平衡（一般最多不超过8个字符），禁止使用Unicode字符与特殊字符等转义困难的字符以避免在后续引用资源时中渲染异常。
5. 禁止输出任何额外的分析说明与解释。

示例：
{
  "image_regions":[
    {"id":"region_1","type":"图示","bbox":[50,120,400,350],"description":"二次函数图像","confidence":0.93},
    {"id":"region_2","type":"table","bbox":[100,500,700,800],"description":"复杂数据表格","confidence":0.64}
  ]
}

注意：仅关注明显的图示区域和复杂表格，而不是纯文本区域、数学符号、数学表达式、矩阵表达式或简单表格。
"""
LOCALIZATION_USER_INSTRUCTION = "请分析此页面图像，识别所有重要的非文本内容区域，包括图示（如图形、插图等）和复杂表格。请务必区分'table'类型和其他图像类型，并以指定的JSON格式返回它们的精确位置和简短描述。"

# 转写任务描述
TASK_DESCRIPTION = "文档页面" # 更通用的描述，适用于PDF和PPT

_SYSTEM_PROMPT_BASE = """你是一名专业的文档转写助手，任务是将【%s】完整且准确地转写为【Obsidian Flavored Markdown】格式。
最终目标是生成一份结构清晰、内容无误、便于阅读和在Obsidian中使用的Markdown文档。

请严格遵循以下规范进行转写：
1.  **内容完整性优先**：
    *   必须转写页面上所有可见的文本内容，包括题目、题号、选项、解题步骤、数学公式、图表相关的文字说明（如图注、"见图X"等）。
    *   页眉、页脚中的重要上下文信息（如试卷年份、科目名称、页码等）应予以保留。
    *   保留原有的内容布局空间关系，不要随意调整。对于可能存在的双栏或多栏布局文本，请捕捉转写为正常的阅读排版顺序。
    *   对于可能因文档缺陷而无法辨认的专业文本内容，请尽可能通过上下文推理补全。
    *   **核心要求**：确保不丢失任何来自原始文档的信息。

2.  **Markdown结构化输出**：
    *   **标题层级**：使用Markdown的标题标记（例如 `# 一级标题`，`## 二级标题`，`### 三级标题`）来反映原文的章节、题型（如"选择题"、"解答题"）和大题号的层级结构。例如，'**## 二、填空题**'，'**### 1.**'。
    *   **列表与选项**：
        * 对于纯文本选择题选项（如 A、B、C、D），可以使用Markdown列表格式列出，每个选项占一行。
        * **特别注意**：当选项包含图片时，不要使用列表格式（不要在选项前添加"-"或"*"等列表符号），而是直接使用 "(A)"、"(B)"等标识符，后跟图片markdown语法。例如：`(A) ![描述](图片路径)`。这样可以确保包含图片的选项在Obsidian中正确渲染。
    *   **段落与换行**：自然段落间应有空行。公式和文本的布局应尽可能还原原始逻辑。
    *   **PPT特有元素**：对于PPT幻灯片，标题通常应转换为一级或二级标题，子标题应相应调整。幻灯片中的项目符号列表应保持原有的层级结构转换为Markdown列表。

3.  **数学表达式 (LaTeX)**：
    *   所有数学公式、单个数学符号和数学表达式【必须】准确无误地转写为【合法 LaTeX】。
    *   **行内公式**：使用单美元符号 `$...$` 包裹，例如：函数 $f(x) = ax + b$。
    *   **行间公式（块级公式）**：重要的、独立成行的公式使用双美元符号 `$$...$$` 包裹，且不论原文档中如何组织公式，行间公式必须独立为一行来保证块级公式正确渲染。
    *   **表格结构化**：【优先处理】对于任何识别为表格的区域，【务必优先尝试】将其完整且准确地转写为结构化的【Obsidian Flavored Markdown表格】。请保持原始结构（LaTeX表内符号、行列、角分割逻辑、合并单元格的逻辑、对齐等）。只有在表格内容极度模糊、结构异常复杂到完全无法进行任何有意义的结构化转写，或者该区域明确不是表格（例如，纯粹的图示或照片）时，才可使用 `[图片]` 占位符。避免因微小的瑕疵或略微复杂的布局就放弃结构化尝试。
    *   **特别注意**：仔细区分普通文本中的特殊字符（如 `*`, `_`, `{`, `}`）与LaTeX命令中的这些字符，避免错误转义或格式冲突。

4.  **文本与格式化细节**：
    *   **忽略视觉样式**：原始文档的字体、字号、颜色、具体缩进和精确布局等视觉表现信息通常应忽略，转而专注于内容的语义和逻辑结构。
    *   **语义强调**：如果原文通过【加粗】或【斜体】来强调特定术语、变量、定理名称或关键步骤，请在Markdown中使用相应的 `**加粗**` 或 `*斜体*` 来保留这种语义强调。
    *   **水印处理**：【请务必忽略】页面背景中任何形式的水印（文字、图案、logo等），绝对不要将水印内容转写出来。
    *   **图表处理**：对于【非表格类的纯图像、图示、照片等】，请转写其标题或相关描述文字，并在其位置插入 `[图片]` 占位符。对于【表格类内容】，请严格遵循上述表格处理要求，【优先结构化】。

5.  **转写核心原则**：
    *   【严禁编造、摘要或解释】任何内容。输出必须与原版文档在文字和数学公式上【逐字逐式高度一致】。
    *   专注于"转写"，而非"创作解答"或"理解复述"。

6.  **输出格式要求**：
    *   输出内容【仅能包含纯粹的 Markdown 文本】。不要在Markdown文本的开头或结尾添加任何如 ```markdown ... ``` 这样的代码块包裹。

请仔细分析每个页面的结构和内容，确保转写质量达到最高标准。
"""
SYSTEM_PROMPT = _SYSTEM_PROMPT_BASE % TASK_DESCRIPTION

USER_INSTRUCTION = "请将此页转写为 Obsidian Markdown。若有 LaTeX 数学符号需正确转义。"

REFINEMENT_SYSTEM_PROMPT = """你是一名专业的Obsidian Markdown文档编辑和优化助手。
你将收到一份从PDF或PPT幻灯片逐页转写并初步合并的Markdown文本，这份文本可能存在以下问题：
1. 重复的页眉或页脚：由于逐页转写，原文文档中的页眉页脚可能在合并文本中反复出现。
2. 格式不一致：不同页面转写的内容可能在Markdown格式（如标题、列表）上存在细微差异，同一逻辑层级的标题却被编码为不同的Markdown层次，例如"## 一、选择题 # 二、填空题 ### 三、解答题"应当修正为"# 一、选择题 # 二、填空题 # 三、解答题"。
3. 逻辑中断：较长的段落、题目或解题步骤可能因为跨页而被切断。
4. 图片占位符：文中可能包含 `[图片]` 或 `[图片:一些描述]` 这样的占位符。你的任务是确保这些占位符周围的文本格式正确且逻辑连贯。
5. PPT特有格式问题：对于PPT幻灯片转写的内容，可能存在标题层级不一致或多余的分页标记。
6. 加粗符号：**内容**的**前后最好加一个空格，例如"**【答案】**6"修正为"**【答案】** 6"这样加粗才会正确渲染。
7. 选项格式问题：带图片的选择题选项可能错误地使用了列表格式（如"- (A) ![图片](...)"），这会导致图片无法正确渲染。需要删除选项前的"-"等列表标记，仅保留"(A) ![图片](...)"格式。

你的任务是：
A. **智能识别并移除重复的元数据**。请注意保留有意义的、仅出现一次的文档标题或章节信息，不要误删。
B. **统一和规范化Markdown格式**：
   - 确保标题层级（#，##，### ...）在整个文档中一致且符合逻辑。
   - 统一列表（有序、无序）的格式。
   - 确保数学公式（行内 $...$ 和块级 $$...$$）的 LaTeX 语法正确且一致 (如果存在)。保证$$...$$包裹每个可能存在的\begin{align*}$...\end{align*}环境。
   - 确保行间公式（块级公式）独立为一行输出来保证块级公式正确渲染，即行间公式前后不要有其他内容连续。
   - **修正带图片的选项格式**：找出所有以"- (A)"、"- (B)"等开头且后面跟着图片语法"![..."的行，删除开头的列表符号"-"，确保这些选项直接以"(A)"、"(B)"等开头。
C. **提升内容连贯性**：
   - 尽力识别并逻辑上连接那些因分节（分页或分幻灯片）而被中断的内容。例如，如果一个段落或步骤明显在下一部分继续，请尝试平滑地将它们整合。
   - 修正因分节导致的突兀的换行或段落中断。
D. **保持内容准确性**：在进行格式调整和结构优化的同时，【绝对不能修改原始文本的语义内容或数学公式的准确性】。你的核心是优化结构和移除冗余元信息，而非重写、删减或解释内容，请保证内容的完整性与一致性。
E. **图片引用和占位符处理**：
   - 如果用户在输入中提供了图片元数据（一个描述到路径的映射），请利用这些信息。
   - 当你在文本中遇到 `[图片:描述]` 这样的占位符时，如果元数据中有对应的"描述"，请将其替换转义为正确的Markdown图片链接 `![描述](路径)`。
   - 如果遇到有描述的占位符 `[图片:描述]`，但在元数据中找不到该"描述"，或者遇到无描述的 `[图片]` 占位符，请【删除】这些无法解析的占位符，因为它们在最终文档中无法渲染。
   - 你的主要任务是格式化和连贯性。在删除无法解析的占位符后，请确保周围文本的逻辑和可读性不受影响。
F. **输出纯净文本**： 输出内容【仅能包含纯粹的 Markdown 文本】，不包含任何额外的解释或代码块包裹。请删除任何在Markdown文本的开头或结尾出现的任何如 ```markdown ... ``` 这样的代码块包裹。

请仔细分析输入文本，并输出一份高质量、结构清晰、阅读流畅的最终Obsidian Markdown文档。
"""

# --- Qwen-VL图像定位API调用 ---
@tenacity.retry(
    wait=tenacity.wait_fixed(2),
    stop=tenacity.stop_after_attempt(5),
    retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,)),
    retry_error_callback=lambda retry_state: {"image_regions": []} if retry_state.outcome.failed else retry_state.outcome.result(),
)
def call_vision_api_for_localization(base64_img: str) -> Dict[str, Any]:
    """调用Qwen-VL API进行图像区域定位
    
    参数:
        base64_img: base64编码的图像
        
    返回:
        Dict[str, Any]: 包含image_regions的字典
    """
    try:
        payload = {
            "model": MODEL,
            "messages": [
                {"role": "system", "content": LOCALIZATION_SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/webp;base64,{base64_img}",
                                "detail": "high"
                            }
                        },
                        {"type": "text", "text": LOCALIZATION_USER_INSTRUCTION}
                    ]
                }
            ],
            "stream": False,
            "temperature": 0.1,
            "max_tokens": 4096,
            "enable_thinking": False
        }
        resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=180)
        resp.raise_for_status()
        response_json = resp.json()
        text_content = response_json["choices"][0]["message"]["content"]
        
        # 尝试解析JSON
        try:
            result = json.loads(text_content)
            # 安全检查，确保返回值是一个字典且包含image_regions键
            if not isinstance(result, dict):
                logger.warning(f"API返回值不是字典: {result}")
                return {"image_regions": []}
            if "image_regions" not in result:
                logger.warning(f"API返回值缺少image_regions键: {result}")
                return {"image_regions": []}
            # 确保image_regions是列表，且所有元素都是字典
            image_regions = result["image_regions"]
            if not isinstance(image_regions, list):
                logger.warning(f"image_regions不是列表: {image_regions}")
                return {"image_regions": []}
            # 过滤掉非字典元素
            valid_regions = []
            for region in image_regions:
                if not isinstance(region, dict):
                    logger.warning(f"区域不是字典: {region}")
                    continue
                valid_regions.append(region)
            result["image_regions"] = valid_regions
            logger.info(f"成功解析VLM返回的区域信息，共找到{len(valid_regions)}个区域")
            return result
        except json.JSONDecodeError:
            logger.warning("Response is not pure JSON, attempting to extract JSON from markdown...")
            # 打印部分原始响应内容，帮助理解
            preview_len = min(500, len(text_content))
            logger.info(f"原始响应前{preview_len}个字符: {text_content[:preview_len]}")
            extracted_result = extract_json_from_text(text_content)
            if extracted_result:
                region_count = len(extracted_result.get("image_regions", []))
                logger.info(f"从文本中成功提取出JSON，找到{region_count}个区域")
                return extracted_result
            logger.warning("无法提取有效JSON，返回空结果")
            return {"image_regions": []}
    except Exception as e:
        logger.error(f"调用定位API时出错: {e}")
        return {"image_regions": []}

# 新增提取JSON的辅助函数
def extract_json_from_text(text_content: str) -> Dict[str, Any]:
    """从文本中提取JSON数据
    
    参数:
        text_content: 可能包含JSON的文本
        
    返回:
        Dict[str, Any]: 解析和验证后的JSON数据
    """
    try:
        # 尝试找到JSON代码块
        json_start = text_content.find("```json")
        if json_start != -1:
            logger.info("在Markdown中找到```json代码块")
            json_start = text_content.find("\n", json_start) + 1
            json_end = text_content.find("```", json_start)
            json_str = text_content[json_start:json_end].strip()
            return extract_and_validate_json(json_str)
        
        # 尝试找到任何代码块
        json_start = text_content.find("```")
        if json_start != -1:
            logger.info("在Markdown中找到通用代码块")
            json_start = text_content.find("\n", json_start) + 1
            json_end = text_content.find("```", json_start)
            json_str = text_content[json_start:json_end].strip()
            return extract_and_validate_json(json_str)
        
        # 尝试找到JSON对象
        json_start = text_content.find("{")
        json_end = text_content.rfind("}") + 1
        if json_start != -1 and json_end > json_start:
            logger.info(f"找到可能的JSON对象，位置从{json_start}到{json_end}")
            json_str = text_content[json_start:json_end].strip()
            # 显示一部分提取的JSON字符串
            preview_len = min(200, len(json_str))
            logger.info(f"提取的JSON片段: {json_str[:preview_len]}...")
            return extract_and_validate_json(json_str)
        
        # 所有尝试都失败
        logger.warning("无法从响应中提取JSON，原因：未找到有效的JSON对象或代码块")
        logger.info("尝试搜索特定关键词'image_regions'")
        if "image_regions" in text_content:
            position = text_content.find("image_regions")
            context_start = max(0, position - 100)
            context_end = min(len(text_content), position + 200)
            logger.info(f"找到'image_regions'关键词，上下文: {text_content[context_start:context_end]}")
        
        return {"image_regions": []}
    except Exception as e:
        logger.error(f"提取JSON时出错: {e}")
        return {"image_regions": []}

# --- 工具函数 ---
def convert_pdf_page_to_image(pdf_path: Path, page_number: int, dpi: int = 300) -> Image.Image:
    images = convert_from_path(
        pdf_path, 
        dpi=dpi, 
        first_page=page_number + 1, 
        last_page=page_number + 1,
        fmt='png',
        thread_count=1
    )
    if not images:
        raise ValueError(f"Failed to convert page {page_number} of {pdf_path}")
    return images[0]

def image_to_base64(img: Union[Image.Image, Path], format="PNG") -> str:
    """将图像转换为base64字符串
    
    参数:
        img: PIL图像对象或图像文件路径
        format: 图像格式，默认为PNG
        
    返回:
        str: base64编码的图像字符串
    """
    try:
        if isinstance(img, Path):
            # 打开图像文件
            try:
                img = Image.open(img)
            except Exception as e:
                logger.error(f"无法打开图像文件 {img}: {e}")
                # 返回一个小的空白图像的base64编码以避免进一步的错误
                blank_img = Image.new('RGB', (100, 100), color='white')
                buffer = BytesIO()
                blank_img.save(buffer, format="PNG")
                buffer.seek(0)
                return base64.b64encode(buffer.read()).decode()
        
        # 确保图像使用RGB模式（解决一些PNG文件的RGBA或其他颜色模式问题）
        if img.mode != 'RGB':
            try:
                img = img.convert('RGB')
            except Exception as e:
                logger.warning(f"无法将图像转换为RGB模式: {e}")
                # 继续尝试处理
        
        # 尝试使用不同的格式，如果指定的格式失败
        formats_to_try = [format, "JPEG", "PNG"]
        
        for fmt in formats_to_try:
            try:
                buffer = BytesIO()
                img.save(buffer, format=fmt, quality=95)
                buffer.seek(0)
                return base64.b64encode(buffer.read()).decode()
            except Exception as e:
                logger.warning(f"无法使用{fmt}格式保存图像: {e}")
                # 尝试下一个格式
                continue
        
        # 如果所有格式都失败，尝试重新采样图像
        try:
            resized_img = img.resize((img.width // 2, img.height // 2), Image.LANCZOS)
            buffer = BytesIO()
            resized_img.save(buffer, format="JPEG", quality=85)
            buffer.seek(0)
            logger.warning("已将图像降级处理以适应转换需求")
            return base64.b64encode(buffer.read()).decode()
        except Exception as e:
            logger.error(f"所有图像转换尝试均失败: {e}")
            # 创建一个小的错误提示图像
            error_img = Image.new('RGB', (200, 100), color='red')
            buffer = BytesIO()
            error_img.save(buffer, format="JPEG")
            buffer.seek(0)
            return base64.b64encode(buffer.read()).decode()
            
    except Exception as e:
        logger.error(f"图像转base64过程中发生致命错误: {e}")
        # 返回一个最小的有效base64图像
        return "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC"

# --- Qwen-VL智能定位图片提取 ---
def _crop_regions_and_get_paths(
    page_image: Image.Image, 
    page_number: int, 
    image_regions: List[Dict], 
    output_dir: Path, 
    visualize: bool = False
) -> Tuple[List[str], List[Dict]]:
    """
    Crops regions from a given page image based on pre-determined image_regions.
    
    Args:
        page_image: PIL Image of the full page.
        page_number: 0-indexed page number.
        image_regions: List of dictionaries, each describing a region with 'bbox'.
        output_dir: Directory to save cropped images.
        visualize: Whether to save a visualization for debugging.
    
    Returns:
        - List[str]: Relative paths to the saved images.
        - List[Dict]: The input image_regions list, with an 'image_path' key added to each region dict.
    """
    saved_image_relative_paths = []
    
    if not image_regions:
        logger.debug(f"No regions to crop for page {page_number+1}.")
        return saved_image_relative_paths, []
    
    # Make sure output_dir exists
    output_dir.mkdir(parents=True, exist_ok=True)
    
    updated_regions = []
    
    for i, region in enumerate(image_regions):
        if not isinstance(region, dict):
            logger.warning(f"Region {i} for page {page_number+1} is not a dictionary: {region}")
            updated_regions.append(region) # Keep original if not processable
            continue
        
        if "bbox" not in region:
            logger.warning(f"Region {i} for page {page_number+1} has no bbox: {region}")
            updated_regions.append(region)
            continue
        
        bbox = region["bbox"]
        if not isinstance(bbox, list) or len(bbox) != 4:
            logger.warning(f"Invalid bbox format for region {i} on page {page_number+1}: {bbox}")
            updated_regions.append(region)
            continue
        
        try:
            # Get region information
            x1, y1, x2, y2 = bbox
            
            # Make sure coordinates are within image bounds and in the right order
            width, height = page_image.size
            x1 = max(0, min(x1, width))
            y1 = max(0, min(y1, height))
            x2 = max(0, min(x2, width))
            y2 = max(0, min(y2, height))
            
            # Ensure x1 < x2 and y1 < y2
            if x1 > x2: x1, x2 = x2, x1
            if y1 > y2: y1, y2 = y2, y1
            
            # Skip very small regions (could be noise)
            if x2 - x1 < 5 or y2 - y1 < 5:
                logger.warning(f"Region {i} on page {page_number+1} is too small: {bbox}")
                updated_regions.append(region)
                continue
                
            # Create a copy of the region dict to add an image_path key
            current_region_copy = region.copy()
            
            # Crop the image
            region_image = page_image.crop((x1, y1, x2, y2))
            
            # Get or generate a descriptive name for the region
            region_type = region.get("type", "区域")
            region_description = region.get("description", region_type)
            safe_desc = "".join(c if c.isalnum() or c in "- " else "_" for c in region_description).strip().replace(" ", "_")
            if not safe_desc:
                safe_desc = f"region_{i+1}"
            
            # 修改：直接使用 output_dir 而不是创建子目录
            image_filename = f"page{page_number+1}_{safe_desc}_{i+1}.png"
            image_path = output_dir / image_filename
            
            # Save the cropped image
            region_image.save(image_path)
            
            # 必须包含assets目录名，确保从Markdown文件可以正确引用
            # 获取output_dir的名称（通常是assets目录名）
            output_dir_name = output_dir.name
            # 构建完整的相对路径
            rel_path = f"{output_dir_name}/{image_filename}".replace("\\", "/")
            saved_image_relative_paths.append(rel_path)
            
            # Add path to region info - 也要使用完整相对路径
            current_region_copy["image_path"] = rel_path
            updated_regions.append(current_region_copy)
            logger.debug(f"Saved region {i+1} from page {page_number+1} to {rel_path}")
        except Exception as e:
            logger.error(f"Error cropping region {i+1} from page {page_number+1}: {e}")
            updated_regions.append(region) # Add original region if save fails
    
    if visualize and any(isinstance(r, dict) and r.get("bbox") for r in image_regions): # only save if there were valid regions to draw
        debug_path = output_dir / f"page{page_number+1}_regions_debug.png"
        debug_img = page_image.copy()
        draw = ImageDraw.Draw(debug_img)
        for i, region in enumerate(image_regions):
            if isinstance(region, dict) and region.get("bbox"):
                x1, y1, x2, y2 = region["bbox"]
                draw.rectangle([x1, y1, x2, y2], outline="red", width=2)
                # Try to use TrueType font but fall back to default if not available
                try:
                    font = ImageFont.truetype("arial.ttf", 12)
                except IOError:
                    font = ImageFont.load_default()
                draw.text((x1, y1-10), str(i+1), fill="red", font=font)
        debug_img.save(debug_path)
        logger.debug(f"Saved visualization of {len(image_regions)} regions to {debug_path}")
    
    return saved_image_relative_paths, updated_regions

def extract_images_from_page_with_qwen_vl(
    pdf_path: Path, 
    page_number: int, 
    output_dir: Path, 
    dpi: int = 300, 
    visualize: bool = False
) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    page_image = convert_pdf_page_to_image(pdf_path, page_number, dpi)
    base64_image = image_to_base64(page_image)
    localization_result = call_vision_api_for_localization(base64_image)
    image_regions = localization_result.get("image_regions", [])
    if not image_regions:
        logger.info(f"No image regions detected on page {page_number} of {pdf_path.name}")
        return []
    if visualize:
        debug_image = page_image.copy()
        draw = ImageDraw.Draw(debug_image)
    
    # 获取output_dir的名称（通常是assets目录名）
    output_dir_name = output_dir.name
    saved_image_paths = []
    
    for i, region in enumerate(image_regions):
        bbox = region.get("bbox")
        if not bbox or len(bbox) != 4:
            continue
        x1, y1, x2, y2 = [int(coord) for coord in bbox]
        x1 = max(0, x1)
        y1 = max(0, y1)
        x2 = min(page_image.width, x2)
        y2 = min(page_image.height, y2)
        if x2 <= x1 or y2 <= y1 or (x2 - x1) < 20 or (y2 - y1) < 20:
            continue
        cropped_image = page_image.crop((x1, y1, x2, y2))
        # 获取区域描述，默认为区域编号
        region_desc = region.get("description", f"region_{i+1}")
        safe_desc = "".join(c if c.isalnum() or c in "- " else "_" for c in region_desc).strip()
        filename = f"page{page_number+1}_{safe_desc}_{i+1}.png"
        output_path = output_dir / filename
        cropped_image.save(output_path, format="PNG")
        
        # 修复：返回包含assets目录名的相对路径
        rel_path = f"{output_dir_name}/{filename}".replace("\\", "/")
        saved_image_paths.append(rel_path)
        
        if visualize:
            draw.rectangle([x1, y1, x2, y2], outline="red", width=3)
            draw.text((x1, y1-15), f"{i+1}", fill="red")
    if visualize and image_regions:
        debug_path = output_dir / f"page{page_number+1}_regions_debug.png"
        debug_image.save(debug_path)
        logger.info(f"Saved visualization with bounding boxes to {debug_path}")
    return saved_image_paths

# --- PyMuPDF图片提取 ---
def extract_images_from_pdf_with_pymupdf(pdf_path: Path, output_dir: Path, min_dimension: int = 32) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    total_pages = doc.page_count
    candidate_images = [] # 修改： 使用列表存储元组 (page_num, img_index, img_bytes, img_ext, img_hash, width, height, xref)
    image_hashes_on_pages = defaultdict(list) # 修改：存储哈希对应的页码列表
    
    page_image_counter = defaultdict(int) # 新增：跟踪每页已保存的图片数量，用于生成唯一文件名，如果PyMuPDF也按页保存
    
    saved_image_paths_relative = [] # 修改：存储相对路径

    for page_num in range(total_pages):
        page = doc.load_page(page_num)
        img_list = page.get_images(full=True)
        for img_index, img_info in enumerate(img_list): # 修改变量名
            xref = img_info[0]
            if xref == 0: continue

            try:
                base_image = doc.extract_image(xref)
            except Exception as e:
                logger.debug(f"Page {page_num + 1}, Img {img_index} (xref: {xref}): Failed to extract image - {e}")
                continue
                
            img_bytes = base_image["image"] # 修改变量名
            img_ext = base_image["ext"] # 修改变量名
            
            try:
                image = Image.open(io.BytesIO(img_bytes))
                width, height = image.size
            except Exception as e:
                logger.warning(f"Could not open image with xref {xref} on page {page_num+1}: {e}")
                continue


            if width < min_dimension or height < min_dimension:
                logger.debug(f"Page {page_num + 1}, Img {img_index} (xref: {xref}): Filtered out by size ({width}x{height}).")
                continue
            
            img_hash = hashlib.md5(img_bytes).digest() # 使用MD5哈希，更稳定
            candidate_images.append((xref, page_num, width, height, img_bytes, img_ext, img_hash)) #调整顺序和内容
            image_hashes_on_pages[img_hash].append(page_num) #存储哈希对应的页码列表
    
    doc.close()

    if not candidate_images:
        logger.info(f"No candidate images found by PyMuPDF for {pdf_path.name} after size filter.")
        return []

    logger.info(f"PyMuPDF found {len(candidate_images)} candidate images for {pdf_path.name} after size filtering.")
    
    final_images_to_save = [] # List of (page_num, img_index_on_page, image_bytes, img_ext)
    
    # Sort candidate images by page number and then by original xref (to maintain some order)
    candidate_images.sort(key=lambda x: (x[1], x[0])) # x[1] is page_num, x[0] is xref

    for xref, page_num, width, height, img_bytes, img_ext, img_hash in candidate_images: # 解包调整
        num_occurrences = len(image_hashes_on_pages[img_hash])
        
        # 借鉴pdf2obsidian.py的过滤逻辑
        is_watermark = num_occurrences > (total_pages * 0.5) if total_pages > 0 else False
        
        if is_watermark:
            logger.debug(f"Page {page_num + 1} (xref: {xref}): Filtered out as watermark. Occurrences: {num_occurrences}/{total_pages}.")
            continue

        current_img_index_on_page = page_image_counter[page_num]
        page_image_counter[page_num] = current_img_index_on_page + 1
        final_images_to_save.append((page_num, current_img_index_on_page, img_bytes, img_ext))

    if not final_images_to_save:
        logger.info(f"No images to save for {pdf_path.name} after all PyMuPDF filters (size, watermark).")
        return []
    
    logger.info(f"{len(final_images_to_save)} images selected for saving by PyMuPDF after all filters.")

    for page_num, img_idx_on_page, img_bytes, img_ext in final_images_to_save:
        # 生成与pdf2obsidian.py一致的文件名格式
        image_filename = f"page{page_num + 1}_img{img_idx_on_page + 1}.{img_ext}"
        image_filename_abs = output_dir / image_filename
        
        try:
            with open(image_filename_abs, "wb") as img_file:
                img_file.write(img_bytes)
            
            relative_image_path = str(Path(output_dir.name) / image_filename).replace("\\", "/") #确保asset_dir.name
            saved_image_paths_relative.append(relative_image_path) #存储相对路径
        except Exception as e:
            logger.warning(f"Could not save image {image_filename} to {output_dir.name} using PyMuPDF: {e}")

    if saved_image_paths_relative:
        logger.info(f"PyMuPDF successfully filtered and saved {len(saved_image_paths_relative)} images for {pdf_path.name} to {output_dir.name}.")
    elif candidate_images:
        logger.info(f"All {len(candidate_images)} PyMuPDF candidate images for {pdf_path.name} were filtered out.")
    else:
        logger.info(f"No suitable images found by PyMuPDF to extract from {pdf_path.name} after initial filtering.")
        
    return saved_image_paths_relative # 返回相对路径列表

# --- VLM转写API调用 ---
@tenacity.retry(
    wait=tenacity.wait_exponential(multiplier=2, min=5, max=60), # 修改：增加初始等待时间至5秒，最大等待时间至60秒，倍率为2
    stop=tenacity.stop_after_attempt(6), # 修改：减少最大重试次数以避免过长等待
    retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,)), # Retry on any request exception
    reraise=True, # 确保重试后仍会抛出原始异常
)
def call_vlm_for_markdown(base64_img: str, custom_instruction: str = None) -> str:
    """
    给定一个 base64 Image 调用 Qwen-VL 并返回 Markdown
    可选参数 custom_instruction 允许注入自定义指令（如图片区域信息）
    """
    instruction = custom_instruction if custom_instruction else USER_INSTRUCTION
    
    # 全局变量用于追踪API调用状态，如果不存在则初始化
    global _consecutive_api_failures, _last_error_code, _last_error_time
    if '_consecutive_api_failures' not in globals():
        _consecutive_api_failures = 0
        _last_error_code = None
        _last_error_time = 0
    
    # 自适应延迟逻辑
    current_time = time.time()
    time_since_last_error = current_time - _last_error_time
    
    # 如果有连续失败且时间间隔过短，增加额外等待
    if _consecutive_api_failures > 0 and time_since_last_error < 5:
        # 根据连续失败次数和上次错误类型动态计算等待时间
        adaptive_wait = min(5 * _consecutive_api_failures, 30)
        
        # 对503错误特殊处理，给服务器更多恢复时间
        if _last_error_code == 503:
            adaptive_wait = min(10 * _consecutive_api_failures, 60)
        
        logger.info(f"检测到连续失败 ({_consecutive_api_failures}次), 自适应等待 {adaptive_wait}秒后重试")
        time.sleep(adaptive_wait)
    
    logger.info(f"等待API信号量 (可用: {API_SEMAPHORE._value}/{MAX_CONCURRENT_API_CALLS})")
    with API_SEMAPHORE: # Acquire semaphore before making an API call
        logger.info(f"已获取API信号量，开始调用VLM API进行Markdown转写, 指令长度: {len(instruction)}")
        start_time = time.time()
        
        payload = {
            "model": MODEL,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/webp;base64,{base64_img}",
                                "detail": "high"
                            }
                        },
                        {"type": "text", "text": instruction}
                    ]
                }
            ],
            "stream": False,
            "temperature": 0.1,
            "max_tokens": 4096,
            "enable_thinking": False
        }
    
        try:
            # 添加更长的超时时间，以处理大图像或复杂请求
            logger.info(f"发送API请求到URL: {API_URL}, 使用模型: {MODEL}")
            resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=600)
            
            # 记录响应时间
            elapsed_time = time.time() - start_time
            logger.info(f"VLM API响应用时: {elapsed_time:.2f}秒, 状态码: {resp.status_code}")
            
            # 成功响应，重置连续失败计数
            _consecutive_api_failures = 0
            _last_error_code = None
            
            # 处理HTTP错误
            resp.raise_for_status() # This will raise an HTTPError for bad responses (4xx or 5xx)
            
            # 解析响应
            try:
                response_json = resp.json()
                logger.info(f"成功获取API响应JSON，开始提取内容")
            except json.JSONDecodeError as e:
                logger.error(f"API响应解析错误: {e}, 响应内容: {resp.text[:200]}...")
                raise
            
            # 检查响应格式
            if "choices" not in response_json or len(response_json["choices"]) == 0:
                logger.error(f"API返回无效响应: 缺少choices字段, 响应: {response_json}")
                return f"[API返回格式错误: 缺少choices字段]"
            
            if "message" not in response_json["choices"][0]:
                logger.error(f"API返回无效响应: 缺少message字段, 响应: {response_json['choices'][0]}")
                return f"[API返回格式错误: 缺少message字段]"
            
            if "content" not in response_json["choices"][0]["message"]:
                logger.error(f"API返回无效响应: 缺少content字段, 响应: {response_json['choices'][0]['message']}")
                return f"[API返回格式错误: 缺少content字段]"
            
            text_content = response_json["choices"][0]["message"]["content"]
            
            # 检查内容是否为空
            if not text_content or text_content.strip() == "":
                logger.warning("API返回了空内容")
                return "[API返回了空内容]"
            
            # 检查内容长度是否符合预期
            if len(text_content) < 10:  # 任意小的阈值，用于检测异常短的回复
                logger.warning(f"API返回内容可能异常简短: {text_content}")
            
            logger.debug(f"成功获取VLM转写结果，内容长度: {len(text_content)}")
            return text_content
        
        except requests.exceptions.Timeout:
            # 更新失败状态
            _consecutive_api_failures += 1
            _last_error_code = "timeout"
            _last_error_time = time.time()
            
            logger.error(f"VLM API调用超时 (超过600秒), 连续失败次数: {_consecutive_api_failures}")
            # 超时通常意味着服务端负载过高，增加等待时间
            adaptive_wait = min(8 * _consecutive_api_failures, 40)
            logger.info(f"超时后延迟 {adaptive_wait} 秒")
            time.sleep(adaptive_wait)
            raise # Reraise for tenacity to handle
            
        except requests.exceptions.ConnectionError as e:
            # 更新失败状态
            _consecutive_api_failures += 1
            _last_error_code = "connection"
            _last_error_time = time.time()
            
            logger.error(f"VLM API连接错误: {e}, 连续失败次数: {_consecutive_api_failures}")
            # 连接错误可能是网络或服务端问题，适当增加等待
            adaptive_wait = min(5 * _consecutive_api_failures, 30)
            logger.info(f"连接错误后延迟 {adaptive_wait} 秒")
            time.sleep(adaptive_wait)
            raise # Reraise for tenacity to handle
            
        except requests.exceptions.HTTPError as e:
            # 更新失败状态
            _consecutive_api_failures += 1
            _last_error_code = e.response.status_code if e.response else "http"
            _last_error_time = time.time()
            
            logger.error(f"VLM API HTTP错误: {e}, 状态码: {_last_error_code}, 连续失败次数: {_consecutive_api_failures}")
            
            if e.response and e.response.status_code == 429: # Rate limit
                logger.warning("可能遇到API速率限制 (429)")
                # 对于429速率限制错误，指数级增加等待
                adaptive_wait = min(15 * _consecutive_api_failures, 120)
                logger.info(f"速率限制后延迟 {adaptive_wait} 秒")
                time.sleep(adaptive_wait)
                
            elif e.response and e.response.status_code == 503: # Service unavailable
                logger.warning("服务不可用 (503)")
                # 对于503服务不可用错误，使用更长的等待时间
                adaptive_wait = min(20 * _consecutive_api_failures, 180)
                logger.info(f"服务不可用后延迟 {adaptive_wait} 秒")
                time.sleep(adaptive_wait)
                
            else:
                # 其他HTTP错误，基础等待
                adaptive_wait = min(5 * _consecutive_api_failures, 30)
                logger.info(f"HTTP错误后延迟 {adaptive_wait} 秒")
                time.sleep(adaptive_wait)
                
            raise # Reraise for tenacity to handle
            
        except Exception as e:
            # 更新失败状态
            _consecutive_api_failures += 1
            _last_error_code = "unknown"
            _last_error_time = time.time()
            
            logger.error(f"VLM API调用过程中发生未知错误: {type(e).__name__} - {e}, 连续失败次数: {_consecutive_api_failures}")
            # 未知错误使用中等等待时间
            adaptive_wait = min(7 * _consecutive_api_failures, 35)
            logger.info(f"未知错误后延迟 {adaptive_wait} 秒")
            time.sleep(adaptive_wait)
            raise # Reraise for tenacity to handle

# 新增：用于第二阶段精炼的LLM调用函数
@tenacity.retry(wait=tenacity.wait_fixed(5), stop=tenacity.stop_after_attempt(3),
                retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,)))
def call_llm_for_refinement(markdown_text: str, model_to_use: str, image_metadata: Optional[Dict[str, str]] = None) -> str:
    """
    给定一个 Markdown 文本调用大语言模型（如 DeepSeek）并返回精炼后的 Markdown
    新增参数 image_metadata 允许传入图片描述到路径的映射，用于辅助LLM处理图片引用。
    """
    user_content = markdown_text
    if image_metadata:
        # 将图片元数据格式化为文本，附加到用户输入内容的末尾，或以特定方式提示LLM
        metadata_prompt = "\n\n--- 图片元数据 (用于上下文参考, 并非必须使用) ---\n"
        for desc, path in image_metadata.items():
            metadata_prompt += f"- 描述: \"{desc}\" -> 路径: \"{path}\"\n"
        user_content += metadata_prompt
        logger.info(f"向精炼模型传入了 {len(image_metadata)} 条图片元数据。")
        logger.debug(f"传入精炼模型的图片元数据详情:\n{metadata_prompt}")

    payload = {
        "model": model_to_use,
        "messages": [
            {"role": "system", "content": REFINEMENT_SYSTEM_PROMPT},
            {"role": "user", "content": user_content}
        ],
        "stream": False,
        "temperature": 0.2,
        "max_tokens": 4096, # 保持足够大的max_tokens以处理元数据和文本
        "enable_thinking": False
    }
    # 使用 REFINEMENT_API_URL 和 REFINEMENT_HEADERS，添加超时设置
    logger.info(f"调用精炼API ({model_to_use})，输入内容长度(含元数据): {len(user_content)}")
    resp = requests.post(REFINEMENT_API_URL, headers=REFINEMENT_HEADERS, json=payload, timeout=300)
    resp.raise_for_status()
    refined_text = resp.json()["choices"][0]["message"]["content"]
    logger.info(f"精炼API ({model_to_use}) 返回内容长度: {len(refined_text)}")
    return refined_text

# --- 新增：PPTX 直接转图片 ---
def convert_presentation_to_images_with_libreoffice(presentation_path: Path, output_folder: Path, dpi: int = 300) -> List[Path]:
    """
    使用LibreOffice将演示文稿文件（PPT/PPTX）转换为PDF，然后提取为图像
    返回生成的图像路径列表
    """
    logger.info(f"使用LibreOffice转换演示文稿为图像: {presentation_path.name}")
    
    # 确保输出目录存在
    temp_slides_folder = output_folder / f"{presentation_path.stem}_slides"
    temp_slides_folder.mkdir(parents=True, exist_ok=True)
    
    # 创建临时目录用于LibreOffice输出
    with tempfile.TemporaryDirectory(prefix=f"libreoffice_pres_{presentation_path.stem}_") as temp_dir:
        # 确定LibreOffice可执行文件路径
        soffice_path = None
        common_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"D:\Program Files\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/opt/libreoffice/program/soffice"
        ]
        
        # 检查常见路径
        for path in common_paths:
            if os.path.exists(path):
                soffice_path = path
                break
        
        # 如果没有找到确切路径，使用系统PATH中的soffice命令
        if not soffice_path:
            soffice_path = "soffice"
        
        temp_dir_path = Path(temp_dir)
        slide_image_paths = []
        temp_pdf_path = None  # 记录临时PDF文件路径
        
        # 直接转换为PDF
        pdf_cmd = [
            soffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(temp_dir_path),
            str(presentation_path)
        ]
        
        try:
            logger.info(f"执行命令: {' '.join(pdf_cmd)}")
            result = subprocess.run(pdf_cmd, capture_output=True, text=True, check=False, timeout=120) # Added timeout
            
            if result.returncode != 0:
                logger.error(f"LibreOffice转换为PDF失败 ({presentation_path.name}): {result.stderr}")
                return []
            
            pdf_files = list(temp_dir_path.glob("*.pdf"))
            if not pdf_files:
                logger.error(f"未找到LibreOffice为 {presentation_path.name} 生成的PDF文件")
                return []
            
            temp_pdf_path = pdf_files[0]
            
            try:
                doc = fitz.open(str(temp_pdf_path))
                total_pages = doc.page_count
                logger.info(f"从 {presentation_path.name} 生成的PDF文件共有 {total_pages} 页")
                
                for page_num in range(total_pages):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(dpi=dpi)
                    slide_filename = f"slide_{page_num+1:03d}.png"
                    slide_path = temp_slides_folder / slide_filename
                    pix.save(str(slide_path))
                    slide_image_paths.append(slide_path)
                    # Reduced log verbosity here, logged once after loop
                doc.close()
                logger.info(f"从PDF为 {presentation_path.name} 提取并保存了 {len(slide_image_paths)} 张幻灯片图像到 {temp_slides_folder}")
                
                # Copy the intermediate PDF to output_folder for potential inspection/debugging
                # This can be controlled by a flag like --no-clean-temp later if needed
                output_pdf_path = output_folder / f"{presentation_path.stem}_intermediate.pdf"
                shutil.copy2(temp_pdf_path, output_pdf_path)
                logger.info(f"已保存 {presentation_path.name} 的中间PDF文件: {output_pdf_path.name}")
                
                return slide_image_paths
            except Exception as e:
                logger.error(f"从PDF提取图像失败 ({presentation_path.name}): {e}")
                return []
        
        except subprocess.TimeoutExpired:
            logger.error(f"LibreOffice转换 ({presentation_path.name}) 超时。")
            return []
        except Exception as e:
            logger.error(f"调用LibreOffice时出错 ({presentation_path.name}): {e}")
            return []

# --- 修改：PPTX处理主函数 ---
def process_presentation(
    presentation_path: Path, 
    output_dir: Path,
    dpi: int = 300,
    visualize_localization: bool = False,
    enable_refinement: bool = False
) -> bool:
    """
    直接处理演示文稿文件（PPT/PPTX），提取幻灯片为图片并转换为Markdown
    """
    # 检查是否可以处理演示文稿 (comtypes is only for Windows + MS Office automation, LibreOffice is primary)
    # The main check for LibreOffice availability is now in `process_document` / `main`.
    
    presentation_stem = presentation_path.stem
    safe_stem = "".join(c if c.isalnum() or c in "- " else "_" for c in presentation_stem).strip().replace(" ", "_")
    # Assets dir for this specific presentation
    assets_dir_for_presentation = output_dir / f"{safe_stem}_assets"
    assets_dir_for_presentation.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"处理演示文稿《{presentation_path.name}》")
    
    # 导出幻灯片为图像 using the generalized function
    # Pass assets_dir_for_presentation as the base for slide images and intermediate PDF
    slide_images = convert_presentation_to_images_with_libreoffice(presentation_path, assets_dir_for_presentation, dpi)
    if not slide_images:
        logger.error(f"无法从演示文稿文件 '{presentation_path.name}' 提取幻灯片图像，处理终止。")
        return False
    
    # 使用VLM处理每张幻灯片图像
    total_slides = len(slide_images)
    logger.info(f"开始处理 {total_slides} 张幻灯片 (从 {presentation_path.name})...")
    
    slide_md_results = [None] * total_slides
    # Store {index: regions_data} from localization for each slide
    slide_localized_regions_map = {} 

    with ThreadPoolExecutor(max_workers=SF_MAX_WORKERS) as executor:
        localization_futures = {} # future -> slide_idx
        transcription_futures = {} # future -> slide_idx
        
        # Prepare image data for all slides first (base64 encoding)
        slide_image_data_list = [] # List of dicts: {id: int, path: Path, b64: str, error: bool}
        for idx, slide_image_path_obj in enumerate(slide_images):
            entry = {"id": idx, "path": slide_image_path_obj, "b64": None, "error": False}
            try:
                with open(slide_image_path_obj, "rb") as img_file:
                    entry["b64"] = base64.b64encode(img_file.read()).decode("utf-8")
            except Exception as e:
                logger.error(f"读取或编码幻灯片图像 {slide_image_path_obj.name} 失败: {e}")
                entry["error"] = True
                slide_md_results[idx] = f"[第{idx+1}张幻灯片图像准备失败: {e}]"
            slide_image_data_list.append(entry)

        # Phase 1: Submit localization tasks (Qwen-VL for slides can also identify text/table regions)
        # For presentations, localization can help guide the VLM even if we don't crop separate sub-images.
        logger.info(f"演示文稿 {presentation_path.name}: 提交 {total_slides} 张幻灯片的区域检测任务...")
        for idx, data_entry in enumerate(slide_image_data_list):
            if data_entry["error"] or not data_entry["b64"]:
                continue
            future = executor.submit(call_vision_api_for_localization, data_entry["b64"])
            localization_futures[future] = idx
        
        # Phase 1.5: Process localization and queue transcriptions
        logger.info(f"演示文稿 {presentation_path.name}: 处理区域检测结果并提交转写任务...")
        for loc_future in tqdm(as_completed(localization_futures), total=len(localization_futures), desc=f"幻灯片区域检测 ({presentation_path.name})"):
            slide_idx = localization_futures[loc_future]
            data_entry = slide_image_data_list[slide_idx]
            if data_entry["error"]: continue

            try:
                localization_result = loc_future.result()
                regions = localization_result.get("image_regions", [])
                slide_localized_regions_map[slide_idx] = regions

                # 新增：裁剪检测到的图像区域
                extracted_img_rel_paths = []
                updated_regions = regions
                if regions:
                    try:
                        img_pil = Image.open(data_entry["path"])
                        # 使用与PDF相同的区域裁剪函数
                        extracted_img_rel_paths, updated_regions = _crop_regions_and_get_paths(
                            img_pil,
                            slide_idx,
                            regions,
                            assets_dir_for_presentation,
                            visualize_localization
                        )
                        slide_localized_regions_map[slide_idx] = updated_regions
                        logger.info(f"幻灯片 {slide_idx+1}: 裁剪了 {len(extracted_img_rel_paths)} 个区域")
                    except Exception as crop_e:
                        logger.warning(f"幻灯片 {slide_idx+1} 区域裁剪失败: {crop_e}")
                    
                # Prepare instruction for transcription
                slide_instruction = f"请转写这张幻灯片中的所有文字内容，并保持原始格式。这是演示文稿 '{presentation_stem}' 的第 {slide_idx+1} 张幻灯片（共 {total_slides} 张）。"
                
                if regions:
                    slide_instruction += " 此幻灯片识别出以下内容区域：\n"
                    for j, region_data in enumerate(regions):
                        if isinstance(region_data, dict):
                            region_type = region_data.get("type", "内容")
                            region_desc = region_data.get("description", f"区域{j+1}")
                            slide_instruction += f"  - {region_type}: {region_desc}\n"
                    
                    slide_instruction += "\n当你在转写中提到这些区域时，请使用[图片:区域描述]的格式作为占位符，我将在后续处理中自动替换为对应的图片引用。例如，如果你需要引用'图表:季度销售数据'，请在文本中插入[图片:季度销售数据]。"
                
                trans_future = executor.submit(call_vlm_for_markdown, data_entry["b64"], slide_instruction)
                transcription_futures[trans_future] = slide_idx
                logger.info(f"PNG {slide_idx+1} ({data_entry['path'].name}): 已提交转写任务，当前队列长度: {len(transcription_futures)}")
                
                # 控制同时提交的转写任务数量，避免API过载
                active_futures = [f for f in transcription_futures.keys() if not f.done() and not f.cancelled()]
                if len(active_futures) >= MAX_CONCURRENT_API_CALLS:
                    # 等待至少一个任务完成后再继续
                    logger.debug(f"已达到最大并发API调用数量({MAX_CONCURRENT_API_CALLS})，等待任务完成...")
                    try:
                        completed_future = as_completed(active_futures, timeout=None).__next__()
                        slide_idx_completed = transcription_futures[completed_future]
                        try:
                            result = completed_future.result()
                            slide_md_results[slide_idx_completed] = result
                            logger.debug(f"幻灯片 {slide_idx_completed+1} 转写完成，继续提交新任务")
                            # 添加额外等待，避免触发API速率限制
                            time.sleep(2)
                        except Exception as e:
                            logger.error(f"幻灯片 {slide_idx_completed+1} 在等待时处理失败: {e}")
                            slide_md_results[slide_idx_completed] = f"[第{slide_idx_completed+1}张幻灯片转写失败: {e}]"
                    except Exception as wait_error:
                        logger.error(f"等待任务完成时出错: {wait_error}")
                        time.sleep(5)  # 出错时增加更长等待
                        
            except Exception as e:
                logger.error(f"幻灯片 {slide_idx+1}: 区域检测后处理或转写任务提交失败: {e}")
                slide_md_results[slide_idx] = f"[第{slide_idx+1}张幻灯片后处理失败: {e}]"

        # Phase 2: Collect transcription results
        logger.info(f"演示文稿 {presentation_path.name}: 收集转写结果...")
        collected_count = 0
        for trans_future in tqdm(as_completed(transcription_futures), total=len(transcription_futures), desc=f"VLM转写幻灯片 ({presentation_path.name})"):
            slide_idx = transcription_futures[trans_future]
            data_entry = slide_image_data_list[slide_idx]
            # Skip if error already recorded during prep or post-localization
            if slide_md_results[slide_idx] is not None and "失败" in slide_md_results[slide_idx]:
                logger.debug(f"幻灯片 {slide_idx+1}: 跳过已标记失败的幻灯片")
                continue
            try:
                logger.debug(f"幻灯片 {slide_idx+1}: 尝试获取转写结果...")
                md_text = trans_future.result(timeout=180)  # 添加超时参数
                if not md_text or md_text.strip() == "":
                    logger.warning(f"幻灯片 {slide_idx+1} 转写内容为空，设置为默认内容")
                    slide_md_results[slide_idx] = f"[幻灯片 {slide_idx+1} 内容为空]"
                else:
                    slide_md_results[slide_idx] = md_text
                    collected_count += 1
                    logger.debug(f"幻灯片 {slide_idx+1} 转写完成，当前已完成: {collected_count}/{len(transcription_futures)}")
            except TimeoutError:
                logger.error(f"幻灯片 {slide_idx+1} 转写结果获取超时")
                slide_md_results[slide_idx] = f"[第{slide_idx+1}张幻灯片转写超时]"
            except concurrent.futures.CancelledError:
                logger.error(f"幻灯片 {slide_idx+1} 转写任务被取消")
                slide_md_results[slide_idx] = f"[第{slide_idx+1}张幻灯片转写任务被取消]"
            except requests.exceptions.RequestException as e:
                logger.error(f"幻灯片 {slide_idx+1} 网络请求错误: {e}")
                slide_md_results[slide_idx] = f"[第{slide_idx+1}张幻灯片网络请求错误: {e}]"
            except json.JSONDecodeError as e:
                logger.error(f"幻灯片 {slide_idx+1} API响应JSON解析错误: {e}")
                slide_md_results[slide_idx] = f"[第{slide_idx+1}张幻灯片API响应解析错误]"
            except Exception as e:
                logger.error(f"幻灯片 {slide_idx+1} VLM转写失败: {type(e).__name__} - {e}")
                logger.exception(f"幻灯片 {slide_idx+1} 详细错误堆栈")
                slide_md_results[slide_idx] = f"[第{slide_idx+1}张幻灯片转写失败: {type(e).__name__}]"
    
    # 合并Markdown并添加标题
    md_fragments = []
    cropped_regions_data = {}  # 存储裁剪区域数据用于后续处理
    
    for slide_idx, md_text in enumerate(slide_md_results):
        # Construct relative path for the slide image from the final .md file
        # Final .md is in output_dir (e.g. ./output/myPres.md)
        # Slide image is in assets_dir_for_presentation / slides_dir_name / slide_X.png
        # (e.g. ./output/myPres_assets/myPres_slides/slide_X.png)
        # So, relative path is myPres_assets/myPres_slides/slide_X.png
        
        # Find the original data_entry for this slide_idx to get copied_path
        slide_img_path_obj = Path("error_path.png") # Default
        for data_entry in slide_image_data_list: # slide_image_data_list should be available
            if data_entry["id"] == slide_idx:
                slide_img_path_obj = data_entry["path"] # This is already the copied path in assets structure
                break

        # assets_dir_for_presentation.name gives e.g. "myPres_assets"
        # slide_img_path_obj.parent.name gives e.g. "myPres_slides"
        # slide_img_path_obj.name gives e.g. "slide_001.png"
        rel_slide_image_path = f"{assets_dir_for_presentation.name}/{slide_img_path_obj.parent.name}/{slide_img_path_obj.name}".replace("\\", "/")
        
        slide_header = f"\n\n## 幻灯片 {slide_idx + 1}\n\n"
        # Add the image link to the markdown for each slide
        slide_md = slide_header + f"![幻灯片 {slide_idx + 1}]({rel_slide_image_path})\n\n" + (md_text if md_text else f"[幻灯片{slide_idx+1}内容转写失败]")
        
        # 处理该幻灯片的裁剪区域图片
        regions = slide_localized_regions_map.get(slide_idx, [])
        if regions:
            cropped_regions_data[slide_idx] = {
                "regions": regions,
                "slide_num": slide_idx + 1,
                "asset_base": f"{assets_dir_for_presentation.name}",
                "slide_rel_path": rel_slide_image_path
            }
            
        md_fragments.append(slide_md)
    
    raw_md_text = f"# {presentation_stem}\n\n" + "\n\n".join(fragment for fragment in md_fragments if fragment)
    
    # 处理图片区域占位符替换
    processed_md_text = raw_md_text
    for slide_idx, data in cropped_regions_data.items():
        regions = data["regions"]
        for region_idx, region in enumerate(regions):
            if isinstance(region, dict) and region.get("type") and "image_path" in region: # 确保image_path存在
                # 构建相对路径 - 直接使用 region["image_path"]，它已经是正确的相对路径
                region_rel_path = region["image_path"].replace("\\", "/")
                region_type = region.get("type", "图片")
                region_desc = region.get("description", f"区域{region_idx+1}")
                
                # 替换占位符，格式: [图片:描述]
                placeholder = f"[图片:{region_desc}]"
                replacement = f"![{region_type}: {region_desc}]({region_rel_path})"
                
                processed_md_text = processed_md_text.replace(placeholder, replacement)
    
    final_md = processed_md_text
    if enable_refinement and REFINEMENT_MODEL_NAME:
        logger.info(f"正在使用 {REFINEMENT_MODEL_NAME} 精炼优化 {presentation_path.name} 的Markdown...")
        raw_markdown_len = len(processed_md_text) # 原始待精炼Markdown的长度
        try:
            image_metadata = {}
            for slide_idx, data in cropped_regions_data.items():
                regions = data.get("regions", [])
                for region in regions:
                    if isinstance(region, dict) and "description" in region and "image_path" in region:
                        image_metadata[region["description"]] = region["image_path"] 
            
            for i, data_entry in enumerate(slide_image_data_list):
                if "path" in data_entry and not data_entry.get("error", False):
                    slide_img_path_obj = data_entry["path"]
                    try:
                        abs_slide_path = slide_img_path_obj.resolve()
                        abs_output_dir = output_dir.resolve()
                        rel_path_for_metadata = str(abs_slide_path.relative_to(abs_output_dir)).replace("\\", "/")
                    except ValueError:
                        logger.warning(f"无法为幻灯片 {i+1} 的元数据计算相对路径，尝试拼接。")
                        rel_path_for_metadata = f"{assets_dir_for_presentation.name}/{slide_img_path_obj.parent.name}/{slide_img_path_obj.name}".replace("\\", "/")

                    image_metadata[f"幻灯片 {i+1}"] = rel_path_for_metadata
                    image_metadata[slide_img_path_obj.name] = rel_path_for_metadata 
            
            metadata_prompt_str = ""
            if image_metadata:
                metadata_prompt_parts = ["\n\n--- 图片元数据 (用于上下文参考, 并非必须使用) ---\n"]
                for desc, path in image_metadata.items():
                    metadata_prompt_parts.append(f"- 描述: \"{desc}\" -> 路径: \"{path}\"\n")
                metadata_prompt_str = "".join(metadata_prompt_parts)
            
            input_to_llm_len = raw_markdown_len + len(metadata_prompt_str)
            logger.info(f"待精炼Markdown长度: {raw_markdown_len}字符, 注入元数据长度: {len(metadata_prompt_str)}字符, LLM总输入长度: {input_to_llm_len}字符 ({presentation_path.name})")

            refined_md_text_from_llm = call_llm_for_refinement(processed_md_text, REFINEMENT_MODEL_NAME, image_metadata)
            llm_output_len = len(refined_md_text_from_llm)
            
            cleaned_md_text = clean_markdown_code_wrappers(refined_md_text_from_llm)
            final_md = fix_image_references(cleaned_md_text, image_metadata) 
            final_output_len = len(final_md)

            logger.info(f"Markdown精炼各阶段长度: 初始Markdown={raw_markdown_len}, LLM输出={llm_output_len}, 清理与修正后最终输出={final_output_len}字符 ({presentation_path.name})")
            if final_output_len < raw_markdown_len * 0.5:
                logger.warning(f"警告: 最终输出的Markdown内容相较于初始转写内容大幅减少! 初始内容长度={raw_markdown_len}字符, 最终输出长度={final_output_len}字符 ({presentation_path.name})")
            logger.info(f"Markdown精炼完成，并已修正代码包裹和图像引用 ({presentation_path.name})")
        except Exception as e:
            logger.warning(f"Markdown精炼失败 ({presentation_path.name}): {e}。使用未精炼的版本。")
            final_md = processed_md_text
    
    md_path = output_dir / f"{safe_stem}.md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(final_md)
    
    logger.info(f"✅ 完成：{presentation_path.name} → {md_path}")
    return True

# --- 主流程 ---
def process_pdf(
    pdf_path: Path, 
    output_dir: Path, 
    image_extraction_method: str = "qwen_vl", 
    dpi: int = 300,
    visualize_localization: bool = False,
    enable_refinement: bool = False # 新增参数
):
    doc = fitz.open(str(pdf_path))
    total_pages = doc.page_count
    doc.close()
    pdf_stem = pdf_path.stem
    safe_stem = "".join(c if c.isalnum() or c in "- " else "_" for c in pdf_stem).strip()
    assets_dir = output_dir / f"{safe_stem}_assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    md_fragments = []
    image_paths_by_page = {}
    # 全局图片列表，用于PyMuPDF模式或当Qwen-VL提取的图片需要全局管理时
    all_extracted_image_paths_relative = []
    # 新增：存储每页的图像区域信息
    image_regions_by_page = {}
    
    # 初始化page_md_results，防止UnboundLocalError
    page_md_results = [f"[第{i+1}页默认错误]" for i in range(total_pages)]

    logger.info(f"处理PDF《{pdf_path.name}》，共{total_pages}页，图片提取方式：{image_extraction_method}")
    
    # 第一阶段：图像定位和提取 (对于qwen_vl，此阶段也包含转写)
    if image_extraction_method == "qwen_vl":
        
        # Define the per-page pipeline worker function
        def process_single_page_pipeline_qwen_vl(
            page_num_idx: int, # 0-indexed
            current_pdf_path: Path, 
            current_assets_dir: Path, 
            current_dpi: int, 
            current_visualize_localization: bool
        ) -> Dict[str, Any]:
            """
            Processes a single PDF page: converts, localizes, crops regions, transcribes.
            Returns a dictionary with results for this page.
            """
            page_results = {
                "page_num": page_num_idx,
                "md_text": f"[第{page_num_idx+1}页处理时发生错误]",
                "extracted_images_relative_paths": [],
                "updated_regions_with_paths": []
            }
            try:
                # Step 1: Convert page to PIL Image and base64 string (once per page)
                logger.debug(f"Page {page_num_idx+1}: Converting to PIL image and base64.")
                page_image_pil = convert_pdf_page_to_image(current_pdf_path, page_num_idx, current_dpi)
                page_b64_str = image_to_base64(page_image_pil)

                # Step 2: Localize regions on the page
                logger.debug(f"Page {page_num_idx+1}: Calling vision API for localization.")
                localization_result = call_vision_api_for_localization(page_b64_str)
                # Ensure regions is always a list, even if key is missing or API fails
                regions_data = localization_result.get("image_regions", [])
                if not isinstance(regions_data, list): # Additional safety check
                    logger.warning(f"Page {page_num_idx+1}: image_regions from API was not a list: {regions_data}. Treating as no regions.")
                    regions_data = []

                for region_idx, raw_region_info in enumerate(regions_data):
                     logger.info(f"Page {page_num_idx + 1}, Raw Region {region_idx + 1} from localization: {raw_region_info}")


                # Step 3: Crop images from localized regions and save them
                extracted_img_rel_paths_for_page = []
                updated_regions_for_page = regions_data # Start with original regions
                if regions_data:
                    logger.debug(f"Page {page_num_idx+1}: Cropping {len(regions_data)} identified regions.")
                    # Use the new _crop_regions_and_get_paths helper
                    extracted_img_rel_paths_for_page, updated_regions_for_page = _crop_regions_and_get_paths(
                        page_image_pil, 
                        page_num_idx, 
                        regions_data, 
                        current_assets_dir, # This should be the specific assets dir for *this* PDF
                        current_visualize_localization
                    )
                    page_results["extracted_images_relative_paths"] = extracted_img_rel_paths_for_page
                    page_results["updated_regions_with_paths"] = updated_regions_for_page
                else:
                    logger.info(f"Page {page_num_idx+1}: No regions to crop.")
                    page_results["updated_regions_with_paths"] = [] # Ensure it's an empty list

                # Step 4: Transcribe the page to Markdown, passing pre-converted image data and region info
                logger.debug(f"Page {page_num_idx+1}: Calling VLM for Markdown transcription.")
                md_text_for_page = process_page_to_markdown(
                    current_pdf_path, 
                    page_num_idx, 
                    current_dpi, 
                    image_regions=updated_regions_for_page, # Use regions possibly updated with image_path
                    page_image_pil=page_image_pil,
                    page_b64_str=page_b64_str
                )
                page_results["md_text"] = md_text_for_page
                logger.info(f"Page {page_num_idx+1}: Successfully processed.")

            except Exception as e:
                logger.error(f"Page {page_num_idx+1}: Error in single page pipeline: {type(e).__name__} - {e}", exc_info=True)
                # page_results["md_text"] will retain its default error message or last error
            return page_results

        # Submit all page processing pipelines to the executor
        # 使用 MAX_CONCURRENT_API_CALLS 控制并发的页面处理任务数量
        with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_API_CALLS) as executor:
            pipeline_futures = {}
            logger.info(f"Submitting {total_pages} page processing pipelines for {pdf_path.name} (Qwen-VL mode) with {MAX_CONCURRENT_API_CALLS} concurrent workers...")
            for i in range(total_pages):
                # assets_dir is already defined as output_dir / f"{safe_stem}_assets"
                # It's the correct directory for this PDF's assets.
                future = executor.submit(
                    process_single_page_pipeline_qwen_vl, 
                    i, 
                    pdf_path, 
                    assets_dir, # Pass the main assets_dir for this PDF
                    dpi, 
                    visualize_localization
                )
                pipeline_futures[future] = i
            
            # Collect results as they complete
            logger.info(f"Collecting results for {pdf_path.name}...")
            # Initialize page_md_results before the loop
            page_md_results = [f"[第{i+1}页默认错误]" for i in range(total_pages)]

            for future in tqdm(as_completed(pipeline_futures), total=len(pipeline_futures), desc=f"Processing Pages ({pdf_path.name})"):
                page_num_completed = pipeline_futures[future]
                try:
                    single_page_results = future.result()
                    
                    page_md_results[page_num_completed] = single_page_results["md_text"]
                    
                    # Aggregate extracted image paths and region info
                    if single_page_results["extracted_images_relative_paths"]:
                        all_extracted_image_paths_relative.extend(single_page_results["extracted_images_relative_paths"])
                    
                    # Store updated region info for the page
                    # Ensure that image_regions_by_page is initialized if not already
                    if "image_regions_by_page" not in locals() and "image_regions_by_page" not in globals():
                        image_regions_by_page = {} # type: ignore
                    image_regions_by_page[page_num_completed] = single_page_results["updated_regions_with_paths"]

                except Exception as e:
                    logger.error(f"Page {page_num_completed+1}: Future processing failed: {e}", exc_info=True)
                    page_md_results[page_num_completed] = f"[第{page_num_completed+1}页处理失败: {e}]"
        
        # Deduplicate all_extracted_image_paths_relative just in case, though should be unique per page run
        all_extracted_image_paths_relative = sorted(list(set(all_extracted_image_paths_relative)))
        if all_extracted_image_paths_relative:
            logger.info(f"Qwen-VL mode: Aggregated {len(all_extracted_image_paths_relative)} extracted image paths for {pdf_path.name}.")
        # Qwen-VL 模式下，page_md_results 已经被填充，不需要第二阶段的转写

    elif image_extraction_method == "pymupdf":
        # PyMuPDF提取的图片是全局的，存储在 all_extracted_image_paths_relative
        all_extracted_image_paths_relative = extract_images_from_pdf_with_pymupdf(pdf_path, assets_dir)
        if all_extracted_image_paths_relative:
             logger.info(f"PyMuPDF extracted {len(all_extracted_image_paths_relative)} images globally for {pdf_path.name}.")

        # 对于 PyMuPDF 模式，现在执行转写
        # page_md_results 此时包含初始的错误信息
        logger.info(f"Starting transcription phase for PyMuPDF mode for {pdf_path.name}...")
        # 使用 MAX_CONCURRENT_API_CALLS 控制并发的转写任务数量
        with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_API_CALLS) as executor:
            page_futures = {}
            for page_num in range(total_pages):
                # PyMuPDF模式下，我们没有预先提取的区域信息传给 process_page_to_markdown
                # process_page_to_markdown 将自行处理页面到图像的转换
                regions = image_regions_by_page.get(page_num, []) # 通常为空
                
                future = executor.submit(
                    process_page_to_markdown, pdf_path, page_num, dpi, regions
                )
                page_futures[future] = page_num
            
            for future in tqdm(as_completed(page_futures), total=len(page_futures), desc=f"VLM转写Markdown ({pdf_path.name}, PyMuPDF)"):
                page_num_completed = page_futures[future]
                try:
                    md_text_for_page = future.result()
                    page_md_results[page_num_completed] = md_text_for_page # 填充转写结果
                except Exception as e:
                    logger.error(f"第{page_num_completed+1}页VLM转写主流程失败 (PyMuPDF mode): {e}")
                    page_md_results[page_num_completed] = f"[第{page_num_completed+1}页转写失败 due to: {e}]"

    # 按页码顺序组合Markdown片段
    # page_md_results 此处应包含来自正确处理路径（Qwen-VL 或 PyMuPDF）的结果
    for page_num, md_text in enumerate(page_md_results):
        if md_text:
             md_fragments.append(md_text)
        else: # 理论上应该在上面被捕获了
             md_fragments.append(f"[第{page_num+1}页转写数据丢失]")

    raw_md_text = "\n\n".join(fragment.strip() for fragment in md_fragments if fragment)

    # 图片占位符替换逻辑调整
    final_md_text_before_refinement = raw_md_text
    current_global_image_index = 0 # 用于 PyMuPDF 模式或全局图片列表

    # 统计总的[图片]占位符数量和含描述的占位符
    total_placeholder_count = final_md_text_before_refinement.count("[图片]")
    
    # 检查是否有格式为[图片:描述]的占位符
    desc_placeholder_pattern = re.compile(r'\[图片:([^\]]+)\]')
    desc_placeholders = desc_placeholder_pattern.findall(final_md_text_before_refinement)
    
    if desc_placeholders:
        logger.info(f"找到 {len(desc_placeholders)} 个带描述的'[图片:描述]'占位符。")
        
        # 创建描述到图像路径的映射
        desc_to_image = {}
        
        # 从image_regions_by_page中提取所有描述和对应的图像路径
        for page_num, regions in image_regions_by_page.items():
            for region in regions:
                # 类型检查以防止"'str' object has no attribute 'get'"错误
                if not isinstance(region, dict):
                    continue
                    
                if "description" in region and "image_path" in region:
                    desc = region["description"]
                    path = region["image_path"]
                    desc_to_image[desc] = path
        
        # 替换所有带描述的占位符
        for desc in desc_placeholders:
            if desc in desc_to_image:
                path = desc_to_image[desc]
                placeholder = f"[图片:{desc}]"
                replacement = f"![{desc}]({path})"
                final_md_text_before_refinement = final_md_text_before_refinement.replace(placeholder, replacement)
            else:
                # 如果没有完全匹配，尝试部分匹配
                best_match = None
                best_score = 0.6  # 相似度阈值
                
                for known_desc in desc_to_image.keys():
                    # 使用简单的子字符串匹配，或者可以引入更复杂的相似度算法
                    if desc.lower() in known_desc.lower() or known_desc.lower() in desc.lower():
                        score = len(set(desc.lower()) & set(known_desc.lower())) / len(set(desc.lower()) | set(known_desc.lower()))
                        if score > best_score:
                            best_score = score
                            best_match = known_desc
                
                if best_match:
                    path = desc_to_image[best_match]
                    placeholder = f"[图片:{desc}]"
                    replacement = f"![{desc}]({path})"
                    final_md_text_before_refinement = final_md_text_before_refinement.replace(placeholder, replacement)
    
    # 处理普通[图片]占位符
    if total_placeholder_count > 0:
        logger.info(f"在{pdf_path.name}的初始Markdown中找到 {total_placeholder_count} 个简单的'[图片]'占位符。")
        
        if image_extraction_method == "qwen_vl" or image_extraction_method == "pymupdf":
            # 对于无描述的占位符，继续使用全局顺序替换方法
            temp_md_parts = []
            last_pos = 0
            replaced_count = 0
            for _ in range(total_placeholder_count):
                if current_global_image_index >= len(all_extracted_image_paths_relative):
                    logger.warning(f"提取的图像数量({len(all_extracted_image_paths_relative)})不足以替换所有简单占位符({total_placeholder_count})。")
                    break
                
                placeholder_pos = final_md_text_before_refinement.find("[图片]", last_pos)
                if placeholder_pos == -1:
                    break
                
                temp_md_parts.append(final_md_text_before_refinement[last_pos:placeholder_pos])
                img_rel_path = all_extracted_image_paths_relative[current_global_image_index]
                # 尝试从图像路径提取描述信息
                img_filename = Path(img_rel_path).name
                img_desc = img_filename.split("_")[1] if len(img_filename.split("_")) > 1 else "图片"
                temp_md_parts.append(f"![{img_desc}]({img_rel_path})")
                
                last_pos = placeholder_pos + len("[图片]")
                current_global_image_index += 1
                replaced_count += 1
            
            temp_md_parts.append(final_md_text_before_refinement[last_pos:])
            final_md_text_before_refinement = "".join(temp_md_parts)
            logger.info(f"替换了 {replaced_count} 个简单占位符。")

    # 如需精炼，调用LLM API进行Markdown优化
    final_md_text = final_md_text_before_refinement
    if enable_refinement and REFINEMENT_MODEL_NAME:
        try:
            logger.info(f"使用 {REFINEMENT_MODEL_NAME} 精炼优化Markdown中...")
            raw_markdown_len = len(final_md_text_before_refinement)
            
            comprehensive_image_metadata = {}
            for page_num, regions_on_page in image_regions_by_page.items():
                for region in regions_on_page:
                    if isinstance(region, dict) and "description" in region and "image_path" in region:
                        comprehensive_image_metadata[region["description"]] = region["image_path"]
            
            if image_extraction_method == "pymupdf":
                for i, path in enumerate(all_extracted_image_paths_relative):
                    img_filename = Path(path).name
                    potential_desc_key = f"来自page_{Path(path).stem.split('_')[0][4:]}_img_{Path(path).stem.split('_')[1][3:]}" 
                    if not any(existing_path == path for existing_path in comprehensive_image_metadata.values()):
                        if img_filename not in comprehensive_image_metadata:
                             comprehensive_image_metadata[img_filename] = path
                        elif potential_desc_key not in comprehensive_image_metadata: 
                             comprehensive_image_metadata[potential_desc_key] = path

            metadata_prompt_str = ""
            if comprehensive_image_metadata:
                metadata_prompt_parts = ["\n\n--- 图片元数据 (用于上下文参考, 并非必须使用) ---\n"]
                for desc, path in comprehensive_image_metadata.items():
                    metadata_prompt_parts.append(f"- 描述: \"{desc}\" -> 路径: \"{path}\"\n")
                metadata_prompt_str = "".join(metadata_prompt_parts)
            
            input_to_llm_len = raw_markdown_len + len(metadata_prompt_str)
            logger.info(f"待精炼Markdown长度: {raw_markdown_len}字符, 注入元数据长度: {len(metadata_prompt_str)}字符, LLM总输入长度: {input_to_llm_len}字符 ({pdf_path.name})")

            refined_md_text_from_llm = call_llm_for_refinement(final_md_text_before_refinement, REFINEMENT_MODEL_NAME, comprehensive_image_metadata)
            llm_output_len = len(refined_md_text_from_llm)
            
            cleaned_md_text = clean_markdown_code_wrappers(refined_md_text_from_llm)
            final_md_text = fix_image_references(cleaned_md_text, comprehensive_image_metadata) # 使用 comprehensive_image_metadata
            final_output_len = len(final_md_text)

            logger.info(f"Markdown精炼各阶段长度: 初始Markdown={raw_markdown_len}, LLM输出={llm_output_len}, 清理与修正后最终输出={final_output_len}字符 ({pdf_path.name})")
            if final_output_len < raw_markdown_len * 0.5:
                logger.warning(f"警告: 最终输出的Markdown内容相较于初始转写内容大幅减少! 初始内容长度={raw_markdown_len}字符, 最终输出长度={final_output_len}字符 ({pdf_path.name})")
            logger.info(f"Markdown精炼完成，并已修正代码包裹和图像引用 ({pdf_path.name})")
        except KeyboardInterrupt:
            logger.warning("精炼过程被手动中断。使用未精炼的原始Markdown。")
            final_md_text = final_md_text_before_refinement
        except requests.exceptions.Timeout:
            logger.warning("精炼过程超时。使用未精炼的原始Markdown。")
            final_md_text = final_md_text_before_refinement
        except requests.exceptions.RequestException as e:
            logger.warning(f"精炼过程API请求失败: {e}。使用未精炼的原始Markdown。")
            final_md_text = final_md_text_before_refinement
        except Exception as e:
            logger.warning(f"Markdown精炼过程遇到错误: {e}。使用未精炼的原始Markdown。")
            final_md_text = final_md_text_before_refinement

    # 合并并保存Markdown
    md_path = output_dir / f"{safe_stem}.md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(final_md_text)
    logger.info(f"✅ 完成：{pdf_path.name} → {md_path}")
    logger.info(f"  - 图片资源目录: {assets_dir.resolve()}")

# 新增辅助函数，用于处理单页的VLM转写，包含图像区域信息注入
def process_page_to_markdown(pdf_path: Path, page_num: int, dpi: int, image_regions: Optional[List[Dict]] = None, page_image_pil: Optional[Image.Image] = None, page_b64_str: Optional[str] = None) -> str:
    """
    将单页PDF转为图片，然后调用VLM获取Markdown
    如果提供了image_regions，会将图像区域信息注入到用户指令中
    如果提供了page_image_pil and page_b64_str, 则跳过图像转换
    """
    logger.debug(f"Page {page_num + 1}: 开始处理Markdown转写。")
    try:
        if page_image_pil is None or page_b64_str is None:
            logger.debug(f"Page {page_num + 1}: 转换页面为图像...")
            page_image_pil = convert_pdf_page_to_image(pdf_path, page_num, dpi)
            page_b64_str = image_to_base64(page_image_pil)
            logger.debug(f"Page {page_num + 1}: 页面图像转换完成。")
        else:
            logger.debug(f"Page {page_num + 1}: 使用预转换的页面图像。")
        
        # 构建可能带有图像信息的自定义指令
        custom_instruction = USER_INSTRUCTION
        if image_regions and len(image_regions) > 0:
            logger.info(f"Page {page_num + 1}: 检测到 {len(image_regions)} 个图像区域，准备注入提示词。")
            instruction_parts = [USER_INSTRUCTION, "\n\n本页包含以下特殊区域，请按以下要求处理：\n"]
            for i, region in enumerate(image_regions):
                if not isinstance(region, dict):
                    logger.warning(f"Page {page_num + 1}: 区域 {i+1} 不是预期的字典格式，跳过此区域。")
                    continue
                
                region_type = region.get("type", "").lower()
                desc = region.get("description", f"区域{i+1}")
                bbox_str = str(region.get("bbox", "[未知BBOX]")) # 转换为字符串以确保可记录
                image_path_in_region = region.get("image_path")
                
                logger.debug(f"Page {page_num + 1}: 区域 {i+1} - 类型: {region_type}, 描述: {desc}, BBOX: {bbox_str}, 路径: {image_path_in_region}")
                
                placeholder_desc = desc
                # placeholder_desc目前未使用image_path_in_region，因为VLM主要通过描述来关联

                if "table" in region_type:
                    instruction_parts.append(f"{i+1}. [表格区域] 坐标: {bbox_str} - 描述: {desc}\n")
                    instruction_parts.append(f"   请优先尝试将此表格转换为结构化的Markdown表格格式。\n")
                    instruction_parts.append(f"   只有在表格结构特别复杂，无法准确转换时，才使用 [图片:{placeholder_desc}] 占位符。\n")
                else:
                    instruction_parts.append(f"{i+1}. [图像区域] 坐标: {bbox_str} - 描述: {desc}\n")
                    instruction_parts.append(f"   请使用 [图片:{placeholder_desc}] 占位符。\n")
            
            instruction_parts.append("\n对于表格区域，请尽可能地转换为Markdown表格，确保保留原始表格的所有内容和格式。")
            instruction_parts.append("\n只有在表格非常复杂（如包含复杂角分割逻辑、合并单元格、多重嵌套表格等）无法准确转换时，才使用图片占位符。")
            custom_instruction = "".join(instruction_parts)
            logger.info(f"Page {page_num + 1}: 图像区域提示词构建完成，总长度: {len(custom_instruction)}。")
            logger.debug(f"Page {page_num + 1}: 注入的完整提示词: \n{custom_instruction}") # 记录完整提示词
        else:
            logger.info(f"Page {page_num + 1}: 未检测到图像区域，使用标准提示词。")
        
        # 调用VLM进行转写，传入自定义指令
        logger.debug(f"Page {page_num + 1}: 调用VLM API进行转写...")
        md_text = call_vlm_for_markdown(page_b64_str, custom_instruction)
        logger.info(f"Page {page_num + 1}: VLM转写成功，获得Markdown文本长度: {len(md_text)}。")
        
        # Debug: 检查转写文本是否包含预期的占位符
        if image_regions and len(image_regions) > 0:
            placeholders_found = 0
            placeholders_missing = []
            for region in image_regions:
                if isinstance(region, dict):
                    desc = region.get("description", f"区域{i+1}") # 修正i的范围问题
                    expected_placeholder = f"[图片:{desc}]"
                    if expected_placeholder in md_text:
                        placeholders_found += 1
                    else:
                        placeholders_missing.append(expected_placeholder)
            
            logger.debug(f"Page {page_num + 1}: 转写文本占位符检查 - 找到: {placeholders_found}, 期望总数: {len(image_regions)}")
            if placeholders_missing:
                logger.warning(f"Page {page_num + 1}: 以下占位符在转写文本中缺失: {placeholders_missing}")
        
        return md_text
    except tenacity.RetryError as e:
        last_exception = e.last_attempt.exception()
        error_details = f"RetryError on VLM call for page {page_num + 1}"
        if isinstance(last_exception, requests.exceptions.HTTPError):
            status_code = last_exception.response.status_code
            response_text = last_exception.response.text
            error_details += f": Last HTTPError {status_code} - {response_text[:200]}"
        elif last_exception:
            error_details += f": Last error - {type(last_exception).__name__}: {str(last_exception)[:200]}"
        logger.error(f"❌ VLM ERROR for page {page_num + 1}: {error_details}")
        return f"[第{page_num+1}页VLM转写失败: {error_details}]"
    except Exception as e:
        error_details = f"Unexpected error during VLM processing for page {page_num + 1}: {type(e).__name__} - {str(e)[:200]}"
        logger.error(f"❌ VLM ERROR for page {page_num + 1}: {error_details}", exc_info=True) # 添加exc_info
        return f"[第{page_num+1}页VLM转写失败: {error_details}]"

# 在适当位置添加检测LibreOffice和调用它进行转换的函数
def is_libreoffice_available():
    """检查LibreOffice是否可用"""
    # 检查常见的LibreOffice安装路径
    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"D:\Program Files\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice"
    ]
    
    # 检查环境变量PATH中是否存在soffice
    try:
        # 使用which/where命令查找可执行文件
        if IS_WINDOWS:
            result = subprocess.run(["where", "soffice"], 
                                    capture_output=True, text=True, check=False)
        else:
            result = subprocess.run(["which", "soffice"], 
                                    capture_output=True, text=True, check=False)
        
        if result.returncode == 0:
            return True
    except Exception:
        pass
    
    # 检查常见路径
    for path in common_paths:
        if os.path.exists(path):
            return True
    
    return False

def convert_pptx_to_pdf_with_libreoffice(pptx_path, output_dir):
    """使用LibreOffice将PPTX转换为PDF"""
    logger.info(f"使用LibreOffice转换PPTX: {pptx_path}")
    
    # 确保输出目录存在
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 确定LibreOffice可执行文件路径
    soffice_path = None
    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"D:\Program Files\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice"
    ]
    
    # 检查常见路径
    for path in common_paths:
        if os.path.exists(path):
            soffice_path = path
            break
    
    # 如果没有找到确切路径，使用系统PATH中的soffice命令
    if not soffice_path:
        soffice_path = "soffice"
    
    # PDF输出路径
    pdf_output_path = output_dir / f"{pptx_path.stem}.pdf"
    
    # 构建命令
    cmd = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(pptx_path)
    ]
    
    try:
        # 运行命令
        logger.info(f"执行命令: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        
        if result.returncode != 0:
            logger.error(f"LibreOffice转换失败: {result.stderr}")
            return None
        
        # 检查输出文件是否存在
        if not pdf_output_path.exists():
            logger.error(f"LibreOffice转换后的PDF文件不存在: {pdf_output_path}")
            return None
        
        logger.info(f"PPTX成功转换为PDF: {pdf_output_path}")
        return pdf_output_path
    
    except Exception as e:
        logger.error(f"调用LibreOffice时出错: {e}")
        return None

# --- 新增：通用文档处理函数 ---
def process_document(
    doc_path: Path, 
    output_dir: Path, 
    doc_type: Optional[str] = None, # Made doc_type optional
    image_extraction_method: str = "qwen_vl",
    dpi: int = 300,
    visualize_localization: bool = False,
    enable_refinement: bool = False
) -> bool:
    """
    通用文档处理函数，处理PDF、PPT、PPTX文件或PNG系列
    
    参数:
        doc_path: 文档路径
        output_dir: 输出目录
        doc_type: 文档类型，如果为None则自动从文件扩展名判断
        image_extraction_method: 图片提取方式，仅对PDF有效
        dpi: 图像DPI
        visualize_localization: 是否可视化图像定位结果
        enable_refinement: 是否启用Markdown精炼
        
    返回:
        bool: 处理是否成功
    """
    # 如果未指定文档类型，从文件扩展名判断
    if doc_type is None:
        file_ext = doc_path.suffix.lower()
        if file_ext == ".pdf":
            doc_type = "pdf"
        elif file_ext in [".pptx", ".ppt"]:
            doc_type = "presentation" # Unified type for both ppt and pptx
        elif file_ext == ".png":
            doc_type = "png"
        else:
            logger.error(f"不支持的文档类型: {file_ext} for file {doc_path.name}")
            return False
    
    # 根据文档类型调用对应的处理函数
    try:
        if doc_type == "pdf":
            logger.info(f"开始处理PDF文件: {doc_path.name}")
            # Call process_pdf directly, it returns bool but we don't assign it here
            process_pdf(
                pdf_path=doc_path,
                output_dir=output_dir,
                image_extraction_method=image_extraction_method,
                dpi=dpi,
                visualize_localization=visualize_localization if image_extraction_method == "qwen_vl" else False,
                enable_refinement=enable_refinement
            )
            logger.info(f"处理完成: {doc_path.name}") # Assuming success if no exception
            return True # Return True if process_pdf completes without error
        
        elif doc_type == "presentation": # Handles both .ppt and .pptx
            logger.info(f"开始处理演示文稿文件: {doc_path.name}")
            if not is_libreoffice_available():
                logger.error(f"无法处理演示文稿文件: {doc_path.name}，因为未检测到LibreOffice。请安装LibreOffice并确保其在系统PATH中，或通过常见安装路径可访问。")
                return False
                
            return process_presentation( # This function returns bool
                presentation_path=doc_path,
                output_dir=output_dir,
                dpi=dpi,
                visualize_localization=visualize_localization, # Pass along
                enable_refinement=enable_refinement
            )
            # No explicit logger.info for completion here, process_presentation handles it

        elif doc_type == "png":
            logger.info(f"开始处理PNG文件: {doc_path.name} (作为单文件系列)")
            # For a single PNG, treat it as a series of one file.
            # process_png_series returns the path to MD or None
            result_path = process_png_series(
                png_files=[doc_path],
                output_dir=output_dir,
                series_name=doc_path.stem,
                image_extraction_method=image_extraction_method,
                dpi=dpi, # Pass DPI
                visualize_localization=visualize_localization if image_extraction_method == "qwen_vl" else False,
                enable_refinement=enable_refinement
            )
            return result_path is not None # True if a file path was returned
        else:
            # This case should ideally be caught by the initial doc_type inference
            logger.error(f"内部错误：无法识别的文档类型进行处理: {doc_type} for file {doc_path.name}")
            return False
    except Exception as e:
        logger.error(f"处理文档 {doc_path.name} (类型: {doc_type}) 时发生严重错误: {e}", exc_info=True)
        return False

# --- 新增：PNG系列处理函数 ---
def process_png_series(
    png_files: List[Path], # Changed from List[str] to List[Path]
    output_dir: Path,      # Changed from str to Path
    series_name: str,
    enable_refinement: bool = False,
    image_extraction_method: str = "qwen_vl",
    visualize_localization: bool = False,
    dpi: int = 300 # Added dpi, though less relevant for PNGs, for consistency if used by helpers
) -> Optional[Path]: # Return Path or None
    """处理一系列PNG文件并生成Markdown

    Args:
        png_files: PNG文件路径列表 (Path objects)
        output_dir: 输出目录 (Path object)
        series_name: 系列名称
        enable_refinement: 是否启用Markdown内容优化
        image_extraction_method: 图片提取方法
        visualize_localization: 是否可视化定位结果
        dpi: DPI for any image processing (if applicable)

    Returns:
        Path: 生成的Markdown文件路径, or None if failed
    """
    safe_series_name = "".join(c if c.isalnum() or c in "- " else "_" for c in series_name).strip().replace(" ", "_")
    logger.info(f"开始处理PNG系列: {safe_series_name}，共{len(png_files)}张图片，使用{image_extraction_method}提取方法")
    
    md_output_file = output_dir / f"{safe_series_name}.md"
    assets_dir_for_series = output_dir / f"{safe_series_name}_assets"
    # 删除regions_dir，直接使用assets_dir作为裁剪图像存储目录
    
    assets_dir_for_series.mkdir(parents=True, exist_ok=True)
    
    # 修改：复制PNG文件到assets目录，并更新image_data_list中的路径信息
    image_data_list = [] # List of dicts: {"id": i, "original_path": Path, "copied_asset_path": Path, "b64": str, "error": bool}
    logger.info(f"准备PNG图像数据 {safe_series_name} (复制并进行base64编码)...")
    for i, original_png_path_obj in enumerate(png_files):
        copied_asset_path = assets_dir_for_series / original_png_path_obj.name
        entry = {"id": i, "original_path": original_png_path_obj, "copied_asset_path": copied_asset_path, "b64": None, "error": False}
        try:
            shutil.copy2(original_png_path_obj, copied_asset_path) # 复制文件
            logger.debug(f"已将 {original_png_path_obj.name} 复制到 {copied_asset_path}")
            # 从复制后的文件读取base64
            with open(copied_asset_path, "rb") as img_file:
                entry["b64"] = base64.b64encode(img_file.read()).decode("utf-8")
        except Exception as e:
            logger.error(f"无法复制或编码PNG {original_png_path_obj.name} for series {safe_series_name}: {e}")
            entry["error"] = True
        image_data_list.append(entry)

    markdown_contents = [None] * len(image_data_list)
    # To store {index: regions_data} from localization
    localized_regions_map = {}
    cropped_regions_by_image = {}

    with ThreadPoolExecutor(max_workers=SF_MAX_WORKERS) as executor:
        localization_futures = {} # future -> image_index
        transcription_futures = {} # future -> image_index

        # Phase 1: Submit localization tasks (if qwen_vl)
        if image_extraction_method == "qwen_vl":
            logger.info(f"Series {safe_series_name}: Submitting localization tasks...")
            for i, data_entry in enumerate(image_data_list):
                if data_entry["error"] or not data_entry["b64"]:
                    markdown_contents[i] = f"[图片 {i+1} ({data_entry['original_path'].name}) 准备失败]"
                    continue
                future = executor.submit(call_vision_api_for_localization, data_entry["b64"])
                localization_futures[future] = i
        
        # Phase 1.5: Process localization results and submit transcription tasks (if qwen_vl)
        # OR: Submit transcription tasks directly if not qwen_vl
        if image_extraction_method == "qwen_vl":
            logger.info(f"Series {safe_series_name}: Processing localization and queueing transcriptions...")
            for loc_future in tqdm(as_completed(localization_futures), total=len(localization_futures), desc=f"Localizing ({safe_series_name})"):
                i = localization_futures[loc_future]
                data_entry = image_data_list[i]
                if data_entry["error"]: 
                    continue
                
                regions = [] # 默认没有区域
                try:
                    localization_result = loc_future.result()
                    regions = localization_result.get("image_regions", [])
                    localized_regions_map[i] = regions 
                    
                    # 新增：如果检测到区域，就裁剪并保存区域图像
                    if regions:
                        logger.info(f"PNG {i+1} ({data_entry['original_path'].name}): 检测到 {len(regions)} 个区域，开始裁剪")
                        try:
                            # 打开原始图像 - 直接从原始路径打开
                            img_pil = Image.open(data_entry["original_path"])
                            
                            # 裁剪并保存每个区域 - 修改为直接使用assets_dir而不是regions_dir
                            saved_region_paths, updated_regions = _crop_regions_and_get_paths(
                                page_image=img_pil,
                                page_number=i,  # 使用图片索引作为页码
                                image_regions=regions,
                                output_dir=assets_dir_for_series,  # 使用assets_dir而不是regions_dir
                                visualize=visualize_localization
                            )
                            
                            # 存储已裁剪的区域信息
                            cropped_regions_by_image[i] = updated_regions
                            logger.info(f"PNG {i+1} ({data_entry['original_path'].name}): 成功裁剪并保存 {len(saved_region_paths)} 个区域")
                        except Exception as crop_err:
                            logger.error(f"PNG {i+1} ({data_entry['original_path'].name}): 裁剪区域失败: {crop_err}")
                    
                    if visualize_localization and regions:
                        try:
                            img_pil = Image.open(data_entry["original_path"])
                            draw = ImageDraw.Draw(img_pil)
                            for region_idx, region_info in enumerate(regions):
                                if isinstance(region_info, dict) and region_info.get("bbox"):
                                    x1, y1, x2, y2 = region_info["bbox"]
                                    draw.rectangle([x1,y1,x2,y2], outline="red", width=2)
                                    font = ImageFont.load_default()
                                    try: font = ImageFont.truetype("arial.ttf", 12)
                                    except IOError: pass
                                    draw.text((x1,y1-12 if y1 > 12 else y1+2), f"{region_idx+1}:{region_info.get('type','reg')}", fill="red", font=font)
                            # 修改为在assets_dir_for_series目录中保存可视化结果
                            vis_path = assets_dir_for_series / f"{data_entry['original_path'].name.split('.')[0]}_viz.png"
                            img_pil.save(vis_path)
                            logger.info(f"Saved localization visualization for {data_entry['original_path'].name} to {vis_path.name}")
                        except Exception as viz_e:
                            logger.warning(f"Visualization failed for {data_entry['original_path'].name}: {viz_e}")
                except Exception as e:
                    logger.error(f"PNG {i+1} ({data_entry['original_path'].name}): Localization processing failed: {e}")
                    # 即使定位失败，也尝试提交转写，但记录错误
                    markdown_contents[i] = f"[图片 {i+1} ({data_entry['original_path'].name}) 定位失败: {e}]"

                # 准备 instruction for transcription (无论是否有 regions)
                slide_instruction = f"这是PNG系列 '{safe_series_name}' 中的第{i+1}张图片 (共{len(image_data_list)}张)。请准确转写图片中的所有内容。"
                
                # 如果有剪切的区域，在指令中加入区域信息
                if i in cropped_regions_by_image and cropped_regions_by_image[i]:
                    slide_instruction += " 图片识别出以下区域：\n"
                    for j, region_data in enumerate(cropped_regions_by_image[i]):
                        if isinstance(region_data, dict):
                            region_type = region_data.get("type", "区域")
                            region_desc = region_data.get("description", f"区域{j+1}")
                            region_path = region_data.get("image_path", "")
                            slide_instruction += f"  - {region_type}: {region_desc} (已裁剪并保存为 {region_path})\n"
                elif regions: # 如果有检测区域但没有裁剪成功，仍然提供区域信息
                    slide_instruction += " 图片识别出以下区域：\n"
                    for j, region_data in enumerate(regions):
                        if isinstance(region_data, dict):
                            region_type = region_data.get("type", "区域")
                            region_desc = region_data.get("description", f"区域{j+1}")
                            slide_instruction += f"  - {region_type}: {region_desc}\n"
                
                # 确保 data_entry["b64"] 有效才提交任务
                if data_entry["b64"]:
                    try:
                        trans_future = executor.submit(call_vlm_for_markdown, data_entry["b64"], slide_instruction)
                        transcription_futures[trans_future] = i
                        logger.info(f"PNG {i+1} ({data_entry['original_path'].name}): 已提交转写任务，当前队列长度: {len(transcription_futures)}")
                    except Exception as submit_e:
                        logger.error(f"PNG {i+1} ({data_entry['original_path'].name}): Error submitting transcription task: {submit_e}")
                        markdown_contents[i] = f"[图片 {i+1} ({data_entry['original_path'].name}) 转写任务提交失败: {submit_e}]"
                else:
                    logger.error(f"PNG {i+1} ({data_entry['original_path'].name}): Base64 data is missing, cannot submit transcription task.")
                    markdown_contents[i] = f"[图片 {i+1} ({data_entry['original_path'].name}) Base64数据缺失]"
        else: # Not qwen_vl, submit transcriptions directly
            logger.info(f"Series {safe_series_name}: Submitting transcription tasks directly (no localization)...")
            for i, data_entry in enumerate(image_data_list):
                if data_entry["error"] or not data_entry["b64"]:
                    # Error already logged and markdown_contents[i] set
                    continue
                slide_instruction = f"这是PNG系列 '{safe_series_name}' 中的第{i+1}张图片 (共{len(image_data_list)}张)。请准确转写图片中的所有内容。"
                trans_future = executor.submit(call_vlm_for_markdown, data_entry["b64"], slide_instruction)
                transcription_futures[trans_future] = i

        # Phase 2: Collect all transcription results
        logger.info(f"Series {safe_series_name}: Collecting transcription results...")
        
        # 添加这行记录总共需要收集的任务数量
        logger.info(f"需要收集的转写任务总数: {len(transcription_futures)}")
        
        # 初始化收集计数器
        collected_count = 0
        
        for trans_future in tqdm(as_completed(transcription_futures), total=len(transcription_futures), desc=f"Transcribing ({safe_series_name})"):
            i = transcription_futures[trans_future]
            data_entry = image_data_list[i]
            if data_entry["error"]: # Skip if image preparation had an error
                continue 
            if markdown_contents[i] is not None and "失败" in markdown_contents[i]: # If error already recorded from previous stage
                 continue
            try:
                logger.debug(f"PNG {i+1} ({data_entry['original_path'].name}): 尝试获取转写结果...")
                md_text = trans_future.result()
                # 添加以下日志信息
                if not md_text or md_text.strip() == "":
                    logger.error(f"PNG {i+1} ({data_entry['original_path'].name}) VLM转写内容为空.")
                    markdown_contents[i] = f"[图片 {i+1} ({data_entry['original_path'].name}) VLM转写内容为空]"
                else:
                    # 新增详细日志，记录内容长度和前100个字符
                    text_preview = md_text[:100].replace('\n', ' ')
                    logger.info(f"PNG {i+1} ({data_entry['original_path'].name}) VLM返回内容长度: {len(md_text)}, 前100字符: {text_preview}...")
                    markdown_contents[i] = md_text
                    collected_count += 1
                    logger.debug(f"PNG {i+1} ({data_entry['original_path'].name}) 转写完成，当前已完成: {collected_count}/{len(transcription_futures)}")
            except Exception as e:
                logger.error(f"PNG {i+1} ({data_entry['original_path'].name}) VLM转写失败: {e}")
                markdown_contents[i] = f"[图片 {i+1} ({data_entry['original_path'].name}) VLM转写失败: {e}]"

    # Combine Markdown fragments
    final_md_content_parts = []
    final_md_content_parts.append(f"# {safe_series_name}\n")
    
    # 记录整个列表状态
    for i, content in enumerate(markdown_contents):
        if content is None:
            logger.error(f"markdown_contents[{i}] 是 None")
        elif content.startswith("["):
            logger.warning(f"markdown_contents[{i}] 是错误信息: {content}")
        else:
            logger.info(f"markdown_contents[{i}] 有效内容长度: {len(content)} 字符")
            
    for i, data_entry in enumerate(image_data_list):
        # 详细记录每个片段的状态
        logger.info(f"处理片段 {i+1}: 原始路径={data_entry['original_path'].name}, 资源路径={data_entry['copied_asset_path'].name}")
        rel_image_path_str = f"{assets_dir_for_series.name}/{data_entry['copied_asset_path'].name}".replace("\\", "/")
        
        content = markdown_contents[i] if markdown_contents[i] is not None else f"[内容处理失败 for {data_entry['original_path'].name}]"
        # 记录内容状态
        if content.startswith("[内容处理失败") or content.startswith("[图片") and "失败" in content:
            logger.warning(f"片段 {i+1} 内容状态: 失败 - {content}")
        else:
            content_preview = content[:50].replace('\n', ' ')
            logger.info(f"片段 {i+1} 内容状态: 有效内容, 长度={len(content)}, 前50字符={content_preview}...")
        
        final_md_content_parts.append(f"## 图片 {i+1}: {data_entry['original_path'].name}\n")
        final_md_content_parts.append(f"![{data_entry['original_path'].name}]({rel_image_path_str})\n")
        
        # 新增：如果有裁剪的区域，添加区域图像的引用
        if i in cropped_regions_by_image and cropped_regions_by_image[i]:
            final_md_content_parts.append(f"\n### 检测到的区域：\n\n")
            for j, region_data in enumerate(cropped_regions_by_image[i]):
                if isinstance(region_data, dict) and "image_path" in region_data:
                    region_type = region_data.get("type", "区域")
                    region_desc = region_data.get("description", f"区域{j+1}")
                    region_path = region_data.get("image_path", "")
                    final_md_content_parts.append(f"#### {region_type}: {region_desc}\n\n")
                    final_md_content_parts.append(f"![{region_desc}]({region_path})\n\n")
        
        final_md_content_parts.append(f"{content}\n\n---\n")
    
    raw_md_text = "".join(final_md_content_parts)
    final_md_to_write = raw_md_text

    if enable_refinement and REFINEMENT_MODEL_NAME:
        try:
            logger.info(f"Series {safe_series_name}: 使用 {REFINEMENT_MODEL_NAME} 精炼优化Markdown中...")
            raw_markdown_len = len(raw_md_text)  # 原始待精炼Markdown的长度
            
            image_metadata = {}
            for img_idx, regions in cropped_regions_by_image.items(): 
                for region in regions:
                    if isinstance(region, dict) and "description" in region and "image_path" in region:
                        image_metadata[region["description"]] = region["image_path"] 
            
            for i, data_entry in enumerate(image_data_list):
                rel_image_path_str = f"{assets_dir_for_series.name}/{data_entry['copied_asset_path'].name}".replace("\\", "/")
                image_metadata[f"图片 {i+1}"] = rel_image_path_str
                image_metadata[data_entry['copied_asset_path'].name] = rel_image_path_str
            
            # 记录元数据本身的长度，以便了解注入了多少额外信息
            metadata_prompt_str = ""
            if image_metadata:
                metadata_prompt_parts = ["\n\n--- 图片元数据 (用于上下文参考, 并非必须使用) ---\n"]
                for desc, path in image_metadata.items():
                    metadata_prompt_parts.append(f"- 描述: \"{desc}\" -> 路径: \"{path}\"\n")
                metadata_prompt_str = "".join(metadata_prompt_parts)
            
            input_to_llm_len = raw_markdown_len + len(metadata_prompt_str)
            logger.info(f"待精炼Markdown长度: {raw_markdown_len}字符, 注入元数据长度: {len(metadata_prompt_str)}字符, LLM总输入长度: {input_to_llm_len}字符")

            refined_md_text_from_llm = call_llm_for_refinement(raw_md_text, REFINEMENT_MODEL_NAME, image_metadata)
            llm_output_len = len(refined_md_text_from_llm)
            
            cleaned_md_text = clean_markdown_code_wrappers(refined_md_text_from_llm)
            final_md_to_write = fix_image_references(cleaned_md_text, image_metadata)
            final_output_len = len(final_md_to_write)
            
            logger.info(f"Markdown精炼各阶段长度: 初始Markdown={raw_markdown_len}, LLM输出={llm_output_len}, 清理与修正后最终输出={final_output_len}字符")
            
            # 警告判断基于初始Markdown内容和最终输出内容的比较
            if final_output_len < raw_markdown_len * 0.5:
                logger.warning(f"警告: 最终输出的Markdown内容相较于初始转写内容大幅减少! 初始内容长度={raw_markdown_len}字符, 最终输出长度={final_output_len}字符")
                raw_preview = raw_md_text[:200].replace('\n', ' ')
                final_preview = final_md_to_write[:200].replace('\n', ' ')
                logger.warning(f"原始内容前200字符: {raw_preview}...")
                logger.warning(f"精炼后内容前200字符: {final_preview}...")
            
            logger.info(f"Series {safe_series_name}: Markdown精炼完成，并已修正代码包裹和图像引用")
        except KeyboardInterrupt:
            logger.warning("精炼过程被手动中断。使用未精炼的原始Markdown。")
            final_md_to_write = raw_md_text # 使用 raw_md_text
        except requests.exceptions.Timeout:
            logger.warning("精炼过程超时。使用未精炼的原始Markdown。")
            final_md_to_write = raw_md_text # 使用 raw_md_text
        except requests.exceptions.RequestException as e:
            logger.warning(f"精炼过程API请求失败: {e}。使用未精炼的原始Markdown。")
            final_md_to_write = raw_md_text # 使用 raw_md_text
        except Exception as e:
            logger.warning(f"Series {safe_series_name}: Markdown精炼失败: {e}。使用未精炼版本。")
            final_md_to_write = raw_md_text # 使用 raw_md_text

    with open(md_output_file, "w", encoding="utf-8") as f:
        f.write(final_md_to_write)
    
    logger.info(f"PNG系列 '{safe_series_name}' 处理完成，已生成Markdown文件: {md_output_file}")
    return md_output_file

# --- 新增：PNG系列检测和分组函数 ---
def detect_png_series(png_files: List[Path]) -> Dict[str, List[Path]]:
    """
    检测和分组PNG系列图片
    
    参数:
        png_files: PNG文件路径列表
        
    返回:
        Dict[str, List[Path]]: 系列名称到PNG文件列表的映射
    """
    if not png_files:
        return {}
    
    # 尝试找到PNG系列的分组模式
    series_groups = {}
    
    # 查找常见的命名模式：基本名称后跟_01, _02, _1, _2等
    pattern = re.compile(r'^(.+?)(?:_0*(\d+))?\.png$', re.IGNORECASE)
    
    for png_file in png_files:
        match = pattern.match(png_file.name)
        if match:
            base_name = match.group(1)
            # 如果找不到编号，假设是单个文件系列
            if not match.group(2):
                group_name = base_name
                series_groups.setdefault(group_name, []).append(png_file)
            else:
                group_name = base_name
                series_groups.setdefault(group_name, []).append(png_file)
        else:
            # 对于不匹配模式的文件，将其视为独立系列
            group_name = png_file.stem
            series_groups.setdefault(group_name, []).append(png_file)
    
    # 对每个系列内的文件按名称排序
    for group_name, files in series_groups.items():
        series_groups[group_name] = sorted(files, key=lambda f: f.name.lower())
    
    # 只保留含有2个或更多文件的组作为"系列"，或者以特定模式命名的单文件
    valid_series = {}
    for group_name, files in series_groups.items():
        if len(files) >= 2:
            valid_series[group_name] = files
    
    return valid_series

# --- CLI入口 ---
def main():
    global SF_MAX_WORKERS # 允许通过CLI修改
    parser = argparse.ArgumentParser(description="any2md: 任何PDF、PPT/PPTX或PNG系列转Obsidian Markdown，支持Qwen-VL智能图片定位和可选的DeepSeek精炼")
    parser.add_argument("input_path", type=str, help="PDF/PPT/PPTX文件路径、PNG文件或系列所在目录路径，或包含这些文件的目录")
    parser.add_argument("-o", "--output", type=str, default="./output", help="输出目录 (默认: ./output)")
    parser.add_argument("--dpi", type=int, default=300, help="PDF/演示文稿转图片DPI(默认300)")
    parser.add_argument("--image-extraction-method", type=str, choices=["qwen_vl", "pymupdf"], default="qwen_vl", help="图片提取方式 (PDF/PNG系列): qwen_vl(智能定位) 或 pymupdf(嵌入图片) (默认: qwen_vl)")
    parser.add_argument("--visualize-localization", action="store_true", help="可视化Qwen-VL定位结果 (PDF/演示文稿/PNG系列)")
    parser.add_argument("--enable-refinement", action="store_true", help="启用第二阶段的Markdown精炼优化 (使用 REFINEMENT_MODEL)")
    parser.add_argument("-w", "--workers", type=int, default=SF_MAX_WORKERS, help=f"并行工作线程数 (默认: {SF_MAX_WORKERS} 或 .env中 SF_MAX_WORKERS)")
    parser.add_argument("--no-clean-temp", action="store_true", help="禁用处理完成后自动清理LibreOffice生成的临时PDF等文件")

    args = parser.parse_args()
    
    SF_MAX_WORKERS = args.workers

    target_path = Path(args.input_path)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not API_KEY:
        logger.error("错误：SF_API_KEY 环境变量未设置。请在 .env 文件或环境变量中提供API密钥。")
        return

    if args.enable_refinement and not REFINEMENT_API_KEY:
        logger.error("错误：精炼已启用，但 REFINEMENT_API_KEY 未设置。如果与主SF_API_KEY相同，请确保已设置或在.env中 REFINEMENT_API_KEY=你的主KEY。")
        return
    if args.enable_refinement and not REFINEMENT_MODEL_NAME:
        logger.warning("警告：精炼已启用，但 REFINEMENT_MODEL 环境变量未设置。将跳过精炼步骤。")
        args.enable_refinement = False

    pdf_files_to_process: List[Path] = []
    presentation_files_to_process: List[Path] = [] # Combined list for .ppt and .pptx
    png_files_to_process: List[Path] = []

    if target_path.is_file():
        ext = target_path.suffix.lower()
        if ext == ".pdf":
            pdf_files_to_process.append(target_path)
        elif ext in [".pptx", ".ppt"]:
            presentation_files_to_process.append(target_path)
        elif ext == ".png":
            png_files_to_process.append(target_path)
        else:
            logger.warning(f"不支持的文件类型: {target_path.name}。只支持 .pdf, .ppt, .pptx, .png 文件。")
    elif target_path.is_dir():
        logger.info(f"扫描目录 {target_path} 以查找支持的文件...")
        pdf_files_to_process.extend(sorted(target_path.rglob("*.pdf"))) 
        pdf_files_to_process.extend(sorted(target_path.rglob("*.PDF"))) 
        presentation_files_to_process.extend(sorted(target_path.rglob("*.pptx")))
        presentation_files_to_process.extend(sorted(target_path.rglob("*.PPTX")))
        presentation_files_to_process.extend(sorted(target_path.rglob("*.ppt")))
        presentation_files_to_process.extend(sorted(target_path.rglob("*.PPT")))
        png_files_to_process.extend(sorted(target_path.rglob("*.png")))
        png_files_to_process.extend(sorted(target_path.rglob("*.PNG")))
        
        pdf_files_to_process = sorted(list(set(pdf_files_to_process)))
        presentation_files_to_process = sorted(list(set(presentation_files_to_process)))
        png_files_to_process = sorted(list(set(png_files_to_process)))
    else:
        logger.error(f"提供的输入路径 '{target_path}' 不是有效的文件或目录。")
        return

    if not pdf_files_to_process and not presentation_files_to_process and not png_files_to_process:
        logger.error(f"在 '{target_path}' 中未找到PDF、PPT/PPTX或PNG文件。")
        return

    logger.info(f"找到 {len(pdf_files_to_process)} 个PDF, {len(presentation_files_to_process)} 个演示文稿 (PPT/PPTX), 和 {len(png_files_to_process)} 个PNG文件待处理。输出目录：{output_dir.resolve()}")

    # 检查LibreOffice是否可用
    libreoffice_available = is_libreoffice_available()
    
    # 如果有PPTX文件但LibreOffice不可用，显示警告
    if presentation_files_to_process and not libreoffice_available:
        logger.warning("========================================================")
        logger.warning("未检测到LibreOffice，无法处理PPTX文件。")
        logger.warning("为获得最佳结果，请安装LibreOffice:")
        logger.warning(" - Windows: https://www.libreoffice.org/download/download/")
        logger.warning(" - Linux: 使用包管理器安装，如: sudo apt install libreoffice")
        logger.warning(" - macOS: 使用Homebrew安装: brew install libreoffice")
        logger.warning("========================================================")
    
    # 使用通用文档处理函数处理所有文件
    processed_files = []
    
    # 处理PDF文件
    for pdf_file in pdf_files_to_process:
        result = process_document(
            doc_path=pdf_file,
            output_dir=output_dir,
            doc_type="pdf",
            image_extraction_method=args.image_extraction_method,
            dpi=args.dpi,
            visualize_localization=args.visualize_localization if args.image_extraction_method == "qwen_vl" else False,
            enable_refinement=args.enable_refinement
        )
        if result:
            processed_files.append(pdf_file)
    
    # 处理PPTX文件
    for pptx_file in presentation_files_to_process:
        result = process_document(
            doc_path=pptx_file,
            output_dir=output_dir,
            doc_type="presentation",
            dpi=args.dpi,
            visualize_localization=args.visualize_localization,
            enable_refinement=args.enable_refinement
        )
        if result:
            processed_files.append(pptx_file)
    
    # 新增：处理PNG系列
    if png_files_to_process:
        # 首先尝试检测PNG系列
        png_series = detect_png_series(png_files_to_process)
        
        if png_series:
            logger.info(f"检测到 {len(png_series)} 个PNG图片系列")
            
            for series_name, png_files in png_series.items():
                logger.info(f"处理PNG系列: {series_name} ({len(png_files)} 张图片)")
                result = process_png_series(
                    png_files=png_files,
                    output_dir=output_dir,
                    series_name=series_name,
                    image_extraction_method=args.image_extraction_method,  # 传递参数
                    visualize_localization=args.visualize_localization if args.image_extraction_method == "qwen_vl" else False,
                    enable_refinement=args.enable_refinement
                )
                if result:
                    processed_files.extend(png_files)
        else:
            # 如果没有检测到系列，单独处理每个PNG文件
            logger.info("未检测到PNG系列，将单独处理每个PNG文件")
            for png_file in png_files_to_process:
                result = process_document(
                    doc_path=png_file,
                    output_dir=output_dir,
                    doc_type="png",
                    image_extraction_method=args.image_extraction_method,  # 传递参数
                    visualize_localization=args.visualize_localization if args.image_extraction_method == "qwen_vl" else False,
                    enable_refinement=args.enable_refinement
                )
                if result:
                    processed_files.append(png_file)

    # 增强的临时文件清理
    # 默认在处理结束后执行清理，除非明确使用--no-clean-temp禁用
    if not args.no_clean_temp:
        try:
            logger.info("清理生成的临时文件...")
            # 清理明确的临时PDF文件
            clean_temp_files()
            
            # 对于PPTX处理，清理可能生成的临时文件
            for pptx_file in presentation_files_to_process:
                try:
                    # 注意：我们不删除原始PPTX文件，因为它们是用户的资源
                    temp_pdf_path = output_dir / f"{pptx_file.stem}.pdf"
                    if temp_pdf_path.exists():
                        logger.info(f"正在删除PPTX转换的临时PDF文件: {temp_pdf_path.name}")
                        os.remove(temp_pdf_path)
                except Exception as e:
                    logger.warning(f"无法删除临时PDF文件: {e}")
        except Exception as e:
            logger.warning(f"清理临时文件过程中发生错误: {e}")
    else:
        logger.info("已禁用自动清理临时文件 (--no-clean-temp)")
    
    # 处理总结
    if processed_files:
        logger.info(f"成功处理了 {len(processed_files)} 个文件")
    else:
        logger.warning("没有成功处理任何文件")
                
    logger.info("处理完成。")

# 在处理结束后执行，可以显式清理临时目录
def clean_temp_files():
    """清理系统临时目录中可能的LibreOffice临时文件"""
    try:
        # 获取临时目录路径
        temp_dir = tempfile.gettempdir()
        # 查找与LibreOffice相关的临时文件和目录
        libreoffice_temp_patterns = [
            "lu*", # LibreOffice User 临时文件
            "libreoffice_*",
            "LibreOffice_*",
            "tmp_*_libreoffice*",
            "ppt_*", # 可能的PPT相关临时文件
            "libreoffice_pptx_*" # 我们自己创建的临时目录前缀
        ]
        
        now = time.time()
        # 最小文件年龄 (1小时前)，避免清理正在使用的文件
        min_age = 3600
        
        for pattern in libreoffice_temp_patterns:
            for item in Path(temp_dir).glob(pattern):
                try:
                    # 检查文件年龄
                    if (now - item.stat().st_mtime) > min_age:
                        if item.is_dir():
                            logger.info(f"清理临时目录: {item}")
                            shutil.rmtree(item, ignore_errors=True)
                        else:
                            logger.info(f"清理临时文件: {item}")
                            os.remove(item)
                except Exception as e:
                    logger.warning(f"清理临时项目 {item} 失败: {e}")
    except Exception as e:
        logger.warning(f"临时文件清理过程出错: {e}")

# 新增辅助函数，用于验证和净化JSON数据
def extract_and_validate_json(json_str: str) -> Dict[str, Any]:
    """从字符串中提取、解析和验证JSON数据，确保返回格式统一且安全"""
    try:
        result = json.loads(json_str)
        # 安全检查，确保返回值是一个字典且包含image_regions键
        if not isinstance(result, dict):
            logger.warning(f"提取的JSON不是字典: {result}")
            return {"image_regions": []}
        if "image_regions" not in result:
            logger.warning(f"提取的JSON缺少image_regions键: {result}")
            return {"image_regions": []}
        # 确保image_regions是列表，且所有元素都是字典
        image_regions = result["image_regions"]
        if not isinstance(image_regions, list):
            logger.warning(f"image_regions不是列表: {image_regions}")
            return {"image_regions": []}
        # 过滤掉非字典元素
        valid_regions = []
        for region in image_regions:
            if not isinstance(region, dict):
                logger.warning(f"区域不是字典: {region}")
                continue
            valid_regions.append(region)
        result["image_regions"] = valid_regions
        return result
    except json.JSONDecodeError as e:
        logger.error(f"JSON解析错误: {e}")
        return {"image_regions": []}
    except Exception as e:
        logger.error(f"验证JSON时出错: {e}")
        return {"image_regions": []}

# --- 新增：精炼后的安全处理函数 ---
def clean_markdown_code_wrappers(markdown_text: str) -> str:
    """
    删除LLM精炼可能添加的```markdown ... ```样式的代码块包裹。
    
    参数:
        markdown_text: 可能包含代码块包裹的Markdown文本
        
    返回:
        str: 清理后的Markdown文本
    """
    # 检查开头是否有代码块标记
    if markdown_text.lstrip().startswith("```markdown") or markdown_text.lstrip().startswith("```md"):
        # 找到第一个代码块标记的结束位置
        start_fence_end = markdown_text.find("\n", markdown_text.find("```"))
        if start_fence_end != -1:
            # 找到结束代码块标记的开始位置
            end_fence_start = markdown_text.rfind("```")
            if end_fence_start > start_fence_end:
                # 提取代码块之间的内容
                return markdown_text[start_fence_end+1:end_fence_start].strip()
    
    return markdown_text

def fix_image_references(markdown_text: str, image_metadata: Dict[str, str]) -> str:
    """
    使用已知的图像元数据修正Markdown文本中的图像引用。
    
    参数:
        markdown_text: 待处理的Markdown文本
        image_metadata: 描述到正确图像路径的映射字典
    
    返回:
        str: 修正后的Markdown文本
    """
    # 用于查找图像引用的正则表达式
    image_pattern = re.compile(r'!\[(.*?)\]\((.*?)\)')
    
    # 找到所有图像引用
    image_matches = image_pattern.findall(markdown_text)
    
    # 处理每个匹配项
    for desc, path in image_matches:
        # 如果路径已经正确，跳过
        if path in image_metadata.values():
            continue
            
        # 尝试通过描述找到正确的路径
        if desc in image_metadata:
            correct_path = image_metadata[desc]
            # 替换不正确的路径为正确的路径
            old_ref = f"![{desc}]({path})"
            new_ref = f"![{desc}]({correct_path})"
            markdown_text = markdown_text.replace(old_ref, new_ref)
            continue
            
        # 如果精确描述匹配失败，尝试模糊匹配
        best_match = None
        best_score = 0.6  # 相似度阈值
        
        for known_desc in image_metadata:
            # 简单的子字符串匹配或标准化字符串比较
            if (desc.lower() in known_desc.lower() or 
                known_desc.lower() in desc.lower()):
                # Jaccard相似度（共同字符的比例）
                score = len(set(desc.lower()) & set(known_desc.lower())) / len(set(desc.lower()) | set(known_desc.lower()))
                if score > best_score:
                    best_score = score
                    best_match = known_desc
        
        if best_match:
            correct_path = image_metadata[best_match]
            old_ref = f"![{desc}]({path})"
            new_ref = f"![{desc}]({correct_path})"
            markdown_text = markdown_text.replace(old_ref, new_ref)
            
    return markdown_text

if __name__ == "__main__":
    # 为PyMuPDF添加一个导入检查，以防用户没有安装它但选择了pymupdf方法
    try:
        import fitz # PyMuPDF
    except ImportError:
        logger.warning("PyMuPDF (fitz) 未安装。'pymupdf' 图片提取方法将不可用。")
    
    # 检查Windows系统上的comtypes支持
    if IS_WINDOWS:
        try:
            import comtypes.client
            COMTYPES_SUPPORT = True
        except ImportError:
            COMTYPES_SUPPORT = False
            logger.warning("comtypes未安装。在Windows上使用PowerPoint导出PPTX将不可用。请使用 pip install comtypes 安装。")
    
    # 检查pyautogui支持
    try:
        import pyautogui
        PYAUTOGUI_SUPPORT = True
    except ImportError:
        PYAUTOGUI_SUPPORT = False
        logger.warning("pyautogui未安装。部分替代截图方法将不可用。")
    
    main()
