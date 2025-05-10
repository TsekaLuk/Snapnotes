#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试脚本：利用Qwen2.5-VL的视觉定位能力提取PDF页面中的图表区域

此脚本演示如何：
1. 将PDF页面转换为图像
2. 使用Qwen2.5-VL模型获取图像中所有图表的坐标
3. 根据坐标裁剪图像
4. 将裁剪后的图像保存到指定目录
"""

import os
import io
import json
import base64
import argparse
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
import logging

import requests
import fitz  # PyMuPDF
from PIL import Image, ImageDraw
from pdf2image import convert_from_path
from dotenv import load_dotenv
import tenacity
from tqdm import tqdm

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# 加载环境变量
load_dotenv()

# 从环境变量获取配置
API_URL = os.getenv("SF_API_URL", "https://api.siliconflow.cn/v1/chat/completions")
API_KEY = os.getenv("SF_API_KEY")
MODEL = os.getenv("SF_MODEL", "Qwen/Qwen2.5-VL-72B-Instruct")

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

# 特殊的提示词，要求Qwen2.5-VL识别图表区域并返回位置信息
LOCALIZATION_SYSTEM_PROMPT = """
你是一个专门帮助用户分析PDF文档页面图像并定位其中图表的视觉分析助手。

你的任务是：
1. 确定页面图像中所有包含数据、图形或插图的图表区域（如柱状图、曲线图、示意图、插图、流程图等）。
2. 对于找到的每个图表区域，提供其精确的像素坐标边界框（bounding box）。

返回格式要求：
1. 你**必须**以JSON格式返回结果，包含一个"image_regions"数组。
2. 对于每个识别到的图表区域，都在数组中添加一个对象，必须包含以下属性：
   - "id": 区域的唯一标识符，如"region_1"、"region_2"等
   - "type": 区域的类型，如"图表"、"插图"、"表格"等
   - "bbox": 一个包含4个整数值的数组[x1, y1, x2, y2]，表示左上角和右下角的坐标
   - "description": 对区域内容的简短描述（不超过10个中文字）

3. 此外，如果能够判断，请添加：
   - "confidence": 0.0-1.0之间的置信度

示例输出：
```json
{
  "image_regions": [
    {
      "id": "region_1",
      "type": "图表",
      "bbox": [50, 120, 400, 350],
      "description": "二次函数图像",
      "confidence": 0.95
    },
    {
      "id": "region_2",
      "type": "表格",
      "bbox": [80, 500, 500, 650],
      "description": "数据对照表",
      "confidence": 0.85
    }
  ]
}
```

注意：
- 坐标系原点是图像的左上角，x增大方向是向右，y增大方向是向下。
- 如果页面中没有发现任何图表区域，则返回空数组。
- 确保所有坐标都是整数，且在图像尺寸范围内有效。
- 小于文档页面5%面积的小图形或标记通常不需要识别（如页码、小图标等）。
- 仅关注明显的图表区域，而不是纯文本区域。
"""

LOCALIZATION_USER_INSTRUCTION = """请分析此页面图像，识别所有包含数据、图形或插图的重要图表区域，并以指定的JSON格式返回它们的精确位置和简短描述。你的输出应该只包含JSON，没有其他文本。"""

@tenacity.retry(
    wait=tenacity.wait_fixed(2),
    stop=tenacity.stop_after_attempt(5),
    retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,)),
    retry_error_callback=lambda retry_state: f"Error after {retry_state.attempt_number} attempts: {retry_state.outcome.exception()}" if retry_state.outcome.failed else retry_state.outcome.result(),
)
def call_vision_api_for_localization(base64_img: str) -> Dict[str, Any]:
    """调用视觉模型API来获取图像的区域定位信息"""
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
        "temperature": 0.1,  # 使用低温度以获得更确定性的结果
        "max_tokens": 2000,
        "enable_thinking": False
    }
    
    # 设置更长的超时时间，因为图像处理可能需要更多时间
    resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=180)
    resp.raise_for_status()
    
    response_json = resp.json()
    text_content = response_json["choices"][0]["message"]["content"]
    
    # 从响应文本中提取JSON
    try:
        # 尝试直接解析回复内容（假设是纯JSON）
        parsed_data = json.loads(text_content)
        return parsed_data
    except json.JSONDecodeError:
        # 如果不是纯JSON，尝试查找```json```标记提取出JSON部分
        logger.warning("Response is not pure JSON, attempting to extract JSON from markdown...")
        try:
            # 尝试提取```json```和```之间的内容
            json_start = text_content.find("```json")
            if json_start != -1:
                json_start = text_content.find("\n", json_start) + 1
                json_end = text_content.find("```", json_start)
                json_str = text_content[json_start:json_end].strip()
                return json.loads(json_str)
            
            # 尝试提取```和```之间的内容（假设是JSON）
            json_start = text_content.find("```")
            if json_start != -1:
                json_start = text_content.find("\n", json_start) + 1
                json_end = text_content.find("```", json_start)
                json_str = text_content[json_start:json_end].strip()
                return json.loads(json_str)
                
            # 最后尝试查找首个{和最后一个}之间的内容
            json_start = text_content.find("{")
            json_end = text_content.rfind("}") + 1
            if json_start != -1 and json_end > json_start:
                json_str = text_content[json_start:json_end].strip()
                return json.loads(json_str)
                
            raise ValueError("Could not extract JSON from response")
        except Exception as e:
            logger.error(f"Failed to extract JSON from response: {e}")
            logger.error(f"Original response: {text_content}")
            # 返回一个空的结构以避免引发错误
            return {"image_regions": []}

def convert_pdf_page_to_image(pdf_path: Path, page_number: int, dpi: int = 300) -> Image.Image:
    """将PDF页面转换为PIL图像"""
    # 页码从0开始计算，但pdf2image从1开始计算
    images = convert_from_path(
        pdf_path, 
        dpi=dpi, 
        first_page=page_number + 1, 
        last_page=page_number + 1,
        fmt='png',  # 使用PNG以保持质量
        thread_count=1
    )
    if not images:
        raise ValueError(f"Failed to convert page {page_number} of {pdf_path}")
    return images[0]

def image_to_base64(img: Image.Image, format: str = "WEBP") -> str:
    """将PIL图像转换为base64编码的字符串"""
    buffer = io.BytesIO()
    img.save(buffer, format=format, quality=95)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")

def extract_image_regions(
    pdf_path: Path, 
    output_dir: Path, 
    page_number: int, 
    dpi: int = 300, 
    visualize: bool = False
) -> List[Path]:
    """提取PDF页面中的图像区域，返回保存的图像文件路径列表"""
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 1. 将PDF页面转换为图像
    logger.info(f"Converting page {page_number} of {pdf_path.name} to image...")
    page_image = convert_pdf_page_to_image(pdf_path, page_number, dpi)
    
    # 2. 将图像转换为base64
    base64_image = image_to_base64(page_image)
    
    # 3. 调用API获取图像区域位置信息
    logger.info(f"Calling vision API to locate image regions...")
    try:
        localization_result = call_vision_api_for_localization(base64_image)
    except Exception as e:
        logger.error(f"Error calling vision API: {e}")
        return []
    
    # 4. 从API结果中提取区域信息
    image_regions = localization_result.get("image_regions", [])
    if not image_regions:
        logger.info(f"No image regions detected on page {page_number} of {pdf_path.name}")
        return []
    
    logger.info(f"Found {len(image_regions)} image regions on page {page_number} of {pdf_path.name}")
    
    # 创建可视化的调试图像
    if visualize:
        debug_image = page_image.copy()
        draw = ImageDraw.Draw(debug_image)
        
    # 5. 裁剪并保存每个区域
    saved_image_paths = []
    
    for i, region in enumerate(image_regions):
        bbox = region.get("bbox")
        if not bbox or len(bbox) != 4:
            logger.warning(f"Invalid bbox for region {i}: {bbox}")
            continue
        
        try:
            # 确保坐标是整数且在图像范围内
            x1, y1, x2, y2 = [int(coord) for coord in bbox]
            x1 = max(0, x1)
            y1 = max(0, y1)
            x2 = min(page_image.width, x2)
            y2 = min(page_image.height, y2)
            
            # 确保区域有合理大小
            if x2 <= x1 or y2 <= y1 or (x2 - x1) < 20 or (y2 - y1) < 20:
                logger.warning(f"Region {i} has invalid dimensions: {x1},{y1},{x2},{y2}")
                continue
                
            # 裁剪图像
            cropped_image = page_image.crop((x1, y1, x2, y2))
            
            # 生成输出文件名
            region_type = region.get("type", "图表")
            region_desc = region.get("description", f"region_{i+1}")
            # 安全文件名
            safe_desc = "".join(c if c.isalnum() or c in "- " else "_" for c in region_desc).strip()
            filename = f"page{page_number+1}_{safe_desc}_{i+1}.png"
            output_path = output_dir / filename
            
            # 保存裁剪后的图像
            cropped_image.save(output_path, format="PNG")
            saved_image_paths.append(output_path)
            
            logger.info(f"Saved region {i+1}: {region_type} - {region_desc} to {output_path}")
            
            # 在调试图像上绘制边界框
            if visualize:
                draw.rectangle([x1, y1, x2, y2], outline="red", width=3)
                draw.text((x1, y1-15), f"{i+1}: {region_type}", fill="red")
            
        except Exception as e:
            logger.error(f"Error extracting region {i}: {e}")
    
    # 保存调试图像（带有边界框标记）
    if visualize and image_regions:
        debug_path = output_dir / f"page{page_number+1}_regions_debug.png"
        debug_image.save(debug_path)
        logger.info(f"Saved visualization with bounding boxes to {debug_path}")
    
    return saved_image_paths

def process_pdf(
    pdf_path: Path, 
    output_dir: Path, 
    page_range: Optional[Tuple[int, int]] = None,
    dpi: int = 300,
    visualize: bool = False
) -> Dict[int, List[Path]]:
    """处理PDF的指定页面范围，提取所有图像区域"""
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    
    # 确定处理的页面范围
    try:
        doc = fitz.open(str(pdf_path))
        total_pages = doc.page_count
        doc.close()
        
        if page_range is None:
            start_page, end_page = 0, total_pages - 1
        else:
            start_page = max(0, page_range[0])
            end_page = min(total_pages - 1, page_range[1])
    except Exception as e:
        logger.error(f"Error opening PDF with PyMuPDF: {e}")
        raise
    
    logger.info(f"Processing {pdf_path.name} pages {start_page}-{end_page} (0-indexed)")
    
    # 创建以PDF名称为基础的输出子目录
    pdf_stem = pdf_path.stem
    safe_stem = "".join(c if c.isalnum() or c in "- " else "_" for c in pdf_stem).strip()
    pdf_output_dir = output_dir / f"{safe_stem}_extracted_images"
    pdf_output_dir.mkdir(parents=True, exist_ok=True)
    
    # 处理每一页并收集结果
    results = {}
    for page_num in tqdm(range(start_page, end_page + 1), desc="Processing PDF pages"):
        try:
            image_paths = extract_image_regions(
                pdf_path, 
                pdf_output_dir, 
                page_num, 
                dpi=dpi,
                visualize=visualize
            )
            if image_paths:
                results[page_num] = image_paths
        except Exception as e:
            logger.error(f"Error processing page {page_num}: {e}")
    
    # 生成结果摘要信息
    total_regions = sum(len(paths) for paths in results.values())
    total_pages_with_regions = len(results)
    
    logger.info(f"✅ 完成处理 {pdf_path.name}:")
    logger.info(f"  - 共处理页面: {end_page - start_page + 1} 页")
    logger.info(f"  - 包含图表的页面: {total_pages_with_regions} 页")
    logger.info(f"  - 提取图表总数: {total_regions} 个")
    logger.info(f"  - 输出目录: {pdf_output_dir}")
    
    return results

def main():
    parser = argparse.ArgumentParser(description="使用Qwen2.5-VL从PDF页面中提取图像区域")
    parser.add_argument("pdf_path", type=str, help="PDF文件路径")
    parser.add_argument("-o", "--output", type=str, default="./extracted_images", help="输出目录")
    parser.add_argument("-p", "--pages", type=str, help="要处理的页面范围 (例如: '0-5' 或 '1,3,5')")
    parser.add_argument("--dpi", type=int, default=300, help="PDF转图像的DPI (默认: 300)")
    parser.add_argument("-v", "--visualize", action="store_true", help="生成带有边界框的可视化图像")
    
    args = parser.parse_args()
    
    # 解析页面范围
    page_range = None
    if args.pages:
        if "-" in args.pages:
            # 范围形式 "0-5"
            start, end = map(int, args.pages.split("-"))
            page_range = (start, end)
        elif "," in args.pages:
            # 列表形式 "0,2,4"
            pages = sorted(map(int, args.pages.split(",")))
            if not pages:
                raise ValueError("Invalid page list")
            page_range = (min(pages), max(pages))
        else:
            # 单个页面
            page = int(args.pages)
            page_range = (page, page)
    
    pdf_path = Path(args.pdf_path)
    output_dir = Path(args.output)
    
    # 启动处理
    process_pdf(
        pdf_path=pdf_path,
        output_dir=output_dir,
        page_range=page_range,
        dpi=args.dpi,
        visualize=args.visualize
    )

if __name__ == "__main__":
    main()
