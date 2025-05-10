#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Batch PDF → (VLM) → Obsidian‑ready Markdown
"""
import os, sys, base64, io, argparse, pathlib, json, textwrap
from concurrent.futures import ProcessPoolExecutor, as_completed, ThreadPoolExecutor
from functools import partial
from pdf2image import convert_from_path
from PIL import Image
import requests, tqdm, tenacity
from dotenv import load_dotenv
import fitz  # PyMuPDF
import hashlib
import collections

##############################
# 0 环境变量 & 常量
##############################
load_dotenv()                                           # .env 里放 API_KEY
API_URL  = os.getenv("SF_API_URL",  "https://api.siliconflow.cn/v1/chat/completions")
API_KEY  = os.getenv("SF_API_KEY")      # export SF_API_KEY='sk-....'
MODEL    = os.getenv("SF_MODEL", "Qwen/Qwen2.5-VL-72B-Instruct")
POST_PROCESS_MODEL = os.getenv("POST_PROCESS_MODEL", "Pro/deepseek-ai/DeepSeek-V3-1226")
SF_MAX_WORKERS = int(os.getenv("SF_MAX_WORKERS", "4")) # Default to 4 if not set in .env

HEADERS  = {"Authorization": f"Bearer {API_KEY}",
            "Content-Type":  "application/json"}
TASK_DESCRIPTION = "PDF页面"

SYSTEM_PROMPT = (
    "你是一名专业的文档转写助手，任务是将【" + TASK_DESCRIPTION + "】完整且准确地转写为【Obsidian Flavored Markdown】格式。"
    "最终目标是生成一份结构清晰、内容无误、便于阅读和在Obsidian中使用的Markdown文档。"
    "\n\n请严格遵循以下规范进行转写：\n"
    "1.  **内容完整性优先**：\n"
    "    *   必须转写页面上所有可见的文本内容，包括题目、题号、选项、解题步骤、数学公式、图表相关的文字说明（如图注、“见图X”等）。\n"
    "    *   页眉、页脚中的重要上下文信息（如试卷年份、科目名称、页码等）应予以保留。\n"
    "    *   **核心要求：确保不丢失任何来自原始PDF的信息。**\n\n"
    "2.  **Markdown结构化输出**：\n"
    "    *   **标题层级**：使用Markdown的标题标记（例如 `# 一级标题`，`## 二级标题`，`### 三级标题`）来反映原文的章节、题型（如“选择题”、“解答题”）和大题号的层级结构。例如，‘**## 二、填空题**’，‘**### 1.**’。\n"
    "    *   **列表与选项**：选择题的选项（如 A、B、C、D）应清晰列出，通常每个选项占一行，可考虑使用Markdown列表。\n"
    "    *   **段落与换行**：自然段落间应有空行。公式和文本的布局应尽可能还原原始逻辑。\n\n"
    "3.  **数学表达式 (LaTeX)**：\n"
    "    *   所有数学公式、单个数学符号和数学表达式【必须】准确无误地转写为【合法 LaTeX】。\n"
    "    *   **行内公式**：使用单美元符号 `$...$` 包裹，例如：函数 $f(x) = ax + b$。\n"
    "    *   **行间公式（块级公式）**：重要的、独立成行的公式使用双美元符号 `$$...$$` 包裹。\n"
    "    *   **表格**：表格的转写需要尽可能保持表格的原始结构，包括列宽、对齐方式、表内符号、角分割等。\n"
    "    *   **特别注意**：仔细区分普通文本中的特殊字符（如 `*`, `_`, `{`, `}`）与LaTeX命令中的这些字符，避免错误转义或格式冲突。\n\n"
    "4.  **文本与格式化细节**：\n"
    "    *   **忽略视觉样式**：原始PDF的字体、字号、颜色、具体缩进和精确布局等视觉表现信息通常应忽略，转而专注于内容的语义和逻辑结构。\n"
    "    *   **语义强调**：如果原文通过【加粗】或【斜体】来强调特定术语、变量、定理名称或关键步骤，请在Markdown中使用相应的 `**加粗**` 或 `*斜体*` 来保留这种语义强调。\n"
    "    *   **水印处理**：【请务必忽略】页面背景中任何形式的水印（文字、图案、logo等），绝对不要将水印内容转写出来。\n"
    "    *   **图示示例**：如果原文中包含图片、图表或插图，请首先尝试转写其标题或围绕该图片的任何描述性文字。然后在图片应该出现的位置插入文本占位符：`[图片]`。不要尝试描述图片内容本身。\n\n"
    "5.  **转写核心原则**：\n"
    "    *   【严禁编造、摘要或解释】任何内容。输出必须与原版PDF在文字和数学公式上【逐字逐式高度一致】。\n"
    "    *   专注于“转写”，而非“创作”或“理解内容并复述”。\n\n"
    "6.  **输出格式要求**：\n"
    "    *   输出内容【仅能包含纯粹的 Markdown 文本】。不要在Markdown文本的开头或结尾添加任何如 ```markdown ... ``` 这样的代码块包裹。\n\n"
    "请仔细分析每个页面的结构和内容，确保转写质量达到最高标准。"
)

USER_INSTRUCTION = "请将此页转写为 Obsidian Markdown。若有 LaTeX 数学符号需正确转义。"

REFINEMENT_SYSTEM_PROMPT = (
    "你是一名专业的Obsidian Markdown文档编辑和优化助手。"
    "你将收到一份从PDF逐页转写并初步合并的Markdown文本，这份文本可能存在以下问题：\n"
    "1. 重复的页眉或页脚：由于逐页转写，原文PDF中的页眉页脚可能在合并文本中反复出现。\n"
    "2. 格式不一致：不同页面转写的内容可能在Markdown格式（如标题、列表）上存在细微差异。\n"
    "3. 逻辑中断：较长的段落、题目或解题步骤可能因为跨页而被切断。\n"
    "4. 图片占位符：文中可能包含 `[图片]` 占位符，这些占位符本身不需要你处理，它们将在后续步骤中被替换为实际的图片链接。你的任务是确保这些占位符周围的文本格式正确且逻辑连贯。\n\n"
    "5. 加粗符号：**内容**的**前后最好加一个空格，例如“**【答案】**6”修正为“**【答案】** 6”这样加粗才会正确渲染。\n"
    "你的任务是：\n"
    "A. **智能识别并移除重复的页眉和页脚**。请注意保留有意义的、仅出现一次的文档标题或章节信息，不要误删。\n"
    "B. **统一和规范化Markdown格式**：\n"
    "   - 确保标题层级（#，##，### ...）在整个文档中一致且符合逻辑。\n"
    "   - 统一列表（有序、无序）的格式。\n"
    "   - 确保数学公式（行内 $...$ 和块级 $$...$$）的 LaTeX 语法正确且一致。\n"
    "C. **提升内容连贯性**：\n"
    "   - 尽力识别并逻辑上连接那些因分页而被中断的内容。例如，如果一个段落或解题步骤明显在下一页继续，请尝试平滑地将它们整合。\n"
    "   - 修正因分页导致的突兀的换行或段落中断。\n"
    "D. **保持内容准确性**：在进行格式调整和结构优化的同时，【绝对不能修改原始文本的语义内容或数学公式的准确性】。你的核心是优化结构和移除冗余元信息，而非重写、删减或解释内容，请保证内容的完整性与一致性。\n"
    "E. **输出纯净Markdown**：最终输出必须是纯粹的Markdown文本，不包含任何额外的解释或代码块包裹。\n\n"
    "请仔细分析输入文本，并输出一份高质量、结构清晰、阅读流畅的最终Obsidian Markdown文档。"
)


##############################
# 1 LLM 调用
##############################
@tenacity.retry(wait=tenacity.wait_fixed(2), stop=tenacity.stop_after_attempt(5),
                retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,)))
def call_vlm(base64_img: str, enable_thinking=False, temperature=0.2, max_tokens=4096) -> str:
    """
    给定一个 base64 Image 调用 Qwen‑VL 并返回 Markdown
    """
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {
                        "url": f"data:image/webp;base64,{base64_img}",
                        "detail": "high"
                    }},
                    {"type": "text", "text": USER_INSTRUCTION}
                ]
            }
        ],
        "stream": False,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "enable_thinking": enable_thinking
    }
    resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=120)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

@tenacity.retry(wait=tenacity.wait_fixed(5), stop=tenacity.stop_after_attempt(3),
                retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,)))
def call_llm_for_refinement(markdown_text: str, model_to_use: str) -> str:
    """
    给定一个 Markdown 文本调用 DeepSeek‑VL 并返回精炼后的 Markdown
    """
    payload = {
        "model": model_to_use,
        "messages": [
            {"role": "system", "content": REFINEMENT_SYSTEM_PROMPT},
            {"role": "user", "content": markdown_text}
        ],
        "stream": False,
        "temperature": 0.2,
        "max_tokens": 4096,
        "enable_thinking": False
    }
    resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=600)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

def image_to_base64(img: Image.Image) -> str:
    """PIL.Image → base64 webp string"""
    buf = io.BytesIO()
    img.save(buf, format="WEBP", quality=95)
    return base64.b64encode(buf.getvalue()).decode("utf-8")

##############################
# 2 核心转换函数
##############################
def extract_images_from_pdf(pdf_path: pathlib.Path, assets_dir: pathlib.Path) -> list[str]:
    """Extracts content images from PDF, filters watermarks/small images, and saves them. Returns a list of relative image paths."""
    
    candidate_images_info = [] 
    min_dimension = 50  # Minimum width/height in pixels for an image to be considered content

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        print(f"⚠️ WARNING: Could not open PDF {pdf_path.name} with PyMuPDF: {e}. Skipping image extraction.")
        return []

    print(f"DEBUG: Starting image extraction for {pdf_path.name}. Total pages: {doc.page_count}")

    image_hashes_on_pages = collections.defaultdict(list)
    candidate_images = [] # List of (xref, page_num, width, height, image_bytes)

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        image_list = page.get_images(full=True)
        if image_list:
            print(f"DEBUG: Page {page_num + 1}: Found {len(image_list)} raw image entries.")
        for img_index, img_info in enumerate(image_list):
            xref = img_info[0]
            if xref == 0: # Skip if xref is 0, can indicate an invalid image entry
                print(f"DEBUG: Page {page_num + 1}, Img {img_index}: Skipping xref 0.")
                continue
            
            try:
                base_image = doc.extract_image(xref)
            except Exception as e:
                print(f"DEBUG: Page {page_num + 1}, Img {img_index} (xref: {xref}): Failed to extract image - {e}")
                continue
                
            image_bytes = base_image["image"]
            width = base_image["width"]
            height = base_image["height"]

            if width < min_dimension or height < min_dimension:
                print(f"DEBUG: Page {page_num + 1}, Img {img_index} (xref: {xref}): Filtered out by size ({width}x{height}).")
                continue
            
            # Store all images that pass the size filter along with their page number
            candidate_images.append((xref, page_num, width, height, image_bytes))
            # For watermark detection, hash the image and record its presence on this page
            img_hash = hashlib.md5(image_bytes).hexdigest()
            image_hashes_on_pages[img_hash].append(page_num)
    
    print(f"DEBUG: Found {len(candidate_images)} candidate images after size filtering.")
    # 过滤掉水印（出现在太多页上的图像）
    # 还过滤掉即使不是 >50% 页数也太常见的图像（例如现在 5 页）
    # 这有助于小 logo 或反复出现的小图标，如果 total_pages 较小。
    # MAX_OCCURRENCES_FOR_NON_WATERMARK = 5 
    
    final_images_to_save = [] # List of (page_num, img_index_on_page, image_bytes)
    page_image_counter = collections.defaultdict(int)
    saved_image_paths_relative = []

    # Sort candidate images by page number and then by original xref (to maintain some order)
    candidate_images.sort(key=lambda x: (x[1], x[0]))

    for xref, page_num, width, height, image_bytes in candidate_images:
        img_hash = hashlib.md5(image_bytes).hexdigest()
        num_occurrences = len(image_hashes_on_pages[img_hash])
        
        is_watermark = num_occurrences > (doc.page_count * 0.5)
        # is_too_common = num_occurrences > MAX_OCCURRENCES_FOR_NON_WATERMARK
        
        if is_watermark: # or is_too_common:
            print(f"DEBUG: Page {page_num + 1} (xref: {xref}): Filtered out as watermark/too_common. Occurrences: {num_occurrences}/{doc.page_count}.")
            continue

        # If it passes all filters, add to final list
        current_img_index_on_page = page_image_counter[page_num]
        page_image_counter[page_num] += 1
        final_images_to_save.append((page_num, current_img_index_on_page, image_bytes))

    if not final_images_to_save:
        print(f"INFO: No images to save for {pdf_path.name} after all filters (size, watermark).")
        doc.close()
        return []
    
    print(f"DEBUG: {len(final_images_to_save)} images selected for saving after all filters.")

    for page_num, img_idx_on_page, image_bytes in final_images_to_save:
        image_filename = f"page{page_num + 1}_img{img_idx_on_page + 1}.png"
        image_filename_abs = assets_dir / image_filename
        
        try:
            with open(image_filename_abs, "wb") as img_file:
                img_file.write(image_bytes)
            
            # Ensure asset_dir.name is used for the relative path correctly
            relative_image_path = str(pathlib.Path(assets_dir.name) / image_filename).replace("\\", "/")
            saved_image_paths_relative.append(relative_image_path)
        except Exception as e:
            print(f"⚠️ WARNING: Could not save image {image_filename} to {assets_dir.name}: {e}")
    
    doc.close()
    if saved_image_paths_relative:
        print(f"INFO: Successfully filtered and saved {len(saved_image_paths_relative)} images for {pdf_path.name}.")
    elif candidate_images: # Had candidates, but all were filtered
        print(f"INFO: All {len(candidate_images)} candidate images for {pdf_path.name} were filtered out (e.g. as watermarks/small).")
    else: # No candidates to begin with
        print(f"INFO: No suitable images found to extract from {pdf_path.name} after initial filtering.")
        
    return saved_image_paths_relative

def process_pdf(pdf_path: str, out_dir: str, dpi=300) -> pathlib.Path:
    """
    单个 PDF → Obsidian Markdown
    """
    pdf_path = pathlib.Path(pdf_path)
    out_dir = pathlib.Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Sanitize PDF stem for directory naming
    # More aggressive sanitization: allow only alphanumeric, underscore, hyphen.
    # Replace spaces and other problematic characters with a single underscore.
    # Collapse multiple underscores.
    temp_stem = []
    last_char_was_underscore = False
    for char_code in pdf_path.stem.encode('utf-8', 'ignore').decode('utf-8'): # Attempt to handle unicode chars better
        if char_code.isalnum() or char_code == '-':
            temp_stem.append(char_code)
            last_char_was_underscore = False
        else:
            if not last_char_was_underscore:
                temp_stem.append('_')
            last_char_was_underscore = True
    sanitized_stem = "".join(temp_stem).strip('_') # Remove leading/trailing underscores
    if not sanitized_stem: # Handle cases where the stem becomes empty after sanitization
        sanitized_stem = "pdf_document"

    assets_dir = out_dir / (sanitized_stem + "_assets")
    extracted_image_paths = [] # Initialize here
    try:
        assets_dir.mkdir(parents=True, exist_ok=True)
        # Stage 0: Extract images from PDF only if assets_dir was created
        print(f"INFO: Attempting to extract images from {pdf_path.name} into {assets_dir.resolve()}")
        extracted_image_paths = extract_images_from_pdf(pdf_path, assets_dir)
        if extracted_image_paths:
            print(f"INFO: Extracted {len(extracted_image_paths)} images for {pdf_path.name}.")
        # else: # Removed the 'no images found' print from here to avoid confusion if filtering is aggressive
            # print(f"INFO: No images found or extracted from {pdf_path.name} after filtering.")
            
    except OSError as e:
        print(f"❌ ERROR: Could not create or access assets directory {assets_dir.resolve()}: {e}. Skipping image extraction for this PDF.")
    except ValueError as e: # Specifically for relative_to issues
        print(f"❌ ERROR: Path issue with assets directory {assets_dir.resolve()}: {e}. Skipping image extraction for this PDF.")

    # Stage 1: Convert PDF pages to images and get initial Markdown transcription
    md_parts = []
    num_pages = 0
    try:
        # Temporarily open with pdf2image to get page count for tqdm
        temp_images = convert_from_path(pdf_path, dpi=10, thread_count=1) # Low DPI just for count
        num_pages = len(temp_images)
    except Exception as e:
        print(f"❌ ERROR: Could not get page count for {pdf_path.name}: {e}. Skipping this PDF.")
        return None # Indicate failure
    if num_pages == 0:
        print(f"❌ ERROR: PDF {pdf_path.name} has 0 pages or is unreadable. Skipping.")
        return None

    with ThreadPoolExecutor(max_workers=SF_MAX_WORKERS) as executor: # Max 4 concurrent API calls for politeness
        futures = [executor.submit(process_pdf_page_to_image_and_text, pdf_path, i, dpi)
                   for i in range(num_pages)]
        for fut in tqdm.tqdm(as_completed(futures), total=len(futures),
                             desc=f"{pdf_path.name} (pages)", unit="page"):
            try:
                page_no, text = fut.result()
                md_parts.append((page_no, text))
            except Exception as e:
                # This exception should ideally be caught within process_pdf_page_to_image_and_text
                # but as a fallback:
                page_no_approx = -1 # Placeholder if page_no cannot be determined
                md_parts.append((page_no_approx, f"> **Critical error processing a page: {e}**"))

    md_parts.sort(key=lambda p: p[0])
    raw_md_text = "\n\n".join(txt.strip() for _, txt in md_parts)

    # Stage 2: Refine the raw Markdown text if POST_PROCESS_MODEL is set
    refined_md_text = raw_md_text
    if POST_PROCESS_MODEL and POST_PROCESS_MODEL.strip():
        print(f"INFO: Attempting to refine Markdown for {pdf_path.name} using {POST_PROCESS_MODEL}...")
        try:
            refined_md_text = call_llm_for_refinement(raw_md_text, POST_PROCESS_MODEL)
            print(f"INFO: Markdown refinement successful for {pdf_path.name}.")
        except Exception as e:
            print(f"⚠️ WARNING: Refinement failed for {pdf_path.name}: {e}. Using raw Markdown instead.")
    else:
        print(f"INFO: Skipping refinement for {pdf_path.name} as POST_PROCESS_MODEL is not set.")

    print(f"DEBUG: Initial MD Transcription for {pdf_path.name} (first 500 chars):\n{raw_md_text[:500]}")

    # Stage 3: Replace image placeholders with actual image links
    final_md_text = refined_md_text
    placeholder_count = final_md_text.count("[图片]")
    if extracted_image_paths and placeholder_count > 0:
        print(f"INFO: Replacing {placeholder_count} image placeholders in {pdf_path.name}...")
        current_img_index = 0
        temp_md_parts = []
        last_pos = 0
        for _ in range(placeholder_count):
            if current_img_index >= len(extracted_image_paths):
                break 
            placeholder_pos = final_md_text.find("[图片]", last_pos)
            if placeholder_pos == -1:
                break 
            temp_md_parts.append(final_md_text[last_pos:placeholder_pos])
            temp_md_parts.append(f"![]({extracted_image_paths[current_img_index]})")
            last_pos = placeholder_pos + len("[图片]")
            current_img_index += 1
        temp_md_parts.append(final_md_text[last_pos:])
        final_md_text = "".join(temp_md_parts)
        if current_img_index < placeholder_count:
             print(f"INFO: Replaced {current_img_index} of {placeholder_count} image placeholders for {pdf_path.name} (extracted: {len(extracted_image_paths)}).")
        # elif placeholder_count < len(extracted_image_paths):
            # print(f"INFO: Found {placeholder_count} image placeholders and {len(extracted_image_paths)} extracted images for {pdf_path.name}. Some extracted images may not be referenced.")

    # Use the same sanitized_stem for the markdown filename
    md_file_path = out_dir / (sanitized_stem + ".md") 
    try:
        md_file_path.write_text(final_md_text, encoding="utf-8")
    except OSError as e:
        print(f"❌ ERROR: Could not write Markdown file {md_file_path.resolve()}: {e}")
        return None # Indicate failure to write MD file
    return md_file_path

def process_pdf_page_to_image_and_text(pdf_path_str: str, page_no_0_indexed: int, dpi: int) -> tuple[int, str]:
    """Converts a single PDF page to a base64 WEBP image and then to Markdown text via API."""
    try:
        # Convert PDF page to a single image
        images = convert_from_path(
            pdf_path_str, 
            dpi=dpi, 
            first_page=page_no_0_indexed + 1, 
            last_page=page_no_0_indexed + 1,
            fmt='webp', # Use webp directly if supported and preferred
            thread_count=1 # Process one page at a time
        )
        if not images:
            return page_no_0_indexed, f"> **Error: Could not convert page {page_no_0_indexed + 1} to image.**"
        
        img = images[0]
        
        # Convert image to base64
        buffered = io.BytesIO()
        img.save(buffered, format="WEBP") # Save as WEBP
        base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
        
        # Call the Vision API
        page_md_content = call_vlm(base64_image, page_no_0_indexed + 1)
        return page_no_0_indexed, page_md_content
    
    except tenacity.RetryError as e:
        last_exception = e.last_attempt.exception()
        error_details = f"RetryError on page {page_no_0_indexed + 1}"
        if isinstance(last_exception, requests.exceptions.HTTPError):
            status_code = last_exception.response.status_code
            response_text = last_exception.response.text
            error_details += f": Last HTTPError {status_code} - {response_text[:500]}" # Truncate response text
        elif last_exception:
            error_details += f": Last error - {type(last_exception).__name__}: {str(last_exception)[:500]}"
        else:
            error_details += ": No specific last error information available in RetryError."
        print(f"❌ DETAILED ERROR for page {page_no_0_indexed + 1}: {error_details}")
        return page_no_0_indexed, f"> **Error processing page {page_no_0_indexed + 1}: {error_details}**"
    
    except Exception as e:
        # Catch any other unexpected errors during page processing
        error_details = f"Unexpected error on page {page_no_0_indexed + 1}: {type(e).__name__} - {str(e)[:500]}"
        print(f"❌ DETAILED ERROR for page {page_no_0_indexed + 1}: {error_details}")
        return page_no_0_indexed, f"> **Error processing page {page_no_0_indexed + 1}: {error_details}**"


@tenacity.retry(
    wait=tenacity.wait_fixed(2), 
    stop=tenacity.stop_after_attempt(5),
    retry=tenacity.retry_if_exception_type((requests.exceptions.RequestException,))
)
def call_vlm(base64_img: str, page_no_1_indexed: int) -> str:
    """
    给定一个 base64 Image 调用 Qwen‑VL 并返回 Markdown
    """
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {
                        "url": f"data:image/webp;base64,{base64_img}",
                        "detail": "high"
                    }},
                    {"type": "text", "text": USER_INSTRUCTION}
                ]
            }
        ],
        "stream": False,
        "temperature": 0.2,
        "max_tokens": 4096,
        "enable_thinking": False
    }
    resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=120)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

##############################
# 3 CLI
##############################
def parse_args():
    parser = argparse.ArgumentParser(
        description="Batch convert PDF (Kaoyan Math I) to Obsidian Markdown via Qwen‑VL")
    parser.add_argument("target", help="pdf file or directory containing pdfs")
    parser.add_argument("-o", "--output", default="obsidian_md",
                        help="output directory (default: ./obsidian_md)")
    parser.add_argument("--dpi", type=int, default=300, help="pdf2image DPI")
    parser.add_argument("-w", "--workers", type=int, default=os.cpu_count()//2 or 1,
                        help="parallel pdf workers")
    return parser.parse_args()

def main():
    args = parse_args()
    target = pathlib.Path(args.target)
    pdf_files = [target] if target.is_file() else list(target.rglob("*.pdf"))
    if not pdf_files:
        print("❌ No PDF found.")
        sys.exit(1)
    with ProcessPoolExecutor(max_workers=args.workers) as pool:
        jobs = {pool.submit(process_pdf, str(p), args.output, args.dpi): p for p in pdf_files}
        for fut in tqdm.tqdm(as_completed(jobs), total=len(jobs), desc="All PDFs", unit="doc"):
            p = jobs[fut]
            try:
                md_path_result = fut.result()
                if md_path_result:
                    try:
                        print(f"✅ {p.name}  →  {md_path_result.relative_to(pathlib.Path.cwd().resolve())}")
                    except ValueError:
                        # If relative_to fails, print the absolute path as a fallback
                        print(f"✅ {p.name}  →  {md_path_result.resolve()} (Note: Path is absolute)")
                else:
                    print(f"⚠️ {p.name} processing completed, but no Markdown file path was returned (check logs for errors).")
            except Exception as e:
                # Log the full exception for debugging
                import traceback
                print(f"❌ {p.name} failed with an unexpected error during future processing: {e}\n{traceback.format_exc()}")

if __name__ == "__main__":
    main()
