#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档文本提取CLI插件
支持从URL提取PDF/DOC/DOCX/WPS文档文本并保存为TXT
使用方式: python file_extractor.py --url <文档地址> --api-key <你的API密钥>
"""

import base64
import datetime
import re
import sys
import time
import shutil
import argparse
import os

from colorama import Fore, Style
import io
from PIL import Image
import fitz  # PyMuPDF
import tempfile
import requests

# 处理doc/docx/wps的核心库
try:
    from docx import Document
except ImportError:
    print(f"{Fore.YELLOW}⚠️ 未安装python-docx，请执行: pip install python-docx{Style.RESET_ALL}")
    sys.exit(1)

# Windows下必须的pywin32库（处理doc/wps）
try:
    import win32com.client
except ImportError:
    print(f"{Fore.RED}❌ 缺少pywin32库，请执行: pip install pywin32{Style.RESET_ALL}")
    sys.exit(1)

# 全局配置
CUSTOM_TEMP_DIR = os.path.join(os.getcwd(), 'temp_files')
os.makedirs(CUSTOM_TEMP_DIR, exist_ok=True)

def download_file_from_url(url, save_path=None):
    """从URL下载文件（支持PDF/DOC/DOCX/WPS）"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        url = url.replace('.doc.docx', '.doc')
        suffix = url.split('.')[-1].lower()
        valid_suffixes = ['pdf', 'doc', 'docx', 'wps']
        if suffix not in valid_suffixes:
            suffix = 'bin'

        if save_path:
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            with open(save_path, 'wb') as f:
                f.write(response.content)
            f.flush()
            os.fsync(f.fileno())
            return save_path
        else:
            temp_file = tempfile.NamedTemporaryFile(
                suffix=f'.{suffix}',
                delete=False,
                dir=CUSTOM_TEMP_DIR
            )
            with temp_file as f:
                f.write(response.content)
                f.flush()
                os.fsync(f.fileno())
            temp_file_path = temp_file.name
            
            if os.path.exists(temp_file_path) and os.path.getsize(temp_file_path) > 0:
                return temp_file_path
            else:
                print(f"{Fore.RED}❌ 临时文件创建失败或为空{Style.RESET_ALL}")
                return None

    except Exception as e:
        print(f"{Fore.RED}❌ 下载文件失败: {str(e)}{Style.RESET_ALL}")
        return None

def extract_text_from_docx(file_path):
    """从docx文件提取文本"""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        print(f"{Fore.RED}❌ 文件无效: {file_path}{Style.RESET_ALL}")
        return ""

    try:
        for retry in range(2):
            try:
                doc = Document(file_path)
                full_text = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
                text = '\n'.join(full_text)
                if text:
                    return text
                else:
                    print(f"{Fore.YELLOW}⚠️ DOCX文件无有效文本，第{retry + 1}次重试...{Style.RESET_ALL}")
                    time.sleep(1)
            except Exception as e:
                print(f"{Fore.YELLOW}⚠️ 第{retry + 1}次读取docx失败: {str(e)}{Style.RESET_ALL}")
                time.sleep(1)
                continue
        return ""
    except Exception as e:
        print(f"{Fore.RED}❌ 读取docx失败: {str(e)}{Style.RESET_ALL}")
        try:
            shutil.copy2(file_path, file_path + ".backup")
        except:
            pass
        return ""

def extract_text_from_doc_wps(file_path):
    """Windows下从doc/wps文件提取文本"""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        print(f"{Fore.RED}❌ 文件无效: {file_path}{Style.RESET_ALL}")
        return ""

    try:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except:
            word = win32com.client.Dispatch("Kwps.Application")

        word.Visible = False
        word.DisplayAlerts = 0

        file_path = os.path.abspath(file_path)
        doc = word.Documents.Open(file_path)
        full_text = doc.Content.Text.strip()

        doc.Close(SaveChanges=0)
        word.Quit()

        return full_text
    except Exception as e:
        print(f"{Fore.RED}❌ 读取doc/wps失败: {str(e)}{Style.RESET_ALL}")
        try:
            word.Quit()
        except:
            pass
        return ""

def extract_text_from_document(file_path):
    """通用文档文本提取入口"""
    file_ext = os.path.splitext(file_path)[1].lower()

    if file_ext == '.docx':
        return extract_text_from_docx(file_path)
    elif file_ext in ['.doc', '.wps']:
        return extract_text_from_doc_wps(file_path)
    else:
        print(f"{Fore.YELLOW}⚠️ 不支持的文件格式: {file_ext}{Style.RESET_ALL}")
        return ""

def pdf_to_images_with_fitz(pdf_path, dpi=200):
    """PDF转图片"""
    image_paths = []

    try:
        doc = fitz.open(pdf_path)
        print(f"{Fore.CYAN}📄 PDF 总页数: {len(doc)} 页{Style.RESET_ALL}")

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            with tempfile.NamedTemporaryFile(suffix='.png', delete=False, dir=CUSTOM_TEMP_DIR) as temp_file:
                img_path = temp_file.name

            img = Image.open(io.BytesIO(pix.tobytes("png")))
            img.save(img_path, 'PNG')
            image_paths.append(img_path)

            print(f"{Fore.CYAN}✅ 已转换第 {page_num + 1}/{len(doc)} 页{Style.RESET_ALL}")

        doc.close()
        return image_paths

    except Exception as e:
        print(f"{Fore.RED}❌ PDF转换失败: {str(e)}{Style.RESET_ALL}")
        return []

def deepseek_ocr_image(image_path, api_key):
    """调用OCR API识别图片文本"""
    url = "https://api.siliconflow.cn/v1/chat/completions"
    model = "deepseek-ai/DeepSeek-OCR"

    if not api_key:
        print(f"{Fore.RED}❌ API Key未配置{Style.RESET_ALL}")
        return None

    with open(image_path, "rb") as image_file:
        image_data = base64.b64encode(image_file.read()).decode("utf-8")

    image_url = f"data:image/png;base64,{image_data}"

    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": image_url}
                    },
                    {
                        "type": "text",
                        "text": "<image>\n<|grounding|>OCR this image."
                    }
                ]
            }
        ]
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    for i in range(5):
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()

            if "choices" in data and data["choices"]:
                content = data["choices"][0]["message"]["content"]
                texts = re.findall(r'<\|ref\|>(.*?)<\|\/ref\|>', content)
                return texts
            else:
                print(f"{Fore.YELLOW}⚠️ 未获取到识别结果，第 {i + 1} 次重试...{Style.RESET_ALL}")
                time.sleep(1)
        except requests.exceptions.Timeout:
            print(f"{Fore.YELLOW}⚠️ 请求超时，第 {i + 1} 次重试...{Style.RESET_ALL}")
            time.sleep(2)
        except Exception as e:
            print(f"{Fore.RED}❌ 请求异常: {str(e)}，第 {i + 1} 次重试...{Style.RESET_ALL}")
            time.sleep(1)

    print(f"{Fore.RED}💥 重试 5 次后仍失败，放弃请求{Style.RESET_ALL}")
    return None

def process_file_from_url(file_url, api_key, dpi=200):
    """处理指定URL的文档并提取文本"""
    print(f"{Fore.CYAN}📂 开始下载文件: {file_url}{Style.RESET_ALL}")

    temp_file_path = download_file_from_url(file_url)
    if not temp_file_path:
        print(f"{Fore.RED}❌ 文件下载失败{Style.RESET_ALL}")
        return None

    print(f"{Fore.GREEN}✅ 文件下载完成: {temp_file_path}{Style.RESET_ALL}")

    file_ext = os.path.splitext(temp_file_path)[1].lower()
    full_text = ""

    try:
        if file_ext in ['.pdf']:
            image_paths = pdf_to_images_with_fitz(temp_file_path, dpi)
            if not image_paths:
                return None

            for idx, image_path in enumerate(image_paths):
                page_num = idx + 1
                print(f"{Fore.CYAN}--- 正在识别第 {page_num}/{len(image_paths)} 页 ---{Style.RESET_ALL}")

                texts = deepseek_ocr_image(image_path, api_key)
                if texts:
                    full_text += "\n".join(texts) + "\n"
                    print(f"{Fore.GREEN}✅ 第 {page_num} 页识别成功{Style.RESET_ALL}")
                else:
                    print(f"{Fore.RED}❌ 第 {page_num} 页识别失败{Style.RESET_ALL}")

                try:
                    os.unlink(image_path)
                except:
                    pass
        elif file_ext in ['.doc', '.wps','.docx']:
            print(f"{Fore.CYAN}🔄 正在提取{file_ext}文件文本...{Style.RESET_ALL}")
            full_text = extract_text_from_document(temp_file_path)
            if full_text:
                print(f"{Fore.GREEN}✅ 文本提取成功，共 {len(full_text)} 个字符{Style.RESET_ALL}")
            else:
                print(f"{Fore.RED}❌ 文本提取失败{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}❌ 不支持的文件类型: {file_ext}{Style.RESET_ALL}")
            return None
    finally:
        try:
            if os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
                print(f"{Fore.CYAN}🗑️  已清理临时文件{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.YELLOW}⚠️ 清理临时文件失败: {str(e)}{Style.RESET_ALL}")

    return full_text.strip()

def save_text_to_file(text, output_filename=None):
    """保存文本到TXT文件"""
    if not text:
        print(f"{Fore.RED}❌ 无文本内容可保存{Style.RESET_ALL}")
        return None
    
    if not output_filename:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"extracted_text_{timestamp}.txt"
    
    output_dir = "extracted_texts"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"{Fore.GREEN}✅ 文本已保存到: {output_path}{Style.RESET_ALL}")
        return output_path
    except Exception as e:
        print(f"{Fore.RED}❌ 保存文件失败: {str(e)}{Style.RESET_ALL}")
        return None

def clean_temp_dir():
    """清理临时目录"""
    try:
        for filename in os.listdir(CUSTOM_TEMP_DIR):
            file_path = os.path.join(CUSTOM_TEMP_DIR, filename)
            if os.path.isfile(file_path):
                os.unlink(file_path)
        print(f"{Fore.GREEN}✅ 已清理自定义临时目录{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.YELLOW}⚠️ 清理临时目录失败: {str(e)}{Style.RESET_ALL}")

def main():
    """CLI主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description='文档文本提取工具 - 支持PDF/DOC/DOCX/WPS格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python doc_extractor.py --url "https://example.com/file.pdf" --api-key "sk-xxx"
  python doc_extractor.py --url "https://example.com/report.docx" --api-key "sk-xxx" --output "结果.txt"
        """
    )
    
    # 核心参数
    parser.add_argument('--url', '-u', required=True, help='文档的URL地址（必填）')
    parser.add_argument('--api-key', '-k', required=True, help='DeepSeek OCR API密钥（必填）')
    parser.add_argument('--output', '-o', default=None, help='输出的TXT文件名（可选，默认自动生成）')
    parser.add_argument('--dpi', type=int, default=200, help='PDF转图片的DPI（默认200）')
    
    # 解析参数
    args = parser.parse_args()
    
    # 验证必填参数
    if not args.url or not args.api_key:
        parser.print_help()
        sys.exit(1)
    
    # 执行提取流程
    start_time = datetime.datetime.now()
    try:
        # 提取文本
        full_text = process_file_from_url(args.url, args.api_key, args.dpi)
        
        # 保存文本
        if full_text:
            save_text_to_file(full_text, args.output)
            elapsed_time = (datetime.datetime.now() - start_time).total_seconds()
            print(f"\n{Fore.GREEN}{'=' * 60}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}🎉 提取完成！总耗时: {elapsed_time:.2f}秒{Style.RESET_ALL}")
            print(f"{Fore.GREEN}📝 提取文本长度: {len(full_text)} 字符{Style.RESET_ALL}")
            print(f"{Fore.GREEN}{'=' * 60}{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}❌ 文本提取失败{Style.RESET_ALL}")
            sys.exit(1)
    finally:
        # 清理临时文件
        clean_temp_dir()

if __name__ == "__main__":
    main()