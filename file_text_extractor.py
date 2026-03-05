#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档文本提取类
支持从URL提取PDF/DOC/DOCX/WPS文档文本并返回JSON格式结果
"""

import base64
import datetime
import re
import sys
import time
import shutil
import os
import json
from typing import Optional, Dict, Any

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


class DocumentTextExtractor:
    """文档文本提取类"""

    def __init__(self, api_key: str, dpi: int = 200):
        """
        初始化提取器
        :param api_key: DeepSeek OCR API密钥
        :param dpi: PDF转图片的DPI值，默认200
        """
        self.api_key = api_key
        self.dpi = dpi
        self.custom_temp_dir = os.path.join(os.getcwd(), 'temp_files')
        os.makedirs(self.custom_temp_dir, exist_ok=True)

    def _download_file_from_url(self, url: str) -> Optional[str]:
        """
        内部方法：从URL下载文件
        :param url: 文档URL地址
        :return: 临时文件路径或None
        """
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

            temp_file = tempfile.NamedTemporaryFile(
                suffix=f'.{suffix}',
                delete=False,
                dir=self.custom_temp_dir
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

    def _extract_text_from_docx(self, file_path: str) -> str:
        """
        内部方法：从docx文件提取文本
        :param file_path: 文件路径
        :return: 提取的文本
        """
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

    def _extract_text_from_doc_wps(self, file_path: str) -> str:
        """
        内部方法：Windows下从doc/wps文件提取文本
        :param file_path: 文件路径
        :return: 提取的文本
        """
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

    def _extract_text_from_document(self, file_path: str) -> str:
        """
        内部方法：通用文档文本提取入口
        :param file_path: 文件路径
        :return: 提取的文本
        """
        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext == '.docx':
            return self._extract_text_from_docx(file_path)
        elif file_ext in ['.doc', '.wps']:
            return self._extract_text_from_doc_wps(file_path)
        else:
            print(f"{Fore.YELLOW}⚠️ 不支持的文件格式: {file_ext}{Style.RESET_ALL}")
            return ""

    def _pdf_to_images_with_fitz(self, pdf_path: str) -> list:
        """
        内部方法：PDF转图片
        :param pdf_path: PDF文件路径
        :return: 图片路径列表
        """
        image_paths = []

        try:
            doc = fitz.open(pdf_path)
            print(f"{Fore.CYAN}📄 PDF 总页数: {len(doc)} 页{Style.RESET_ALL}")

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                mat = fitz.Matrix(self.dpi / 72, self.dpi / 72)
                pix = page.get_pixmap(matrix=mat, alpha=False)

                with tempfile.NamedTemporaryFile(suffix='.png', delete=False, dir=self.custom_temp_dir) as temp_file:
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

    def _deepseek_ocr_image(self, image_path: str) -> Optional[list]:
        """
        内部方法：调用OCR API识别图片文本
        :param image_path: 图片路径
        :return: 识别的文本列表
        """
        url = "https://api.siliconflow.cn/v1/chat/completions"
        model = "deepseek-ai/DeepSeek-OCR"

        if not self.api_key:
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
            "Authorization": f"Bearer {self.api_key}",
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

    def _clean_temp_files(self, file_path: str = None):
        """
        内部方法：清理临时文件
        :param file_path: 要删除的文件路径（可选）
        """
        # 删除指定文件
        if file_path and os.path.exists(file_path):
            try:
                os.unlink(file_path)
                print(f"{Fore.CYAN}🗑️  已清理临时文件: {file_path}{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.YELLOW}⚠️ 清理临时文件失败: {str(e)}{Style.RESET_ALL}")

        # 清理临时图片文件
        for filename in os.listdir(self.custom_temp_dir):
            if filename.endswith('.png'):
                file_path = os.path.join(self.custom_temp_dir, filename)
                try:
                    os.unlink(file_path)
                except:
                    pass

    def extract(self, url: str) -> Dict[str, Any]:
        """
        核心方法：从指定URL提取文档文本
        :param url: 文档URL地址
        :return: JSON格式的结果字典，包含url和full_text字段
        """
        result = {
            "url": url,
            "full_text": "",
            "status": "failed",
            "message": ""
        }

        print(f"{Fore.CYAN}📂 开始处理文档: {url}{Style.RESET_ALL}")

        # 下载文件
        temp_file_path = self._download_file_from_url(url)
        if not temp_file_path:
            result["message"] = "文件下载失败"
            return result

        print(f"{Fore.GREEN}✅ 文件下载完成: {temp_file_path}{Style.RESET_ALL}")

        file_ext = os.path.splitext(temp_file_path)[1].lower()
        full_text = ""

        try:
            # 处理PDF文件
            if file_ext in ['.pdf']:
                image_paths = self._pdf_to_images_with_fitz(temp_file_path)
                if not image_paths:
                    result["message"] = "PDF转换图片失败"
                    return result

                for idx, image_path in enumerate(image_paths):
                    page_num = idx + 1
                    print(f"{Fore.CYAN}--- 正在识别第 {page_num}/{len(image_paths)} 页 ---{Style.RESET_ALL}")

                    texts = self._deepseek_ocr_image(image_path)
                    if texts:
                        full_text += "\n".join(texts) + "\n"
                        print(f"{Fore.GREEN}✅ 第 {page_num} 页识别成功{Style.RESET_ALL}")
                    else:
                        print(f"{Fore.RED}❌ 第 {page_num} 页识别失败{Style.RESET_ALL}")

                    try:
                        os.unlink(image_path)
                    except:
                        pass

            # 处理Office文档
            elif file_ext in ['.doc', '.wps', '.docx']:
                print(f"{Fore.CYAN}🔄 正在提取{file_ext}文件文本...{Style.RESET_ALL}")
                full_text = self._extract_text_from_document(temp_file_path)
                if not full_text:
                    result["message"] = "Office文档文本提取失败"
                    return result

            # 不支持的格式
            else:
                result["message"] = f"不支持的文件类型: {file_ext}"
                return result

            # 处理提取结果
            full_text = full_text.strip()
            if full_text:
                result["full_text"] = full_text
                result["status"] = "success"
                result["message"] = "文本提取成功"
                print(f"{Fore.GREEN}✅ 文本提取成功，共 {len(full_text)} 个字符{Style.RESET_ALL}")
            else:
                result["message"] = "提取的文本为空"

        except Exception as e:
            result["message"] = f"处理文件时出错: {str(e)}"
            print(f"{Fore.RED}❌ 处理文件失败: {str(e)}{Style.RESET_ALL}")

        finally:
            # 清理临时文件
            self._clean_temp_files(temp_file_path)

        return result

    def save_to_file(self, result: Dict[str, Any], output_filename: str = None) -> Optional[str]:
        """
        将提取结果保存为JSON文件（或TXT文件）
        :param result: extract方法返回的结果字典
        :param output_filename: 输出文件名（可选）
        :return: 保存的文件路径
        """
        if result["status"] != "success" or not result["full_text"]:
            print(f"{Fore.RED}❌ 无有效文本可保存{Style.RESET_ALL}")
            return None

        # 生成默认文件名
        if not output_filename:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"extracted_result_{timestamp}.json"

        output_dir = "extracted_texts"
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_filename)

        try:
            # 如果是json后缀，保存为JSON格式；否则保存为TXT
            if output_filename.lower().endswith('.json'):
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=4)
            else:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(result["full_text"])

            print(f"{Fore.GREEN}✅ 结果已保存到: {output_path}{Style.RESET_ALL}")
            return output_path
        except Exception as e:
            print(f"{Fore.RED}❌ 保存文件失败: {str(e)}{Style.RESET_ALL}")
            return None

    def clean_temp_dir(self):
        """清理所有临时文件"""
        try:
            for filename in os.listdir(self.custom_temp_dir):
                file_path = os.path.join(self.custom_temp_dir, filename)
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            print(f"{Fore.GREEN}✅ 已清理自定义临时目录{Style.RESET_ALL}")
            try:
                os.rmdir(self.custom_temp_dir)
                print(f"{Fore.GREEN}✅ 已删除临时目录根目录: {self.custom_temp_dir}{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.YELLOW}⚠️ 删除根目录失败（可能非空）: {str(e)}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.YELLOW}⚠️ 清理临时目录失败: {str(e)}{Style.RESET_ALL}")
