# 文档文本提取工具 (GitHub 插件版)
一个轻量级CLI工具，只需输入文档URL和API Key，即可提取PDF/DOC/DOCX/WPS文档的文本内容并保存为TXT文件。

## 功能特点
- 🚀 极简使用：仅需文档URL和API Key即可运行
- 📄 多格式支持：PDF/DOC/DOCX/WPS
- 🖥️ Windows专属：完美支持doc/wps格式解析
- 🧹 自动清理：自动清理临时文件，避免冗余
- 📝 编码友好：输出UTF-8编码的TXT文件

## 环境要求
- Windows 系统（依赖pywin32处理doc/wps）
- Python 3.7+
- 有效的DeepSeek OCR API Key

## 安装步骤
1. 克隆仓库
```bash
git clone https://github.com/你的用户名/file-extractor.git
cd file-extractor
```
2. 安装依赖
```bash
pip install -r requirements.txt
```
## 使用方法
### 基础使用

```bash
python doc_extractor.py --url "https://example.com/your-document.pdf" --api-key "你的API密钥"
```
指定输出文件名
```bash

python doc_extractor.py -u "https://example.com/report.docx" -k "sk-xxx" -o "提取结果.txt"
```
自定义 PDF 转图片 DPI
```bash
python doc_extractor.py -u "https://example.com/file.pdf" -k "sk-xxx" --dpi 300
```
完整参数说明

参数	简写	说明	是否必填
--url	-u	文档的 URL 地址	✅ 是
--api-key	-k	DeepSeek OCR API 密钥	✅ 是
--output	-o	输出的 TXT 文件名	❌ 否（默认自动生成）
--dpi	-	PDF 转图片的 DPI 值	❌ 否（默认 200）
输出说明
提取的文本默认保存在 extracted_texts 目录下
临时文件保存在 temp_files 目录，程序结束后自动清理
输出文件编码为 UTF-8，兼容所有中文场景
常见问题
doc/wps 文件解析失败：确保安装了 Microsoft Word 或 WPS，且以管理员身份运行
API 调用失败：检查 API Key 是否有效，网络是否能访问硅基流动 API
PDF 识别乱码：提高 DPI 参数（如 --dpi 300）重试
依赖安装失败：
bash
运行
pip install --upgrade pip
pip install -r requirements.txt --user
