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
git clone https://github.com/TuDoing/file_extractor.git
cd file-extractor
```
2. 安装依赖
```bash
pip install -r requirements.txt
```

## 命令行使用方法

基础使用
```bash
python file_text_extractor_cli.py --url "https://example.com/your-document.pdf" --api-key "你的API密钥"
```

指定输出文件名
```bash
python file_text_extractor_cli.py -u "https://example.com/report.docx" -k "sk-xxx" -o "提取结果.txt"
```

自定义 PDF 转图片 DPI
```bash
python file_text_extractor_cli.py -u "https://example.com/file.pdf" -k "sk-xxx" --dpi 300
```

## 直接调用方式参考test_use.py文件


### 完整参数说明
| 参数 | 简写 | 说明 | 是否必填 |
|------|------|------|----------|
| --url | -u | 文档的URL地址 | ✅ 是 |
| --api-key | -k | DeepSeek OCR API密钥 | ✅ 是 |
| --output | -o | 输出的TXT文件名 | ❌ 否（默认自动生成） |
| --dpi | - | PDF转图片的DPI值 | ❌ 否（默认200） |

## 输出说明
- 提取的文本默认保存在 `extracted_texts` 目录下
- 临时文件保存在 `temp_files` 目录，程序结束后自动清理
- 输出文件编码为UTF-8，兼容所有中文场景

## 常见问题
1. **doc/wps文件解析失败**：确保安装了Microsoft Word或WPS，且以管理员身份运行
2. **API调用失败**：检查API Key是否有效，网络是否能访问硅基流动API
3. **PDF识别乱码**：提高DPI参数（如--dpi 300）重试
4. **依赖安装失败**：
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt --user
   ```



## 项目结构
```
file_extractor/
├── file_text_extractor.py     # 核心文本提取类
├── file_text_extractor_cli.py # 文本提取命令行工具
├── test_use.py                # 使用示例
├── requirements.txt           # 依赖文件
└── README.md                  # 项目说明
```

## 许可证
```
MIT License

Copyright (c) 2024 

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


