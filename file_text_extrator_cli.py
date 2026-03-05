# 示例使用
import datetime
import json
import sys

from colorama import Fore, Style

from file_text_extractor import DocumentTextExtractor


def main():
    """示例：使用DocumentTextExtractor类提取文档文本"""
    import argparse

    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description='文档文本提取工具 - 支持PDF/DOC/DOCX/WPS格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python file_text_extractor.py --url "https://example.com/file.pdf" --api-key "sk-xxx"
  python file_text_extractor.py --url "https://example.com/report.docx" --api-key "sk-xxx" --output "结果.json"
        """
    )

    # 核心参数
    parser.add_argument('--url', '-u', required=True, help='文档的URL地址（必填）')
    parser.add_argument('--api-key', '-k', required=True, help='DeepSeek OCR API密钥（必填）')
    parser.add_argument('--output', '-o', default=None, help='输出文件名（可选，默认自动生成JSON文件）')
    parser.add_argument('--dpi', type=int, default=200, help='PDF转图片的DPI（默认200）')

    # 解析参数
    args = parser.parse_args()

    # 初始化提取器
    extractor = DocumentTextExtractor(api_key=args.api_key, dpi=args.dpi)

    try:
        # 提取文本
        start_time = datetime.datetime.now()
        result = extractor.extract(args.url)

        # 打印JSON结果
        print(f"\n{Fore.CYAN}📋 提取结果 (JSON格式):{Style.RESET_ALL}")
        print(json.dumps(result, ensure_ascii=False, indent=4))

        # 保存结果
        if result["status"] == "success":
            extractor.save_to_file(result, args.output)

            # 计算耗时
            elapsed_time = (datetime.datetime.now() - start_time).total_seconds()
            print(f"\n{Fore.GREEN}{'=' * 60}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}🎉 提取完成！总耗时: {elapsed_time:.2f}秒{Style.RESET_ALL}")
            print(f"{Fore.GREEN}📝 提取文本长度: {len(result['full_text'])} 字符{Style.RESET_ALL}")
            print(f"{Fore.GREEN}{'=' * 60}{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}❌ 提取失败: {result['message']}{Style.RESET_ALL}")
            sys.exit(1)

    finally:
        # 清理临时文件
        extractor.clean_temp_dir()

if __name__ == "__main__":
    # 运行主函数（命令行模式）
    main()