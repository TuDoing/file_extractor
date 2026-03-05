from file_text_extractor import DocumentTextExtractor
# 单独调用类的示例
def simple_example():
    """简单示例：直接调用类"""
    # 1. 初始化提取器
    api_key = "token"
    extractor = DocumentTextExtractor(api_key=api_key, dpi=200)

    # 2. 提取文本
    TARGET_URL_docx = "https://www.csrc.gov.cn/guangdong/c104558/c7448028/7448028/files/%E4%B8%AD%E5%9B%BD%E8%AF%81%E5%88%B8%E7%9B%91%E7%9D%A3%E7%AE%A1%E7%90%86%E5%A7%94%E5%91%98%E4%BC%9A%E5%B9%BF%E4%B8%9C%E7%9B%91%E7%AE%A1%E5%B1%80%E8%A1%8C%E6%94%BF%E5%A4%84%E7%BD%9A%E5%86%B3%E5%AE%9A%E4%B9%A6%E3%80%942023%E3%80%9527%E5%8F%B7.docx"  # 替换为实际的文档链接
    TARGET_URL_pdf = "https://www.csrc.gov.cn/henan/c104282/c7510079/7510079/files/%E4%B8%AD%E5%9B%BD%E8%AF%81%E5%88%B8%E7%9B%91%E7%9D%A3%E7%AE%A1%E7%90%86%E5%A7%94%E5%91%98%E4%BC%9A%E6%B2%B3%E5%8D%97%E7%9B%91%E7%AE%A1%E5%B1%80%E8%A1%8C%E6%94%BF%E5%A4%84%E7%BD%9A%E5%86%B3%E5%AE%9A%E4%B9%A6(%E6%B2%B3%E5%8D%97%E5%B9%BF%E5%AE%89%E7%94%9F%E7%89%A9%E7%A7%91%E6%8A%80%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%E3%80%81%E9%AB%98%E5%A4%A9%E5%A2%9E).pdf"
    TARGET_URL_wps = "https://www.csrc.gov.cn/guangdong/c104558/c7479932/7479932/files/%E4%B8%AD%E5%9B%BD%E8%AF%81%E5%88%B8%E7%9B%91%E7%9D%A3%E7%AE%A1%E7%90%86%E5%A7%94%E5%91%98%E4%BC%9A%E5%B9%BF%E4%B8%9C%E7%9B%91%E7%AE%A1%E5%B1%80%E8%A1%8C%E6%94%BF%E5%A4%84%E7%BD%9A%E5%86%B3%E5%AE%9A%E4%B9%A6%E3%80%942024%E3%80%9524%E5%8F%B7.wps"
    TARGET_URL_doc = "https://www.csrc.gov.cn/henan/c104282/c7572545/7572545/files/%E4%B8%AD%E5%9B%BD%E8%AF%81%E5%88%B8%E7%9B%91%E7%9D%A3%E7%AE%A1%E7%90%86%E5%A7%94%E5%91%98%E4%BC%9A%E6%B2%B3%E5%8D%97%E7%9B%91%E7%AE%A1%E5%B1%80%E8%A1%8C%E6%94%BF%E5%A4%84%E7%BD%9A%E5%86%B3%E5%AE%9A%E4%B9%A6(%E7%8E%8B%E8%95%99).doc"
    TARGET_URL_doc_docx = "https://www.csrc.gov.cn/guangdong/c104548/c1351636/1351636/files/%E4%B8%AD%E5%9B%BD%E8%AF%81%E5%88%B8%E7%9B%91%E7%9D%A3%E7%AE%A1%E7%90%86%E5%A7%94%E5%91%98%E4%BC%9A%E5%B9%BF%E4%B8%9C%E7%9B%91%E7%AE%A1%E5%B1%80%E8%A1%8C%E6%94%BF%E5%A4%84%E7%BD%9A%E5%86%B3%E5%AE%9A%E4%B9%A6%EF%BC%88%E5%96%BB%E7%AD%A0%EF%BC%89%E3%80%942020%E3%80%9516%E5%8F%B7.doc.docx"
    result = extractor.extract(TARGET_URL_doc_docx)

    # 3. 获取结果
    print("URL:", result["url"])
    print("结果:", result)

    # 4. 保存结果
    # extractor.save_to_file(result, "output.json")

    # 5. 清理临时文件
    extractor.clean_temp_dir()


if __name__ == "__main__":
    # 如果想运行简单示例，取消下面的注释
    simple_example()