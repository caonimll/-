from PyPDF2 import PdfReader

# 带水印的PDF路径
PDF_FILE = r"D:\python\pythonProject\pdf_watermark\带隐形水印.pdf"

def extract_watermark():
    reader = PdfReader(PDF_FILE)
    metadata = reader.metadata

    if hasattr(metadata, '/HiddenWatermark'):
        print("🔍 成功提取水印：")
        print("==================================")
        print(metadata['/HiddenWatermark'])
        print("==================================")
    else:
        print("🔍 未找到水印，但我帮你显示所有元数据：")
        for key in metadata:
            print(f"{key}: {metadata[key]}")

if __name__ == "__main__":
    extract_watermark()