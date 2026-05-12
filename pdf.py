from PyPDF2 import PdfReader, PdfWriter

# ========== 路径 ==========
INPUT_PDF  = r"D:\python\pythonProject\pdf_watermark\云主机逻辑卷LVM实验.pdf"
OUT_PDF    = r"D:\python\pythonProject\pdf_watermark\带隐形水印.pdf"
WATERMARK  = "wei"
# ======================================

def add_hidden_watermark():
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()

    # 把所有页面复制过去
    for page in reader.pages:
        writer.add_page(page)

    # 加入【别人看不见、你能提取】的隐藏水印
    writer.add_metadata({
        '/HiddenWatermark': WATERMARK
    })

    with open(OUT_PDF, "wb") as f:
        writer.write(f)

    print("✅ 成功添加不可见水印！输出文件：", OUT_PDF)

if __name__ == "__main__":
    add_hidden_watermark()