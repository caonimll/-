# -*- coding: utf-8 -*-
"""
Word 选择题 → Excel 自动提取脚本
支持：A/B/C/D 选项、题号自动识别、空行自动跳过
"""

import docx
import pandas as pd

# ====================== 【请修改这里】 ======================
WORD_FILE = "题库.docx"  # 改成你真实的文件名！
EXCEL_FILE = "题库结果.xlsx"
# ============================================================

def extract_questions_from_word(doc_path):
    """从Word中提取题目和选项"""
    doc = docx.Document(doc_path)
    lines = []

    # 读取所有非空行
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)

    questions = []
    current_q = {"题目": "", "A": "", "B": "", "C": "", "D": ""}
    opt_types = ["A.", "A、", "B.", "B、", "C.", "C、", "D.", "D、"]

    for line in lines:
        # 判断是不是新题目（以数字+. 开头）
        if line[0].isdigit() and ". " in line[:5]:
            # 保存上一题
            if current_q["题目"]:
                questions.append(current_q.copy())
            # 重置新题目
            current_q = {"题目": line, "A": "", "B": "", "C": "", "D": ""}

        # 判断是不是选项
        elif any(line.startswith(opt) for opt in opt_types):
            if line.startswith(("A.", "A、")):
                current_q["A"] = line
            elif line.startswith(("B.", "B、")):
                current_q["B"] = line
            elif line.startswith(("C.", "C、")):
                current_q["C"] = line
            elif line.startswith(("D.", "D、")):
                current_q["D"] = line

    # 加入最后一题
    if current_q["题目"]:
        questions.append(current_q)

    return questions


def save_to_excel(data, save_path):
    """保存到Excel"""
    df = pd.DataFrame(data)
    df.to_excel(save_path, index=False)  # 这里去掉了 encoding！
    print(f"✅ 转换完成！共提取 {len(data)} 道题")
    print(f"📄 文件已保存到：{save_path}")


if __name__ == "__main__":
    try:
        q_list = extract_questions_from_word(WORD_FILE)
        save_to_excel(q_list, EXCEL_FILE)
    except Exception as e:
        print(f"❌ 出错了：{str(e)}")
        print("\n💡 常见解决办法：")
        print("1. Word文件必须是 .docx 格式（不是 .doc）")
        print("2. 把Word文件和脚本放在同一个文件夹")
        print("3. 检查文件名是否正确（包括大小写）")