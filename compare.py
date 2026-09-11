# -*- coding: utf-8 -*-
"""
文件内容比较工具（控制台版）
功能：比较两个文件（txt/docx/xlsx/pptx/pdf）的文本内容，
      输出相似度百分比，并列出具体差异位置及上下文。

使用方法：
1. 修改下面的 FILE1_PATH 和 FILE2_PATH 为你的文件路径
2. 修改 MODE 选择空白字符处理模式（见下方注释）
3. 运行脚本，结果打印在控制台
"""

import os
import re
import difflib

# ========== 用户修改区 ==========

# 第一个文件路径，包含扩展名
FILE1_PATH = r"D:\PersonalTools\1.txt"

# 第二个文件路径，包含扩展名
FILE2_PATH = r"D:\PersonalTools\ds.txt"

# 空白字符处理模式（只保留换行符，忽略其他空白）
# 可选值：'strict' , 'ignore_spaces', 'ignore_all_whitespace'
MODE = 'ignore_spaces'   # 推荐：忽略空格/制表符，保留换行符

# ==============================

# ---------- 文本提取模块（带位置信息）----------
# 为了输出差异位置，我们需要记录每个字符所在的“位置标签”
# 位置标签可以是：页码、幻灯片号、工作表名、行号等。
# 本代码采用“文本块”方式：每个块有 (位置标签, 文本内容)

try:
    from docx import Document
    from openpyxl import load_workbook
    from pptx import Presentation
    import pdfplumber
    LIBRARIES_OK = True
except ImportError as e:
    print(f"警告：缺少必要的库 -> {e}")
    print("请运行：pip install python-docx openpyxl python-pptx pdfplumber")
    LIBRARIES_OK = False

def extract_text_blocks(file_path):
    """根据文件扩展名，返回 [(位置标签, 文本片段), ...] 列表"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.txt':
        return extract_blocks_from_txt(file_path)
    elif ext == '.docx':
        return extract_blocks_from_docx(file_path)
    elif ext == '.xlsx':
        return extract_blocks_from_xlsx(file_path)
    elif ext == '.pptx':
        return extract_blocks_from_pptx(file_path)
    elif ext == '.pdf':
        return extract_blocks_from_pdf(file_path)
    else:
        raise ValueError(f"不支持的文件类型: {ext}")

def extract_blocks_from_txt(file_path):
    """文本文件：每行作为一个块，位置为行号"""
    blocks = []
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    for i, line in enumerate(lines, start=1):
        # 保留换行符，便于后续比较
        blocks.append((f"行 {i}", line.rstrip('\n')))
    return blocks

def extract_blocks_from_docx(file_path):
    """Word文档：每个段落作为一个块，位置为段落编号"""
    doc = Document(file_path)
    blocks = []
    for i, para in enumerate(doc.paragraphs, start=1):
        text = para.text.strip()
        if text:
            blocks.append((f"段落 {i}", text))
    # 提取表格内容（每个单元格作为一个块，位置标明表格行列）
    for table_idx, table in enumerate(doc.tables, start=1):
        for row_idx, row in enumerate(table.rows, start=1):
            for col_idx, cell in enumerate(row.cells, start=1):
                cell_text = cell.text.strip()
                if cell_text:
                    blocks.append((f"表格{table_idx} 第{row_idx}行第{col_idx}列", cell_text))
    return blocks

def extract_blocks_from_xlsx(file_path):
    """Excel：每个单元格作为一个块，位置为工作表名+行列"""
    wb = load_workbook(file_path, data_only=True)
    blocks = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            for col_idx, cell_value in enumerate(row, start=1):
                if cell_value is not None:
                    text = str(cell_value).strip()
                    if text:
                        blocks.append((f"工作表'{sheet_name}' 第{row_idx}行第{col_idx}列", text))
    return blocks

def extract_blocks_from_pptx(file_path):
    """PPT：每张幻灯片的每个文本框/表格单元格为一个块，位置为幻灯片编号"""
    prs = Presentation(file_path)
    blocks = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                blocks.append((f"幻灯片 {slide_idx}", shape.text.strip()))
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                for row_idx, row in enumerate(table.rows, start=1):
                    for col_idx, cell in enumerate(row.cells, start=1):
                        cell_text = cell.text.strip()
                        if cell_text:
                            blocks.append((f"幻灯片 {slide_idx} 表格第{row_idx}行第{col_idx}列", cell_text))
    return blocks

def extract_blocks_from_pdf(file_path):
    """PDF：按页提取文本，每页的每行作为一个块（按换行分割）"""
    blocks = []
    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            for line_idx, line in enumerate(lines, start=1):
                line = line.strip()
                if line:
                    blocks.append((f"第{page_num}页 行{line_idx}", line))
    return blocks

# ---------- 空白字符规范化函数 ----------
def normalize_text(text, mode):
    if mode == 'strict':
        return text
    elif mode == 'ignore_spaces':
        # 保留换行符，删除所有空格和制表符
        return re.sub(r'[ \t]+', '', text)
    elif mode == 'ignore_all_whitespace':
        # 删除所有空白（包括换行）
        return re.sub(r'\s+', '', text)
    else:
        raise ValueError("未知模式，可选 strict / ignore_spaces / ignore_all_whitespace")

# ---------- 差异输出（基于文本块比较） ----------
def compare_blocks(blocks1, blocks2, mode):
    """
    比较两个块列表，输出差异。
    blocks1, blocks2: [(位置, 文本), ...]
    mode: 空白处理模式
    """
    # 将每个块的文本规范化后拼接，同时保留原始文本和位置
    norm_text1 = "\n".join(normalize_text(block[1], mode) for block in blocks1)
    norm_text2 = "\n".join(normalize_text(block[1], mode) for block in blocks2)
    
    # 整体相似度（字符级）
    matcher = difflib.SequenceMatcher(None, norm_text1, norm_text2)
    ratio = matcher.ratio()
    total_chars = len(norm_text1) + len(norm_text2)
    matched_chars = int(ratio * total_chars / 2) if total_chars > 0 else 0
    print("=" * 70)
    print(f"文件1: {FILE1_PATH}")
    print(f"文件2: {FILE2_PATH}")
    print(f"空白处理模式: {mode}")
    print(f"相似度: {ratio*100:.4f}%  (匹配字符 {matched_chars} / 总字符 {total_chars})")
    print("-" * 70)
    
    if ratio == 1.0:
        print("✓ 两个文件在所选模式下完全一致")
        return
    
    # 找出差异：基于行的比较更易读（因为块内文本较长，按块比较可能漏掉内部差异）
    # 这里改用“按行”比较，同时保留每行的位置来源
    lines1 = []
    lines1_pos = []
    for pos, text in blocks1:
        # 将每个块内的文本按换行分割（如果块内本身有换行，例如PDF的一页内有多个行）
        sublines = text.split('\n')
        for sub in sublines:
            if sub.strip():  # 忽略空行（但为了位置准确，保留空行？根据需求，忽略空行）
                lines1.append(normalize_text(sub, mode))
                lines1_pos.append(pos)
            else:
                # 保留空行（仅占位，不参与比较？difflib比较时空行会影响，这里选择保留）
                lines1.append("")
                lines1_pos.append(pos)
    
    lines2 = []
    lines2_pos = []
    for pos, text in blocks2:
        sublines = text.split('\n')
        for sub in sublines:
            if sub.strip():
                lines2.append(normalize_text(sub, mode))
                lines2_pos.append(pos)
            else:
                lines2.append("")
                lines2_pos.append(pos)
    
    # 使用 difflib 比较行列表
    differ = difflib.SequenceMatcher(None, lines1, lines2)
    diff_count = 0
    for tag, i1, i2, j1, j2 in differ.get_opcodes():
        if tag == 'replace':
            print(f"\n[差异] 替换:")
            print(f"  位置1: {lines1_pos[i1]} (行 {i1+1}) 内容: {lines1[i1][:60]}...")
            print(f"  位置2: {lines2_pos[j1]} (行 {j1+1}) 内容: {lines2[j1][:60]}...")
            diff_count += 1
        elif tag == 'delete':
            print(f"\n[差异] 删除:")
            for idx in range(i1, i2):
                print(f"  位置1: {lines1_pos[idx]} (行 {idx+1}) 内容: {lines1[idx][:60]}")
            diff_count += (i2 - i1)
        elif tag == 'insert':
            print(f"\n[差异] 插入:")
            for idx in range(j1, j2):
                print(f"  位置2: {lines2_pos[idx]} (行 {idx+1}) 内容: {lines2[idx][:60]}")
            diff_count += (j2 - j1)
    
    if diff_count == 0:
        print("\n未检测到行级差异（可能差异在行内字符）")
        # 进一步输出字符级差异样例（取第一个差异块）
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != 'equal':
                context1 = norm_text1[max(0, i1-20):i2+20]
                context2 = norm_text2[max(0, j1-20):j2+20]
                print(f"\n字符级差异示例（tag={tag}）:")
                print(f"  文件1: ...{context1}...")
                print(f"  文件2: ...{context2}...")
                break

# ---------- 主程序 ----------
def main():
    if not LIBRARIES_OK:
        print("请安装缺少的库后重试。")
        return
    
    try:
        blocks1 = extract_text_blocks(FILE1_PATH)
        blocks2 = extract_text_blocks(FILE2_PATH)
        compare_blocks(blocks1, blocks2, MODE)
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    main()