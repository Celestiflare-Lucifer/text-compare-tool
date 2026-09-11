# -*- coding: utf-8 -*-
"""
文件内容比较工具（图形界面版）
@Author:星辰蝶语 (Celestiflare Lucifer
支持格式：.txt, .docx, .xlsx, .pptx, .pdf
功能：计算相似度，输出差异位置及上下文
字体：统一使用宋体 (SimSun)，支持动态调整大小（小于14时可能回退到系统默认字体）
"""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import difflib
import tkinter.font as tkfont

# ---------- 导入可选库 ----------
try:
    from docx import Document
except ImportError:
    Document = None
try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None
try:
    from pptx import Presentation
except ImportError:
    Presentation = None
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# ---------- 文本提取模块（带位置信息）----------
def extract_blocks_from_txt(file_path):
    blocks = []
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    for i, line in enumerate(lines, start=1):
        text = line.rstrip('\n')
        if text.strip():
            blocks.append((f"行 {i}", text))
        else:
            blocks.append((f"行 {i}", ""))
    return blocks

def extract_blocks_from_docx(file_path):
    if Document is None:
        raise ImportError("缺少 python-docx 库")
    doc = Document(file_path)
    blocks = []
    for i, para in enumerate(doc.paragraphs, start=1):
        text = para.text.strip()
        if text:
            blocks.append((f"段落 {i}", text))
        else:
            blocks.append((f"段落 {i}", ""))
    for t_idx, table in enumerate(doc.tables, start=1):
        for r_idx, row in enumerate(table.rows, start=1):
            for c_idx, cell in enumerate(row.cells, start=1):
                text = cell.text.strip()
                if text:
                    blocks.append((f"表{t_idx} 第{r_idx}行第{c_idx}列", text))
    return blocks

def extract_blocks_from_xlsx(file_path):
    if load_workbook is None:
        raise ImportError("缺少 openpyxl 库")
    wb = load_workbook(file_path, data_only=True)
    blocks = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for r_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            for c_idx, val in enumerate(row, start=1):
                if val is not None:
                    text = str(val).strip()
                    if text:
                        blocks.append((f"工作表'{sheet_name}' R{r_idx}C{c_idx}", text))
    return blocks

def extract_blocks_from_pptx(file_path):
    if Presentation is None:
        raise ImportError("缺少 python-pptx 库")
    prs = Presentation(file_path)
    blocks = []
    for s_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                blocks.append((f"幻灯片 {s_idx}", shape.text.strip()))
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows, start=1):
                    for c_idx, cell in enumerate(row.cells, start=1):
                        text = cell.text.strip()
                        if text:
                            blocks.append((f"幻灯片{s_idx} 表格R{r_idx}C{c_idx}", text))
    return blocks

def extract_blocks_from_pdf(file_path):
    if pdfplumber is None:
        raise ImportError("缺少 pdfplumber 库")
    blocks = []
    with pdfplumber.open(file_path) as pdf:
        for p_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for l_num, line in enumerate(lines, start=1):
                    line = line.strip()
                    if line:
                        blocks.append((f"第{p_num}页 行{l_num}", line))
    return blocks

def extract_blocks(file_path):
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

# ---------- 空白字符规范化 ----------
def normalize_text(text, mode):
    if mode == 'strict':
        return text
    elif mode == 'ignore_spaces':
        return re.sub(r'[ \t]+', '', text)
    elif mode == 'ignore_all':
        return re.sub(r'\s+', '', text)
    else:
        return text

# ---------- 比较与差异输出 ----------
def compare_and_get_diff(blocks1, blocks2, mode):
    norm1 = [normalize_text(block[1], mode) for block in blocks1]
    norm2 = [normalize_text(block[1], mode) for block in blocks2]
    full_text1 = "".join(norm1)
    full_text2 = "".join(norm2)
    matcher = difflib.SequenceMatcher(None, full_text1, full_text2)
    ratio = matcher.ratio()
    differ = difflib.SequenceMatcher(None, norm1, norm2)
    diff_lines = []
    for tag, i1, i2, j1, j2 in differ.get_opcodes():
        if tag == 'equal':
            continue
        pos1 = blocks1[i1][0] if i1 < len(blocks1) else "末尾"
        pos2 = blocks2[j1][0] if j1 < len(blocks2) else "末尾"
        if tag == 'replace':
            diff_lines.append(f"[替换] {pos1} ↔ {pos2}")
            diff_lines.append(f"  文件1: {norm1[i1][:80]}")
            diff_lines.append(f"  文件2: {norm2[j1][:80]}")
        elif tag == 'delete':
            diff_lines.append(f"[删除] {pos1}")
            diff_lines.append(f"  内容: {norm1[i1][:80]}")
        elif tag == 'insert':
            diff_lines.append(f"[插入] {pos2}")
            diff_lines.append(f"  内容: {norm2[j1][:80]}")
    if not diff_lines and ratio < 1.0:
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != 'equal':
                ctx1 = full_text1[max(0, i1-30):i2+30]
                ctx2 = full_text2[max(0, j1-30):j2+30]
                diff_lines.append("字符级差异示例（无法定位到具体行）:")
                diff_lines.append(f"  文件1: ...{ctx1}...")
                diff_lines.append(f"  文件2: ...{ctx2}...")
                break
    diff_text = "\n".join(diff_lines) if diff_lines else "未检测到明显差异（可能仅空格/格式不同）"
    return ratio * 100, diff_text

# ---------- 图形界面 ----------
class CompareApp:
    def __init__(self, root):
        self.root = root
        root.title("文件内容比较工具")
        root.geometry("800x600")
        root.resizable(True, True)

        # 自定义字体：宋体，初始大小14（小四）
        self.custom_font = tkfont.Font(family="SimSun", size=14)
        self.current_font_size = 14  # 记录当前字号

        # 创建菜单栏，直接放置顶级菜单项
        menubar = tk.Menu(root)
        menubar.add_command(label="关于", command=self.show_about)
        menubar.add_command(label="增大字体", command=self.increase_font)
        menubar.add_command(label="减小字体", command=self.decrease_font)
        root.config(menu=menubar)

        # 文件选择行
        tk.Label(root, text="文件1：", font=self.custom_font, anchor='w').grid(row=0, column=0, padx=5, pady=10, sticky='w')
        self.file1_var = tk.StringVar()
        tk.Entry(root, textvariable=self.file1_var, font=self.custom_font, width=60).grid(row=0, column=1, padx=5, pady=10)
        tk.Button(root, text="浏览", font=self.custom_font, command=self.browse_file1).grid(row=0, column=2, padx=5)

        tk.Label(root, text="文件2：", font=self.custom_font, anchor='w').grid(row=1, column=0, padx=5, pady=10, sticky='w')
        self.file2_var = tk.StringVar()
        tk.Entry(root, textvariable=self.file2_var, font=self.custom_font, width=60).grid(row=1, column=1, padx=5, pady=10)
        tk.Button(root, text="浏览", font=self.custom_font, command=self.browse_file2).grid(row=1, column=2, padx=5)

        # 模式选择
        tk.Label(root, text="比较模式：", font=self.custom_font, anchor='w').grid(row=2, column=0, padx=5, pady=10, sticky='w')
        self.mode_var = tk.StringVar(value="忽略空格/制表符（保留换行）")
        mode_combo = ttk.Combobox(root, textvariable=self.mode_var, values=[
            "严格模式（保留所有字符）",
            "忽略空格/制表符（保留换行）",
            "忽略所有空白（包括换行）"
        ], width=40, state="readonly", font=self.custom_font)
        mode_combo.grid(row=2, column=1, columnspan=2, padx=5, pady=10, sticky='w')

        # 比较按钮
        self.compare_btn = tk.Button(root, text="开始比较", font=self.custom_font, bg="#4CAF50", fg="white", command=self.start_compare)
        self.compare_btn.grid(row=3, column=0, columnspan=3, pady=15)

        # 结果显示区域
        self.result_text = tk.Text(root, wrap=tk.WORD, font=self.custom_font, height=20)
        scrollbar = tk.Scrollbar(root, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=4, column=3, sticky='ns')
        root.grid_rowconfigure(4, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪 | 提示：字体小于14时可能无法显示宋体")
        status_label = tk.Label(root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, font=self.custom_font)
        status_label.grid(row=5, column=0, columnspan=3, sticky='ew')

    def browse_file1(self):
        path = filedialog.askopenfilename(filetypes=[("支持的文件", "*.txt *.docx *.xlsx *.pptx *.pdf"), ("所有文件", "*.*")])
        if path:
            self.file1_var.set(path)

    def browse_file2(self):
        path = filedialog.askopenfilename(filetypes=[("支持的文件", "*.txt *.docx *.xlsx *.pptx *.pdf"), ("所有文件", "*.*")])
        if path:
            self.file2_var.set(path)

    def show_about(self):
        about_text = (
            "文件内容比较工具\n\n"
            "版本：1.0\n\n"
            "作者：星辰蝶语 (Celestiflare Lucifer)\n"
            "本工具基于 DeepSeek + Python + Tkinter 制作，\n"
            "使用 PyInstaller 打包为独立应用。\n"
            "支持比较 .txt / .docx / .xlsx / .pptx / .pdf\n"
            "三种空白处理模式，输出差异位置。\n\n"
            "感谢使用！"
        )
        messagebox.showinfo("关于", about_text)

    def increase_font(self):
        """增大字体大小（每次+1）"""
        self.current_font_size += 1
        self.custom_font.config(size=self.current_font_size)
        self.status_var.set(f"字体大小：{self.current_font_size}")

    def decrease_font(self):
        """减小字体大小（每次-1，最小8）"""
        if self.current_font_size > 8:
            self.current_font_size -= 1
            self.custom_font.config(size=self.current_font_size)
            if self.current_font_size < 14:
                self.status_var.set(f"字体大小：{self.current_font_size}（小于14，可能无法显示宋体）")
            else:
                self.status_var.set(f"字体大小：{self.current_font_size}")
        else:
            self.status_var.set("字体已是最小（8）")

    def start_compare(self):
        file1 = self.file1_var.get().strip()
        file2 = self.file2_var.get().strip()
        if not file1 or not file2:
            messagebox.showerror("错误", "请选择两个文件")
            return
        mode_name = self.mode_var.get()
        if "严格" in mode_name:
            mode = "strict"
        elif "忽略空格/制表符" in mode_name:
            mode = "ignore_spaces"
        else:
            mode = "ignore_all"

        self.compare_btn.config(state=tk.DISABLED, text="比较中...")
        self.result_text.delete(1.0, tk.END)
        self.status_var.set("正在提取文本...")
        thread = threading.Thread(target=self._compare_worker, args=(file1, file2, mode))
        thread.daemon = True
        thread.start()

    def _compare_worker(self, file1, file2, mode):
        try:
            blocks1 = extract_blocks(file1)
            blocks2 = extract_blocks(file2)
            self.root.after(0, lambda: self.status_var.set("正在比较..."))
            percent, diff_text = compare_and_get_diff(blocks1, blocks2, mode)
            result_str = f"相似度：{percent:.4f}%\n\n差异详情：\n{diff_text}"
            self.root.after(0, lambda: self._show_result(result_str))
            self.root.after(0, lambda: self.status_var.set("比较完成"))
        except Exception as e:
            self.root.after(0, lambda: self._show_error(str(e)))
        finally:
            self.root.after(0, lambda: self.compare_btn.config(state=tk.NORMAL, text="开始比较"))

    def _show_result(self, text):
        self.result_text.insert(1.0, text)

    def _show_error(self, err_msg):
        messagebox.showerror("错误", f"比较失败：{err_msg}")
        self.status_var.set("错误")

# ---------- 主程序 ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = CompareApp(root)
    root.mainloop()