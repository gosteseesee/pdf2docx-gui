#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Word 转换器 - 简化版（确保按钮可见）
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from pdf2docx import Converter
except ImportError:
    print("错误：未安装 pdf2docx")
    sys.exit(1)


class PDFToWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 转 Word 工具")
        self.root.geometry("500x350")
        
        self.selected_file = None
        self.output_dir = os.getcwd()
        
        self.setup_ui()
        
    def setup_ui(self):
        """创建界面"""
        # 标题
        title = tk.Label(
            self.root,
            text="📄 PDF → Word 转换器",
            font=("Microsoft YaHei", 14, "bold"),
            pady=10
        )
        title.pack()
        
        # 选择文件按钮
        select_btn = tk.Button(
            self.root,
            text="1. 选择 PDF 文件",
            font=("Microsoft YaHei", 10),
            width=30,
            height=2,
            command=self.select_file
        )
        select_btn.pack(pady=10)
        
        # 文件信息标签
        self.file_label = tk.Label(
            self.root,
            text="未选择文件",
            font=("Microsoft YaHei", 9),
            fg="gray"
        )
        self.file_label.pack(pady=5)
        
        # 选择输出目录按钮
        output_btn = tk.Button(
            self.root,
            text="2. 选择输出目录（可选）",
            font=("Microsoft YaHei", 10),
            width=30,
            command=self.select_output
        )
        output_btn.pack(pady=10)
        
        # 输出目录标签
        self.output_label = tk.Label(
            self.root,
            text=f"输出到: {self.output_dir}",
            font=("Microsoft YaHei", 8),
            fg="gray"
        )
        self.output_label.pack(pady=5)
        
        # 转换按钮（关键！）
        self.convert_btn = tk.Button(
            self.root,
            text="3. 开始转换",
            font=("Microsoft YaHei", 12, "bold"),
            width=30,
            height=2,
            bg="#4CAF50",
            fg="white",
            state="disabled",
            command=self.start_convert
        )
        self.convert_btn.pack(pady=20)
        
        # 进度条
        self.progress = ttk.Progressbar(
            self.root,
            mode="indeterminate"
        )
        self.progress.pack(fill="x", padx=20, pady=10)
        
        # 状态标签
        self.status_label = tk.Label(
            self.root,
            text="请先选择 PDF 文件",
            font=("Microsoft YaHei", 9),
            fg="gray"
        )
        self.status_label.pack()
        
    def select_file(self):
        """选择 PDF 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(
                text=f"已选择: {filename}",
                fg="blue"
            )
            self.convert_btn.config(state="normal")
            self.status_label.config(text="就绪，点击'开始转换'")
            
    def select_output(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir = dir_path
            self.output_label.config(text=f"输出到: {dir_path}")
            
    def start_convert(self):
        """开始转换"""
        if not self.selected_file:
            return
            
        # 禁用按钮，开始进度条
        self.convert_btn.config(state="disabled", text="转换中...")
        self.progress.start(10)
        self.status_label.config(text="正在转换...")
        
        # 在新线程中转换
        thread = threading.Thread(
            target=self.convert_pdf,
            args=(self.selected_file, self.output_dir)
        )
        thread.daemon = True
        thread.start()
        
    def convert_pdf(self, pdf_path, output_dir):
        """执行转换"""
        try:
            filename = os.path.basename(pdf_path)
            name = os.path.splitext(filename)[0]
            output_path = os.path.join(output_dir, f"{name}.docx")
            
            # 转换
            cv = Converter(pdf_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            
            # 成功
            self.root.after(0, self.on_success, output_path)
            
        except Exception as e:
            self.root.after(0, self.on_error, str(e))
            
    def on_success(self, output_path):
        """转换成功"""
        self.progress.stop()
        self.convert_btn.config(state="normal", text="开始转换")
        self.status_label.config(text="✅ 转换成功！")
        messagebox.showinfo("成功", f"转换完成！\n\n保存位置: {output_path}")
        
        # 打开输出目录
        try:
            os.startfile(os.path.dirname(output_path))
        except:
            pass
            
    def on_error(self, error_msg):
        """转换失败"""
        self.progress.stop()
        self.convert_btn.config(state="normal", text="开始转换")
        self.status_label.config(text="❌ 转换失败")
        messagebox.showerror("错误", f"转换失败！\n\n{error_msg}")


def main():
    root = tk.Tk()
    app = PDFToWordConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
