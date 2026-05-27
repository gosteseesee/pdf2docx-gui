#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Word 转换器 - 傻瓜版
拖入 PDF 文件，自动转换为 Word (DOCX)
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    USE_DND = True
except ImportError:
    USE_DND = False

# 检查依赖
try:
    from pdf2docx import Converter
except ImportError:
    print("错误：未安装 pdf2docx")
    print("请运行：pip install pdf2docx")
    sys.exit(1)


class PDFToWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 转 Word 工具 v1.0")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.setup_ui()
        
    def setup_ui(self):
        # 标题
        title_frame = tk.Frame(self.root, bg="#4A90E2", height=60)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📄 PDF → Word 转换器",
            font=("Microsoft YaHei", 16, "bold"),
            bg="#4A90E2",
            fg="white"
        )
        title_label.pack(pady=15)
        
        # 主区域
        main_frame = tk.Frame(self.root, bg="#F5F5F5")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 拖拽区域
        self.drop_frame = tk.Frame(
            main_frame,
            bg="white",
            relief="groove",
            borderwidth=3,
            height=150
        )
        self.drop_frame.pack(fill="x", pady=(0, 20))
        self.drop_frame.pack_propagate(False)
        
        # 拖拽提示文字
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📂 拖拽 PDF 文件到这里\n或点击选择文件",
            font=("Microsoft YaHei", 12),
            bg="white",
            fg="#666666",
            cursor="hand2"
        )
        self.drop_label.pack(expand=True)
        
        # 绑定拖拽事件
        if USE_DND:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        
        # 绑定点击事件
        self.drop_label.bind("<Button-1>", lambda e: self.select_file())
        self.drop_frame.bind("<Button-1>", lambda e: self.select_file())
        
        # 文件信息
        self.file_info_frame = tk.Frame(main_frame, bg="#F5F5F5")
        self.file_info_frame.pack(fill="x", pady=(0, 20))
        
        self.file_label = tk.Label(
            self.file_info_frame,
            text="未选择文件",
            font=("Microsoft YaHei", 9),
            bg="#F5F5F5",
            fg="#999999",
            anchor="w"
        )
        self.file_label.pack(fill="x")
        
        # 输出目录选择
        output_frame = tk.Frame(main_frame, bg="#F5F5F5")
        output_frame.pack(fill="x", pady=(0, 20))
        
        output_label = tk.Label(
            output_frame,
            text="输出目录：",
            font=("Microsoft YaHei", 9),
            bg="#F5F5F5",
            anchor="w"
        )
        output_label.pack(fill="x", pady=(0, 5))
        
        output_select_frame = tk.Frame(output_frame, bg="#F5F5F5")
        output_select_frame.pack(fill="x")
        
        self.output_path_var = tk.StringVar(value=os.getcwd())
        self.output_entry = tk.Entry(
            output_select_frame,
            textvariable=self.output_path_var,
            font=("Microsoft YaHei", 9),
            state="readonly"
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        output_btn = tk.Button(
            output_select_frame,
            text="选择目录",
            font=("Microsoft YaHei", 9),
            bg="#4A90E2",
            fg="white",
            relief="flat",
            padx=15,
            command=self.select_output_dir
        )
        output_btn.pack(side="right")
        
        # 转换按钮
        self.convert_btn = tk.Button(
            main_frame,
            text="开始转换",
            font=("Microsoft YaHei", 11, "bold"),
            bg="#34C759",
            fg="white",
            relief="flat",
            height=2,
            state="disabled",
            command=self.start_convert
        )
        self.convert_btn.pack(fill="x", pady=(0, 20))
        
        # 进度条
        self.progress = ttk.Progressbar(
            main_frame,
            mode="indeterminate"
        )
        self.progress.pack(fill="x", pady=(0, 10))
        
        # 状态栏
        self.status_label = tk.Label(
            main_frame,
            text="等待操作...",
            font=("Microsoft YaHei", 8),
            bg="#F5F5F5",
            fg="#999999",
            anchor="w"
        )
        self.status_label.pack(fill="x")
        
        # 当前选中的文件
        self.selected_file = None
        
    def select_file(self):
        """选择 PDF 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.set_file(file_path)
            
    def on_drop(self, event):
        """拖拽文件回调"""
        files = event.data
        if files:
            # 处理拖拽路径（可能带花括号）
            file_path = files.strip()
            if file_path.startswith("{") and file_path.endswith("}"):
                file_path = file_path[1:-1]
            
            if file_path.lower().endswith(".pdf"):
                self.set_file(file_path)
            else:
                messagebox.showwarning("警告", "请选择 PDF 文件！")
                
    def set_file(self, file_path):
        """设置选中的文件"""
        self.selected_file = file_path
        filename = os.path.basename(file_path)
        filesize = os.path.getsize(file_path) / 1024 / 1024  # MB
        
        self.file_label.config(
            text=f"已选择：{filename} ({filesize:.2f} MB)",
            fg="#333333"
        )
        self.drop_label.config(
            text=f"✅ {filename}",
            fg="#34C759"
        )
        self.convert_btn.config(state="normal")
        self.status_label.config(text="就绪，点击'开始转换'")
        
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_path_var.set(dir_path)
            
    def start_convert(self):
        """开始转换"""
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择 PDF 文件！")
            return
            
        output_dir = self.output_path_var.get()
        if not output_dir or not os.path.exists(output_dir):
            messagebox.showwarning("警告", "请选择有效的输出目录！")
            return
            
        # 禁用按钮，开始进度条
        self.convert_btn.config(state="disabled", text="转换中...")
        self.progress.start(10)
        self.status_label.config(text="正在转换...")
        
        # 在新线程中执行转换
        thread = threading.Thread(
            target=self.convert_pdf,
            args=(self.selected_file, output_dir)
        )
        thread.daemon = True
        thread.start()
        
    def convert_pdf(self, pdf_path, output_dir):
        """执行 PDF 转 Word"""
        try:
            filename = os.path.basename(pdf_path)
            name_without_ext = os.path.splitext(filename)[0]
            output_path = os.path.join(output_dir, f"{name_without_ext}.docx")
            
            # 转换
            cv = Converter(pdf_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            
            # 成功
            self.root.after(0, self.on_convert_success, output_path)
            
        except Exception as e:
            # 失败
            self.root.after(0, self.on_convert_error, str(e))
            
    def on_convert_success(self, output_path):
        """转换成功回调"""
        self.progress.stop()
        self.convert_btn.config(state="normal", text="开始转换")
        self.status_label.config(text="✅ 转换成功！")
        
        messagebox.showinfo(
            "成功",
            f"转换完成！\n\n保存位置：\n{output_path}"
        )
        
        # 打开输出目录
        try:
            os.startfile(os.path.dirname(output_path))
        except:
            pass
            
    def on_convert_error(self, error_msg):
        """转换失败回调"""
        self.progress.stop()
        self.convert_btn.config(state="normal", text="开始转换")
        self.status_label.config(text="❌ 转换失败")
        
        messagebox.showerror(
            "错误",
            f"转换失败！\n\n错误信息：\n{error_msg}"
        )


def main():
    """主函数"""
    if USE_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("警告：未安装 tkinterdnd2，拖拽功能不可用")
        print("请运行：pip install tkinterdnd2")
    
    app = PDFToWordConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
