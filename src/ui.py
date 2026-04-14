import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from src.services.excel_reader import read_excel_records
    from src.services.ppt_exporter import replace_and_export_row
except ModuleNotFoundError:
    from services.excel_reader import read_excel_records
    from services.ppt_exporter import replace_and_export_row


class PPTReplaceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT批量替换工具（Excel数据源）")
        self.root.geometry("800x850")
        self.root.resizable(False, False)

        self.df = None
        self.column_names = []
        self.selected_fields = []
        self.processing = False

        frame_path = tk.Frame(root, padx=20, pady=10)
        frame_path.pack(fill=tk.X)

        tk.Label(frame_path, text="PPT模板文件：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.ppt_path_var = tk.StringVar()
        entry_ppt = tk.Entry(frame_path, textvariable=self.ppt_path_var, width=50, font=("微软雅黑", 9))
        entry_ppt.grid(row=0, column=1, padx=10, pady=5)
        btn_ppt = tk.Button(frame_path, text="选择文件", command=self.select_ppt, width=10, font=("微软雅黑", 9))
        btn_ppt.grid(row=0, column=2, padx=5, pady=5)

        tk.Label(frame_path, text="Excel数据文件：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.excel_path_var = tk.StringVar()
        entry_excel = tk.Entry(frame_path, textvariable=self.excel_path_var, width=50, font=("微软雅黑", 9))
        entry_excel.grid(row=1, column=1, padx=10, pady=5)
        btn_excel = tk.Button(frame_path, text="选择文件", command=self.select_excel, width=10, font=("微软雅黑", 9))
        btn_excel.grid(row=1, column=2, padx=5, pady=5)

        tk.Label(frame_path, text="图片输出目录：", font=("微软雅黑", 10)).grid(row=2, column=0, sticky="w", pady=5)
        self.output_dir_var = tk.StringVar()
        entry_output = tk.Entry(frame_path, textvariable=self.output_dir_var, width=50, font=("微软雅黑", 9))
        entry_output.grid(row=2, column=1, padx=10, pady=5)
        btn_output = tk.Button(frame_path, text="选择目录", command=self.select_output_dir, width=10, font=("微软雅黑", 9))
        btn_output.grid(row=2, column=2, padx=5, pady=5)

        frame_excel = tk.Frame(root, padx=20, pady=10)
        frame_excel.pack(fill=tk.X)

        tk.Label(frame_excel, text="Excel表头行号：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.header_row_var = tk.StringVar(value="1")
        entry_header = tk.Entry(frame_excel, textvariable=self.header_row_var, width=10, font=("微软雅黑", 9))
        entry_header.grid(row=0, column=1, padx=10, pady=5)
        tk.Label(frame_excel, text="（按Excel显示的行号，如第2行输入2）", font=("微软雅黑", 9), fg="gray").grid(row=0, column=2, sticky="w", pady=5)

        btn_read_excel = tk.Button(frame_excel, text="读取Excel数据", command=self.read_excel, width=15, font=("微软雅黑", 9), bg="#4CAF50", fg="white")
        btn_read_excel.grid(row=0, column=3, padx=20, pady=5)

        frame_fields = tk.Frame(root, padx=20, pady=10)
        frame_fields.pack(fill=tk.X)

        tk.Label(frame_fields, text="选择要替换的字段：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=5)

        self.field_listbox = tk.Listbox(frame_fields, selectmode=tk.MULTIPLE, width=70, height=6, font=("微软雅黑", 9))
        self.field_listbox.grid(row=1, column=0, columnspan=2, pady=5)

        btn_confirm_fields = tk.Button(frame_fields, text="确认选择", command=self.confirm_fields, width=15, font=("微软雅黑", 9), bg="#2196F3", fg="white")
        btn_confirm_fields.grid(row=1, column=2, padx=10, pady=5)

        frame_progress = tk.Frame(root, padx=20, pady=10)
        frame_progress.pack(fill=tk.X)

        tk.Label(frame_progress, text="处理进度：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame_progress, variable=self.progress_var, maximum=100, length=500)
        self.progress_bar.grid(row=0, column=1, padx=10, pady=5)
        self.progress_label = tk.Label(frame_progress, text="0%", font=("微软雅黑", 9))
        self.progress_label.grid(row=0, column=2, padx=5, pady=5)

        frame_log = tk.Frame(root, padx=20, pady=10)
        frame_log.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame_log, text="运行日志：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.log_text = tk.Text(frame_log, width=90, height=15, font=("微软雅黑", 9), state=tk.DISABLED)
        scrollbar = tk.Scrollbar(frame_log, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=1, column=0, sticky="nsew", pady=5)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=5)

        frame_btn = tk.Frame(root, padx=20, pady=10)
        frame_btn.pack(fill=tk.X)

        self.btn_start = tk.Button(frame_btn, text="开始批量替换", command=self.start_process, width=20, font=("微软雅黑", 10), bg="#FF9800", fg="white")
        self.btn_start.grid(row=0, column=0, pady=10)

        self.log("✅ 程序已启动，请按步骤操作：")
        self.log("1. 选择PPT模板、Excel数据文件、输出目录")
        self.log("2. 设置Excel表头行号，点击【读取Excel数据】")
        self.log("3. 在字段列表中选择要替换的字段，点击【确认选择】")
        self.log("4. 点击【开始批量替换】执行操作\n")

    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def select_ppt(self):
        path = filedialog.askopenfilename(
            title="选择PPT模板文件",
            filetypes=[("PPT文件", "*.pptx;*.ppt"), ("所有文件", "*.*")],
        )
        if path:
            self.ppt_path_var.set(os.path.abspath(path))
            self.log(f"📄 已选择PPT文件：{path}")

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="选择Excel数据文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.excel_path_var.set(os.path.abspath(path))
            self.log(f"📄 已选择Excel文件：{path}")

    def select_output_dir(self):
        path = filedialog.askdirectory(title="选择图片输出目录")
        if path:
            self.output_dir_var.set(os.path.abspath(path))
            self.log(f"📂 已选择输出目录：{path}")

    def read_excel(self):
        excel_path = self.excel_path_var.get().strip()
        if not excel_path:
            messagebox.showerror("错误", "请先选择Excel文件！")
            return

        try:
            header_row = int(self.header_row_var.get().strip())
            if header_row < 1:
                messagebox.showerror("错误", "表头行号必须≥1！")
                return
        except ValueError:
            messagebox.showerror("错误", "表头行号必须是数字！")
            return

        self.log(f"📖 开始读取Excel数据（表头行：{header_row}）...")
        try:
            records, column_names = read_excel_records(excel_path, header_row)
            self.df = records
            self.column_names = column_names

            self.field_listbox.delete(0, tk.END)
            for idx, col in enumerate(column_names):
                self.field_listbox.insert(idx, col)

            self.log(f"✅ Excel读取成功！共{len(records)}行有效数据，包含字段：{', '.join(column_names)}")
            messagebox.showinfo("成功", f"Excel读取成功！\n共{len(records)}行数据，{len(column_names)}个字段")
        except Exception as error:
            self.log(f"❌ Excel读取失败：{error}")
            messagebox.showerror("错误", f"Excel读取失败：{error}")

    def confirm_fields(self):
        if not self.column_names:
            messagebox.showerror("错误", "请先读取Excel数据！")
            return

        selected_indices = self.field_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请至少选择一个替换字段！")
            return

        self.selected_fields = [self.column_names[idx] for idx in selected_indices]
        self.log(f"🎯 已确认替换字段：{', '.join(self.selected_fields)}")
        messagebox.showinfo("成功", f"已选择以下字段：\n{', '.join(self.selected_fields)}")

    def start_process(self):
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请等待！")
            return
        if not self.ppt_path_var.get():
            messagebox.showerror("错误", "请选择PPT文件！")
            return
        if not self.output_dir_var.get():
            messagebox.showerror("错误", "请选择输出目录！")
            return
        if self.df is None:
            messagebox.showerror("错误", "请先读取Excel数据！")
            return
        if not self.selected_fields:
            messagebox.showerror("错误", "请先选择替换字段！")
            return

        os.makedirs(self.output_dir_var.get(), exist_ok=True)
        self.processing = True
        self.btn_start.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        threading.Thread(target=self.process_thread, daemon=True).start()

    def process_thread(self):
        total_rows = len(self.df)
        self.log(f"🚀 开始批量处理，共{total_rows}行数据...")
        try:
            for row_idx in range(total_rows):
                self.log(f"🔄 正在处理第{row_idx + 1}/{total_rows}行数据...")
                slide_count = replace_and_export_row(
                    ppt_path=self.ppt_path_var.get(),
                    output_dir=self.output_dir_var.get(),
                    row_data=self.df[row_idx],
                    selected_fields=self.selected_fields,
                    row_idx=row_idx,
                    log_func=self.log,
                )
                self.log(f"✅ 第{row_idx + 1}行处理完成，导出{slide_count}张图片")

                progress = (row_idx + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"{progress:.1f}%")

            self.log("\n🎉 所有数据处理完成！")
            self.log(f"📂 图片已保存至：{self.output_dir_var.get()}")
            messagebox.showinfo("完成", "所有数据处理完成！\n已打开输出目录")
            os.startfile(self.output_dir_var.get())
        except Exception as error:
            self.log(f"❌ 处理失败：{error}")
            messagebox.showerror("错误", f"处理失败：{error}")
        finally:
            self.processing = False
            self.btn_start.config(state=tk.NORMAL)
            self.progress_var.set(100)
            self.progress_label.config(text="100%")
