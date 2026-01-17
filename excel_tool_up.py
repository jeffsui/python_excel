# excel_tool_up.py - GUI版本
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading

import pandas as pd


def merge_excels(file_paths, output_path, merge_by="rows", sheet_name=None, log_callback=None):
    """
    合并多个 Excel 文件
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    dfs = []
    for fp in file_paths:
        try:
            data = pd.read_excel(fp, sheet_name=sheet_name, engine='openpyxl')

            if isinstance(data, dict):
                sheet_names = list(data.keys())
                if sheet_names:
                    df = data[sheet_names[0]]
                    if not isinstance(df, pd.DataFrame):
                        log(f"⚠️  文件格式异常 {fp}: 第一个 sheet 不是 DataFrame")
                        continue
                else:
                    log(f"⚠️  文件为空 {fp}")
                    continue
            elif isinstance(data, pd.DataFrame):
                df = data
            else:
                log(f"⚠️  文件格式异常 {fp}: 返回类型 {type(data)}")
                continue

            df["_source_file"] = Path(fp).name
            dfs.append(df)
            log(f"✅ 已加载: {fp} ({df.shape[0]} 行 × {df.shape[1]} 列)")
        except Exception as e:
            log(f"⚠️  加载失败 {fp}: {e}")
            continue

    if not dfs:
        log("❌ 无有效数据，合并终止。")
        return False

    if merge_by == "rows":
        result = pd.concat(dfs, ignore_index=True)
    elif merge_by == "columns":
        result = pd.concat(dfs, axis=1)
    else:
        log("❌ merge_by 必须是 'rows' 或 'columns'")
        return False

    result.to_excel(output_path, index=False)
    log(f"✅ 合并完成！结果已保存至: {output_path}")
    return True


def statistics_excel(file_path, group_cols=None, agg_cols=None, output_path=None, log_callback=None):
    """
    对 Excel 做统计分析
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        log(f"📊 加载数据: {df.shape[0]} 行 × {df.shape[1]} 列")
        log(f"📋 所有列名: {', '.join(df.columns.tolist())}")

        if output_path is None:
            stem = Path(file_path).stem
            output_path = f"{stem}_stats.xlsx"

        with pd.ExcelWriter(output_path) as writer:
            desc = df.describe(include="all").T
            desc.to_excel(writer, sheet_name="summary")

            if group_cols or agg_cols:
                all_cols = df.columns.tolist()
                missing_cols = []

                if group_cols:
                    missing_cols.extend([col for col in group_cols if col not in all_cols])
                if agg_cols:
                    missing_cols.extend([col for col in agg_cols if col not in all_cols])

                if missing_cols:
                    log(f"⚠️  以下列名未找到: {', '.join(missing_cols)}")
                    log(f"   可用的列名: {', '.join(all_cols)}")
                else:
                    if group_cols and agg_cols:
                        try:
                            grouped = (
                                df.groupby(group_cols)[agg_cols]
                                .agg(["sum", "mean", "count"])
                                .reset_index()
                            )
                            if isinstance(grouped.columns, pd.MultiIndex):
                                grouped.columns = ['_'.join(col).strip() for col in grouped.columns.values]
                            grouped.to_excel(writer, sheet_name="grouped_stats", index=False)
                            log(f"✅ 分组统计已生成")
                        except Exception as e:
                            log(f"⚠️  分组统计失败: {e}")

            counts = df.count().to_frame("non_null_count")
            counts.to_excel(writer, sheet_name="counts")

        log(f"✅ 统计完成！结果已保存至: {output_path}")
        return True

    except Exception as e:
        log(f"❌ 统计失败: {e}")
        return False


class ExcelToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Excel 处理工具")
        self.root.geometry("900x700")

        self.merge_files = []
        self.stats_file = ""
        self.current_columns = []

        self.create_widgets()

    def create_widgets(self):
        # 创建 Notebooks（选项卡）
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 合并选项卡
        merge_frame = ttk.Frame(notebook)
        notebook.add(merge_frame, text="📋 合并 Excel")

        # 统计选项卡
        stats_frame = ttk.Frame(notebook)
        notebook.add(stats_frame, text="📊 统计分析")

        # 创建合并界面
        self.create_merge_ui(merge_frame)

        # 创建统计界面
        self.create_stats_ui(stats_frame)

    def create_merge_ui(self, parent):
        # 文件选择区域
        file_frame = ttk.LabelFrame(parent, text="选择 Excel 文件", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        # 添加文件按钮
        add_btn = ttk.Button(file_frame, text="📁 添加文件", command=self.add_merge_files)
        add_btn.pack(side=tk.LEFT, padx=5)

        # 清空按钮
        clear_btn = ttk.Button(file_frame, text="🗑️ 清空", command=self.clear_merge_files)
        clear_btn.pack(side=tk.LEFT, padx=5)

        # 文件列表
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.merge_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.merge_listbox.yview)

        # 合并选项
        option_frame = ttk.LabelFrame(parent, text="合并选项", padding=10)
        option_frame.pack(fill=tk.X, padx=10, pady=10)

        # 合并方式
        ttk.Label(option_frame, text="合并方式:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.merge_by_var = tk.StringVar(value="rows")
        ttk.Radiobutton(option_frame, text="纵向拼接 (rows)", variable=self.merge_by_var, value="rows").grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Radiobutton(option_frame, text="横向拼接 (columns)", variable=self.merge_by_var, value="columns").grid(row=0, column=2, sticky=tk.W, padx=5)

        # Sheet 名称
        ttk.Label(option_frame, text="Sheet 名称 (可选):").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.sheet_var = tk.StringVar()
        ttk.Entry(option_frame, textvariable=self.sheet_var, width=30).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5)

        # 输出文件
        ttk.Label(option_frame, text="输出文件:").grid(row=2, column=0, sticky=tk.W, padx=5)
        self.merge_output_var = tk.StringVar(value="merged_output.xlsx")
        ttk.Entry(option_frame, textvariable=self.merge_output_var, width=40).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Button(option_frame, text="浏览", command=self.browse_merge_output).grid(row=2, column=2, padx=5)

        # 合并按钮
        ttk.Button(parent, text="🚀 开始合并", command=self.run_merge, width=20).pack(pady=10)

    def create_stats_ui(self, parent):
        # 文件选择
        file_frame = ttk.LabelFrame(parent, text="选择 Excel 文件", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(file_frame, text="文件路径:").pack(side=tk.LEFT, padx=5)
        self.stats_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.stats_file_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="浏览", command=self.browse_stats_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="加载列名", command=self.load_columns).pack(side=tk.LEFT, padx=5)

        # 列名显示
        cols_frame = ttk.LabelFrame(parent, text="可用列名", padding=10)
        cols_frame.pack(fill=tk.X, padx=10, pady=5)

        self.cols_label = ttk.Label(cols_frame, text="请先加载 Excel 文件", wraplength=800)
        self.cols_label.pack(fill=tk.BOTH)

        # 统计选项
        option_frame = ttk.LabelFrame(parent, text="统计选项", padding=10)
        option_frame.pack(fill=tk.X, padx=10, pady=10)

        # 分组列
        ttk.Label(option_frame, text="分组列 (可选):").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.group_cols_var = tk.StringVar()
        group_entry = ttk.Entry(option_frame, textvariable=self.group_cols_var, width=50)
        group_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Label(option_frame, text="（逗号或空格分隔）").grid(row=0, column=2, sticky=tk.W, padx=5)

        # 聚合列
        ttk.Label(option_frame, text="聚合列 (可选):").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.agg_cols_var = tk.StringVar()
        agg_entry = ttk.Entry(option_frame, textvariable=self.agg_cols_var, width=50)
        agg_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(option_frame, text="（逗号或空格分隔）").grid(row=1, column=2, sticky=tk.W, padx=5)

        # 输出文件
        ttk.Label(option_frame, text="输出文件:").grid(row=2, column=0, sticky=tk.W, padx=5)
        self.stats_output_var = tk.StringVar()
        ttk.Entry(option_frame, textvariable=self.stats_output_var, width=50).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Button(option_frame, text="浏览", command=self.browse_stats_output).grid(row=2, column=2, padx=5)

        # 统计按钮
        ttk.Button(parent, text="📊 开始统计", command=self.run_stats, width=20).pack(pady=10)

    def add_merge_files(self):
        files = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if files:
            self.merge_files.extend(files)
            for f in files:
                self.merge_listbox.insert(tk.END, f)

    def clear_merge_files(self):
        self.merge_files = []
        self.merge_listbox.delete(0, tk.END)

    def browse_merge_output(self):
        file = filedialog.asksaveasfilename(
            title="保存合并文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file:
            self.merge_output_var.set(file)

    def browse_stats_file(self):
        file = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file:
            self.stats_file_var.set(file)
            self.load_columns()

    def browse_stats_output(self):
        file = filedialog.asksaveasfilename(
            title="保存统计文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file:
            self.stats_output_var.set(file)

    def load_columns(self):
        file_path = self.stats_file_var.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return

        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            self.current_columns = df.columns.tolist()
            self.cols_label.config(text=f"可用的列名: {', '.join(self.current_columns)}")
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {e}")

    def parse_cols(self, text):
        """解析列名，支持逗号和空格分隔"""
        if not text:
            return None
        cols = [col.strip() for col in text.replace(',', ' ').split()]
        return cols if cols else None

    def run_merge(self):
        if len(self.merge_files) < 2:
            messagebox.showwarning("警告", "请至少选择两个 Excel 文件")
            return

        output_path = self.merge_output_var.get()
        if not output_path:
            messagebox.showwarning("警告", "请指定输出文件路径")
            return

        sheet_name = self.sheet_var.get() or None

        # 创建进度窗口
        self.show_progress_window()

        # 在线程中执行
        thread = threading.Thread(target=self._run_merge_thread, args=(output_path, sheet_name))
        thread.daemon = True
        thread.start()

    def _run_merge_thread(self, output_path, sheet_name):
        result = merge_excels(
            self.merge_files,
            output_path,
            merge_by=self.merge_by_var.get(),
            sheet_name=sheet_name,
            log_callback=self.log_message
        )
        self.root.after(0, lambda: self.close_progress_window(result))

    def run_stats(self):
        file_path = self.stats_file_var.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return

        group_cols = self.parse_cols(self.group_cols_var.get())
        agg_cols = self.parse_cols(self.agg_cols_var.get())
        output_path = self.stats_output_var.get()

        # 创建进度窗口
        self.show_progress_window()

        # 在线程中执行
        thread = threading.Thread(target=self._run_stats_thread, args=(file_path, group_cols, agg_cols, output_path))
        thread.daemon = True
        thread.start()

    def _run_stats_thread(self, file_path, group_cols, agg_cols, output_path):
        result = statistics_excel(
            file_path,
            group_cols=group_cols,
            agg_cols=agg_cols,
            output_path=output_path,
            log_callback=self.log_message
        )
        self.root.after(0, lambda: self.close_progress_window(result))

    def show_progress_window(self):
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("处理中...")
        self.progress_window.geometry("600x400")
        # 设置为模态窗口，防止创建多个窗口
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()

        # 日志显示
        log_frame = ttk.Frame(self.progress_window)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(log_frame, text="处理日志:").pack(anchor=tk.W)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log_message(self, msg):
        if hasattr(self, 'log_text') and self.log_text.winfo_exists():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, msg + '\n')
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')

    def close_progress_window(self, success):
        if hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            self.progress_window.destroy()

        if success:
            messagebox.showinfo("成功", "操作完成！")
        else:
            messagebox.showerror("失败", "操作失败，请查看日志了解详情")


def main():
    root = tk.Tk()
    app = ExcelToolGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
