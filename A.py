import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


def format_hms(td):
    """将 Timedelta 对象统一格式化为 HH:MM:SS 字符串"""
    if pd.isna(td):
        return ""
    total_seconds = int(td.total_seconds())
    sign = "-" if total_seconds < 0 else ""
    total_seconds = abs(total_seconds)

    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{sign}{hours:02d}:{minutes:02d}:{seconds:02d}"


class RaceDataProcessorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("铁人三项赛分段成绩计算器")
        self.root.geometry("680x460")
        self.root.resizable(False, False)

        self.create_widgets()

    def create_widgets(self):
        # 1. 原始文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="文件设置", padding=15)
        file_frame.pack(fill="x", padx=15, pady=10)

        ttk.Label(file_frame, text="待处理表格:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        self.input_entry = ttk.Entry(file_frame, width=48)
        self.input_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(
            file_frame, text="浏览...", command=self.select_input_file
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(file_frame, text="输出新文件:").grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.output_entry = ttk.Entry(file_frame, width=48)
        self.output_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(
            file_frame, text="另存为...", command=self.select_output_file
        ).grid(row=1, column=2, padx=5, pady=5)

        # 2. 操作按钮
        btn_frame = ttk.Frame(self.root, padding=5)
        btn_frame.pack(fill="x", padx=15)

        self.run_btn = ttk.Button(
            btn_frame, text="开始处理并导出", command=self.process_data
        )
        self.run_btn.pack(pady=5, ipadx=10, ipady=4)

        # 3. 运行日志提示区
        log_frame = ttk.LabelFrame(self.root, text="运行状态与日志", padding=10)
        log_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.log_text = tk.Text(log_frame, height=8, wrap="word", font=("Menlo", 11))
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def select_input_file(self):
        filetypes = [
            ("表格文件", "*.xlsx *.xls *.csv"),
            ("Excel 文件", "*.xlsx *.xls"),
            ("CSV 文件", "*.csv"),
            ("所有文件", "*.*"),
        ]
        filename = filedialog.askopenfilename(
            title="选择比赛数据表", filetypes=filetypes
        )
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)

            # 自动生成同目录下的导出文件名
            dir_name, base_name = os.path.split(filename)
            name, ext = os.path.splitext(base_name)
            default_out = os.path.join(dir_name, f"{name}_已分段{ext}")
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, default_out)
            self.log(f"已选定输入文件: {base_name}")

    def select_output_file(self):
        filetypes = [("Excel 工作簿", "*.xlsx"), ("CSV 文件", "*.csv")]
        filename = filedialog.asksaveasfilename(
            title="设置输出文件路径",
            filetypes=filetypes,
            defaultextension=".xlsx",
        )
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)

    def process_data(self):
        in_path = self.input_entry.get().strip()
        out_path = self.output_entry.get().strip()

        if not in_path:
            messagebox.showwarning("提示", "请先选择需要处理的数据文件！")
            return
        if not out_path:
            messagebox.showwarning("提示", "请指定导出文件的路径！")
            return

        try:
            self.log("正在读取文件...")
            if in_path.endswith(".csv"):
                df = pd.read_csv(in_path)
            else:
                df = pd.read_excel(in_path)

            # 检查必需列
            required_cols = ["cp1", "cp1_1", "cp1_2", "cp1_3", "cp2_1", "枪时"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise ValueError(
                    f"数据表缺失以下必要列名: {', '.join(missing_cols)}"
                )

            self.log("正在执行各分段耗时计算...")
            parsed = {}
            for col in required_cols:
                parsed[col] = pd.to_timedelta(df[col].astype(str), errors="coerce")

            # 业务公式计算
            td_swim = parsed["cp1"]
            td_t1 = parsed["cp1_2"] - parsed["cp1_1"]
            td_bike = parsed["cp1_3"] - parsed["cp1_2"]
            td_t2 = parsed["cp2_1"] - parsed["cp1_3"]
            td_run = parsed["枪时"] - parsed["cp1"] - td_t1 - td_bike - td_t2

            # 格式化为 HH:MM:SS
            df["游泳"] = td_swim.apply(format_hms)
            df["T1"] = td_t1.apply(format_hms)
            df["自行车"] = td_bike.apply(format_hms)
            df["T2"] = td_t2.apply(format_hms)
            df["跑步"] = td_run.apply(format_hms)

            self.log(f"正在保存文件至: {out_path}")
            if out_path.endswith(".csv"):
                df.to_csv(out_path, index=False, encoding="utf-8-sig")
            else:
                df.to_excel(out_path, index=False)

            self.log(">>> 处理完成！导出成功。")
            messagebox.showinfo("成功", f"处理完成！\n新文件已保存至:\n{out_path}")

        except Exception as e:
            self.log(f"[错误] {str(e)}")
            messagebox.showerror("运行出错", f"发生错误:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = RaceDataProcessorApp(root)
    root.mainloop()
