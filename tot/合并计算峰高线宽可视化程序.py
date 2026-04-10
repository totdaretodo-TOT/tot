import os
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
import threading
from datetime import datetime
from scipy.signal import find_peaks, peak_widths
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class DataProcessorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("数据处理工具 - 合并 & 峰高线宽分析")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)

        self.data_dir = tk.StringVar(value=r'D:\电场实验数据汇总2024.6.12\拉比频率测试程序')
        self.file_prefix = tk.StringVar(value='tek')
        self.file_suffix = tk.StringVar(value='ALL.csv')
        self.skip_rows = tk.IntVar(value=21)
        self.output_name = tk.StringVar(value='合并结果')
        self.t_sample = tk.DoubleVar(value=0.008)
        self.prominence = tk.DoubleVar(value=0.01)
        self.peak_width = tk.IntVar(value=1)
        self.is_absorption = tk.BooleanVar(value=False)

        self.combined_data = None
        self.analysis_results = None

        self._setup_ui()

    def _setup_ui(self):
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        tk.Label(
            title_frame, text="数据处理工具 - 合并 & 峰高线宽分析",
            font=("Microsoft YaHei", 16, "bold"),
            bg="#2c3e50", fg="white"
        ).pack(pady=15)

        config_frame = tk.LabelFrame(
            self.root, text=" 参数配置 ",
            font=("Microsoft YaHei", 11),
            padx=15, pady=10
        )
        config_frame.pack(fill="x", padx=15, pady=(15, 5))

        row1 = tk.Frame(config_frame)
        row1.pack(fill="x", pady=5)
        tk.Label(row1, text="数据文件夹:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row1, textvariable=self.data_dir, width=50, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)
        tk.Button(row1, text="浏览...", command=self.browse_folder, width=8).pack(side="left", padx=5)

        row2 = tk.Frame(config_frame)
        row2.pack(fill="x", pady=5)
        tk.Label(row2, text="文件前缀:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row2, textvariable=self.file_prefix, width=15, font=("Microsoft YaHei", 10)).pack(side="left", padx=(5, 15))
        tk.Label(row2, text="文件后缀:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row2, textvariable=self.file_suffix, width=15, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)
        tk.Label(row2, text="跳过行数:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row2, textvariable=self.skip_rows, width=8, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)
        tk.Label(row2, text="采样时间:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row2, textvariable=self.t_sample, width=10, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)

        row3 = tk.Frame(config_frame)
        row3.pack(fill="x", pady=5)
        tk.Label(row3, text="输出文件名:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row3, textvariable=self.output_name, width=25, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)
        tk.Checkbutton(row3, text="吸收峰(负峰)", variable=self.is_absorption, font=("Microsoft YaHei", 10)).pack(side="left", padx=20)
        tk.Label(row3, text="峰 prominence:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row3, textvariable=self.prominence, width=10, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)
        tk.Label(row3, text="峰 width:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(row3, textvariable=self.peak_width, width=8, font=("Microsoft YaHei", 10)).pack(side="left", padx=5)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        self.merge_btn = tk.Button(
            btn_frame, text="📂 合并数据", command=self.run_merge,
            font=("Microsoft YaHei", 11, "bold"),
            bg="#3498db", fg="white", width=15, height=2,
            cursor="hand2", relief="flat"
        )
        self.merge_btn.pack(side="left", padx=10)

        self.analyze_single_btn = tk.Button(
            btn_frame, text="📄 分析单文件", command=self.analyze_single_file,
            font=("Microsoft YaHei", 11, "bold"),
            bg="#9b59b6", fg="white", width=15, height=2,
            cursor="hand2", relief="flat"
        )
        self.analyze_single_btn.pack(side="left", padx=10)

        self.analyze_btn = tk.Button(
            btn_frame, text="📊 分析峰高线宽", command=self.run_analyze,
            font=("Microsoft YaHei", 11, "bold"),
            bg="#27ae60", fg="white", width=15, height=2,
            cursor="hand2", relief="flat", state="disabled"
        )
        self.analyze_btn.pack(side="left", padx=10)

        self.plot_btn = tk.Button(
            btn_frame, text="📈 可视化", command=self.show_plot,
            font=("Microsoft YaHei", 11, "bold"),
            bg="#e74c3c", fg="white", width=15, height=2,
            cursor="hand2", relief="flat", state="disabled"
        )
        self.plot_btn.pack(side="left", padx=10)

        self.save_btn = tk.Button(
            btn_frame, text="💾 保存结果", command=self.save_results,
            font=("Microsoft YaHei", 10),
            width=12, height=2,
            cursor="hand2", relief="flat", state="disabled"
        )
        self.save_btn.pack(side="left", padx=10)

        self.status_label = tk.Label(
            self.root, text="就绪", font=("Microsoft YaHei", 10),
            anchor="w", padx=15, pady=5
        )
        self.status_label.pack(fill="x")

        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill="both", expand=True, padx=15, pady=(5, 5))

        columns = ("文件名", "峰高(mV)", "线宽(FWHM)", "峰高/线宽")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12)

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.tag_configure("even", background="#f8f9fa")
        self.tree.tag_configure("odd", background="white")

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=15, pady=(0, 5))

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.data_dir.set(folder)
            self.status_label.config(text=f"已选择文件夹: {folder}")

    def analyze_single_file(self):
        file_path = filedialog.askopenfilename(
            title="选择合并后的CSV文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        self.analyze_single_btn.config(state="disabled", text="分析中...")
        self.progress.start(10)
        self.status_label.config(text=f"正在分析: {os.path.basename(file_path)}")

        def worker():
            try:
                prominence = self.prominence.get()
                width = self.peak_width.get()
                is_absorption = self.is_absorption.get()

                df = pd.read_csv(file_path)
                if 'TIME' not in df.columns:
                    self.root.after(0, lambda: messagebox.showerror("错误", "CSV文件缺少TIME列！"))
                    self.root.after(0, lambda: self.on_error("CSV文件缺少TIME列"))
                    return

                time_values = df['TIME'].values
                dx = time_values[1] - time_values[0]
                if pd.isna(dx) or dx == 0:
                    dx = self.t_sample.get()

                results = []

                for col in df.columns:
                    if col == 'TIME':
                        continue

                    y_values = df[col].values

                    if is_absorption:
                        y_values = -y_values

                    peaks, properties = find_peaks(y_values, prominence=prominence, width=width)

                    if len(peaks) > 0:
                        peak_idx = peaks[np.argmax(properties['prominences'])]
                    else:
                        peak_idx = np.argmax(y_values)

                    widths, _, _, _ = peak_widths(y_values, [peak_idx], rel_height=0.5)

                    peak_height = (np.max(y_values) - np.min(y_values)) * 1000

                    if len(widths) > 0 and widths[0] > 0:
                        fwhm = widths[0] * dx
                    else:
                        fwhm = np.nan

                    results.append({
                        '文件名': col,
                        '峰高(mV)': round(peak_height, 4),
                        '线宽(FWHM)': round(fwhm, 6) if not np.isnan(fwhm) else np.nan,
                        '峰高/线宽': round(peak_height / fwhm, 2) if fwhm > 0 else np.nan
                    })

                self.analysis_results = pd.DataFrame(results)
                self.root.after(0, self.on_analyze_complete)

            except Exception as e:
                self.root.after(0, self.on_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def run_merge(self):
        data_dir = self.data_dir.get()
        prefix = self.file_prefix.get()
        suffix = self.file_suffix.get()
        skip_rows = self.skip_rows.get()
        t_sample = self.t_sample.get()
        output_name = self.output_name.get()

        if not os.path.isdir(data_dir):
            messagebox.showerror("错误", f"文件夹不存在:\n{data_dir}")
            return

        all_files = [f for f in os.listdir(data_dir) if f.startswith(prefix) and f.endswith(suffix)]
        all_files.sort()

        if not all_files:
            messagebox.showwarning("警告", f"未找到匹配的文件！\n前缀: {prefix}\n后缀: {suffix}")
            return

        self.merge_btn.config(state="disabled", text="合并中...")
        self.progress.start(10)
        self.status_label.config(text=f"正在处理 0/{len(all_files)} 个文件...")

        def worker():
            try:
                df_combined = pd.DataFrame()

                for i, file in enumerate(all_files):
                    file_path = os.path.join(data_dir, file)

                    try:
                        df = pd.read_csv(file_path, skiprows=skip_rows, header=None, usecols=[2], names=['CH2'])
                        df['CH2'] = pd.to_numeric(df['CH2'], errors='coerce')
                        df = df.dropna().reset_index(drop=True)

                        if i == 0:
                            df_combined['TIME'] = np.arange(len(df)) * t_sample

                        col_name = file.replace('.csv', '')
                        df_combined[col_name] = df['CH2']

                        self.root.after(0, self.status_label.config,
                                      {"text": f"正在处理 {i+1}/{len(all_files)} 个文件... {file}"})

                    except Exception as e:
                        print(f"解析 {file} 失败: {e}")

                self.combined_data = df_combined
                save_dir = data_dir
                output_file = os.path.join(save_dir, f"{output_name}.csv")
                output_excel = os.path.join(save_dir, f"{output_name}.xlsx")

                df_combined.to_csv(output_file, index=False, encoding='utf-8-sig')
                df_combined.to_excel(output_excel, index=False)

                self.root.after(0, self.on_merge_complete, len(all_files), output_file, output_excel)

            except Exception as e:
                self.root.after(0, self.on_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def on_merge_complete(self, count, csv_file, excel_file):
        self.progress.stop()
        self.merge_btn.config(state="normal", text="📂 合并数据")
        self.analyze_btn.config(state="normal")
        self.status_label.config(
            text=f"✅ 合并完成！成功处理了 {count} 个文件。\n"
                 f"📁 CSV: {csv_file}\n"
                 f"📁 Excel: {excel_file}"
        )
        messagebox.showinfo("成功", f"✅ 合并完成！\n\n成功处理了 {count} 个文件。\n\n"
                                 f"📊 CSV: {os.path.basename(csv_file)}\n"
                                 f"📊 Excel: {os.path.basename(excel_file)}")

    def on_error(self, error_msg):
        self.progress.stop()
        self.merge_btn.config(state="normal", text="📂 合并数据")
        self.analyze_single_btn.config(state="normal", text="📄 分析单文件")
        messagebox.showerror("错误", f"处理失败:\n{error_msg}")
        self.status_label.config(text=f"❌ 错误: {error_msg}")

    def run_analyze(self):
        if self.combined_data is None or self.combined_data.empty:
            csv_path = os.path.join(self.data_dir.get(), f"{self.output_name.get()}.csv")
            if os.path.exists(csv_path):
                self.combined_data = pd.read_csv(csv_path)
            else:
                csv_path = filedialog.askopenfilename(
                    title="选择合并后的CSV文件",
                    filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
                )
                if csv_path:
                    self.combined_data = pd.read_csv(csv_path)
                else:
                    return

        self.analyze_btn.config(state="disabled", text="分析中...")
        self.progress.start(10)
        self.status_label.config(text="正在分析峰高线宽...")

        def worker():
            try:
                time_values = self.combined_data['TIME'].values
                dx = time_values[1] - time_values[0]
                if pd.isna(dx) or dx == 0:
                    dx = self.t_sample.get()

                prominence = self.prominence.get()
                width = self.peak_width.get()
                is_absorption = self.is_absorption.get()

                results = []

                for col in self.combined_data.columns:
                    if col == 'TIME':
                        continue

                    y_values = self.combined_data[col].values

                    if is_absorption:
                        y_values = -y_values

                    peaks, properties = find_peaks(y_values, prominence=prominence, width=width)

                    if len(peaks) > 0:
                        peak_idx = peaks[np.argmax(properties['prominences'])]
                    else:
                        peak_idx = np.argmax(y_values)

                    widths, _, _, _ = peak_widths(y_values, [peak_idx], rel_height=0.5)

                    peak_height = (np.max(y_values) - np.min(y_values)) * 1000

                    if len(widths) > 0 and widths[0] > 0:
                        fwhm = widths[0] * dx
                    else:
                        fwhm = np.nan

                    results.append({
                        '文件名': col,
                        '峰高(mV)': round(peak_height, 4),
                        '线宽(FWHM)': round(fwhm, 6) if not np.isnan(fwhm) else np.nan,
                        '峰高/线宽': round(peak_height / fwhm, 2) if fwhm > 0 else np.nan
                    })

                self.analysis_results = pd.DataFrame(results)

                self.root.after(0, self.on_analyze_complete)

            except Exception as e:
                self.root.after(0, self.on_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def on_analyze_complete(self):
        self.progress.stop()
        self.analyze_btn.config(state="normal", text="📊 分析峰高线宽")
        self.plot_btn.config(state="normal")
        self.save_btn.config(state="normal")

        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, row in self.analysis_results.iterrows():
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(row['文件名'], row['峰高(mV)'], row['线宽(FWHM)'], row['峰高/线宽']), tags=(tag,))

        self.tree.insert("", "end", values=("-" * 20, "-" * 10, "-" * 10, "-" * 10))
        avg_peak = self.analysis_results['峰高(mV)'].mean()
        avg_width = self.analysis_results['线宽(FWHM)'].mean()
        avg_ratio = avg_peak / avg_width if avg_width > 0 else np.nan
        self.tree.insert("", "end", values=("平均值", f"{avg_peak:.4f}", f"{avg_width:.6f}", f"{avg_ratio:.2f}"))

        self.status_label.config(text=f"✅ 分析完成！共 {len(self.analysis_results)} 列数据 | 平均峰高: {avg_peak:.4f} mV | 平均线宽: {avg_width:.6f} s")

    def show_plot(self):
        if self.analysis_results is None or len(self.analysis_results) == 0:
            messagebox.showwarning("警告", "没有可显示的数据！")
            return

        plot_window = tk.Toplevel(self.root)
        plot_window.title("峰高/线宽 比值可视化")
        plot_window.geometry("800x500")

        fig = Figure(figsize=(10, 6))
        ax1 = fig.add_subplot(2, 1, 1)
        ax2 = fig.add_subplot(2, 1, 2)

        filenames = self.analysis_results['文件名'].tolist()
        ratios = self.analysis_results['峰高/线宽'].tolist()

        ax1.bar(range(len(filenames)), ratios, color='#3498db', alpha=0.8)
        ax1.set_title('Peak Height / Linewidth Ratio', fontsize=12)
        ax1.set_xlabel('File Index')
        ax1.set_ylabel('Ratio')
        ax1.grid(True, alpha=0.3)

        ax2.plot(filenames, ratios, 'o-', color='#e74c3c', linewidth=2, markersize=6)
        ax2.set_title('Ratio Trend', fontsize=12)
        ax2.set_xlabel('Filename')
        ax2.set_ylabel('Peak Height / Linewidth')
        ax2.tick_params(axis='x', rotation=45)
        ax2.grid(True, alpha=0.3)

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def save_results(self):
        if self.analysis_results is None or len(self.analysis_results) == 0:
            messagebox.showwarning("警告", "没有可保存的数据！")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("所有文件", "*.*")],
            initialfile=f"峰高线宽分析结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )

        if not filepath:
            return

        try:
            if filepath.endswith('.xlsx'):
                self.analysis_results.to_excel(filepath, index=False)
            else:
                self.analysis_results.to_csv(filepath, index=False, encoding='utf-8-sig')

            messagebox.showinfo("成功", f"结果已保存到:\n{filepath}")
            self.status_label.config(text=f"✅ 结果已保存到: {filepath}")

        except Exception as e:
            messagebox.showerror("错误", f"保存失败:\n{e}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = DataProcessorApp()

    def auto_run():
        app.run_merge()
        app.root.after(500, app.run_analyze)
        app.root.after(1000, lambda: app.status_label.config(
            text="✅ 处理完成！可点击「💾 保存结果」导出数据"))

    app.root.after(500, auto_run)
    app.run()
