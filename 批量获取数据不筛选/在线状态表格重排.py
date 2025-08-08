import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import os
from datetime import datetime
import threading
import re


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("采集点数据合并工具")
        self.root.geometry("700x500")
        self.root.resizable(True, True)

        # 初始化变量
        self.file_path = None
        self.output_dir = os.getcwd()

        # 创建UI组件
        self.create_widgets()

        # 日志记录器
        self.log("程序启动 - 准备好选择文件")

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题标签
        title_label = ttk.Label(main_frame, text="采集点数据合并工具", font=("Arial", 14, "bold"))
        title_label.pack(pady=5)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="输入文件选择")
        file_frame.pack(fill=tk.X, pady=10, padx=5)

        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.pack(side=tk.LEFT, padx=5, pady=5)

        browse_btn = ttk.Button(file_frame, text="选择Excel文件", command=self.select_file)
        browse_btn.pack(side=tk.RIGHT, padx=5, pady=5)

        # 输出设置区域
        output_frame = ttk.LabelFrame(main_frame, text="输出设置")
        output_frame.pack(fill=tk.X, pady=10, padx=5)

        self.dir_label = ttk.Label(output_frame, text=f"输出目录: {self.output_dir}")
        self.dir_label.pack(side=tk.LEFT, padx=5, pady=5)

        dir_btn = ttk.Button(output_frame, text="选择输出目录", command=self.select_output_dir)
        dir_btn.pack(side=tk.RIGHT, padx=5, pady=5)

        # 选项设置区域
        options_frame = ttk.LabelFrame(main_frame, text="处理选项")
        options_frame.pack(fill=tk.X, pady=10, padx=5)

        ttk.Label(options_frame, text="空白单元格处理:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.blank_option = ttk.Combobox(options_frame,
                                         values=["保留空白", "删除空白(左移值)", "填充为0", "填充为NULL"],
                                         width=18)
        self.blank_option.current(0)
        self.blank_option.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(options_frame, text="启用状态校验:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.enable_check = ttk.Combobox(options_frame,
                                         values=["是", "否"],
                                         width=5,
                                         state="readonly")
        self.enable_check.current(0)
        self.enable_check.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)

        # 处理按钮
        process_btn = ttk.Button(main_frame, text="开始处理数据", command=self.start_processing)
        process_btn.pack(pady=10)

        # 进度条
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)

    def log(self, message):
        """在日志区域添加消息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"

        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message + "\n")
        self.log_text.see(tk.END)  # 自动滚动到底部
        self.log_text.config(state=tk.DISABLED)

    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择输入Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )

        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.log(f"已选择文件: {file_path}")

            # 预览文件头
            self.preview_file()

    def preview_file(self):
        """预览文件前几行数据"""
        try:
            # 只读取前5行和需要的列以节省内存
            df = pd.read_excel(self.file_path, nrows=5, usecols=lambda c: c in ['采集点编码', '采集点值', '时间戳'])

            # 检查必要列是否存在
            if not all(col in df.columns for col in ['采集点编码', '采集点值', '时间戳']):
                missing = [col for col in ['采集点编码', '采集点值', '时间戳'] if col not in df.columns]
                self.log(f"错误: 文件中缺少必要的列: {', '.join(missing)}")
                return

            self.log("文件预览(前5行):")
            self.log(str(df.head()))

        except Exception as e:
            self.log(f"文件预览失败: {str(e)}")

    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir = dir_path
            self.dir_label.config(text=f"输出目录: {dir_path}")
            self.log(f"输出目录设置为: {dir_path}")

    def start_processing(self):
        """开始处理数据"""
        if not self.file_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return

        # 禁用按钮防止重复点击
        self.root.config(cursor="wait")
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.DISABLED)

        # 在新线程中执行处理，避免GUI冻结
        threading.Thread(target=self.process_data).start()

    def shift_and_compress(self, row, max_length):
        """删除空白单元格并左移值"""
        # 将时间值列保留，采集点编码不动
        point_code = row[0]
        values = row[1:].tolist()

        # 移除空白值
        non_blank_values = [v for v in values if pd.notna(v) and v != ""]

        # 填充到最大长度
        while len(non_blank_values) < max_length - 1:  # -1 因为采集点编码不算在内
            non_blank_values.append(np.nan)

        # 组合结果：采集点编码 + 处理后的时间值
        return [point_code] + non_blank_values[:max_length - 1]

    def check_machine_status(self, row):
        """根据行数据检查机器状态"""
        # 跳过采集点编码和校验结果列
        values = row[2:].tolist()

        # 如果所有值都一样（包括None），则认为机器停机
        if len(values) > 0:
            # 转换None为特殊字符串以便比较
            clean_values = [v if not pd.isna(v) else None for v in values]

            # 移除空白值（如果有）
            non_blank_values = [v for v in clean_values if v is not None and v != ""]

            if len(non_blank_values) == 0:
                return "停机"  # 所有值都是空白
            elif all(v == non_blank_values[0] for v in non_blank_values):
                return "停机"  # 所有非空值都相同
            else:
                return "开机"  # 有至少两个不同的值

        return "无法判断"  # 没有时间值数据

    def process_data(self):
        """数据处理主函数"""
        try:
            self.progress["value"] = 0
            self.log("开始处理数据...")
            blank_option = self.blank_option.get()
            enable_check = self.enable_check.get() == "是"

            # 读取Excel文件
            self.log("正在读取输入文件...")
            df = pd.read_excel(self.file_path)
            self.progress["value"] = 10

            # 验证数据列
            required_cols = ['采集点编码', '采集点值', '时间戳']
            if not all(col in df.columns for col in required_cols):
                missing = [col for col in required_cols if col not in df.columns]
                raise ValueError(f"输入文件缺少必要的列: {', '.join(missing)}")

            # 检查行数
            total_rows = len(df)
            self.log(f"文件读取完成，共 {total_rows} 行")

            # 转换时间戳
            try:
                self.log("正在转换时间戳格式...")
                df['时间戳'] = pd.to_datetime(df['时间戳'], errors='coerce')
                invalid_time_count = df['时间戳'].isnull().sum()
                if invalid_time_count > 0:
                    self.log(f"警告: 发现 {invalid_time_count} 个无效时间戳")
            except Exception as e:
                raise ValueError(f"时间戳转换失败: {str(e)}")

            # 数据预处理 - 删除完全重复的行
            initial_count = len(df)
            df = df.drop_duplicates()
            removed_count = initial_count - len(df)
            if removed_count > 0:
                self.log(f"移除 {removed_count} 条完全重复的记录")

            self.progress["value"] = 20

            # 检查数据量
            point_count = df['采集点编码'].nunique()
            self.log(f"开始处理 {point_count} 个采集点的数据...")

            # 按采集点分组处理
            grouped_data = {}
            groups = df.groupby('采集点编码')

            # 找出最大时间点数量
            max_timepoints = df.groupby('采集点编码').size().max()
            self.log(f"最大时间点数: {max_timepoints}")

            # 处理每组数据
            for i, (group_name, group_df) in enumerate(groups):
                # 删除重复时间戳的记录（保留第一个）
                group_df = group_df.drop_duplicates(subset=['时间戳'], keep='first')

                # 按时间排序
                group_df = group_df.sort_values('时间戳')

                # 准备行数据
                row = [group_name]  # 第一列为采集点编码
                for j, (_, row_data) in enumerate(group_df.iterrows()):
                    row.append(row_data['采集点值'])

                # 添加到分组数据
                grouped_data[group_name] = row

                # 更新进度
                if (i + 1) % 100 == 0 or i == point_count - 1:
                    self.progress["value"] = 20 + (i + 1) / point_count * 70
                    self.log(f"处理进度: 已完成 {i + 1}/{point_count} 个采集点")

            # 创建结果DataFrame
            self.log("正在创建结果数据表...")

            # 创建时间点列名
            max_columns = max(len(v) for v in grouped_data.values())
            columns = ['采集点编码'] + [f'时间{j + 1}' for j in range(max_columns - 1)]

            # 从字典创建DataFrame
            result_df = pd.DataFrame.from_dict(grouped_data, orient='index',
                                               columns=columns).reset_index(drop=True)

            # 处理空白单元格
            if blank_option != "保留空白":
                self.log(f"正在处理空白单元格: {blank_option}")
                max_timepoints = result_df.shape[1] - 1  # 减去采集点编码列

                if blank_option == "删除空白(左移值)":
                    # 删除空白并左移值
                    result_df = result_df.apply(lambda row: pd.Series(self.shift_and_compress(row, max_columns)),
                                                axis=1)
                    result_df.columns = columns[:len(result_df.columns)]
                elif blank_option == "填充为0":
                    result_df = result_df.fillna(0)
                elif blank_option == "填充为NULL":
                    result_df = result_df.fillna("NULL")

                # 删除完全空白的列
                empty_cols = result_df.columns[result_df.isnull().all()]
                if not empty_cols.empty:
                    self.log(f"移除 {len(empty_cols)} 个空白列")
                    result_df = result_df.drop(columns=empty_cols)

            # 添加校验结果列（如有需要）
            if enable_check:
                self.log("正在计算设备状态...")
                # 添加校验结果列，放在第二列位置
                status_col = []
                for _, row in result_df.iterrows():
                    status_col.append(self.check_machine_status(row))

                # 插入校验结果列到第二列
                result_df.insert(1, '校验结果', status_col)

            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"在线状态校验_{timestamp}.xlsx"
            output_path = os.path.join(self.output_dir, output_filename)

            # 保存结果
            self.log(f"正在保存结果到: {output_path}")
            result_df.to_excel(output_path, index=False)
            self.progress["value"] = 100

            # 最终统计
            input_points = df['采集点编码'].nunique()
            output_points = len(result_df)
            input_timepoints = len(df)
            output_columns = len(result_df.columns) - 1  # 减去采集点编码列

            # 状态统计（如有校验）
            status_stats = {}
            if enable_check and '校验结果' in result_df.columns:
                status_counts = result_df['校验结果'].value_counts()
                for status, count in status_counts.items():
                    status_stats[status] = count

                # 添加到日志
                for status, count in status_stats.items():
                    self.log(f"{status}设备: {count} 个")

            self.log("数据处理完成!")
            self.log(f"输入采集点: {input_points} 个")
            self.log(f"输入时间点: {input_timepoints} 个")
            self.log(f"输出采集点: {output_points} 个")
            self.log(f"输出时间点列: {output_columns} 列")
            self.log(f"结果文件已保存到: {output_path}")

            # 显示成功消息
            messagebox.showinfo("处理完成", f"数据合并成功!\n结果文件已保存到:\n{output_path}")

        except Exception as e:
            self.log(f"处理出错: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            # 恢复UI状态
            self.root.config(cursor="")
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.NORMAL)

            self.progress["value"] = 0


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()