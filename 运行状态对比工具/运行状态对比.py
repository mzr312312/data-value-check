import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import pandas as pd
import numpy as np
import time
import os
from datetime import datetime


def process_data():
    # 清空日志
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, "准备处理数据...\n")
    text_area.update()

    # 选择文件
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
    )
    if not file_path:
        text_area.insert(tk.END, "未选择文件，操作已取消\n")
        return

    text_area.insert(tk.END, f"已选择文件: {file_path}\n")
    text_area.insert(tk.END, "开始读取数据...\n")
    text_area.update()

    try:
        # 读取数据
        iot_df = pd.read_excel(file_path, sheet_name='IOT数据')
        eap_df = pd.read_excel(file_path, sheet_name='EAP数据')

        # 检查列名
        required_columns = ['采集点编码', '时间戳', '返回值']
        for col in required_columns:
            if col not in iot_df.columns or col not in eap_df.columns:
                raise ValueError(f"Sheet中缺少必需的列: {col}")

        # 确保时间戳为datetime类型
        iot_df['时间戳'] = pd.to_datetime(iot_df['时间戳'])
        eap_df['时间戳'] = pd.to_datetime(eap_df['时间戳'])

        # 添加原始索引以便后续匹配
        iot_df = iot_df.reset_index(drop=False).rename(columns={'index': '原始索引_IOT'})
        eap_df = eap_df.reset_index(drop=False).rename(columns={'index': '原始索引_EAP'})

        # 创建结果DataFrame
        result_list = []

        # 对每个采集点编码进行处理
        text_area.insert(tk.END, "开始匹配数据...\n")
        text_area.update()

        all_tags = set(iot_df['采集点编码']).union(set(eap_df['采集点编码']))

        for tag in all_tags:
            tag_iot = iot_df[iot_df['采集点编码'] == tag].copy()
            tag_eap = eap_df[eap_df['采集点编码'] == tag].copy()

            if tag_iot.empty or tag_eap.empty:
                continue

            # 排序以加速查找
            tag_iot.sort_values('时间戳', inplace=True)
            tag_eap.sort_values('时间戳', inplace=True)

            # 转换为纳秒（数值类型）
            iot_times_ns = tag_iot['时间戳'].view('int64').values
            eap_times_ns = tag_eap['时间戳'].view('int64').values

            # 对IOT中的每个点查找最近的EAP点
            for i, (idx_iot, row_iot) in enumerate(tag_iot.iterrows()):
                time_val = iot_times_ns[i]

                # 找到最接近的EAP时间戳
                closest_idx = np.argmin(np.abs(eap_times_ns - time_val))
                row_eap = tag_eap.iloc[closest_idx]

                # 添加到结果
                time_diff = abs(row_iot['时间戳'] - row_eap['时间戳'])
                result_list.append({
                    '采集点编码': tag,
                    '时间戳_IOT': row_iot['时间戳'],
                    '返回值_IOT': row_iot['返回值'],
                    '时间戳_EAP': row_eap['时间戳'],
                    '返回值_EAP': row_eap['返回值'],
                    '时间差': time_diff
                })

        # 创建结果DataFrame
        if not result_list:
            text_area.insert(tk.END, "未找到匹配的数据\n")
            return

        result_df = pd.DataFrame(result_list)

        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"同一组数据对比_{timestamp}.xlsx"
        output_path = os.path.join(os.path.dirname(file_path), output_filename)

        # 保存结果
        result_df.to_excel(output_path, index=False)

        text_area.insert(tk.END, f"处理完成! 匹配到 {len(result_df)} 组数据\n")
        text_area.insert(tk.END, f"结果已保存至: {output_path}\n")

    except Exception as e:
        text_area.insert(tk.END, f"处理过程中发生错误: {str(e)}\n")


# 创建主窗口
root = tk.Tk()
root.title("光伏设备数据比对工具")
root.geometry("700x500")

# 创建按钮
button_frame = ttk.Frame(root)
button_frame.pack(pady=20)

process_btn = ttk.Button(button_frame, text="选择Excel文件并处理", command=process_data)
process_btn.pack(side=tk.LEFT, padx=10)

# 创建日志区域
log_frame = ttk.LabelFrame(root, text="处理日志")
log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

text_area = scrolledtext.ScrolledText(log_frame, height=20)
text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
text_area.insert(tk.END, "请点击上方按钮选择Excel文件...\n")

# 运行主循环
root.mainloop()