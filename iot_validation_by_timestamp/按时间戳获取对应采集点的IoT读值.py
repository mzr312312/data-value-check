import pandas as pd
import requests
from datetime import datetime, timedelta
import os
import tkinter as tk
from tkinter import ttk


# 定义输入和输出文件路径
input_file_path = r".\待拉取的编码和时间戳.xlsx"
output_dir = r"."

# 读取各基地 URL 配置文件
base_url_file = "各基地url.txt"
if not os.path.exists(base_url_file):
    raise FileNotFoundError(f"未找到配置文件：{base_url_file}")

# 解析各基地 URL
with open(base_url_file, "r", encoding="utf-8") as f:
    base_urls = dict(line.strip().split("=") for line in f if line.strip())

# 检查是否成功解析出基地 URL
if not base_urls:
    raise ValueError("未能从配置文件中解析出有效的基地 URL")

# GUI 界面相关函数
def fetch_and_process_data():
    selected_base = base_var.get()
    api_url = base_urls.get(selected_base)
    if not api_url:
        result_label.config(text="请选择一个有效的基地！")
        return

    try:
        # 获取用户选择的时间偏移量
        time_offset = int(time_offset_var.get())
        if time_offset < 10 or time_offset > 120:
            raise ValueError("时间偏移量必须在 10 到 120 分钟之间")

        # 读取输入文件
        df_input = pd.read_excel(input_file_path)

        # 确保列名正确
        if "采集点编码" not in df_input.columns or "时间戳" not in df_input.columns:
            raise ValueError("输入文件必须包含'采集点编码'和'时间戳'两列")

        # 定义函数：清理时间戳字符串
        def clean_timestamp(timestamp):
            if isinstance(timestamp, str):
                timestamp = timestamp.strip().replace("*", "").replace("#", "")  # 扩展清理规则
            return timestamp

        # 定义函数：尝试将时间戳转换为 datetime 对象
        def parse_timestamp(timestamp):
            """
            尝试将时间戳字符串转换为 datetime 对象，支持多种常见格式。
            """
            timestamp = clean_timestamp(timestamp)  # 清理时间戳

            # 支持的时间戳格式列表
            time_formats = [
                "%Y/%m/%d %H:%M:%S",  # yyyy/m/d h:mm:ss
                "%Y-%m-%d %H:%M:%S",  # yyyy-m-d h:mm:ss
                "%Y/%m/%d %H:%M",     # yyyy/m/d h:mm
                "%Y-%m-%d %H:%M",     # yyyy-m-d h:mm
                "%Y/%m/%d",           # yyyy/m/d
                "%Y-%m-%d",           # yyyy-m-d
                "%Y%m%d%H%M%S",       # yyyymmddhhmmss
                "%Y%m%d%H%M",         # yyyymmddhhmm
                "%Y/%m/%d %H:%M:%S.%f",  # yyyy/m/d h:mm:ss.fff
                "%Y-%m-%d %H:%M:%S.%f",  # yyyy-m-d h:mm:ss.fff
            ]

            # 如果已经是 datetime 类型，直接返回
            if isinstance(timestamp, datetime):
                return timestamp

            # 如果是 pandas Timestamp 类型，转换为 datetime
            if isinstance(timestamp, pd.Timestamp):
                return timestamp.to_pydatetime()

            # 尝试匹配所有支持的格式
            for fmt in time_formats:
                try:
                    return datetime.strptime(str(timestamp), fmt)
                except ValueError:
                    continue

            # 如果所有格式都不匹配，抛出异常
            raise ValueError(f"无法解析时间戳：'{timestamp}'，请检查格式是否正确")

        # 定义函数：根据时间戳生成 startTime 和 endTime
        def generate_time_range(timestamp, offset_minutes):
            # 转换为 datetime 对象
            timestamp = parse_timestamp(timestamp)

            # 生成时间范围
            start_time = (timestamp - timedelta(minutes=offset_minutes)).strftime("%Y-%m-%d %H:%M:%S")
            end_time = (timestamp + timedelta(minutes=offset_minutes)).strftime("%Y-%m-%d %H:%M:%S")
            return start_time, end_time

        # 定义函数：发送 POST 请求获取数据
        def fetch_data(tag_codes, start_time, end_time):
            body = {
                "tagCodes": tag_codes,
                "startTime": start_time,
                "endTime": end_time
            }
            response = requests.post(api_url, json=body)
            if response.status_code == 200:
                return response.json()
            else:
                raise Exception(f"接口请求失败，状态码：{response.status_code}, 响应内容：{response.text}")

        # 定义函数：找到离目标时间戳最近的数据
        def find_nearest_data(time_series, target_timestamp):
            # 转换为目标 datetime 对象
            target_dt = parse_timestamp(target_timestamp)

            if not time_series:
                return None

            nearest_data = None
            min_diff = float('inf')
            for entry in time_series:
                entry_time = datetime.strptime(entry["time"], "%Y-%m-%d %H:%M:%S")
                diff = abs((entry_time - target_dt).total_seconds())
                if diff < min_diff:
                    min_diff = diff
                    nearest_data = entry
            return nearest_data

        # 按时间戳分组
        grouped = df_input.groupby("时间戳")

        # 存储结果
        results = []

        # 遍历每组数据
        for timestamp, group in grouped:
            tag_codes = group["采集点编码"].tolist()
            start_time, end_time = generate_time_range(timestamp, time_offset)

            # 调用接口获取数据
            try:
                response_data = fetch_data(tag_codes, start_time, end_time)
            except Exception as e:
                print(f"获取数据失败，时间戳：{timestamp}，错误信息：{e}")
                continue

            # 解析返回数据
            if response_data.get("code") != 0:
                print(f"接口返回错误，时间戳：{timestamp}，错误信息：{response_data.get('msg')}")
                continue

            for item in response_data["data"]:
                tag_code = item["tagCode"]
                time_series = item.get("timeSeries", [])

                # 找到离目标时间戳最近的数据
                nearest_data = find_nearest_data(time_series, timestamp)
                if nearest_data:
                    results.append({
                        "采集点编码": tag_code,
                        "返回值": nearest_data["tagValue"],
                        "时间戳": nearest_data["time"]
                    })

        # 将结果保存到 Excel 文件
        if results:
            df_output = pd.DataFrame(results)
            output_file_name = f"多个时间的iot数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
            output_file_path = os.path.join(output_dir, output_file_name)
            df_output.to_excel(output_file_path, index=False)
            result_label.config(text=f"数据已成功保存到文件：{output_file_path}")
        else:
            result_label.config(text="未获取到任何有效数据，未生成输出文件。")

    except Exception as e:
        result_label.config(text=f"发生错误：{str(e)}")


# 创建 GUI 界面
root = tk.Tk()
root.title("IoT 数据获取工具")

# 设置窗口大小
root.geometry("720x520")

# 添加标签
tk.Label(root, text="选择基地：").pack(pady=10)

# 添加基地下拉框
base_var = tk.StringVar()
base_combobox = ttk.Combobox(root, textvariable=base_var, state="readonly")
base_combobox['values'] = list(base_urls.keys())
base_combobox.pack(pady=0.1)

# 添加时间偏移量标签
tk.Label(root, text=f"选择从接口获取IoT数据的时间偏移量（向前/向后，分钟）：\n注意：此项选得太大会导致接口返回缓慢/报错").pack(pady=10)

# 添加时间偏移量下拉框
time_offset_var = tk.StringVar(value="20")  # 默认值为 50 分钟
time_offset_combobox = ttk.Combobox(root, textvariable=time_offset_var, state="readonly")
time_offset_combobox['values'] = [str(i) for i in range(10, 121, 10)]  # 选项为 10, 20, ..., 120
time_offset_combobox.pack(pady=0.1)

# 添加按钮
fetch_button = tk.Button(root, text="开始获取对应时间戳的 IoT 值", command=fetch_and_process_data)
fetch_button.pack(pady=10)

# 添加结果显示标签
result_label = tk.Label(root, text="", wraplength=500, justify="left")
result_label.pack(pady=1)

# 添加使用说明文字
usage_text = """
使用说明：
1.填写“待拉取的编码和时间戳.xlsx”文件，必须有“采集点编码”和“时间戳”两列，列名不能修改；
2.时间戳可以是以下任意格式：
   - yyyy/m/d h:mm:ss （如 2023/12/31 23:59:59）
   - yyyy-m-d h:mm:ss （如 2023-12-31 23:59:59）
   - yyyy/m/d h:mm （如 2023/12/31 23:59）
   - yyyy-m-d h:mm （如 2023-12-31 23:59）
   - yyyy/m/d （如 2023/12/31）
   - yyyy-m-d （如 2023-12-31）
   - yyyymmddhhmmss （如 20231231235959）
   - yyyymmddhhmm （如 202312312359）
   - yyyy/m/d h:mm:ss.fff （如 2023/12/31 23:59:59.123）
   - yyyy-m-d h:mm:ss.fff （如 2023-12-31 23:59:59.456）
3.此程序会找到距离目标时间戳最近的IoT值（范围：指定前后偏移量），后续与基地抄表值进行比对；
4.如果某个采集点没有数据，则不会出现在输出文件中；
"""
usage_label = tk.Label(root, text=usage_text, wraplength=700, justify="left", anchor="w", font=("Arial", 10))
usage_label.pack(pady=1, fill=tk.X)

# 启动主循环
root.mainloop()