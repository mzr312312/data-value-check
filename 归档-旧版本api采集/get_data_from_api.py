import os
import sys
import pandas as pd
import requests
from ast import literal_eval
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox


# ==================== 全局配置区域 ====================
TAG_FILE = "批量输入采集点编码.txt"  # 修改为 .txt 文件
BASE_URL_FILE = "各基地url.txt"  # 基地URL文件名
DEBUG = False
TIME_TOLERANCE_MINUTES = 5  # 时间容差（分钟）


# ====================================================

def parse_time(time_str):
    """
    将时间字符串解析为 datetime 对象。
    """
    return datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")


def load_base_urls(file_path):
    """
    从各基地url.txt文件中加载基地名称和对应的API URL。
    返回一个字典：{基地名称: API_URL}
    """
    base_urls = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if "=" in line:
                    base_name, url = line.split("=", 1)
                    base_urls[base_name.strip()] = url.strip()
    except Exception as e:
        print(f"加载基地URL失败: {str(e)}")
    return base_urls


def load_tags_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        start = content.find('[')
        end = content.rfind(']') + 1
        return literal_eval(content[start:end])
    except Exception as e:
        print(f"加载tag失败: {str(e)}")
        return []


def safe_get(data, key, default):
    try:
        return data.get(key, default)
    except AttributeError:
        return default


def process_api_response(response_json):
    data_list = []
    for item in response_json.get("data", []):
        tag_code = safe_get(item, "tagCode", "")
        tag_unit = safe_get(item, "tagUnit", "")
        time_series = safe_get(item, "timeSeries", [])

        if not isinstance(time_series, (list, tuple)):
            time_series = []

        for point in time_series:
            data_list.append([
                tag_code,
                tag_unit,
                safe_get(point, "time", ""),
                safe_get(point, "tagValue", "")
            ])
    return data_list


def find_nearest_record(tag_code, group, target_time, tolerance_sec):
    if group.empty:
        return {'tagCode': tag_code, 'status': '无数据'}

    group['time'] = pd.to_datetime(group['time'], errors='coerce')
    valid_group = group.dropna(subset=['time'])

    if valid_group.empty:
        return {'tagCode': tag_code, 'status': '无有效时间数据'}

    valid_group['time_diff'] = valid_group['time'].apply(
        lambda x: abs((x - target_time).total_seconds())
    )

    nearest_index = valid_group['time_diff'].idxmin()
    nearest_record = valid_group.loc[nearest_index]

    if nearest_record['time_diff'] > tolerance_sec:
        return {
            'tagCode': tag_code,
            'tagUnit': nearest_record.get('tagUnit', ''),
            'time': nearest_record.get('time', ''),
            'tagValue': nearest_record.get('tagValue', ''),
            'status': '超出容差'
        }
    return {
        'tagCode': tag_code,
        'tagUnit': nearest_record.get('tagUnit', ''),
        'time': nearest_record.get('time', ''),
        'tagValue': nearest_record.get('tagValue', ''),
        'status': '正常'
    }


def process_data(df, target_time, tolerance_minutes):
    tolerance_sec = tolerance_minutes * 60
    results = []

    df['time'] = pd.to_datetime(df['time'], errors='coerce')
    total_tags = len(df['tagCode'].unique())

    for idx, (tag_code, group) in enumerate(df.groupby('tagCode', group_keys=True), 1):
        record = find_nearest_record(tag_code, group, target_time, tolerance_sec)
        results.append(record)

    return pd.DataFrame(results)


def main(api_url, start_time, end_time, target_time, batch_size):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    tag_file_path = os.path.join(current_dir, TAG_FILE)
    tag_list = load_tags_from_file(tag_file_path)

    if not tag_list:
        print("没有可处理的tag，程序退出")
        return

    total_batches = (len(tag_list) - 1) // batch_size + 1
    all_data = []
    failed_batches = []

    for i in range(0, len(tag_list), batch_size):
        batch_num = i // batch_size + 1
        batch = tag_list[i:i + batch_size]

        payload = {
            "tagCodes": batch,
            "startTime": start_time,
            "endTime": end_time
        }

        try:
            response = requests.post(
                api_url,
                json=payload,
                timeout=10
            )
            response.raise_for_status()
            result = response.json()

            if not isinstance(result, dict) or "data" not in result:
                raise ValueError("无效的API响应格式")

            batch_data = process_api_response(result)
            all_data.extend(batch_data)

        except Exception as e:
            failed_batches.append(batch_num)
            if DEBUG:
                print(f"\n批次 {batch_num} 处理失败: {str(e)}")
                print("响应内容：", response.text[:500] if response else "无响应")
            continue

    if failed_batches:
        print(f"\n警告：以下批次处理失败：{failed_batches}")

    if not all_data:
        print("\n未获取到有效数据")
        return

    df = pd.DataFrame(all_data, columns=["tagCode", "tagUnit", "time", "tagValue"])
    filtered_df = process_data(df, parse_time(target_time), TIME_TOLERANCE_MINUTES)

    if not filtered_df.empty:
        filename = f"IoT采集到的数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filtered_df.to_excel(filename, index=False, engine='openpyxl')

        tolerance_exceeded = len(filtered_df[filtered_df['status'] == '超出容差'])
        no_data = len(filtered_df[filtered_df['status'].isin(['无数据', '无有效时间数据'])])

        print(f"\n处理完成，数据已保存到：{os.path.abspath(filename)}")
        print(f"异常统计：")
        print(f"  - 超出容差记录数：{tolerance_exceeded}")
        print(f"  - 无数据记录数：{no_data}")
    else:
        print("\n过滤后无有效数据")


# ==================== GUI 部分 ====================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("数据采集与筛选工具")

        # 加载基地URL
        current_dir = os.path.dirname(os.path.abspath(__file__))
        base_url_file_path = os.path.join(current_dir, BASE_URL_FILE)
        self.base_urls = load_base_urls(base_url_file_path)

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 标签和下拉框（基地选择）
        tk.Label(self.root, text="选择基地：").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.base_var = tk.StringVar()
        self.base_dropdown = ttk.Combobox(self.root, textvariable=self.base_var, state="readonly")
        self.base_dropdown["values"] = list(self.base_urls.keys())
        self.base_dropdown.grid(row=0, column=1, padx=10, pady=5)

        # 输入框（起始时间、终止时间、目标时间、每批次请求的数据点数量）
        tk.Label(self.root, text="起始时间：").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.start_time_entry = tk.Entry(self.root)
        self.start_time_entry.insert(0, "2025-06-10 11:50:00")
        self.start_time_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(self.root, text="终止时间：").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.end_time_entry = tk.Entry(self.root)
        self.end_time_entry.insert(0, "2025-06-10 12:10:00")
        self.end_time_entry.grid(row=2, column=1, padx=10, pady=5)

        tk.Label(self.root, text="目标时间：").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.target_time_entry = tk.Entry(self.root)
        self.target_time_entry.insert(0, "2025-06-10 12:00:00")
        self.target_time_entry.grid(row=3, column=1, padx=10, pady=5)

        tk.Label(self.root, text="每批次请求的数据点数量：").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.batch_size_entry = tk.Entry(self.root)
        self.batch_size_entry.insert(0, "20")
        self.batch_size_entry.grid(row=4, column=1, padx=10, pady=5)

        # 开始按钮
        self.start_button = tk.Button(self.root, text="开始从API获取数据并筛选", command=self.start_process)
        self.start_button.grid(row=5, column=0, columnspan=2, pady=10)

    def start_process(self):
        # 获取用户输入
        selected_base = self.base_var.get()
        start_time = self.start_time_entry.get()
        end_time = self.end_time_entry.get()
        target_time = self.target_time_entry.get()
        batch_size = int(self.batch_size_entry.get())

        if not selected_base:
            messagebox.showerror("错误", "请选择一个基地！")
            return

        # 获取对应API URL
        api_url = self.base_urls.get(selected_base)

        # 执行主程序
        main(api_url, start_time, end_time, target_time, batch_size)
        messagebox.showinfo("完成", f"数据处理完成，请查看生成的Excel文件")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()