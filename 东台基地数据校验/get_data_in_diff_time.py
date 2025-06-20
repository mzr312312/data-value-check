import pandas as pd
import requests
from datetime import datetime, timedelta
import os

# 定义输入和输出文件路径
input_file_path = r".\采集需求_DT.xlsx"
output_dir = r"."

# 读取输入文件
df_input = pd.read_excel(input_file_path)

# 确保列名正确
if "采集点编码" not in df_input.columns or "时间戳" not in df_input.columns:
    raise ValueError("输入文件必须包含'采集点编码'和'时间戳'两列")

# 定义接口地址
api_url = "http://10.52.11.58:8081/japrojecttag/timeseries"


# 定义函数：根据时间戳生成 startTime 和 endTime
def generate_time_range(timestamp):
    # 将 Timestamp 转换为字符串
    if isinstance(timestamp, pd.Timestamp):
        timestamp = timestamp.strftime("%Y/%m/%d %H:%M")
    timestamp_dt = datetime.strptime(timestamp, "%Y/%m/%d %H:%M")
    start_time = (timestamp_dt - timedelta(minutes=20)).strftime("%Y-%m-%d %H:%M:%S")
    end_time = (timestamp_dt + timedelta(minutes=20)).strftime("%Y-%m-%d %H:%M:%S")
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
    # 检查 time_series 是否为空
    if not time_series:
        return None

    target_dt = datetime.strptime(target_timestamp, "%Y/%m/%d %H:%M")
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
    # 获取 tagCodes 和时间范围
    tag_codes = group["采集点编码"].tolist()
    start_time, end_time = generate_time_range(timestamp)

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
        time_series = item.get("timeSeries", [])  # 确保 timeSeries 至少是一个空列表

        # 找到离目标时间戳最近的数据
        nearest_data = find_nearest_data(time_series, timestamp.strftime("%Y/%m/%d %H:%M"))
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
    print(f"数据已成功保存到文件：{output_file_path}")
else:
    print("未获取到任何有效数据，未生成输出文件。")