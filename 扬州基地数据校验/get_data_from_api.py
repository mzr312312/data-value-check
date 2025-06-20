import os
import sys
import pandas as pd
import requests
from ast import literal_eval
from datetime import datetime

"""
#######################################
# 数据采集与处理脚本使用说明
#######################################

一、脚本功能
本脚本用于从指定API接口批量获取设备监测数据，筛选出指定时间点附近的有效数据，并生成Excel格式的分析报告。

二、配置参数说明（在代码开头修改）
------------------------------------------------------------------------------
参数名                  | 说明                                   | 示例值
-----------------------|---------------------------------------|-------------------
START_TIME             | 数据采集开始时间（包含）               | "2025-04-01 13:50:00"
END_TIME               | 数据采集结束时间（包含）               | "2025-04-01 14:10:00"
TARGET_TIME            | 需要筛选的目标时间点                   | "2025-04-01 14:00:00"
TIME_TOLERANCE_MINUTES | 允许的时间偏差（分钟）                 | 5
BATCH_SIZE             | 每批处理的tag数量                      | 20
API_URL                | 数据接口地址                          | "http://172.17.200.155:8081/..."
TAG_FILE               | 本地存储tag列表的Markdown文件名        | "SJ_电聚合_Tag_List.md"
DEBUG                  | 是否开启调试模式（True/False）         | False
------------------------------------------------------------------------------

三、使用步骤
1. 准备工作
   - 安装依赖库：pandas, requests, openpyxl
   - 在同目录下创建包含tag列表的Markdown文件（格式示例）：
     tag_list = [
         "设备A/监测点1",
         "设备B/监测点2",
         ...
     ]

2. 配置参数
   - 根据需求修改代码开头的配置参数

3. 运行脚本
   > python SJ_get_data_from_api.py

四、输出结果
1. 生成Excel文件（过滤后数据_时间戳.xlsx），包含以下列：
   - tagCode: 设备编码
   - tagUnit: 数据单位
   - time: 实际采集时间
   - tagValue: 监测值
   - status: 数据状态（正常/超出容差/无数据）

2. 控制台输出：
   - 批次处理进度（实时更新）
   - 数据处理进度（实时更新）
   - 异常统计报告
   - 文件保存路径

五、注意事项
1. 时间格式必须为"YYYY-MM-DD HH:MM:SS"
2. 确保TAG_FILE文件存在且格式正确
3. 网络异常时会自动跳过失败批次，建议开启DEBUG模式查看详细错误
4. 当数据量较大时，可适当调大BATCH_SIZE提高效率
5. 生成的Excel文件默认保存在脚本同目录下

六、异常处理
1. 常见错误码：
   - 404: 检查API_URL是否正确
   - 500: 检查时间格式或接口参数
   - 超时: 检查网络连接或调大timeout值

2. 无数据时的处理建议：
   - 扩大START_TIME和END_TIME的时间范围
   - 检查tag列表是否有效
"""


# ==================== 用户配置区域 ====================
START_TIME = "2025-05-23 16:50:00"  # 更新时间范围
END_TIME = "2025-05-23 17:10:00"
TARGET_TIME = "2025-05-23 17:00:00"  # 目标时间点
TIME_TOLERANCE_MINUTES = 5  # 容差时间（分钟）
BATCH_SIZE = 100
API_URL = "http://172.17.200.155:8081/japrojecttag/timeseries"
TAG_FILE = "YZ_电_Tag_List.md"
DEBUG = False


# ====================================================

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


def parse_time(time_str):
    return datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")


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


def process_data(df):
    target_time = parse_time(TARGET_TIME)
    tolerance_sec = TIME_TOLERANCE_MINUTES * 60
    results = []

    df['time'] = pd.to_datetime(df['time'], errors='coerce')
    total_tags = len(df['tagCode'].unique())

    for idx, (tag_code, group) in enumerate(df.groupby('tagCode', group_keys=True), 1):
        # 实时进度显示
        sys.stdout.write(f"\r处理数据：{idx}/{total_tags} ({(idx / total_tags) * 100:.1f}%)")
        sys.stdout.flush()

        record = find_nearest_record(tag_code, group, target_time, tolerance_sec)
        results.append(record)

    print()  # 换行
    return pd.DataFrame(results)


def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, TAG_FILE)
    tag_list = load_tags_from_file(file_path)

    if not tag_list:
        print("没有可处理的tag，程序退出")
        return

    total_batches = (len(tag_list) - 1) // BATCH_SIZE + 1
    all_data = []
    failed_batches = []

    for i in range(0, len(tag_list), BATCH_SIZE):
        batch_num = i // BATCH_SIZE + 1
        batch = tag_list[i:i + BATCH_SIZE]

        # 实时进度显示
        sys.stdout.write(f"\r处理批次：{batch_num}/{total_batches} ({(batch_num / total_batches) * 100:.1f}%)")
        sys.stdout.flush()

        payload = {
            "tagCodes": batch,
            "startTime": START_TIME,
            "endTime": END_TIME
        }

        try:
            response = requests.post(
                API_URL,
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

    print()  # 换行

    if failed_batches:
        print(f"\n警告：以下批次处理失败：{failed_batches}")

    if not all_data:
        print("\n未获取到有效数据")
        return

    df = pd.DataFrame(all_data, columns=["tagCode", "tagUnit", "time", "tagValue"])

    print("\n正在处理时间过滤...")
    filtered_df = process_data(df)

    if not filtered_df.empty:
        filename = f"过滤后数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filtered_df.to_excel(filename, index=False, engine='openpyxl')

        tolerance_exceeded = len(filtered_df[filtered_df['status'] == '超出容差'])
        no_data = len(filtered_df[filtered_df['status'].isin(['无数据', '无有效时间数据'])])

        print(f"\n处理完成，数据已保存到：{os.path.abspath(filename)}")
        print(f"异常统计：")
        print(f"  - 超出容差记录数：{tolerance_exceeded}")
        print(f"  - 无数据记录数：{no_data}")
    else:
        print("\n过滤后无有效数据")


if __name__ == "__main__":
    main()