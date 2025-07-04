import os
import sys
import pandas as pd
import requests
from ast import literal_eval
from datetime import datetime


# ==================== 用户配置区域 ====================
START_TIME = "2025-03-29 00:40:00"  # 更新时间范围
END_TIME = "2025-03-29 01:10:00"
TARGET_TIME = "2025-03-29 01:00:00"  # 目标时间点
TIME_TOLERANCE_MINUTES = 5  # 容差时间（分钟）
BATCH_SIZE = 5
API_URL = "http://10.86.6.3:8081/japrojecttag/timeseries"
TAG_FILE = "SJ_62个错误点校验_Tag_List.md"
DEBUG = True  # 开启调试模式以输出详细信息


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
    try:
        # 检查 JSON 数据是否符合预期结构
        if not isinstance(response_json, dict) or "data" not in response_json:
            raise ValueError("无效的API响应格式")

        for item in response_json.get("data", []):
            tag_code = safe_get(item, "tagCode", "")
            tag_unit = safe_get(item, "tagUnit", "")
            time_series = safe_get(item, "timeseries", [])  # 注意字段名称是否为 timeseries

            if not isinstance(time_series, (list, tuple)):
                print(f"警告: tagCode={tag_code} 的 timeseries 字段不是列表类型")
                continue

            for point in time_series:
                # 检查每个时间点的数据是否完整
                time_value = safe_get(point, "time", "")
                tag_value = safe_get(point, "tagValue", "")

                if not time_value or not tag_value:
                    print(f"警告: tagCode={tag_code} 的某条记录缺少 time 或 tagValue")
                    continue

                data_list.append([
                    tag_code,
                    tag_unit,
                    time_value,
                    tag_value
                ])
    except Exception as e:
        print(f"解析 API 响应失败: {str(e)}")
        return []

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
            # 打印调试信息：请求参数
            if DEBUG:
                print("\n" + "=" * 50)
                print(f"批次 {batch_num} 请求参数:")
                print(f"URL: {API_URL}")
                print(f"Payload: {payload}")
                print("=" * 50)

            response = requests.post(
                API_URL,
                json=payload,
                timeout=10
            )

            # 打印调试信息：响应状态码和内容
            if DEBUG:
                print("\n" + "=" * 50)
                print(f"批次 {batch_num} 响应状态码: {response.status_code}")
                print(f"响应内容（前500字符）: {response.text[:500]}")
                print("=" * 50)

            response.raise_for_status()  # 检查 HTTP 错误
            result = response.json()

            # 打印完整 JSON 数据用于调试
            if DEBUG:
                print("\n" + "=" * 50)
                print("完整 JSON 响应内容:")
                print(result)
                print("=" * 50)

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