import pandas as pd
import requests
from datetime import datetime, timedelta
import os
import json
import time
import logging
import re

# ====================== 可调节参数 ======================
API_URL = "http://172.17.200.155:8081/japrojecttag/timeseries"
INPUT_FILE = "待拉取的编码和时间戳.xlsx"  # 输入文件名
OUTPUT_BASE_NAME = "采集数据结果"  # 输出文件基础名（不带后缀）
MISSING_BASE_NAME = "缺失数据点"  # 缺失数据点文件基础名
TIME_RANGE = 30  # 时间范围（分钟） - 以输入时间戳为中心前后扩展的时间
BATCH_SIZE = 50  # 每批次处理的采集点数量
REQUEST_TIMEOUT = 30  # 请求超时时间（秒）
MAX_RETRY_RANGE = 120  # 最大重试时间范围(分钟)
HEADERS = {"Content-Type": "application/json"}  # API请求头
DEBUG_MODE = True  # 调试模式，显示详细日志
ADD_TIMESTAMP = True  # 在输出文件名中添加时间戳


# ======================================================

# 生成安全的文件名（移除无效字符）
def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)


# 生成带时间戳的文件名
def timestamped_filename(base_name, extension=".xlsx"):
    safe_base = safe_filename(base_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if ADD_TIMESTAMP:
        return f"{safe_base}_{timestamp}{extension}"
    else:
        return f"{safe_base}{extension}"


# 配置日志
log_filename = timestamped_filename("api_fetch_debug", ".log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()


def fetch_iot_data():
    """主函数：读取、处理数据并保存结果"""
    try:
        # 生成带时间戳的输出文件名
        output_file = timestamped_filename(OUTPUT_BASE_NAME)
        missing_file = timestamped_filename(MISSING_BASE_NAME)

        # 读取Excel文件
        logger.info(f"读取输入文件: {INPUT_FILE}...")
        df_input = pd.read_excel(INPUT_FILE, engine='openpyxl')

        # 检查必要的列
        if '采集点编码' not in df_input or '时间戳' not in df_input:
            logger.error("错误：输入文件缺少必要列名'采集点编码'或'时间戳'")
            return

        logger.info(f"找到 {len(df_input)} 条记录，开始分组处理...")

        # 按时间戳分组（每个相同的timestamp为一组）
        groups = {}
        for index, row in df_input.iterrows():
            # 处理时间戳格式
            timestamp = row['时间戳']
            if isinstance(timestamp, pd.Timestamp):
                timestamp = timestamp.to_pydatetime()

            # 作为字典键（使用标准化字符串格式）
            key = timestamp.strftime("%Y-%m-%d %H:%M:%S")

            if key not in groups:
                groups[key] = {
                    'timestamp': timestamp,
                    'tag_codes': []
                }
            groups[key]['tag_codes'].append(row['采集点编码'])

        logger.info(f"已分为 {len(groups)} 个时间组")
        results = []
        missing_points = []  # 存储缺失数据的点
        processed_count = 0
        error_count = 0

        # 处理每个组
        for time_str, group in groups.items():
            timestamp = group['timestamp']
            tag_codes = group['tag_codes']

            # 按批次大小切分
            batches = [tag_codes[i:i + BATCH_SIZE] for i in range(0, len(tag_codes), BATCH_SIZE)]
            logger.info(f"时间组 [{time_str}] 有 {len(tag_codes)} 个采集点，分为 {len(batches)} 个批次处理")

            # 计算时间范围（每个组共享相同的时间范围）
            start_time = (timestamp - timedelta(minutes=TIME_RANGE)).strftime("%Y-%m-%d %H:%M:%S")
            end_time = (timestamp + timedelta(minutes=TIME_RANGE)).strftime("%Y-%m-%d %H:%M:%S")

            # 处理每个批次
            for batch_idx, batch in enumerate(batches, 1):
                logger.info(f"  处理批次 {batch_idx}/{len(batches)} | 采集点数量: {len(batch)}")

                payload = {
                    "tagCodes": batch,
                    "startTime": start_time,
                    "endTime": end_time
                }

                # 调试信息：记录请求详情
                if DEBUG_MODE:
                    debug_info = {
                        "request": {
                            "url": API_URL,
                            "payload": payload,
                            "batch_size": len(batch),
                            "time_range": f"{start_time} to {end_time}",
                            "batch_index": f"{batch_idx}/{len(batches)}"
                        }
                    }
                    logger.debug(f"调试信息 - 请求参数: {json.dumps(debug_info, indent=2, ensure_ascii=False)}")

                try:
                    # 发送API请求
                    response = requests.post(API_URL, json=payload, headers=HEADERS, timeout=REQUEST_TIMEOUT)

                    # 调试信息：记录响应详情
                    if DEBUG_MODE:
                        response_info = {
                            "status_code": response.status_code,
                            "response_text": response.text[:1000] + "..." if len(
                                response.text) > 1000 else response.text
                        }
                        logger.debug(f"调试信息 - 响应内容: {json.dumps(response_info, indent=2, ensure_ascii=False)}")

                    if response.status_code != 200:
                        logger.error(f"    API请求失败! 状态码: {response.status_code}")
                        error_count += 1
                        continue

                    try:
                        data = response.json()
                    except json.JSONDecodeError as json_err:
                        logger.error(f"    JSON解析错误: {str(json_err)}")
                        logger.error(f"    响应内容: {response.text[:500]}")
                        error_count += 1
                        continue

                    # 检查API响应代码
                    if data.get('code') != 0:
                        logger.error(f"    API返回错误: code={data.get('code')}, msg={data.get('msg', '未知错误')}")
                        error_count += 1
                        continue

                    # 检查数据部分是否存在且有效
                    api_data = data.get('data')
                    if api_data is None:
                        logger.warning(
                            f"    API返回的data字段为空! 完整响应: {json.dumps(data, ensure_ascii=False)[:500]}")
                        continue

                    # 处理返回数据
                    success_points = 0
                    for item in api_data:
                        tag_code = item.get('tagCode')
                        time_series = item.get('timeSeries', [])

                        if not time_series:
                            logger.warning(f"    采集点 '{tag_code}' 无时间序列数据")
                            missing_points.append({
                                "采集点编码": tag_code,
                                "基准时间": time_str,
                                "查询范围": f"{start_time} 至 {end_time}",
                                "批次": f"{batch_idx}/{len(batches)}",
                                "错误原因": "无时间序列数据"
                            })
                            continue

                        for point in time_series:
                            results.append({
                                "采集点编码": tag_code,
                                "时间戳": point.get('time'),
                                "采集点值": point.get('tagValue')
                            })
                        success_points += 1

                    logger.info(f"    成功处理: {success_points}个采集点, 获取: {len(results) - len(results)}条数据点")
                    processed_count += len(batch)

                except requests.exceptions.RequestException as req_err:
                    logger.error(f"    请求异常: {str(req_err)}")
                    # 记录缺失点信息
                    for tag_code in batch:
                        missing_points.append({
                            "采集点编码": tag_code,
                            "基准时间": time_str,
                            "查询范围": f"{start_time} 至 {end_time}",
                            "批次": f"{batch_idx}/{len(batches)}",
                            "错误原因": str(req_err)
                        })
                    error_count += 1
                except Exception as e:
                    logger.error(f"    处理异常: {str(e)}")
                    # 记录缺失点信息
                    for tag_code in batch:
                        missing_points.append({
                            "采集点编码": tag_code,
                            "基准时间": time_str,
                            "查询范围": f"{start_time} 至 {end_time}",
                            "批次": f"{batch_idx}/{len(batches)}",
                            "错误原因": str(e)
                        })
                    error_count += 1

        # 保存结果
        if results:
            df_output = pd.DataFrame(results)
            df_output.to_excel(output_file, index=False, engine='openpyxl')
            logger.info(f"\n处理完成! 共处理 {processed_count} 个采集点请求")
            logger.info(f"获取到 {len(results)} 条数据记录")
            logger.info(f"数据结果已保存到: {os.path.abspath(output_file)}")

            # 保存缺失数据点信息
            if missing_points:
                df_missing = pd.DataFrame(missing_points)
                df_missing.to_excel(missing_file, index=False, engine='openpyxl')
                logger.info(f"缺失数据点信息已保存到: {os.path.abspath(missing_file)}")
                logger.info(f"共有 {len(missing_points)} 个采集点缺失数据")

            if error_count > 0:
                logger.warning(f"发生 {error_count} 个错误")

        else:
            logger.warning("\n警告: 未获取到有效数据")

    except FileNotFoundError:
        logger.error(f"错误: 找不到输入文件 '{INPUT_FILE}'")
    except Exception as e:
        logger.error(f"处理过程中出错: {str(e)}")


if __name__ == "__main__":
    start_time = time.time()

    print("============ 物联网数据采集脚本 ============")
    print(f"API地址: {API_URL}")
    print(f"输入文件: {INPUT_FILE}")
    print(f"输出文件名模式: {OUTPUT_BASE_NAME}_[timestamp].xlsx")
    print(f"缺失数据文件名模式: {MISSING_BASE_NAME}_[timestamp].xlsx")
    print(f"日志文件: {os.path.abspath(log_filename)}")
    print(f"时间范围: ±{TIME_RANGE} 分钟")
    print(f"最大重试范围: ±{MAX_RETRY_RANGE} 分钟")
    print(f"每批次处理数量: {BATCH_SIZE}")
    print(f"调试模式: {'开启' if DEBUG_MODE else '关闭'}")
    print(f"文件名加时间戳: {'是' if ADD_TIMESTAMP else '否'}")
    print("==========================================")

    fetch_iot_data()
    elapsed_time = time.time() - start_time

    print(f"\n脚本执行完毕! 总耗时: {elapsed_time:.2f}秒")