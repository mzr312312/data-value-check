import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from matplotlib import rcParams

# 设置支持中文的字体
rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
rcParams['axes.unicode_minus'] = False    # 解决负号 '-' 显示为方块的问题

# 定义输入和输出路径
output_dir = r".\outputs"
input_file_pattern = os.path.join(output_dir, "多个时间的iot数据_*.xlsx")  # 匹配所有输出文件

# 检查输出目录是否存在
if not os.path.exists(output_dir):
    raise FileNotFoundError(f"输出目录 {output_dir} 不存在，请先运行 detect_anomalies.py 生成数据！")


# 读取所有输出文件
def read_output_files():
    import glob
    file_paths = glob.glob(input_file_pattern)
    if not file_paths:
        raise FileNotFoundError("未找到任何输出文件，请先运行 detect_anomalies.py 生成数据！")

    # 合并所有输出文件到一个 DataFrame
    dfs = []
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)


# 绘制异常曲线
def plot_anomaly_curves(df):
    # 按采集点编码、请求开始时间、请求结束时间分组
    grouped = df.groupby(["采集点编码", "请求开始时间", "请求结束时间"])

    for group_key, group_data in grouped:
        # 提取分组信息
        tag_code, start_time, end_time = group_key

        # 计算异常点时间戳（取请求开始时间和结束时间的中点）
        start_dt = pd.to_datetime(start_time)
        end_dt = pd.to_datetime(end_time)
        anomaly_timestamp = start_dt + (end_dt - start_dt) / 2

        # 提取横坐标和纵坐标
        timestamps = pd.to_datetime(group_data["时间戳"])
        values = group_data["返回值"]

        # 创建绘图
        plt.figure(figsize=(10, 6))
        plt.plot(timestamps, values, marker='o', linestyle='-', color='b', label="返回值")

        # 标注信息
        plt.title(f"采集点编码: {tag_code}", fontsize=14)
        plt.xlabel("时间戳", fontsize=12)
        plt.ylabel("返回值", fontsize=12)
        plt.grid(True)
        plt.legend()

        # 在左上角添加标注
        plt.text(
            0.02, 0.95,
            f"采集点编码: {tag_code}\n异常点时间戳: {anomaly_timestamp.strftime('%Y-%m-%d %H:%M:%S')}",
            transform=plt.gca().transAxes,
            fontsize=10,
            verticalalignment='top',
            bbox=dict(boxstyle="round", facecolor="white", alpha=0.5)
        )

        # 保存图片
        image_name = f"{tag_code}_{anomaly_timestamp.strftime('%Y%m%d%H%M%S')}.png"
        image_path = os.path.join(output_dir, image_name)
        plt.savefig(image_path, dpi=300, bbox_inches='tight')
        plt.close()


# 主程序
if __name__ == "__main__":
    # 读取输出文件
    df_output = read_output_files()

    # 绘制异常曲线
    plot_anomaly_curves(df_output)