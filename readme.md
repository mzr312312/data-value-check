# 多基地 IoT 数据批量校验工具 (IoT Data Validator)

## 📖 项目简介 (Introduction)
本项目是一个轻量级的 IoT 数据验证工具，旨在解决原生 IoT 平台在数据治理过程中面临的“批量历史数据回溯难”和“跨系统时间戳不一致”问题。

工具通过封装标准的 **RESTful API** 接口，结合时间滑动窗口算法，实现对指定采集点在特定时间点的历史数据进行自动化抓取与比对。目前已应用于全部12个基地的平台数据迁移测试及日常数据质量核查（Data Quality Check）。

## ✨ 核心功能 (Key Features)
* **批量并发查询**：支持通过 Excel 导入海量“采集点编码”与“时间戳”，自动化完成批量数据检索，替代人工单点查询。
* **RESTful 协议集成**：底层基于标准的 **HTTP/RESTful** 协议（POST Method + JSON Payload）实现，直接与 IoT 平台网关交互，轻量高效，无额外 SDK 依赖。
* **智能时间匹配**：针对 IoT 设备上报延迟问题，内置**滑动窗口算法 (Sliding Window)**，在目标时间戳的容差范围（如 ±20分钟）内自动匹配最近邻（Nearest Neighbor）数据。
* **数据清洗与兼容**：内置多种时间格式解析器，自动处理并标准化异构系统产生的时间戳格式（支持 ISO 8601 及常见变体）。

## 📂 目录结构 (Structure)
```text
IoT-Data-Validator/
├── src/
│   └── fetch_iot_data.py    # 主程序：接口调用与数据处理逻辑
├── config/
│   └── base_urls.txt        # 配置文件：各基地 API Endpoint (此文件线下分发，不上传)
├── data/
│   ├── input/               # 输入目录：存放待校验的点位表格
│   └── output/              # 输出目录：自动生成的校验结果
└── requirements.txt         # 项目依赖库
```

## 🚀 快速开始 (Quick Start)

1. 环境准备
确保 Python 3.8+ 环境，并安装依赖：

Bash
pip install -r requirements.txt

2. 配置接口地址
在 config/base_urls.txt 中配置目标 IoT 平台的 API 地址（格式：基地名=URL）：

Plaintext
ShangHai=[http://10.](http://10.)x.x.x:8080/api/v1/getData
Beijing=[http://10.](http://10.)x.x.x:8080/api/v1/getData

3. 准备数据
在 data/input/target_tags.xlsx 中填入需要校验的数据，需包含以下两列（表头固定）：

采集点编码 (Tag Code)
时间戳 (Timestamp)

4. 运行工具
Bash
python src/fetch_iot_data.py
程序运行后将自动弹出 GUI 界面，选择对应基地和时间偏移量即可开始校验。结果将自动保存至 data/output/ 目录。

## 📝 备注
如遇接口鉴权变更，需修改 src/fetch_iot_data.py 中的 Request Header 部分。