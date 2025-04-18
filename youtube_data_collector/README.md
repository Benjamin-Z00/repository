# YouTube数据采集器

这是一个基于Python的YouTube数据采集工具，可以通过YouTube API收集指定频道的视频数据。

## 功能特点

- 通过YouTube API采集频道和视频数据
- 支持代理设置
- 可配置观看量阈值筛选
- 自动将数据保存为Excel格式
- 文件名按照YYYYMMDD格式自动生成

## 使用方式

### 方式一：直接运行可执行文件

1. 从 `dist` 目录下找到并运行 `YouTube数据采集器.exe`

### 方式二：从源代码运行

1. 安装Python依赖：
```bash
pip install -r requirements.txt
```

2. 运行 Python 脚本：
```bash
python youtube_collector.py
```

## 配置说明

1. 获取YouTube API密钥：
- 访问 [Google Cloud Console](https://console.cloud.google.com/)
- 创建新项目或选择现有项目
- 启用YouTube Data API v3
- 创建API密钥

## 使用说明

1. 准备频道数据文件：
- 创建Excel文件，必须包含`channel_id`列
- 在`channel_id`列中填入要采集的YouTube频道ID

2. 运行程序：
```bash
python youtube_collector.py
```

3. 在程序界面中：
- 输入YouTube API Key并保存
- 如需要，配置代理地址
- 设置视频观看量筛选阈值
- 选择频道Excel文件
- 选择数据保存路径
- 点击"开始采集"

## 输出数据

程序会在指定的保存路径生成Excel文件，包含以下字段：
- 频道ID
- 频道名称
- 频道订阅数
- 视频ID
- 视频标题
- 发布时间
- 观看次数
- 点赞数
- 评论数

## 注意事项

1. 请确保API Key的配额足够
2. 大量数据采集可能需要较长时间
3. 建议使用代理以避免IP限制
4. 确保输入文件格式正确
