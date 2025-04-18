import sys
import os
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QFileDialog, QMessageBox, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
from googleapiclient.discovery import build
from dotenv import load_dotenv
import json

class YouTubeCollector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("YouTube数据采集器")
        self.setMinimumWidth(600)
        
        # 加载配置
        self.load_config()
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 添加各个模块
        self.setup_credentials_section(layout)
        self.add_separator(layout)
        self.setup_filter_section(layout)
        self.add_separator(layout)
        self.setup_file_section(layout)
        self.add_separator(layout)
        self.setup_action_section(layout)
        
        self.collector_thread = None
        
    def add_separator(self, layout):
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)
        
    def setup_credentials_section(self, layout):
        # API Key配置
        api_group = QVBoxLayout()
        api_label = QLabel("YouTube API Key:")
        api_row = QHBoxLayout()
        self.api_key_input = QLineEdit()
        if self.config.get('api_key'):
            self.api_key_input.setText(self.config['api_key'])
        
        save_api_btn = QPushButton("保存")
        clear_api_btn = QPushButton("清空")
        
        api_row.addWidget(self.api_key_input)
        api_row.addWidget(save_api_btn)
        api_row.addWidget(clear_api_btn)
        
        api_group.addWidget(api_label)
        api_group.addLayout(api_row)
        
        # 代理配置
        proxy_group = QVBoxLayout()
        proxy_label = QLabel("代理地址:")
        proxy_row = QHBoxLayout()
        self.proxy_input = QLineEdit()
        if self.config.get('proxy'):
            self.proxy_input.setText(self.config['proxy'])
            
        save_proxy_btn = QPushButton("保存")
        clear_proxy_btn = QPushButton("清空")
        
        proxy_row.addWidget(self.proxy_input)
        proxy_row.addWidget(save_proxy_btn)
        proxy_row.addWidget(clear_proxy_btn)
        
        proxy_group.addWidget(proxy_label)
        proxy_group.addLayout(proxy_row)
        
        layout.addLayout(api_group)
        layout.addLayout(proxy_group)
        
        # 绑定事件
        save_api_btn.clicked.connect(lambda: self.save_config('api_key', self.api_key_input.text()))
        clear_api_btn.clicked.connect(lambda: self.api_key_input.clear())
        save_proxy_btn.clicked.connect(lambda: self.save_config('proxy', self.proxy_input.text()))
        clear_proxy_btn.clicked.connect(lambda: self.proxy_input.clear())
        
    def setup_filter_section(self, layout):
        filter_layout = QHBoxLayout()
        filter_label = QLabel("收集观看量大于")
        self.view_threshold = QLineEdit()
        self.view_threshold.setFixedWidth(100)
        self.view_threshold.setText("1000")  # 默认值
        view_suffix = QLabel("的视频信息")
        
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.view_threshold)
        filter_layout.addWidget(view_suffix)
        filter_layout.addStretch()
        
        layout.addLayout(filter_layout)
        
    def setup_file_section(self, layout):
        # 频道源选择
        source_layout = QHBoxLayout()
        source_label = QLabel("选择提供频道的Excel文件:")
        self.source_path = QLineEdit()
        self.source_path.setReadOnly(True)
        browse_btn = QPushButton("浏览")
        
        source_layout.addWidget(source_label)
        source_layout.addWidget(self.source_path)
        source_layout.addWidget(browse_btn)
        
        # 保存路径设置
        save_layout = QHBoxLayout()
        save_label = QLabel("保存路径:")
        self.save_path = QLineEdit()
        self.save_path.setReadOnly(True)
        if self.config.get('save_path'):
            self.save_path.setText(self.config['save_path'])
        select_btn = QPushButton("选择")
        
        save_layout.addWidget(save_label)
        save_layout.addWidget(self.save_path)
        save_layout.addWidget(select_btn)
        
        layout.addLayout(source_layout)
        layout.addLayout(save_layout)
        
        # 绑定事件
        browse_btn.clicked.connect(self.browse_source_file)
        select_btn.clicked.connect(self.select_save_path)
        
    def setup_action_section(self, layout):
        action_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始采集")
        self.start_btn.setFixedHeight(40)
        action_layout.addWidget(self.start_btn)
        
        layout.addLayout(action_layout)
        
        # 绑定事件
        self.start_btn.clicked.connect(self.start_collection)
        
    def load_config(self):
        self.config = {}
        if os.path.exists('config.json'):
            try:
                with open('config.json', 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            except:
                pass
                
    def save_config(self, key, value):
        self.config[key] = value
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)
        QMessageBox.information(self, "成功", f"{key}已保存")
        
    def browse_source_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            self.source_path.setText(file_name)
            
    def select_save_path(self):
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "选择保存目录",
            ""
        )
        if dir_path:
            self.save_path.setText(dir_path)
            self.save_config('save_path', dir_path)
            
    def start_collection(self):
        # 验证输入
        if not self.api_key_input.text():
            QMessageBox.warning(self, "错误", "请输入API Key")
            return
            
        if not self.source_path.text():
            QMessageBox.warning(self, "错误", "请选择频道Excel文件")
            return
            
        if not self.save_path.text():
            QMessageBox.warning(self, "错误", "请选择保存路径")
            return
            
        try:
            view_threshold = int(self.view_threshold.text())
        except ValueError:
            QMessageBox.warning(self, "错误", "观看量阈值必须是数字")
            return
            
        # 创建并启动采集线程
        self.collector_thread = CollectorThread(
            api_key=self.api_key_input.text(),
            source_file=self.source_path.text(),
            save_path=self.save_path.text(),
            view_threshold=view_threshold,
            proxy=self.proxy_input.text()
        )
        
        self.collector_thread.progress_signal.connect(self.update_progress)
        self.collector_thread.finished.connect(self.collection_finished)
        
        self.start_btn.setEnabled(False)
        self.collector_thread.start()
        
    def update_progress(self, message):
        QMessageBox.information(self, "进度", message)
        
    def collection_finished(self):
        self.start_btn.setEnabled(True)
        QMessageBox.information(self, "完成", "数据采集已完成")


class CollectorThread(QThread):
    progress_signal = pyqtSignal(str)
    
    def __init__(self, api_key, source_file, save_path, view_threshold, proxy):
        super().__init__()
        self.api_key = api_key
        self.source_file = source_file
        self.save_path = save_path
        self.view_threshold = view_threshold
        self.proxy = proxy
        
    def run(self):
        try:
            # 设置代理
            if self.proxy:
                os.environ['HTTP_PROXY'] = self.proxy
                os.environ['HTTPS_PROXY'] = self.proxy
                
            # 读取频道列表
            df = pd.read_excel(self.source_file)
            if 'channel_id' not in df.columns:
                self.progress_signal.emit("Excel文件必须包含channel_id列")
                return
                
            # 初始化YouTube API
            youtube = build('youtube', 'v3', developerKey=self.api_key)
            
            # 准备数据存储
            all_videos = []
            
            # 遍历频道
            for channel_id in df['channel_id']:
                self.progress_signal.emit(f"正在收集频道 {channel_id} 的数据...")
                
                try:
                    # 获取频道信息
                    channel_response = youtube.channels().list(
                        part='snippet,statistics',
                        id=channel_id
                    ).execute()
                    
                    if not channel_response['items']:
                        continue
                        
                    channel_info = channel_response['items'][0]
                    
                    # 获取频道的上传播放列表ID
                    channel_response = youtube.channels().list(
                        part='contentDetails',
                        id=channel_id
                    ).execute()
                    
                    if not channel_response['items']:
                        self.progress_signal.emit(f"无法获取频道 {channel_id} 的信息")
                        continue
                        
                    # 获取上传播放列表ID
                    uploads_playlist_id = channel_response['items'][0]['contentDetails']['relatedPlaylists']['uploads']
                    next_page_token = None
                    
                    while True:
                        try:
                            playlist_response = youtube.playlistItems().list(
                                part='snippet',
                                playlistId=uploads_playlist_id,
                                maxResults=50,
                                pageToken=next_page_token
                            ).execute()
                            
                            if 'items' not in playlist_response:
                                self.progress_signal.emit(f"频道 {channel_id} 的播放列表为空")
                                break
                            
                            video_ids = [item['snippet']['resourceId']['videoId'] 
                                       for item in playlist_response['items']]
                            
                            # 批量获取视频详细信息
                            videos_response = youtube.videos().list(
                                part='snippet,statistics',
                                id=','.join(video_ids)
                            ).execute()
                            
                        except Exception as e:
                            self.progress_signal.emit(f"获取频道 {channel_id} 的视频列表时出错: {str(e)}")
                            break
                        
                        # 处理视频数据
                        for video in videos_response['items']:
                            try:
                                view_count = int(video['statistics']['viewCount'])
                                if view_count >= self.view_threshold:
                                    video_data = {
                                        '频道ID': channel_id,
                                        '频道名称': channel_info['snippet']['title'],
                                        '频道订阅数': channel_info['statistics']['subscriberCount'],
                                        '视频ID': video['id'],
                                        '视频标题': video['snippet']['title'],
                                        '发布时间': video['snippet']['publishedAt'],
                                        '观看次数': view_count,
                                        '点赞数': video['statistics'].get('likeCount', 0),
                                        '评论数': video['statistics'].get('commentCount', 0)
                                    }
                                    all_videos.append(video_data)
                            except Exception as e:
                                continue
                        
                        next_page_token = playlist_response.get('nextPageToken')
                        if not next_page_token:
                            break
                            
                except Exception as e:
                    self.progress_signal.emit(f"处理频道 {channel_id} 时出错: {str(e)}")
                    continue
            
            # 保存数据
            if all_videos:
                df_videos = pd.DataFrame(all_videos)
                today = datetime.now().strftime('%Y%m%d')
                save_file = os.path.join(self.save_path, f'youtube_data_{today}.xlsx')
                df_videos.to_excel(save_file, index=False)
                self.progress_signal.emit(f"数据已保存到: {save_file}")
            else:
                self.progress_signal.emit("未找到符合条件的视频数据")
                
        except Exception as e:
            self.progress_signal.emit(f"采集过程出错: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = YouTubeCollector()
    window.show()
    sys.exit(app.exec())
