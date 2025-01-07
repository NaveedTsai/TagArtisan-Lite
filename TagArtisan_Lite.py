#!/usr/bin/env python
# -*- coding: utf-8 -*-
#TagArtisan
"""
TagArtisan Lite - Simple File Tagging Solution
Version: 1.0.0.0
Developed by: NTTech Studio
Copyright © 2025 NTTech Studio. All rights reserved.
Package Identity Name: NTTechStudio.TagArtisanLite
Publisher: CN=9327792A-5810-46DE-9C59-E902D755EB6F
License: MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import json
import shutil
import sys
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, Toplevel, StringVar
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import locale
from languages import LANGUAGES, LANGUAGE_NAMES
import ctypes
from ctypes import windll
import queue
from tkinterdnd2 import DND_FILES, TkinterDnD
import logging
import tempfile
from packaging import version
import traceback
import itertools
import tkinter.colorchooser as colorchooser
from TagDropWindow import TagDropWindow
import random
from ctypes import windll, wintypes

def get_app_data_dir():
    """获取应用数据目录"""
    try:
        # 获取 AppData\Roaming 目录
        roaming = os.path.join(os.getenv('APPDATA'), 'TagArtisan Lite')
        # 确保目录存在
        os.makedirs(roaming, exist_ok=True)
        return roaming
    except Exception as e:
        # 如果出现错误，返回当前目录
        print(f"Error creating app data directory: {e}")
        return os.getcwd()

# 設置日誌
if not getattr(sys, 'frozen', False):
    log_path = os.path.join(get_app_data_dir(), 'TagArtisan Lite.log')
    # 清除旧的日志文件
    if os.path.exists(log_path):
        try:
            os.remove(log_path)
        except:
            pass
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
else:
    logging.basicConfig(
        level=logging.ERROR,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.NullHandler()
        ]
    )
logger = logging.getLogger('TagArtisan')

class AutoUpdater:
    def __init__(self, current_version, app_instance=None):
        self.current_version = current_version
        self.app_instance = app_instance
        # 修改為你的 GitHub 倉庫地址
        self.update_url = "https://api.github.com/repos/naveedtsai/EasyTag_Lite/releases/latest"
        self.temp_dir = tempfile.gettempdir()

    def check_for_updates(self):
        try:
            logger.info(f"Checking for updates, current version: {self.current_version}")
            response = requests.get(self.update_url)
            logger.info(f"Update check response status code: {response.status_code}")
            
            if response.status_code == 200:
                latest_release = response.json()
                latest_version = latest_release['tag_name'].lstrip('v')
                logger.info(f"Found latest version: {latest_version}")
                
                if version.parse(latest_version) > version.parse(self.current_version):
                    logger.info("Found new version, preparing update")
                    return latest_version, latest_release['assets'][0]['browser_download_url']
                else:
                    logger.info("Current version is up to date")
            return None, None
        except Exception as e:
            logger.error(f"Error occurred while checking for updates: {str(e)}")
            return None, None

    def download_update(self, download_url):
        try:
            logger.info(f"Start downloading update: {download_url}")
            response = requests.get(download_url, stream=True)
            if response.status_code == 200:
                update_file = os.path.join(self.temp_dir, "EasyTag_update.msix")
                with open(update_file, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                logger.info(f"Update file downloaded to: {update_file}")
                return update_file
            logger.error(f"Download update failed, status code: {response.status_code}")
            return None
        except Exception as e:
            logger.error(f"Error occurred while downloading update: {str(e)}")
            return None

    def install_update(self, update_file):
        try:
            
            return True
        except Exception as e:
            logger.error(f"Error occurred while installing update: {str(e)}")
            return False

    def check_and_update(self):
        latest_version, download_url = self.check_for_updates()
        if latest_version and self.app_instance:
            # 詢問用戶是否要更新
            if CustomMessageBox.show_question(
                self.app_instance,
                self.app_instance.get_text("update_available"),
                self.app_instance.get_text("update_prompt").format(version=latest_version)
            ):
                update_file = self.download_update(download_url)
                if update_file:
                    if self.install_update(update_file):
                        CustomMessageBox.show_info(
                            self.app_instance,
                            self.app_instance.get_text("success"),
                            self.app_instance.get_text("update_success")
                        )
                        self.app_instance.destroy()
                    else:
                        CustomMessageBox.show_error(
                            self.app_instance,
                            self.app_instance.get_text("error"),
                            self.app_instance.get_text("update_failed")
                        )

class CrashReporter:
    def __init__(self):
        pass
        
    def report_crash(self, error_info):
        """報告崩潰信息"""
        try:
            logger.error(f"Application crashed: {error_info}")
        except Exception as e:
            logger.error(f"Error reporting crash: {str(e)}")

# 自訂的對話框類
class CustomMessageBox:
    def __init__(self, parent, title, message, style="info", width=400, height=200):
        self.parent = parent
        self.result = None

        self.window = Toplevel(parent)
        self.window.title(title)
        self.window.geometry(f"{width}x{height}")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.withdraw()  # 暫時隱藏窗口

        # 信息標籤
        label = tb.Label(self.window, text=message, wraplength=width-20, bootstyle=style)
        label.pack(pady=20, padx=10, expand=True)

        # 按鈕框架
        button_frame = tb.Frame(self.window)
        button_frame.pack(pady=10)

        # 按鈕根據樣式不同而不同
        if style in ["info", "warning", "error"]:
            btn_text = parent.get_text("confirm") if hasattr(parent, 'get_text') else "確定"
            btn_style = {
                "info": "success",
                "warning": "warning",
                "error": "danger"
            }.get(style, "info")
            btn = tb.Button(button_frame, text=btn_text, command=self.ok, bootstyle=btn_style)
            btn.pack()
        elif style == "question":
            yes_text = parent.get_text("yes") if hasattr(parent, 'get_text') else "是"
            no_text = parent.get_text("no") if hasattr(parent, 'get_text') else "否"
            btn_yes = tb.Button(button_frame, text=yes_text, command=self.yes, bootstyle="success")
            btn_yes.pack(side=tk.LEFT, padx=10)
            btn_no = tb.Button(button_frame, text=no_text, command=self.no, bootstyle="danger")
            btn_no.pack(side=tk.RIGHT, padx=10)

        self.window.protocol("WM_DELETE_WINDOW", self.no if style == "question" else self.ok)

        # 將窗口居中
        self.center_window(self.window, width, height, parent)

        self.window.deiconify()  # 顯示窗口
        self.parent.wait_window(self.window)

    def center_window(self, window, width, height, parent=None):
        if parent:
            parent.update_idletasks()
            parent_x = parent.winfo_rootx()
            parent_y = parent.winfo_rooty()
            parent_width = parent.winfo_width()
            parent_height = parent.winfo_height()
        else:
            parent_x = parent_y = 0
            parent_width = window.winfo_screenwidth()
            parent_height = window.winfo_screenheight()

        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def ok(self):
        self.result = True
        self.window.destroy()

    def yes(self):
        self.result = True
        self.window.destroy()

    def no(self):
        self.result = False
        self.window.destroy()

    @staticmethod
    def show_info(parent, title, message, width=400, height=200):
        return CustomMessageBox(parent, title, message, style="info", width=width, height=height).result

    @staticmethod
    def show_warning(parent, title, message, width=400, height=200):
        return CustomMessageBox(parent, title, message, style="warning", width=width, height=height).result

    @staticmethod
    def show_error(parent, title, message, width=400, height=200):
        return CustomMessageBox(parent, title, message, style="error", width=width, height=height).result

    @staticmethod
    def show_question(parent, title, message, width=400, height=200):
        msg_box = CustomMessageBox(parent, title, message, style="question", width=width, height=height)
        return msg_box.result

def create_modal_dialog(parent, title, width, height):
    """
    創建一個模態對話框，支援 DPI 縮放。
    """
    # 根據 DPI 縮放調整視窗大小
    scaled_width = int(width * parent.dpi_factor)
    scaled_height = int(height * parent.dpi_factor)
    
    dialog = Toplevel(parent)
    dialog.withdraw()  # 暫時隱藏
    dialog.title(title)
    dialog.geometry(f"{scaled_width}x{scaled_height}")
    dialog.resizable(False, False)
    dialog.transient(parent)  # 設置為主視窗的子視窗
    dialog.grab_set()  # 模態化
    
    # 確保所有幾何信息都已更新
    dialog.update_idletasks()
    
    # 將彈出視窗居中於主視窗
    parent.center_window(dialog, scaled_width, scaled_height, parent)
    dialog.deiconify()  # 顯示視窗
    return dialog

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, file_manager, callback):
        super().__init__()
        self.file_manager = file_manager
        self.callback = callback
        self.last_event_time = 0
        self.event_delay = 2  # 增加事件延遲到2秒
        self._lock = threading.Lock()
        self._pending_events = set()  # 用於追蹤待處理的事件
        self._event_timer = None

    def on_any_event(self, event):
        # 忽略 .tmp 檔案和隱藏檔案
        if event.src_path.endswith('.tmp') or '/.' in event.src_path:
            return

        current_time = time.time()
        with self._lock:
            # 將事件添加到待處理集合
            self._pending_events.add(event.src_path)
            
            # 如果已經有計時器在運行，取消它
            if self._event_timer:
                self._event_timer.cancel()
            
            # 設置新的計時器
            self._event_timer = threading.Timer(
                2.0,  # 2秒後執行
                self._process_pending_events
            )
            self._event_timer.start()

    def _process_pending_events(self):
        """處理所有待處理的事件"""
        with self._lock:
            if self._pending_events:
                self._pending_events.clear()
                self.callback()
            self._event_timer = None

class FileManager:
    def __init__(self):
        """初始化檔案管理器"""
        self.folder_paths = []
        self.files_tags = {}
        self.db_data = {}
        self.hash_cache = {}
        
        # 获取应用数据目录
        self.app_data_dir = get_app_data_dir()
        
        # 设置数据文件路径
        self.db_file = os.path.join(self.app_data_dir, 'file_tags.json')
        self.backup_dir = os.path.join(self.app_data_dir, 'backups')
        
        self.observers = []  # 初始化observers列表
        self.event_handler = None
        self.tag_colors = {}  # 添加標籤顏色字典
        self.tag_color_timestamps = {}  # 添加標籤顏色修改時間字典
        os.makedirs(self.backup_dir, exist_ok=True)
        self.load_db()
        self.file_monitor = None
        self.monitoring = False
        self.default_color = "#2b3e50"  # 設置預設顏色為 superhero 主題的背景色

    def calculate_file_hash(self, file_path):
        """使用檔案大小和檔案頭尾字節作為檔案的唯一標識"""
        try:
            # 獲取檔案大小
            file_size = os.path.getsize(file_path)
            
            # 如果檔案小於8字節，直接讀取整個檔案內容
            if file_size < 8:
                with open(file_path, 'rb') as f:
                    content = f.read()
                return f"{file_size}_{content.hex()}"
            
            # 讀取檔案頭尾各4個字節
            with open(file_path, 'rb') as f:
                head = f.read(4)
                f.seek(-4, 2)  # 從檔案末尾向前4個字節
                tail = f.read(4)
            
            # 生成唯一標識（使用大小和頭尾字節的組合）
            file_id = f"{file_size}_{head.hex()}_{tail.hex()}"
            
            # 更新緩存
            self.hash_cache[file_path] = file_id
            
            return file_id
        except Exception as e:
            logger.error(f"計算檔案標識時出錯: {str(e)}")
            return None

    def clean_cache(self):
        """清理不存在的檔案的緩存"""
        for file_path in list(self.hash_cache.keys()):
            if not os.path.exists(file_path):
                del self.hash_cache[file_path]

    def set_folder_paths(self, folder_paths):
        """設置要監控的資料夾路徑，更新檔案列表和雜湊值"""
        self.folder_paths = [os.path.abspath(path) for path in folder_paths]
        self.files_tags = {}

        # 掃描所有資料夾中的檔案
        for folder_path in self.folder_paths:
            if not os.path.exists(folder_path):
                continue
            
            for root, _, files in os.walk(folder_path):
                for file in files:
                    full_path = os.path.abspath(os.path.join(root, file))
                    file_name = os.path.basename(full_path)
                    
                    # 檢查資料庫中是否有相同檔名的記錄
                    name_exists = False
                    for db_key, info in self.db_data.items():
                        if info.get("name") == file_name:
                            name_exists = True
                            break
                    
                    # 如果檔案名稱不存在於資料庫中，直接加入檔案列表
                    if not name_exists:
                        self.files_tags[full_path] = {
                            "tags": [],
                            "note": "",
                            "hash": None,
                            "name": file_name
                        }
                        continue
                    
                    current_hash = self.calculate_file_hash(full_path)
                    if not current_hash:
                        continue
                    
                    # 使用檔案名稱和雜湊值組合作為鍵值
                    current_db_key = f"{file_name}_{current_hash}"
                    
                    # 檢查是否存在相同檔名但不同雜湊值的紀錄
                    found_old_record = False
                    for db_key in list(self.db_data.keys()):
                        if file_name in db_key and full_path in self.db_data[db_key]["paths"]:
                            if db_key != current_db_key:  # 雜湊值不同
                                # 從舊紀錄中移除當前路徑
                                self.db_data[db_key]["paths"].remove(full_path)
                                
                                # 如果舊紀錄沒有其他路徑且有標籤或備註，則複製到新紀錄
                                if (not self.db_data[db_key]["paths"] and 
                                    (self.db_data[db_key]["tags"] or self.db_data[db_key]["note"])):
                                    # 建立新紀錄，沿用舊紀錄的標籤和備註
                                    self.db_data[current_db_key] = {
                                        "tags": self.db_data[db_key]["tags"].copy(),
                                        "note": self.db_data[db_key]["note"],
                                        "hash": current_hash,
                                        "paths": [full_path],
                                        "name": file_name
                                    }
                                    # 如果舊紀錄沒有其他路徑，可以刪除
                                    del self.db_data[db_key]
                                else:
                                    # 如果舊紀錄有標籤或備註，建立新紀錄
                                    if self.db_data[db_key]["tags"] or self.db_data[db_key]["note"]:
                                        self.db_data[current_db_key] = {
                                            "tags": self.db_data[db_key]["tags"].copy(),
                                            "note": self.db_data[db_key]["note"],
                                            "hash": current_hash,
                                            "paths": [full_path],
                                            "name": file_name
                                        }
                            
                            found_old_record = True
                            break
                    
                    # 如果沒有找到舊紀錄，則檢查是否有相同雜湊值的紀錄
                    if not found_old_record:
                        if current_db_key in self.db_data:
                            if full_path not in self.db_data[current_db_key]["paths"]:
                                self.db_data[current_db_key]["paths"].append(full_path)
                                self.db_data[current_db_key]["hash"] = current_hash
                                self.db_data[current_db_key]["name"] = file_name
                        else:
                            # 建立新記錄
                            self.db_data[current_db_key] = {
                                "tags": [],
                                "note": "",
                                "hash": current_hash,
                                "paths": [full_path],
                                "name": file_name
                            }
                    
                    # 更新 files_tags
                    if current_db_key in self.db_data:
                        self.files_tags[full_path] = {
                            "tags": self.db_data[current_db_key]["tags"],
                            "note": self.db_data[current_db_key]["note"],
                            "hash": current_hash,
                            "name": file_name
                        }
                    else:
                        self.files_tags[full_path] = {
                            "tags": [],
                            "note": "",
                            "hash": current_hash,
                            "name": file_name
                        }

        # 清理沒有標籤和備註的紀錄
        for key in list(self.db_data.keys()):
            if not self.db_data[key]["tags"] and not self.db_data[key]["note"]:
                del self.db_data[key]

        # 儲存更新後的資料庫
        self.save_db()

        # 在最後加入更新監控的程式碼
        if hasattr(self, 'event_handler') and self.event_handler:
            self.start_monitoring(self.event_handler.callback)

    def load_db(self):
        """載入資料庫，包括標籤顏色信息"""
        if os.path.exists(self.db_file):
            with open(self.db_file, 'r', encoding='utf-8') as file:
                data = json.load(file)
                if isinstance(data, dict) and 'files' in data:
                    self.db_data = data['files']
                    self.tag_colors = data.get('tag_colors', {})
                    self.tag_color_timestamps = data.get('tag_color_timestamps', {})
                else:
                    # 舊版本格式，只有文件數據
                    self.db_data = data
                    self.tag_colors = {}
                    self.tag_color_timestamps = {}
        else:
            self.db_data = {}

    def merge_subfolder_tags(self):
        """合併子資料夾的標籤資訊到上層資料夾"""
        new_data = {}
        
        # 對每個資料夾路徑進行處理
        for folder_path in self.db_data.keys():
            # 將此資料夾的標籤資訊加入到所有上層資料夾
            current_path = folder_path
            while True:
                parent_path = os.path.dirname(current_path)
                # 如果已經到達根目錄或磁碟根目錄，則停止
                if parent_path == current_path or not parent_path:
                    break
                
                # 如果上層資料夾不在資料中，則創建
                if parent_path not in new_data:
                    new_data[parent_path] = {}
                
                # 將當前資料夾的檔案標籤資訊複製到上層資料夾
                for file_path, info in self.db_data[folder_path].items():
                    if file_path not in new_data[parent_path]:
                        new_data[parent_path][file_path] = info.copy()
                    
                current_path = parent_path
            
        # 將新的標籤資合併到原有資料中
        self.db_data.update(new_data)

    def save_db(self):
        """保存資料庫，包括標籤顏色信息"""
        data = {
            'files': self.db_data,
            'tag_colors': self.tag_colors,
            'tag_color_timestamps': self.tag_color_timestamps
        }
        with open(self.db_file, 'w', encoding='utf-8') as file:
            json.dump(data, file, indent=4, ensure_ascii=False)

    def create_restore_point(self):
        """创建数据文件的备份"""
        if os.path.exists(self.db_file):
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            # 只使用文件名而不是完整路径
            backup_filename = f'{timestamp}_file_tags.json'
            backup_file = os.path.join(self.backup_dir, backup_filename)
            shutil.copy2(self.db_file, backup_file)
            self.cleanup_restore_points()

    def cleanup_restore_points(self):
        """清理还原点，只保留最新的50个"""
        backups = []
        
        # 获取所有备份文件并解析其时间
        for backup in os.listdir(self.backup_dir):
            try:
                # 从文件名解析时间 (格式: YYYYMMDD_HHMMSS_file_tags.json)
                if not backup.endswith('_file_tags.json'):
                    continue
                    
                date_str = backup[:8]  # YYYYMMDD
                time_str = backup[9:15]  # HHMMSS
                backup_time = datetime.strptime(f"{date_str}_{time_str}", "%Y%m%d_%H%M%S")
                backups.append((backup, backup_time))
            except Exception:
                continue  # 跳过无法解析时间的文件
        
        # 按时间排序，最新的在前
        backups.sort(key=lambda x: x[1], reverse=True)
        
        # 如果备份数量超过50个，删除多余的
        if len(backups) > 50:
            for backup, _ in backups[50:]:  # 保留前50个，删除其余的
                try:
                    os.remove(os.path.join(self.backup_dir, backup))
                except Exception:
                    continue  # 如果删除失败，继续处理下一个

    def restore_db(self, backup_file):
        """从备份文件恢复数据库"""
        # 如果提供的是相对于backup_dir的文件名，构建完整路径
        if not os.path.isabs(backup_file):
            backup_file = os.path.join(self.backup_dir, backup_file)
            
        if not os.path.exists(backup_file):
            raise FileNotFoundError(f"备份文件不存在: {backup_file}")
            
        shutil.copy2(backup_file, self.db_file)
        self.load_db()

    def get_restore_points(self):
        """获取所有可用的还原点"""
        backups = []
        if os.path.exists(self.backup_dir):
            for backup in os.listdir(self.backup_dir):
                if backup.endswith('_file_tags.json'):
                    backup_path = os.path.join(self.backup_dir, backup)
                    backups.append(backup)
        return sorted(backups, reverse=True)  # 按文件名排序，最新的在前

    def set_folder_paths(self, folder_paths):
        """設置要監控的資料夾路徑，更新檔案列表和雜湊值"""
        self.folder_paths = [os.path.abspath(path) for path in folder_paths]
        self.files_tags = {}

        # 掃描所有資料夾中的檔案
        for folder_path in self.folder_paths:
            if not os.path.exists(folder_path):
                continue
            
            for root, _, files in os.walk(folder_path):
                for file in files:
                    full_path = os.path.abspath(os.path.join(root, file))
                    file_name = os.path.basename(full_path)
                    
                    # 檢查資料庫中是否有相同檔名的記錄
                    name_exists = False
                    for db_key, info in self.db_data.items():
                        if info.get("name") == file_name:
                            name_exists = True
                            break
                    
                    # 如果檔案名稱不存在於資料庫中，直接加入檔案列表
                    if not name_exists:
                        self.files_tags[full_path] = {
                            "tags": [],
                            "note": "",
                            "hash": None,
                            "name": file_name
                        }
                        continue
                    
                    current_hash = self.calculate_file_hash(full_path)
                    if not current_hash:
                        continue
                    
                    # 使用檔案名稱和雜湊值組合作為鍵值
                    current_db_key = f"{file_name}_{current_hash}"
                    
                    # 檢查是否存在相同檔名但不同雜湊值的紀錄
                    found_old_record = False
                    for db_key in list(self.db_data.keys()):
                        if file_name in db_key and full_path in self.db_data[db_key]["paths"]:
                            if db_key != current_db_key:  # 雜湊值不同
                                # 從舊紀錄中移除當前路徑
                                self.db_data[db_key]["paths"].remove(full_path)
                                
                                # 如果舊紀錄沒有其他路徑且有標籤或備註，則複製到新紀錄
                                if (not self.db_data[db_key]["paths"] and 
                                    (self.db_data[db_key]["tags"] or self.db_data[db_key]["note"])):
                                    # 建立新紀錄，沿用舊紀錄的標籤和備註
                                    self.db_data[current_db_key] = {
                                        "tags": self.db_data[db_key]["tags"].copy(),
                                        "note": self.db_data[db_key]["note"],
                                        "hash": current_hash,
                                        "paths": [full_path],
                                        "name": file_name
                                    }
                                    # 如果舊紀錄沒有其他路徑，可以刪除
                                    del self.db_data[db_key]
                                else:
                                    # 如果舊紀錄有標籤或備註，建立新紀錄
                                    if self.db_data[db_key]["tags"] or self.db_data[db_key]["note"]:
                                        self.db_data[current_db_key] = {
                                            "tags": self.db_data[db_key]["tags"].copy(),
                                            "note": self.db_data[db_key]["note"],
                                            "hash": current_hash,
                                            "paths": [full_path],
                                            "name": file_name
                                        }
                            
                            found_old_record = True
                            break
                    
                    # 如果沒有找到舊紀錄，則檢查是否有相同雜湊值的紀錄
                    if not found_old_record:
                        if current_db_key in self.db_data:
                            if full_path not in self.db_data[current_db_key]["paths"]:
                                self.db_data[current_db_key]["paths"].append(full_path)
                                self.db_data[current_db_key]["hash"] = current_hash
                                self.db_data[current_db_key]["name"] = file_name
                        else:
                            # 建立新記錄
                            self.db_data[current_db_key] = {
                                "tags": [],
                                "note": "",
                                "hash": current_hash,
                                "paths": [full_path],
                                "name": file_name
                            }
                    
                    # 更新 files_tags
                    if current_db_key in self.db_data:
                        self.files_tags[full_path] = {
                            "tags": self.db_data[current_db_key]["tags"],
                            "note": self.db_data[current_db_key]["note"],
                            "hash": current_hash,
                            "name": file_name
                        }
                    else:
                        self.files_tags[full_path] = {
                            "tags": [],
                            "note": "",
                            "hash": current_hash,
                            "name": file_name
                        }

        # 清理沒有標籤和備註的紀錄
        for key in list(self.db_data.keys()):
            if not self.db_data[key]["tags"] and not self.db_data[key]["note"]:
                del self.db_data[key]

        # 儲存更新後的資料庫
        self.save_db()

        # 在最後加入更新監控的程式碼
        if hasattr(self, 'event_handler') and self.event_handler:
            self.start_monitoring(self.event_handler.callback)

    def refresh_db(self):
        """只清理已刪除的檔案記錄"""
        if not self.folder_paths:
            return

        # 檢查資料庫中的檔案是
        for file_path in list(self.db_data.keys()):
            if not os.path.exists(file_path):
                del self.db_data[file_path]

        self.save_db()

    def add_tag(self, file_path, tag):
        if tag.strip():
            file_name = os.path.basename(file_path)
            current_hash = self.calculate_file_hash(file_path)
            if not current_hash:
                return
            
            db_key = f"{file_name}_{current_hash}"
            
            # 確保檔案記錄存在
            if file_path not in self.files_tags:
                self.files_tags[file_path] = {
                    "tags": [],
                    "note": "",
                    "hash": current_hash,
                    "name": file_name  # 確保記錄檔案名稱
                }
            
            # 新增標籤
            if tag not in self.files_tags[file_path]["tags"]:
                self.files_tags[file_path]["tags"].append(tag)
                
                # 更新資料庫
                if db_key not in self.db_data:
                    self.db_data[db_key] = {
                        "tags": [tag],
                        "note": "",
                        "hash": current_hash,
                        "paths": [file_path],
                        "name": file_name  # 確保記錄檔案名稱
                    }
                else:
                    if tag not in self.db_data[db_key]["tags"]:
                        self.db_data[db_key]["tags"].append(tag)
                    # 確保雜湊值和檔案名稱是最新的
                    self.db_data[db_key]["hash"] = current_hash
                    self.db_data[db_key]["name"] = file_name
                
                self.save_db()

    def remove_tag(self, file_path, tag):
        file_name = os.path.basename(file_path)
        current_hash = self.calculate_file_hash(file_path)
        if not current_hash:
            return
        
        db_key = f"{file_name}_{current_hash}"
        
        # 從 files_tags 中移除標籤
        if file_path in self.files_tags and tag in self.files_tags[file_path]["tags"]:
            self.files_tags[file_path]["tags"].remove(tag)
            
            # 從 db_data 中移除標籤
            if db_key in self.db_data and tag in self.db_data[db_key]["tags"]:
                self.db_data[db_key]["tags"].remove(tag)
                
                # 如果檔案沒有任何標籤和備註，從資料庫中移除
                if not self.db_data[db_key]["tags"] and not self.db_data[db_key]["note"]:
                    del self.db_data[db_key]
            
            self.save_db()

    def set_note(self, file_path, note):
        if len(note) > 500:
            note = note[:500]
        
        file_name = os.path.basename(file_path)
        current_hash = self.calculate_file_hash(file_path)
        if not current_hash:
            return
        
        db_key = f"{file_name}_{current_hash}"
        
        # 確保檔案記錄存在
        if file_path not in self.files_tags:
            self.files_tags[file_path] = {
                "tags": [], 
                "note": "",
                "hash": current_hash,
                "name": file_name  # 確保記錄檔案名稱
            }
        
        # 設置備註
        self.files_tags[file_path]["note"] = note
        self.files_tags[file_path]["name"] = file_name  # 更新檔案名稱
        
        # 更新資料庫
        if note or (db_key in self.db_data and self.db_data[db_key]["tags"]):
            if db_key not in self.db_data:
                self.db_data[db_key] = {
                    "tags": [],
                    "note": note,
                    "hash": current_hash,
                    "paths": [file_path],
                    "name": file_name  # 確保記錄檔案名稱
                }
            else:
                self.db_data[db_key]["note"] = note
                self.db_data[db_key]["name"] = file_name  # 更新檔案名稱
        elif db_key in self.db_data and not self.db_data[db_key]["tags"]:
            # 如果沒有備註也沒有標籤，從資料庫中移除
            del self.db_data[db_key]
        
        self.save_db()

    def get_note(self, file_path):
        file_name = os.path.basename(file_path)
        current_hash = self.calculate_file_hash(file_path)
        if not current_hash:
            return ""
        
        db_key = f"{file_name}_{current_hash}"
        return self.db_data.get(db_key, {}).get("note", "")

    def search_by_tags(self, tags):
        if not tags:
            # 返回所有檔案的全路徑，使用字典的副本
            return sorted(list(self.files_tags.keys()))
            
        # 將標籤字串分割成列表
        tags_list = [tag.strip() for tag in tags.split(',')]
        
        # 使用字典的副本進行遍歷
        files_tags_copy = self.files_tags.copy()
        
        # 找出同時包含所有指定標籤的檔案
        matching_files = []
        for file_path, info in files_tags_copy.items():
            file_tags = set(info.get("tags", []))
            # 檢查檔案是否包含所有指定的標籤
            if all(tag in file_tags for tag in tags_list):
                matching_files.append(file_path)
        
        return sorted(matching_files)

    def list_untagged_files(self):
        """返回沒有標籤的檔案的全路徑"""
        untagged_files = []
        # 使用字典的副本進行遍歷
        files_tags_copy = self.files_tags.copy()
        for file_path, info in files_tags_copy.items():
            # 確保 tags 存在且為空列表
            tags = info.get("tags", [])
            if not tags:
                untagged_files.append(file_path)
        
        return sorted(untagged_files)

    def get_all_used_tags(self):
        all_tags = set()
        # 使用字典的副本進行遍歷
        files_tags_copy = self.files_tags.copy()
        for info in files_tags_copy.values():
            all_tags.update(info.get("tags", []))
        return sorted(all_tags)

    def get_all_file_types(self):
        """返回所有檔案的擴展名，已排序並去重"""
        file_types = set()
        # 使用字典的副本進行遍歷
        files_tags_copy = self.files_tags.copy()
        for file_path in files_tags_copy.keys():
            _, ext = os.path.splitext(file_path)
            if ext:
                file_types.add(ext.lower())
        return sorted(file_types)

    def rename_tag(self, old_tag, new_tag, merge=False):
        # 保存旧标签的颜色
        old_tag_color = self.get_tag_color(old_tag)
        old_tag_color_timestamp = self.get_tag_color_timestamp(old_tag)
        
        if merge:
            for file_path, info in self.files_tags.items():
                if old_tag in info["tags"]:
                    index = info["tags"].index(old_tag)
                    info["tags"][index] = new_tag
                    if new_tag not in info["tags"]:
                        self.db_data[file_path]["tags"].append(new_tag)
        else:
            for file_path, info in self.files_tags.items():
                if old_tag in info["tags"]:
                    index = info["tags"].index(old_tag)
                    info["tags"][index] = new_tag
                    # 同步更新 db_data
                    if file_path in self.db_data:
                        tags = self.db_data[file_path]["tags"]
                        if old_tag in tags:
                            index = tags.index(old_tag)
                            tags[index] = new_tag
        
        # 如果旧标签有自定义颜色，将其应用到新标签
        if old_tag_color != self.default_color:
            self.set_tag_color(new_tag, old_tag_color)
            # 保持原有的时间戳
            if old_tag_color_timestamp:
                self.tag_color_timestamps[new_tag] = old_tag_color_timestamp
        
        # 删除旧标签的颜色设置
        if old_tag in self.tag_colors:
            del self.tag_colors[old_tag]
        if old_tag in self.tag_color_timestamps:
            del self.tag_color_timestamps[old_tag]
        
        self.save_db()

    def delete_tag(self, tag):
        """從所有檔案中刪除指定的標籤"""
        # 先從 files_tags 中移除標籤
        for file_path, info in self.files_tags.items():
            if tag in info["tags"]:
                info["tags"].remove(tag)
                # 同步更新 db_data
                if file_path in self.db_data:
                    if tag in self.db_data[file_path]["tags"]:
                        self.db_data[file_path]["tags"].remove(tag)
                    # 如果檔案沒有任何標籤和備註，從資料庫中移除
                    if not self.db_data[file_path]["tags"] and not self.db_data[file_path]["note"]:
                        del self.db_data[file_path]
        
        self.save_db()

    # Batch operations
    def add_tags_batch(self, filenames, tags):
        for filename in filenames:
            for tag in tags:
                if tag.strip():
                    self.add_tag(filename, tag)

    def remove_tags_batch(self, filenames, tags):
        for filename in filenames:
            for tag in tags:
                self.remove_tag(filename, tag)

    def start_monitoring(self, callback):
        """開始監控所有資料夾"""
        # 停止現有的監控
        self.stop_monitoring()
        
        # 創建新的事件處理器
        self.event_handler = FileChangeHandler(self, callback)
        
        # 為每個資料夾創建觀察者
        for folder_path in self.folder_paths:
            if os.path.exists(folder_path):
                observer = Observer()
                observer.schedule(self.event_handler, folder_path, recursive=True)
                observer.start()
                self.observers.append(observer)
    
    def stop_monitoring(self):
        """停止所有資料夾的監控"""
        if hasattr(self, 'observers'):
            for observer in self.observers:
                observer.stop()
            for observer in self.observers:
                observer.join()
            self.observers.clear()

    def set_tag_color(self, tag, color):
        """設置標籤的顏色"""
        self.tag_colors[tag] = color
        self.tag_color_timestamps[tag] = datetime.now().timestamp()
        self.save_db()

    def get_tag_color(self, tag):
        """獲取標籤的顏色，如果沒有設置則返回預設顏色"""
        return self.tag_colors.get(tag, self.default_color)  # 使用預設顏色

    def get_tag_color_timestamp(self, tag):
        """獲取標籤顏色的修改時間，如果沒有則返回0"""
        return self.tag_color_timestamps.get(tag, 0)

    def set_default_color(self, color):
        """設置預設標籤顏色"""
        self.default_color = color
        self.save_db()

class BackgroundUpdateManager:
    def __init__(self, app):
        self.app = app
        self.update_queue = queue.Queue()
        self.running = True
        self.is_updating = False
        self.update_lock = threading.Lock()
        self.last_update_time = time.time()
        self.update_interval = 10.0  # 設置更新間隔為10秒
        self.update_thread = threading.Thread(target=self._process_updates, daemon=True)
        self.update_thread.start()
        
    def _process_updates(self):
        """處理更新隊列中的任務"""
        while self.running:
            try:
                if not self.update_queue.empty():
                    current_time = time.time()
                    if current_time - self.last_update_time < self.update_interval:
                        time.sleep(1)  # 增加休眠時間
                        continue

                    with self.update_lock:
                        if self.is_updating:
                            time.sleep(1)  # 增加休眠時間
                            continue
                        
                        self.is_updating = True
                        
                    try:
                        task_type, args = self.update_queue.get()
                        if task_type == "scan_files":
                            self._scan_files(*args)
                        elif task_type == "refresh_tags":
                            self._refresh_tags()
                    finally:
                        self.is_updating = False
                        self.last_update_time = time.time()
                else:
                    time.sleep(1)  # 增加休眠時間
            except Exception as e:
                logger.error(f"處理更新時出錯: {str(e)}")
                time.sleep(1)
    
    def _scan_files(self, folder_paths):
        """掃描檔案並更新UI"""
        try:
            self.app.file_manager.set_folder_paths(folder_paths)
            self.app.after(100, self._update_ui)
        except Exception as e:
            logger.error(f"掃描檔案時出錯: {str(e)}")
    
    def _update_ui(self):
        """更新UI元素"""
        try:
            if self.is_updating:
                return
            
            # 保存當前選中的檔案
            selected_items = self.app.file_tree.selection()
            self.app.selected_paths = []
            for item in selected_items:
                try:
                    file_path = self.app.file_tree.item(item, 'values')[0]
                    self.app.selected_paths.append(file_path)
                except:
                    continue
                
            # 批量更新UI
            self.app.refresh_tag_list()
            self.app.filter_file_list()  # 改用 filter_file_list 來更新檔案列表
            self.app.update_edit_buttons_state()
        except Exception as e:
            logger.error(f"更新UI時出錯: {str(e)}")
    
    def _refresh_tags(self):
        """刷新標籤"""
        try:
            with self.update_lock:
                if self.is_updating:
                    return
                self.app.file_manager.load_db()
                self.app.after(100, self.app.refresh_tag_list)
        except Exception as e:
            logger.error(f"刷新標籤時出錯: {str(e)}")
    
    def queue_update(self, task_type, *args):
        """將更新任務添加到隊列"""
        # 如果隊列中已有相同類型的任務，不再添加
        if self.update_queue.qsize() > 0:
            return
        self.update_queue.put((task_type, args))
    
    def stop(self):
        """停止更新管理器"""
        self.running = False
        if self.update_thread.is_alive():
            self.update_thread.join(timeout=1)

    def set_updating(self, updating):
        """設置更新狀態"""
        self.is_updating = updating

    def load_files_in_batches(self):
        """分批載入檔案到Treeview中以保持UI流暢"""
        if self.is_loading_cancelled:
            return

        if not hasattr(self, 'all_files') or self.all_files is None:
            self.all_files = []
            return

        start = self.current_batch * self.load_batch_size
        end = min(start + self.load_batch_size, len(self.all_files))
        batch = self.all_files[start:end]

        # 如果是第一批，清空現有項目
        if self.current_batch == 0:
            self.file_tree.delete(*self.file_tree.get_children())

        # 禁用重繪以提高性能
        self.file_tree.configure(height=1)  # 臨時減少高度以加快插入
        
        # 批量插入檔案
        items_to_insert = []
        for full_path in batch:
            file_name = os.path.basename(full_path)
            items_to_insert.append(('', 'end', file_name, (full_path,)))

        # 使用批量插入
        for item in items_to_insert:
            item_id = self.file_tree.insert(item[0], item[1], text=item[2], values=item[3])
            if hasattr(self, 'selected_paths') and item[3][0] in self.selected_paths:
                self.file_tree.selection_add(item_id)

        # 恢復正常高度
        self.file_tree.configure(height=20)  # 或其他適當的高度

        self.current_batch += 1
        
        # 如果還有更多檔案要載入，安排下一批
        if end < len(self.all_files) and not self.is_loading_cancelled:
            self.after(100, self.load_files_in_batches)  # 增加延遲時間到100ms以減少UI凍結

class SplashScreen(tk.Toplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.is_closing = False  # 添加標記，避免重複關閉
        self.is_destroyed = False  # 添加標記，避免重複執行destroy後的操作
        
        # 設置窗口屬性
        self.overrideredirect(True)  # 移除標題欄和邊框
        self.attributes('-alpha', 0.0)  # 初始透明度為0
        self.attributes('-topmost', True)  # 保持在最上層
        
        # 獲取螢幕尺寸和DPI縮放因子
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 獲取DPI縮放因子
        try:
            import ctypes
            user32 = ctypes.windll.user32
            user32.SetProcessDPIAware()
            dpi = user32.GetDpiForSystem()
            dpi_factor = dpi / 96.0
        except:
            dpi_factor = self.winfo_fpixels('1i') / 72.0  # 備用方法
        
        # 設置啟動畫面大小（考慮DPI縮放）
        base_width = 600
        base_height = 400
        width = int(base_width * dpi_factor)
        height = int(base_height * dpi_factor)
        
        # 確保窗口不會太大
        max_width = int(screen_width * 0.8)
        max_height = int(screen_height * 0.8)
        width = min(width, max_width)
        height = min(height, max_height)
        
        # 設置窗口位置
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # 設置窗口位置和大小
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # 創建主框架
        self.frame = tk.Frame(self, bg='#1c1c1c', width=width, height=height)
        self.frame.pack(fill=tk.BOTH, expand=True)
        self.frame.pack_propagate(False)
        
        # 計算字體大小（考慮DPI縮放）
        title_font_size = int(36 * dpi_factor)
        version_font_size = int(14 * dpi_factor)
        loading_font_size = int(14 * dpi_factor)
        
        # 添加標題
        title_label = tk.Label(
            self.frame,
            text=f"{self.parent.app_name}",
            font=("Segoe UI", title_font_size, "bold"),
            fg='white',
            bg='#1c1c1c'
        )
        title_label.pack(pady=(int(80 * dpi_factor), int(10 * dpi_factor)))
        
        # 添加版本號
        version_label = tk.Label(
            self.frame,
            text=f"{self.parent.get_text('version')} {self.parent.version}",
            font=("Segoe UI", version_font_size),
            fg='#888888',
            bg='#1c1c1c'
        )
        version_label.pack(pady=5)
        
        # 添加載入進度條
        self.progress = ttk.Progressbar(
            self.frame,
            length=int(400 * dpi_factor),
            mode='determinate'
        )
        self.progress.pack(pady=int(50 * dpi_factor))
        
        # 添加載入文字
        self.loading_label = tk.Label(
            self.frame,
            text=self.parent.get_text('loading'),
            font=("Segoe UI", loading_font_size),
            fg='#888888',
            bg='#1c1c1c'
        )
        self.loading_label.pack(pady=10)
        
        # 設置進度條樣式
        style = ttk.Style()
        style.configure(
            "TProgressbar",
            thickness=int(10 * dpi_factor),
            troughcolor='#2d2d2d',
            background='#007acc'
        )
        
        # 淡入效果
        self.fade_in()
        
        # 開始進度更新
        self.progress_value = 0
        self.update_progress()
    
    def fade_in(self):
        """實現淡入效果"""
        alpha = self.attributes('-alpha')
        if alpha < 1.0:
            alpha += 0.1
            self.attributes('-alpha', alpha)
            self.after(10, self.fade_in)
    
    def update_progress(self):
        """更新進度條"""
        #logger.info("更新進度條: 當前進度=%s", self.progress_value)
        if self.progress_value < 100:
            self.progress_value += 1
            self.progress['value'] = self.progress_value
            self.after(10, self.update_progress)
        elif not self.is_destroyed and not self.is_closing:
            #logger.info("進度條已達100%，等待初始化完成")
            self.check_initialization()

    def check_initialization(self):
        """檢查初始化狀態"""
        if not self.is_closing and not self.is_destroyed:
            if (hasattr(self.parent, 'background_initializer') and 
                self.parent.background_initializer.initialization_complete.is_set() and
                self.parent.background_initializer.ui_update_complete.is_set()):
                logger.info("初始化已完成，準備關閉啟動畫面")
                self.close_splash_screen()
            else:
                #logger.info("等待初始化完成...")
                self.after(100, self.check_initialization)

    def close_splash_screen(self):
        """關閉啟動畫"""
        if not self.is_closing and not self.is_destroyed:
            logger.info("開始關閉啟動畫面")
            self.is_closing = True
            self.parent.deiconify()  # 顯示主視窗
            self.parent.update()
            self.parent.lift()
            self.parent.focus_force()
            logger.info("主視窗已顯示，準備銷毀啟動畫面")
            self.destroy()
            logger.info("啟動畫面已銷毀")

    def wait_for_initialization(self, app):
        """等待初始化完成"""
        #logger.info("等待初始化: initialization_complete=%s, ui_update_complete=%s, progress=%s", 
        #           app.background_initializer.initialization_complete.is_set(),
        #           app.background_initializer.ui_update_complete.is_set(),
        #           self.progress_value)
                   
        if not self.is_closing and not self.is_destroyed:
            if app.background_initializer.initialization_complete.is_set():
                logger.info("初始化完成標記已設置")
                if app.background_initializer.ui_update_complete.is_set():
                    logger.info("UI更新完成標記已設置")
                    if self.progress_value >= 100:
                        self.close_splash_screen()
                    else:
                        self.after(100, lambda: self.wait_for_initialization(app))
                else:
                    logger.info("UI更新尚未完成，等待中...")
                    self.after(100, lambda: self.wait_for_initialization(app))
            else:
                #logger.info("初始化尚未完成，等待中...")
                self.after(100, lambda: self.wait_for_initialization(app))
    
    def finish_initialization(self, app):
        """完成初始化，關閉啟動畫面並顯示主視窗"""
        if not self.is_destroyed:
            self.is_destroyed = True
            self.destroy()  # 關閉啟動畫面

class BackgroundInitializer:
    def __init__(self, app):
        self.app = app
        self.initialization_complete = threading.Event()
        self.ui_update_complete = threading.Event()  # 添加UI更新完成標記
        self.initialization_thread = None
        self.ui_queue = queue.Queue()
        self._process_queue()

    def start(self):
        """開始背景初始化"""
        self.initialization_thread = threading.Thread(target=self._initialize, daemon=True)
        self.initialization_thread.start()

    def _initialize(self):
        """執行耗時的初始化操作"""
        try:
            # 1. 載入資料庫
            self.app.file_manager.load_db()

            # 2. 掃描檔案
            if self.app.last_folders:
                self.app.file_manager.set_folder_paths(self.app.last_folders)

            # 3. 將UI新任務加入隊列
            self.ui_queue.put(self._update_ui)

            # 4. 將檢查更新任務加入隊列
            #self.ui_queue.put(lambda: self.app.after(1000, self.app.check_for_updates))

        except Exception as e:
            logger.error(f"背景初始化失敗: {str(e)}")
            self.ui_queue.put(lambda: CustomMessageBox.show_error(
                self.app,
                self.app.get_text("error"),
                self.app.get_text("unexpected_error")
            ))
        finally:
            self.initialization_complete.set()

    def _process_queue(self):
        """在主線程中處理UI更新隊列"""
        try:
            while True:
                # 非阻塞方式獲取任務
                try:
                    task = self.ui_queue.get_nowait()
                    task()
                    self.ui_queue.task_done()
                except queue.Empty:
                    break
        finally:
            # 每100ms檢查一次隊列
            self.app.after(100, self._process_queue)

    def _update_ui(self):
        """更新UI元素"""
        try:
            # 更新標籤列表
            self.app.refresh_tag_list()
            
            # 讀取上次選擇的標籤
            try:
                with open(self.app.config_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                    last_selected_tags = config.get('last_selected_tags', [])
                    
                    # 檢查這些標籤是否仍然存在
                    all_tags = self.app.file_manager.get_all_used_tags()
                    valid_tags = [tag for tag in last_selected_tags if tag in all_tags]
                    
                    if valid_tags:
                        # 選中這些標籤
                        for tag in valid_tags:
                            for item in self.app.tag_list.get_children():
                                if self.app.tag_list.item(item, 'text') == tag:
                                    self.app.tag_list.selection_add(item)
                        
                        # 更新檔案列表
                        self.app.update_file_list_by_tag_selection()
                        
                        # 更新檔案類型選單
                        self.app.update_file_type_options()
                        
                        # 恢復上次選擇的檔案類型
                        last_file_type = config.get('last_file_type', "ALL")
                        if last_file_type in self.app.file_type_combo['values']:
                            self.app.file_type_var.set(last_file_type)
                        else:
                            self.app.file_type_var.set(self.app.get_text("all_types"))  # 使用翻譯後的文字
                    else:
                        # 如果沒有有效的標籤，則顯示所有檔案
                        self.app.list_untagged_files()
                        self.app.filter_file_list()
            except Exception as e:
                logger.error(f"載入上次選擇的標籤時發生錯誤: {str(e)}")
                self.app.filter_file_list()
            
            # 更新按鈕狀態
            self.app.update_edit_buttons_state()
            
            # 開始檔案監控
            self.app.start_file_monitoring()
            
            # 確保語言已正確載入並更新UI
            self.app.update_ui_language()
            
            # 確保所有UI更新都已完成
            self.app.update_idletasks()
            
            # 檢查是否有選擇的資料夾，如果沒有則自動打開資料夾選擇視窗
            if not self.app.file_manager.folder_paths:
                self.app.after(500, self.app.open_manage_folders_dialog)
            
            # 設置UI更新完成標記
            self.ui_update_complete.set()
            
        except Exception as e:
            logger.error(f"UI更新失敗: {str(e)}")
            self.ui_update_complete.set()  # 即使發生錯誤也要設置完成標記

    def wait_for_completion(self, timeout=None):
        """等待初始化完成"""
        return self.initialization_complete.wait(timeout)

def open_file_in_explorer(file_path):
    """使用 Windows API 在檔案總管中打開並選中檔案"""
    if not os.path.exists(file_path):
        return False
        
    CSIDL_DESKTOP = 0
    SHGFP_TYPE_CURRENT = 0
    
    _SHParseDisplayName = windll.shell32.SHParseDisplayName
    _SHParseDisplayName.argtypes = [wintypes.LPCWSTR, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p), ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong)]
    
    _SHOpenFolderAndSelectItems = windll.shell32.SHOpenFolderAndSelectItems
    _SHOpenFolderAndSelectItems.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p), ctypes.c_ulong]
    
    pidl = ctypes.c_void_p()
    sfgao = ctypes.c_ulong()
    
    if _SHParseDisplayName(file_path, None, ctypes.byref(pidl), 0, ctypes.byref(sfgao)) == 0:
        _SHOpenFolderAndSelectItems(pidl, 0, None, 0)
        windll.ole32.CoTaskMemFree(pidl)
        return True
    return False

class Application(TkinterDnD.Tk):
    def __init__(self, file_manager, config_file=None):
        # 设置DPI感知
        try:
            windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        except Exception:
            try:
                # Windows 8.1 及更早版本
                windll.user32.SetProcessDPIAware()
            except Exception:
                pass
                
        # 添加版本信息
        self.version = "1.0.0.0"
        self.app_name = "TagArtisan Lite"
        
        # 获取应用数据目录
        self.app_data_dir = get_app_data_dir()
        
        # 设置配置文件路径
        if config_file is None:
            config_file = os.path.join(self.app_data_dir, 'config.json')
        
        # 初始化崩潰報告器
        self.crash_reporter = CrashReporter()
        
        # 設置異常處理
        sys.excepthook = self.handle_exception
        
        # 1. 初始化基本配置
        self.config_file = config_file
        self.file_manager = file_manager
        
        # 創建備份並清理舊備份
        logger.info("正在創建標籤數據備份...")
        self.file_manager.create_restore_point()
        logger.info("正在清理舊備份...")
        self.file_manager.cleanup_restore_points()
        
        self.dpi_factor = 1.0
        self.last_folders = []
        self.last_file_type = "ALL"
        self.current_language = 'en_US'  # 設置預設語言
        self.current_theme = 'superhero'  # 設置預設主題
        
        # 2. 在 Windows 系統上啟用 DPI 感知
        if os.name == 'nt':  # Windows 系統
            try:
                windll.shcore.SetProcessDpiAwareness(2)  # PMv2
            except Exception:
                try:
                    windll.shcore.SetProcessDpiAwareness(1)  # PM
                except Exception:
                    windll.user32.SetProcessDPIAware()  # System

            # 獲取當前螢幕的 DPI 因子
            try:
                self.dpi_factor = windll.shcore.GetScaleFactorForDevice(0) / 100
            except Exception:
                self.dpi_factor = 1.0

        # 3. 初始化語言設置
        self.load_config()
            
        # 4. 載入配置
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                    saved_language = config.get('language')
                    if saved_language and saved_language in LANGUAGES:
                        self.current_language = saved_language
                    saved_theme = config.get('theme')
                    if saved_theme:
                        self.current_theme = saved_theme
                    self.last_folders = config.get('last_folders', [])
                    self.last_file_type = config.get('last_file_type', "ALL")
                    # 載入預設標籤顏色
                    saved_default_color = config.get('default_tag_color')
                    if saved_default_color:
                        self.file_manager.default_color = saved_default_color
                    # 載入快捷鍵設置
                    hotkey_config = config.get('hotkeys', {})
                    self.hotkey_modifier = hotkey_config.get('modifier', 'Ctrl')
                    self.hotkey_key = hotkey_config.get('key', '1')
        except Exception as e:
            print(f"Error loading configuration file: {str(e)}")
            
        # 5. 初始化主窗口（必須在置完基本配置後）
        super().__init__()
        
        # 6. 設置 ttkbootstrap 主題
        self.style = tb.Style(theme=self.current_theme)
        self.style.theme_use(self.current_theme)

        # 設置全局字體大小
        default_font = ('TkDefaultFont', 14)
        text_font = ('TkTextFont', 14)
        fixed_font = ('TkFixedFont', 14)
        
        self.option_add('*Font', default_font)
        self.option_add('*Text.font', text_font)  # 使用option_add來設置Text部件的字體
        self.style.configure('.', font=default_font)
        self.style.configure('Treeview', font=default_font)
        self.style.configure('Treeview.Heading', font=default_font)
        self.style.configure('TCombobox', font=default_font)
        self.style.configure('TButton', font=default_font)
        self.style.configure('TLabel', font=default_font)
        self.style.configure('TEntry', font=default_font)
        self.style.configure('TCheckbutton', font=default_font)
        self.style.configure('TRadiobutton', font=default_font)
        self.style.configure('TNotebook.Tab', font=default_font)
        self.style.configure('TMenubutton', font=default_font)
        self.style.configure('TSpinbox', font=default_font)

        # 確保預設標籤顏色與當前主題背景一致
        theme_colors = {
            'superhero': '#2b3e50',
            'darkly': '#222222',
            'cyborg': '#060606',
            'vapor': '#0d0d2b',
            'solar': '#002b36',
            'slate': '#272b30',
            'morph': '#373a3c',
            'journal': '#f7f7f7',
            'litera': '#f8f9fa',
            'lumen': '#ffffff',
            'minty': '#fff',
            'pulse': '#fff',
            'sandstone': '#fff',
            'united': '#fff',
            'yeti': '#fff',
            'cosmo': '#fff',
            'flatly': '#fff',
            'spacelab': '#fff',
            'default': '#ffffff'
        }
        self.file_manager.default_color = theme_colors.get(self.current_theme, '#ffffff')

        # 計置主視窗標題為軟體名稱
        self.title(f"{self.app_name} v{self.version}")

        # 處理圖標路徑
        if getattr(sys, 'frozen', False):
            # 如果是打包後的執行檔
            application_path = sys._MEIPASS
        else:
            # 如果是直接運行 Python 腳本
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        icon_path = os.path.join(application_path, "TagArtisan Lite.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(default=icon_path)

        # 計算初始窗口大小
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 讀取上次的窗口位置和大小
        window_config = {}
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                    window_config = config.get('window', {})
        except Exception:
            pass

        if window_config and all(key in window_config for key in ['width', 'height', 'x', 'y']):
            # 使用保存的窗口大小和位置
            width = window_config['width']
            height = window_config['height']
            x = window_config['x']
            y = window_config['y']
            
            # 確保窗口位置在螢幕範圍內
            if x < 0 or x > screen_width - 100:
                x = (screen_width - width) // 2
            if y < 0 or y > screen_height - 100:
                y = (screen_height - height) // 2
                
            self.geometry(f"{width}x{height}+{x}+{y}")
        else:
            # 首次開啟時，設置為螢幕寬度的 2/3，保持 16:9 比例
            width = int((screen_width * 2 / 3) * self.dpi_factor)
            height = int(width * 9 / 16)
            
            # 確保窗口不會太大
            if height > screen_height * 0.8:
                height = int(screen_height * 0.8)
                width = int(height * 16 / 9)
            
            # 計算置中位置
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            self.geometry(f"{width}x{height}+{x}+{y}")
        
        # 設置最小窗口大小
        self.minsize(1280, 720)

        # 6. 初始化其他變數
        self.file_type_var = StringVar(value=self.last_file_type)
        self.intended_selection = None
        self.current_batch_id = 0
        self.is_loading_cancelled = False
        
        # 7. 初始化後台更新管理器
        self.update_manager = BackgroundUpdateManager(self)
        
        # 8. 設置窗口屬性
        self.resizable(True, True)

        # 9. 創建 UI 素
        self.create_widgets()

        # 10. 立即更新所有 UI 文字
        self.update_ui_language()

        # 初始化背初始化器
        self.background_initializer = BackgroundInitializer(self)
        self.background_initializer.start()

        # 綁定窗口關閉事件
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 創建標籤拖放視窗
        self.tag_drop_window = TagDropWindow(self, self.file_manager.get_all_used_tags(), self.handle_tag_drop)
        # 設置快捷鍵配置
        if hasattr(self, 'hotkey_modifier') and hasattr(self, 'hotkey_key'):
            self.tag_drop_window.update_hotkey(self.hotkey_modifier, self.hotkey_key)
        
        self.language_dialog = None  # 添加語言對話框實例變量
        self.hotkey_dialog = None  # 添加快捷鍵對話框實例變量
        
        # 绑定窗口最小化事件
        self.bind('<Unmap>', self.on_window_minimize)
        
        # 保存所有子窗口的列表
        self.child_windows = []
        
    def handle_exception(self, exc_type, exc_value, exc_traceback):
        """處理未捕獲的異常"""
        # 記錄錯誤
        logger.error("未捕獲的異常:", exc_info=(exc_type, exc_value, exc_traceback))
        
        # 發送崩潰報告
        self.crash_reporter.report_crash({
            "type": str(exc_type),
            "value": str(exc_value),
            "traceback": "".join(traceback.format_tb(exc_traceback))
        })
        
        # 顯示錯誤信息給用戶
        CustomMessageBox.show_error(
            self,
            self.get_text("error"),
            self.get_text("unexpected_error")
        )

    def check_for_updates(self):
        """檢查更新"""
        self.updater.check_and_update()

    def on_closing(self):
        """窗口關閉時保存配置"""
        try:
            # 獲取當前窗口位和大小
            geometry = self.geometry()
            x = self.winfo_x()
            y = self.winfo_y()
            width = self.winfo_width()
            height = self.winfo_height()

            # 讀取現有配置
            config = {}
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)

            # 更新窗口配置
            config['window'] = {
                'width': width,
                'height': height,
                'x': x,
                'y': y
            }

            # 保存置
            with open(self.config_file, 'w', encoding='utf-8') as file:
                json.dump(config, file, ensure_ascii=False, indent=4)

        except Exception as e:
            print(f"保存窗口配置時發生錯: {str(e)}")

        # 調用原有的銷毀方
        self.destroy()

    def initialize_language(self):
        """初始化語言設置"""
        try:
            # 使用Windows API獲取系統顯示語言
            windll.kernel32.GetUserDefaultUILanguage()
            lcid = windll.kernel32.GetUserDefaultUILanguage()
            
            # LCID轉換對照表
            lcid_to_lang = {
                1028: 'zh_TW',  # 繁體中文
                2052: 'zh_CN',  # 簡體中文
                1041: 'ja_JP',  # 日文
                1033: 'en_US'   # 英文
            }
            
            print(f"[Language Init] Windows UI Language LCID: {lcid}")
            system_locale = lcid_to_lang.get(lcid)
            print(f"[Language Init] Mapped language: {system_locale}")
            
            if system_locale in LANGUAGES:
                self.current_language = system_locale
                print(f"[Language Init] Using Windows display language: {self.current_language}")
            else:
                self.current_language = 'en_US'
                print(f"[Language Init] Using default language: {self.current_language} (Windows language {lcid} not supported)")
            
            print(f"[Language Init] Final language setting: {self.current_language}")
            
        except Exception as e:
            print(f"[Language Init] Error detecting Windows language: {str(e)}")
            self.current_language = 'en_US'

    def load_config(self):
        """載入配置文件"""
        try:
            # 初始化預設值
            self.last_folders = []
            self.last_file_type = "ALL"
            self.last_selected_tags = []  # 添加這行
            self.initialize_language()  # 先初始化語言
            
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                    # 讀取語言設置
                    saved_language = config.get('language')
                    if saved_language and saved_language in LANGUAGES:
                        self.current_language = saved_language
                    
                    saved_theme = config.get('theme')
                    if saved_theme:
                        self.current_theme = saved_theme
                    
                    self.last_folders = config.get('last_folders', [])
                    self.last_file_type = config.get('last_file_type', "ALL")
                    self.last_selected_tags = config.get('last_selected_tags', [])  # 添加這行
                    
            # 無論是否存在配置文件，都保存當前配置
            self.save_config()
            
        except Exception as e:
            print(f"載入配置文件時發生錯誤: {str(e)}")
            # 發生錯誤時使用預設值
            self.last_folders = []
            self.last_file_type = "ALL"
            self.last_selected_tags = []  # 添加這行
            self.initialize_language()

    def save_config(self):
        """保存配置"""
        config = {
            'language': self.current_language,
            'theme': self.current_theme,
            'last_folders': self.file_manager.folder_paths,
            'last_file_type': self.file_type_var.get(),
            'last_selected_tags': [self.tag_list.item(item, 'text') for item in self.tag_list.selection()] if hasattr(self, 'tag_list') else [],
            'default_tag_color': self.file_manager.default_color,  # 添加預設標籤顏色
            'hotkeys': {
                'modifier': self.tag_drop_window.hotkey_modifier if hasattr(self, 'tag_drop_window') else 'Ctrl',
                'key': self.tag_drop_window.hotkey_key if hasattr(self, 'tag_drop_window') else '1'
            }
        }
        with open(self.config_file, 'w', encoding='utf-8') as file:
            json.dump(config, file, ensure_ascii=False, indent=4)

    def get_text(self, key):
        """獲取當前語言的文字"""
        return LANGUAGES[self.current_language].get(key, LANGUAGES['en_US'][key])

    def show_language_dialog(self):
        """顯示語言選擇對話框"""
        if self.language_dialog is None:
            # 創建非模態對話框
            self.language_dialog = tk.Toplevel(self)
            self.language_dialog.title(self.get_text("select_language"))
            self.language_dialog.geometry("300x150")
            
            # 設置對話框屬性
            self.language_dialog.transient(self)  # 設置父窗口
            self.language_dialog.resizable(False, False)  # 禁止調整大小
            
            # 創建下拉選單
            language_var = StringVar(value=LANGUAGE_NAMES[self.current_language])
            language_combo = tb.Combobox(
                self.language_dialog,
                textvariable=language_var,
                values=list(LANGUAGE_NAMES.values()),
                state="readonly",
                bootstyle="primary"
            )
            language_combo.pack(pady=20)

            def change_language():
                # 根據顯示名稱獲取語言代碼
                selected_name = language_var.get()
                for code, name in LANGUAGE_NAMES.items():
                    if name == selected_name:
                        self.current_language = code
                        break
                
                # 立即保存配置
                self.save_config()
                self.update_ui_language()
                self.language_dialog.withdraw()  # 隱藏而不是銷毀

            # 保存按鈕
            save_btn = tb.Button(
                self.language_dialog,
                text=self.get_text("save"),
                command=change_language,
                bootstyle="success"
            )
            save_btn.pack(pady=10)

            # 添加關閉視窗的處理
            self.language_dialog.protocol("WM_DELETE_WINDOW", lambda: self.language_dialog.withdraw())
            
            # 設置視窗位置為居中
            self.center_window(self.language_dialog, 300, 150)
        else:
            # 如果對話框已存在，則更新語言顯示並顯示它
            self.language_dialog.deiconify()
            self.center_window(self.language_dialog, 300, 150)

    def update_ui_language(self):
        """更新界面語言"""

        # 更新按钮文本
        self.list_untagged_btn.config(text=self.get_text("list_untagged"))
        self.delete_tag_btn.config(text=self.get_text("delete_tag"))
        self.rename_tag_btn.config(text=self.get_text("rename_tag"))
        self.add_tag_btn.config(text=self.get_text("add_tag"))
        self.remove_tag_btn.config(text=self.get_text("remove_tag"))
        self.edit_note_btn.config(text=self.get_text("edit_note"))
        self.export_files_btn.config(text=self.get_text("export_files"))  # 添加导出文件按钮的文本更新
        self.help_btn.config(text=self.get_text("help"))
        self.settings_btn.config(text=self.get_text("settings"))
        self.change_color_btn.config(text=self.get_text("change_color"))

        # 更新標籤文字
        self.tag_frame.config(text=self.get_text("tags"))
        self.file_frame.config(text=self.get_text("file_list"))
        self.filter_tags_label.config(text=self.get_text("filter_tags"))
        self.filter_files_label.config(text=self.get_text("filter_files"))
        self.file_type_label.config(text=self.get_text("file_type"))
        
        # 更新右鍵選單
        self.context_menu.entryconfigure(0, label=self.get_text("add_tag"))
        self.context_menu.entryconfigure(1, label=self.get_text("remove_tag"))
        self.context_menu.entryconfigure(2, label=self.get_text("edit_note"))
        self.context_menu.entryconfigure(3, label=self.get_text("export_files"))  # 添加导出文件选项
        self.context_menu.entryconfigure(4, label=self.get_text("open_file_location"))
        
        # 更新標籤右鍵選單
        self.tag_context_menu.entryconfigure(0, label=self.get_text("delete_tag"))
        self.tag_context_menu.entryconfigure(1, label=self.get_text("rename_tag"))
        
        # 更新檔案列表標題
        self.file_tree.heading("#0", text=self.get_text("file_name"))
        
        # 更新當前標籤顯示
        if hasattr(self, 'current_tag_selection') and self.current_tag_selection:
            self.label_current_tag.config(text=f"{self.get_text('current_tag')}: {self.current_tag_selection}")

        # 更新檔案類型下拉選單
        current_index = self.file_type_combo.current()  # 保存當前選擇的索引
        all_file_types = self.file_manager.get_all_file_types()
        sorted_file_types = sorted(all_file_types)
        options = [self.get_text("all_types")] + sorted_file_types
        self.file_type_combo['values'] = options
        # 恢復之前的選擇
        if current_index >= 0:
            self.file_type_combo.current(current_index)
        else:
            self.file_type_combo.current(0)  # 預設選擇"全部類型"
        
        # 更新檔案限制提示訊息
        if hasattr(self, 'limit_message_label'):
            total_count = len(self.all_files) if hasattr(self, 'all_files') else 0
            if total_count > 1000:
                self.limit_message_label.config(
                    text=f"{self.get_text('file_limit_message')} ({total_count} {self.get_text('total_files')})"
                )

    def center_window(self, window, width, height, parent=None):
        """
        將指定的窗口居中於父窗口或螢幕援 DPI 縮放。
        """
        if parent:
            parent.update_idletasks()
            parent_x = parent.winfo_rootx()
            parent_y = parent.winfo_rooty()
            parent_width = parent.winfo_width()
            parent_height = parent.winfo_height()
        else:
            parent_x = parent_y = 0
            parent_width = window.winfo_screenwidth()
            parent_height = window.winfo_screenheight()

        # 計算縮放後的位置
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        
        # 確保視窗位置不會超出螢幕範圍
        x = max(0, min(x, window.winfo_screenwidth() - width))
        y = max(0, min(y, window.winfo_screenheight() - height))
        
        window.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        # 架
        self.main_frame = tb.Frame(self)
        self.main_frame.pack(fill=tb.BOTH, expand=True, padx=10, pady=10)

        # 內容框架（標籤列表和檔案列表）
        content_frame = tb.Frame(self.main_frame)
        content_frame.pack(fill=tb.BOTH, expand=True)

        # 使用 grid 佈局管理器，并限制最大高度
        content_frame.grid_rowconfigure(0, weight=3)  # 给内容区域较小的权重
        content_frame.grid_columnconfigure(0, weight=0)  # tag_frame 列，权重设为0以保持固定宽度
        content_frame.grid_columnconfigure(1, weight=8)  # file_frame 列

        # 標籤列表框架
        tag_frame = tb.Labelframe(content_frame, text=self.get_text("tags"), bootstyle="info")
        tag_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
        tag_frame.grid_propagate(False)  # 防止框架自动调整大小
        tag_frame.configure(height=400, width=200)  # 设置固定高度和宽度

        # 標籤選輸入框
        tag_search_frame = tb.Frame(tag_frame)
        tag_search_frame.pack(fill=tk.X, padx=5, pady=(0, 0))

        tag_search_label = tb.Label(tag_search_frame, text=self.get_text("filter_tags"), bootstyle="info")
        tag_search_label.pack(side=tk.LEFT, padx=(0, 5))

        self.tag_search_var = StringVar()
        self.tag_search_var.trace_add("write", self.filter_tag_list)

        tag_search_entry = ttk.Entry(tag_search_frame, 
                                    textvariable=self.tag_search_var, 
                                    font=("Segoe UI", 14))
        tag_search_entry.pack(fill=tk.X, expand=True)

        # 標籤列表
        self.tag_list = ttk.Treeview(
            tag_frame,
            show='tree',
            selectmode='extended',
            height=10
        )
        self.tag_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 設置拖放目標
        self.tag_list.drop_target_register(DND_FILES)
        self.tag_list.dnd_bind('<<Drop>>', self.on_drop)
        self.tag_list.dnd_bind('<<DropEnter>>', self.on_drop_enter)
        self.tag_list.dnd_bind('<<DropLeave>>', self.on_drop_leave)
        self.tag_list.dnd_bind('<<DropPosition>>', self.on_drop_motion)

        tag_scrollbar = tb.Scrollbar(tag_frame, orient=tk.VERTICAL, command=self.tag_list.yview, bootstyle="round")
        self.tag_list.configure(yscrollcommand=tag_scrollbar.set)
        tag_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        self.tag_list.bind('<<TreeviewSelect>>', self.update_file_list_by_tag_selection)
        self.tag_list.bind('<Double-1>', self.edit_tag_context)
        self.tag_list.bind('<Button-3>', self.show_tag_context_menu)

        # 右選單
        self.tag_context_menu = tb.Menu(self, tearoff=0)
        self.tag_context_menu.add_command(label=self.get_text("delete_tag"), command=self.delete_selected_tag)
        self.tag_context_menu.add_command(label=self.get_text("rename_tag"), command=self.rename_selected_tag)

        # 檔案列表框架
        file_frame = tb.Labelframe(content_frame, text=self.get_text("file_list"), bootstyle="info")
        file_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)
        file_frame.grid_propagate(False)  # 防止框架自动调整大小
        file_frame.configure(height=400)  # 设置固定高度

        # 檔案篩選輸入框使用 grid 而非 pack
        file_search_frame = tb.Frame(file_frame)
        file_search_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 0))

        file_search_label = tb.Label(file_search_frame, text=self.get_text("filter_files"), bootstyle="info")
        file_search_label.grid(row=0, column=0, sticky="w", padx=(0, 5))

        self.search_var = StringVar()
        self.search_var.trace_add("write", self.filter_file_list)

        file_search_entry = ttk.Entry(file_search_frame, 
                                    textvariable=self.search_var, 
                                    font=("Segoe UI", 14))
        file_search_entry.grid(row=0, column=1, sticky="ew")
        file_search_frame.grid_columnconfigure(1, weight=1)  # 使輸入框可拉伸

        # 新增檔案類型下拉式選單
        file_type_label = tb.Label(file_search_frame, text=self.get_text("file_type"), bootstyle="info")
        file_type_label.grid(row=0, column=2, sticky="w", padx=(10, 5))

        self.file_type_combo = tb.Combobox(
            file_search_frame,
            textvariable=self.file_type_var,
            state="readonly",
            values=[self.get_text("all_types")],
            bootstyle="primary"
        )
        self.file_type_combo.grid(row=0, column=3, sticky="w")
        self.file_type_combo.bind("<<ComboboxSelected>>", self.on_file_type_selected)

        # **新增：定義樣式並配置 Treeview 的字體**
        style = ttk.Style()
        style.configure("Custom.Treeview", font=("Segoe UI", 14))  # 修改為14
        style.configure("Custom.Treeview.Heading", font=("Segoe UI", 14, "bold"))  # 修改為14

        # 檔案列表使用 Treeview（移除 pack，使用 grid）
        self.file_tree = ttk.Treeview(
            file_frame,
            show='tree',
            selectmode=tk.EXTENDED,  # 確保多選功能啟用
            style="Custom.Treeview"
        )
        self.file_tree.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)

        # 添加垂直滾動條
        file_scrollbar_y = tb.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_tree.yview, bootstyle="round")
        self.file_tree.config(yscrollcommand=file_scrollbar_y.set)
        file_scrollbar_y.grid(row=1, column=1, sticky='ns', pady=0)

        # 添加水平滾動條
        file_scrollbar_x = tb.Scrollbar(file_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview, bootstyle="round")
        self.file_tree.config(xscrollcommand=file_scrollbar_x.set)
        file_scrollbar_x.grid(row=2, column=0, sticky='ew', padx=0, pady=0)

        # 設定 grid 欄位和行的權重
        file_frame.grid_rowconfigure(1, weight=1)
        file_frame.grid_columnconfigure(0, weight=1)

        self.file_tree.bind('<Double-1>', self.on_file_double_click)  # 修改雙擊事件處理函數
        self.file_tree.bind('<Button-3>', self.show_context_menu)
        self.file_tree.bind('<Motion>', self.show_tooltip)
        self.file_tree.bind('<Leave>', self.hide_tooltip)
        self.file_tree.bind('<<TreeviewSelect>>', self.update_edit_buttons_state)
        
        # 設定欄位自動拉伸以填滿空間
        self.file_tree.column("#0", stretch=True, anchor='w')  # 確保欄位能夠伸並左對齊
        self.file_tree.heading("#0", text=self.get_text("file_name"), anchor='w')  # 修改

        # 儲存所有文件以便分批載入
        self.all_files = []
        self.load_batch_size = 100  # 每次載入的文件數量
        self.current_batch = 0

        # 當前選中標籤顯示標籤
        self.label_current_tag = tb.Label(self.main_frame, text="", font=("Segoe UI", 14, "bold"))
        self.label_current_tag.pack(fill=tb.X, padx=10, pady=(10, 5))
        self.label_current_tag.pack_forget()

        # 按框架，將按鈕放在當前選中籤的下方
        button_frame = tb.Frame(self.main_frame)
        button_frame.pack(fill=tb.X, pady=10)

        # 左側按鈕框架
        footer_left = tb.Frame(button_frame)
        footer_left.pack(side=tb.LEFT, padx=10)

        # 右側鈕框架
        footer_right = tb.Frame(button_frame)
        footer_right.pack(side=tb.RIGHT, padx=10)

        # 左側按鈕
        list_untagged_btn = tb.Button(
            footer_left, text=self.get_text("list_untagged"), command=self.list_untagged_files, bootstyle="success-outline"
        )
        list_untagged_btn.pack(side=tb.LEFT, padx=5)

        # 新增「刪除標籤」按鈕
        self.delete_tag_btn = tb.Button(
            footer_left, text=self.get_text("delete_tag"), command=self.delete_selected_tag, bootstyle="danger", state=tk.DISABLED
        )
        self.delete_tag_btn.pack(side=tb.LEFT, padx=5)

        # 新增「重新命名標籤」按鈕
        self.rename_tag_btn = tb.Button(
            footer_left, text=self.get_text("rename_tag"), command=self.rename_selected_tag, bootstyle="danger", state=tk.DISABLED
        )
        self.rename_tag_btn.pack(side=tb.LEFT, padx=5)

        # 新增「修改顏色」按鈕
        self.change_color_btn = tb.Button(
            footer_left, text=self.get_text("change_color"), command=self.change_tag_color, bootstyle="danger", state=tk.DISABLED
        )
        self.change_color_btn.pack(side=tb.LEFT, padx=5)

        # 新增「新增標籤」按鈕
        self.add_tag_btn = tb.Button(
            footer_left, text=self.get_text("add_tag"), command=self.add_tags, bootstyle="primary", state=tk.DISABLED
        )
        self.add_tag_btn.pack(side=tb.LEFT, padx=5)

        # 新增「移除標籤」按鈕
        self.remove_tag_btn = tb.Button(
            footer_left, text=self.get_text("remove_tag"), command=self.remove_tags, bootstyle="primary", state=tk.DISABLED
        )
        self.remove_tag_btn.pack(side=tb.LEFT, padx=5)

        # 新增「編輯備註」按鈕
        self.edit_note_btn = tb.Button(
            footer_left, text=self.get_text("edit_note"), command=self.edit_notes, bootstyle="warning", state=tk.DISABLED
        )
        self.edit_note_btn.pack(side=tb.LEFT, padx=5)

        # 新增「導出文件」按鈕
        self.export_files_btn = tb.Button(
            footer_left, text=self.get_text("export_files"), command=self.export_selected_files, bootstyle="warning", state=tk.DISABLED
        )
        self.export_files_btn.pack(side=tb.LEFT, padx=5)

        # 右側按鈕 - 順序調整，將說明還原按鈕放在最右邊
        # 將說明按鈕移到設定按鈕前面
        self.help_btn = tb.Button(
            footer_right,
            text=self.get_text("help"),
            command=self.show_help,
            bootstyle="primary-outline"
        )
        self.help_btn.pack(side=tk.LEFT, padx=5)

        # 新增設定按鈕到最右邊
        self.settings_btn = tb.Button(
            footer_right,
            text=self.get_text("settings"),
            command=self.show_settings_dialog,
            bootstyle="info-outline"
        )
        self.settings_btn.pack(side=tk.LEFT, padx=5)

        # 获取系统 DPI 缩放
        try:
            import ctypes
            user32 = ctypes.windll.user32
            user32.SetProcessDPIAware()
            dpi = user32.GetDpiForSystem()
            dpi_scale = dpi / 96.0  # 96 是标准 DPI
            if dpi_scale < 1.0:
                dpi_scale = 1.0
        except:
            dpi_scale = 1.0

        # 根据 DPI 缩放调整字体大小和 rowheight
        base_rowheight = 30
        scaled_rowheight = int(base_rowheight * dpi_scale)

        # Tooltip 標籤
        self.tooltip = tb.Label(
            self,
            text="",
            bootstyle=SECONDARY,
            padding=5,
            font=("Segoe UI", 14, "bold"),  # 修改為12
            foreground="white",
            background="#333333",
            borderwidth=1,
            relief="solid"
        )
        self.tooltip.place_forget()

        # 右鍵選單
        self.context_menu = tb.Menu(self, tearoff=0)
        self.context_menu.add_command(label=self.get_text("add_tag"), command=self.add_tags)
        self.context_menu.add_command(label=self.get_text("remove_tag"), command=self.remove_tags)
        self.context_menu.add_command(label=self.get_text("edit_note"), command=self.edit_notes)
        self.context_menu.add_command(label=self.get_text("export_files"), command=self.export_selected_files)  # 添加导出文件选项
        self.context_menu.add_command(label=self.get_text("open_file_location"), command=self.open_file_location)

        self.refresh_tag_list()

        # **新增綁定窗口大小變事件以自動調整欄位寬度**
        self.bind("<Configure>", self.on_window_resize)

        # 保存其他重要的 UI 元素的引用以便後續更
        self.tag_frame = tag_frame
        self.file_frame = file_frame
        self.filter_tags_label = tag_search_label
        self.filter_files_label = file_search_label
        self.file_type_label = file_type_label
        self.list_untagged_btn = list_untagged_btn
        
        tagStyle = ttk.Style()
        tagStyle.configure("Treeview", rowheight=scaled_rowheight, font=("Segoe UI", 14))
        tagStyle.configure("Treeview.Item", rowheight=scaled_rowheight, font=("Segoe UI", 14))

    def on_file_type_selected(self, event):
        """處理檔案類型選擇事件，重新篩選檔列表並保存選"""
        self.filter_file_list()
        self.save_last_folders(self.file_manager.folder_paths, self.file_type_var.get())

    def filter_tag_list(self, *args):
        """根輸入框內篩選標籤列表"""
        search_text = self.tag_search_var.get().strip().lower()
        all_tags = self.file_manager.get_all_used_tags()
        filtered_tags = [tag for tag in all_tags if search_text in tag.lower()]

        self.tag_list.delete(*self.tag_list.get_children())
        for tag in filtered_tags:
            self.tag_list.insert('', tk.END, text=tag)  # 使用 text 代替 values

    def filter_file_list(self, *args):
        """根據輸入框的內容篩選檔案列表，並根據選擇的檔案類型篩選"""
        # 設置更新狀態為True，阻止自動更新
        self.update_manager.set_updating(True)
        
        try:
            # 保存當前選中的檔案路徑
            selected_items = self.file_tree.selection()
            selected_paths = []
            for item in selected_items:
                try:
                    file_path = self.file_tree.item(item, 'values')[0]
                    selected_paths.append(file_path)
                except:
                    continue
            
            # 取消當前的載入過程
            self.cancel_current_loading()
            
            search_text = self.search_var.get().strip().lower()
            selected_file_type = self.file_type_var.get()
            # 使用索引判斷是否選擇了"全部類型"
            is_all_types = self.file_type_combo.current() == 0

            # 使用生成器獲取基礎檔案列表
            def get_base_files():
                if hasattr(self, 'showing_untagged') and self.showing_untagged:
                    yield from self.file_manager.list_untagged_files()
                elif hasattr(self, 'current_tag_selection') and self.current_tag_selection:
                    yield from self.file_manager.search_by_tags(self.current_tag_selection)
                else:
                    yield from self.file_manager.search_by_tags("")

            # 使用生成器進行篩選
            def filter_files():
                for file in get_base_files():
                    if not is_all_types and not file.lower().endswith(selected_file_type):
                        continue
                    if search_text and search_text not in os.path.basename(file).lower():
                        continue
                    yield file

            # 使用 itertools.islice 限制數量並轉換為列表
            filtered_files = list(itertools.islice(filter_files(), 1000))
            
            # 顯示提示訊息
            total_count = sum(1 for _ in get_base_files())  # 使用生成器計算總數
            self.show_file_limit_message(total_count)

            # 保存選中的檔案路徑到實例變量
            self.selected_paths = selected_paths
            
            # 更新 all_files 並開始新的載入
            self.all_files = filtered_files
            self.start_new_loading()

            # 更新檔案類型選項
            self.update_file_type_options()

            # 更新按鈕狀態
            self.update_edit_buttons_state()
            
        finally:
            # 設置更新狀態為False，允許自動更新
            self.update_manager.set_updating(False)

    def show_file_limit_message(self, total_count):
        """顯示檔案數量超過限制的提示訊息"""
        if not hasattr(self, 'limit_message_label'):
            self.limit_message_label = tb.Label(
                self.file_frame,
                text="",
                bootstyle="warning",
                font=("Segoe UI", 14)
            )
            self.limit_message_label.grid(row=3, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        if total_count > 1000:
            self.limit_message_label.config(
                text=f"{self.get_text('file_limit_message')} ({total_count} {self.get_text('total_files')})"
            )
            self.limit_message_label.grid()
        else:
            self.limit_message_label.grid_remove()

    def list_untagged_files(self, event=None):
        """列出未標籤檔案，並應用篩選條件"""
        try:
            # 設置更新狀態為True，阻止自動更新
            self.update_manager.set_updating(True)
            
            # 取消當前的載入過程
            self.cancel_current_loading()
            
            # 清除當標籤選擇，設置未標籤標記
            self.current_tag_selection = None
            self.showing_untagged = True
            self.tag_list.selection_remove(*self.tag_list.selection())
            
            # 清除標籤顯示
            self.label_current_tag.pack_forget()
            
            # 獲取未標籤檔案列表
            untagged_files = self.file_manager.list_untagged_files()

            # 應用當前的檔案類型過濾
            # 使用索引判斷是否選擇了"全部類型"
            is_all_types = self.file_type_combo.current() == 0
            selected_file_type = self.file_type_var.get()
            if not is_all_types:
                untagged_files = [file for file in untagged_files if file.lower().endswith(selected_file_type)]

            # 應用搜尋文字過濾
            search_text = self.search_var.get().strip().lower()
            if search_text:
                untagged_files = [file for file in untagged_files if search_text in os.path.basename(file).lower()]

            # 顯示提示訊息（不論是否超過限制）
            self.show_file_limit_message(len(untagged_files))

            # 限制顯示數量
            if len(untagged_files) > 1000:
                untagged_files = untagged_files[:1000]

            # 更新文件列表
            self.all_files = untagged_files.copy()
            
            # 開始新的載入過程
            self.start_new_loading()

            # 更新文件類型選項（保持基於所有文件）
            self.update_file_type_options()

            # 更新按鈕狀態
            self.update_edit_buttons_state()
            
            # 清空配置中的選中標籤並保存
            self.save_config()
            
            # 確保所有更新都已完成
            self.update_idletasks()
            
        finally:
            # 無論如何都要設置更新狀態為False
            self.update_manager.set_updating(False)
            # 確保不會觸發其他更新
            self.current_tag_selection = None

    def get_all_used_tags(self):
        all_tags = set()
        # 使用字典的副本進行遍歷
        files_tags_copy = self.files_tags.copy()
        for info in files_tags_copy.values():
            all_tags.update(info.get("tags", []))
        return sorted(all_tags)

    def get_all_file_types(self):
        """返回所有檔案的擴展名，已排序並去重"""
        file_types = set()
        # 使用字典的副本進行遍歷
        files_tags_copy = self.files_tags.copy()
        for file_path in files_tags_copy.keys():
            _, ext = os.path.splitext(file_path)
            if ext:
                file_types.add(ext.lower())
        return sorted(file_types)

    def rename_tag(self, old_tag, new_tag, merge=False):
        # 保存旧标签的颜色
        old_tag_color = self.get_tag_color(old_tag)
        old_tag_color_timestamp = self.get_tag_color_timestamp(old_tag)
        
        if merge:
            for file_path, info in self.files_tags.items():
                if old_tag in info["tags"]:
                    index = info["tags"].index(old_tag)
                    info["tags"][index] = new_tag
                    if new_tag not in info["tags"]:
                        self.db_data[file_path]["tags"].append(new_tag)
        else:
            for file_path, info in self.files_tags.items():
                if old_tag in info["tags"]:
                    index = info["tags"].index(old_tag)
                    info["tags"][index] = new_tag
                    # 同步更新 db_data
                    if file_path in self.db_data:
                        tags = self.db_data[file_path]["tags"]
                        if old_tag in tags:
                            index = tags.index(old_tag)
                            tags[index] = new_tag
        
        # 如果旧标签有自定义颜色，将其应用到新标签
        if old_tag_color != self.default_color:
            self.set_tag_color(new_tag, old_tag_color)
            # 保持原有的时间戳
            if old_tag_color_timestamp:
                self.tag_color_timestamps[new_tag] = old_tag_color_timestamp
        
        # 删除旧标签的颜色设置
        if old_tag in self.tag_colors:
            del self.tag_colors[old_tag]
        if old_tag in self.tag_color_timestamps:
            del self.tag_color_timestamps[old_tag]
        
        self.save_db()

    def delete_tag(self, tag):
        """從所有檔案中刪除指定的標籤"""
        # 先從 files_tags 中移除標籤
        for file_path, info in self.files_tags.items():
            if tag in info["tags"]:
                info["tags"].remove(tag)
                # 同步更新 db_data
                if file_path in self.db_data:
                    if tag in self.db_data[file_path]["tags"]:
                        self.db_data[file_path]["tags"].remove(tag)
                    # 如果檔案沒有任何標籤和備註，從資料庫中移除
                    if not self.db_data[file_path]["tags"] and not self.db_data[file_path]["note"]:
                        del self.db_data[file_path]
        
        self.save_db()

    # Batch operations
    def add_tags_batch(self, filenames, tags):
        for filename in filenames:
            for tag in tags:
                if tag.strip():
                    self.add_tag(filename, tag)

    def remove_tags_batch(self, filenames, tags):
        for filename in filenames:
            for tag in tags:
                self.remove_tag(filename, tag)

    def start_monitoring(self, callback):
        """開始監控所有資料夾"""
        # 停止現有的監控
        self.stop_monitoring()
        
        # 創建新的事件處理器
        self.event_handler = FileChangeHandler(self, callback)
        
        # 為每個資料夾創建觀察者
        for folder_path in self.folder_paths:
            if os.path.exists(folder_path):
                observer = Observer()
                observer.schedule(self.event_handler, folder_path, recursive=True)
                observer.start()
                self.observers.append(observer)
    
    def stop_monitoring(self):
        """停止所有資料夾的監控"""
        if hasattr(self, 'observers'):
            for observer in self.observers:
                observer.stop()
            for observer in self.observers:
                observer.join()
            self.observers.clear()

    def update_file_list_by_tag_selection(self, event=None):
        # 設置更新狀態為True，阻止自動更新
        self.update_manager.set_updating(True)
        
        try:
            selection = self.tag_list.selection()
            if selection:
                # 取消當前的載入過程
                self.cancel_current_loading()
                
                # 清除未標籤標記
                self.showing_untagged = False
                
                # 獲取所有選中的標籤
                selected_tags = [self.tag_list.item(item, 'text') for item in selection]
                self.current_tag_selection = ','.join(selected_tags)
                
                # 獲取同時擁有所有選定標籤的檔案
                full_paths = self.file_manager.search_by_tags(self.current_tag_selection)
                
                # 應用當前的檔案類型過濾
                # 使用索引判斷是否選擇了"全部類型"
                is_all_types = self.file_type_combo.current() == 0
                selected_file_type = self.file_type_var.get()
                if not is_all_types:
                    full_paths = [file for file in full_paths if file.lower().endswith(selected_file_type)]
                
                # 應用搜尋文字過濾
                search_text = self.search_var.get().strip().lower()
                if search_text:
                    full_paths = [file for file in full_paths if search_text in os.path.basename(file).lower()]
                
                # 顯示提示訊息（不論是否超過限制）
                self.show_file_limit_message(len(full_paths))
                
                # 限制顯示數量
                if len(full_paths) > 1000:
                    full_paths = full_paths[:1000]
                
                # 清空並重新填充文件列表
                self.all_files = full_paths
                self.start_new_loading()

                # 更新標籤顯示
                if len(selected_tags) == 1:
                    self.label_current_tag.config(text=f"{self.get_text('current_tag')}: {selected_tags[0]}")
                else:
                    self.label_current_tag.config(text=f"{self.get_text('current_tag')}: {', '.join(selected_tags)}")
                self.label_current_tag.pack(fill=tb.X, padx=10, pady=(10, 5))

                # 根據選擇的標籤數量更新按鈕狀態
                if len(selected_tags) == 1:
                    self.delete_tag_btn.config(state=tk.NORMAL)
                    self.rename_tag_btn.config(state=tk.NORMAL)
                    self.change_color_btn.config(state=tk.NORMAL)
                else:
                    self.delete_tag_btn.config(state=tk.NORMAL)
                    self.rename_tag_btn.config(state=tk.DISABLED)
                    self.change_color_btn.config(state=tk.DISABLED)
                    
                # 保存當前選擇的標籤到配置
                self.save_config()
            else:
                # 只有在不是未標籤模式時才清除標記
                if not hasattr(self, 'showing_untagged') or not self.showing_untagged:
                    self.current_tag_selection = None
                    self.showing_untagged = False
                    self.label_current_tag.pack_forget()
                    self.delete_tag_btn.config(state=tk.DISABLED)
                    self.rename_tag_btn.config(state=tk.DISABLED)
                    self.change_color_btn.config(state=tk.DISABLED)
                    
                    # 當沒有選擇標籤時，顯示所有檔案並應用過濾條件
                    full_paths = self.file_manager.search_by_tags("")
                    
                    # 應用當前的檔案類型過濾
                    # 使用索引判斷是否選擇了"全部類型"
                    is_all_types = self.file_type_combo.current() == 0
                    selected_file_type = self.file_type_var.get()
                    if not is_all_types:
                        full_paths = [file for file in full_paths if file.lower().endswith(selected_file_type)]
                    
                    # 應用搜尋文字過濾
                    search_text = self.search_var.get().strip().lower()
                    if search_text:
                        full_paths = [file for file in full_paths if search_text in os.path.basename(file).lower()]
                    
                    # 顯示提示訊息
                    self.show_file_limit_message(len(full_paths))
                    
                    # 限制顯示數量
                    if len(full_paths) > 1000:
                        full_paths = full_paths[:1000]
                    
                    self.all_files = full_paths
                    self.start_new_loading()
                    
                    # 保存當前選擇的標籤到配置（清除選擇）
                    self.save_config()
        finally:
            # 設置更新狀態為False，允許自動更新
            self.update_manager.set_updating(False)

    def cancel_current_loading(self):
        """取消當前的檔案載入過程"""
        self.is_loading_cancelled = True
        self.current_batch = 0

    def start_new_loading(self):
        """開始新的檔案載入過程"""
        self.is_loading_cancelled = False
        self.current_batch_id += 1
        self.current_batch = 0
        self.file_tree.delete(*self.file_tree.get_children())
        self.load_files_in_batches()

    def select_intended_file(self):
        """選中預期選擇的檔案"""
        children = self.file_tree.get_children()
        for child in children:
            if self.file_tree.item(child, 'values')[0] == self.intended_selection:
                self.file_tree.selection_set(child)
                self.file_tree.focus(child)
                self.file_tree.see(child)
                break
        self.intended_selection = None  # 清除記錄

    def show_help(self):
        """顯示使用說明視窗"""
        # 創建使用說明視窗
        help_window = create_modal_dialog(self, self.get_text("help_title"), 600, 600)  # 修改
        
        # 創建包含文字和滾動條的框架
        text_frame = tb.Frame(help_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 添加版本信息標籤
        version_label = tb.Label(
            help_window,
            text=f"{self.app_name} v{self.version}\n© 2025 NTTech Studio",
            font=("Segoe UI", 14),
            bootstyle="info"
        )
        version_label.pack(pady=(0, 10))
        
        # 添加滾動條
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 創建文字區域
        text_area = tk.Text(
            text_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 14),  # 修改為12
            padx=5,
            pady=5
        )
        text_area.pack(fill=tk.BOTH, expand=True)
        
        # 配置滾動條
        scrollbar.config(command=text_area.yview)
        
        # 插入說明文字（使用翻譯）
        text_area.insert(tk.END, self.get_text("help_content"))
        text_area.config(state=tk.DISABLED)  # 設為唯讀
        
        # 添加確定按鈕
        btn_frame = tb.Frame(help_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ok_btn = tb.Button(
            btn_frame,
            text=self.get_text("confirm"),  # 使用翻譯的確定按鈕文字
            command=help_window.destroy,
            bootstyle="success",
            width=10
        )
        ok_btn.pack(pady=5)

    def refresh_tag_list(self):
        """更新標籤列表，並套用過濾條件"""
        # 保存當前選中的標籤
        selected_tags = [self.tag_list.item(item, 'text') for item in self.tag_list.selection()]
        
        # 獲取當前的過濾文字
        search_text = self.tag_search_var.get().strip().lower()
        
        # 清空並重新填充標籤列表
        self.tag_list.delete(*self.tag_list.get_children())
        all_tags = self.file_manager.get_all_used_tags()
        
        # 套用過濾條件
        filtered_tags = [tag for tag in all_tags if search_text in tag.lower()]
        
        # 將標籤分為兩組：有顏色的和沒有顏色的
        colored_tags = []
        normal_tags = []
        for tag in filtered_tags:
            if self.file_manager.get_tag_color(tag) != self.file_manager.default_color:
                colored_tags.append(tag)
            else:
                normal_tags.append(tag)
        
        # 對有顏色的標籤按修改時間排序（最新的在前）
        colored_tags.sort(key=lambda tag: self.file_manager.get_tag_color_timestamp(tag), reverse=True)
        
        # 先插入有顏色的標籤，再插入普通標籤
        for tag in colored_tags + normal_tags:
            color = self.file_manager.get_tag_color(tag)
            item = self.tag_list.insert('', tk.END, text=tag)
            # 為每個標籤創建唯一的標籤樣式
            tag_style = f'tag_{hash(tag)}'
            # 只有當顏色不是預設顏色時才設置標籤樣式
            if color != self.file_manager.default_color:
                self.tag_list.tag_configure(tag_style, foreground=color)
                self.tag_list.item(item, tags=(tag_style,))
        
        # 恢復之前選中的標籤（如果它們仍然存在於過濾後的列表中）
        for tag in selected_tags:
            if tag in filtered_tags:
                for item in self.tag_list.get_children():
                    if self.tag_list.item(item, 'text') == tag:
                        self.tag_list.selection_add(item)
                        self.tag_list.see(item)
        
        # 如果有選中的標籤，更新檔案列表
        if self.tag_list.selection():
            self.update_file_list_by_tag_selection()
            
        # 更新按鈕狀態
        self.update_edit_buttons_state()

        # 更新拖放視窗的標籤列表
        if hasattr(self, 'tag_drop_window'):
            self.tag_drop_window.update_tags(self.file_manager.get_all_used_tags())

    def is_dark_color(self, color):
        """判斷顏色是否為深色"""
        try:
            # 檢查顏色格式
            if not color or not isinstance(color, str):
                return True  # 如果顏色無效，預設返回 True
            
            # 移除 # 符號
            color = color.lstrip('#')
            
            # 檢查顏色字串長度
            if len(color) != 6:
                return True  # 如果格式不正確，預設返回 True
            
            # 將顏色轉換為 RGB 值
            try:
                r = int(color[0:2], 16)
                g = int(color[2:4], 16)
                b = int(color[4:6], 16)
            except ValueError:
                return True  # 如果轉換失敗，預設返回 True
            
            # 計算亮度 (根據 WCAG 標準)
            brightness = (r * 299 + g * 587 + b * 114) / 1000
            
            # 如果亮度小於 128，則認為是深色
            return brightness < 128
        except Exception:
            return True  # 發生任何錯誤時，預設返回 True

    def update_edit_buttons_state(self, event=None):
        selected = self.file_tree.selection()
        if selected:
            self.add_tag_btn.config(state=tk.NORMAL)
            self.edit_note_btn.config(state=tk.NORMAL)
            self.export_files_btn.config(state=tk.NORMAL)
            
            # 檢查選取的檔案中是否有至少一個有標籤
            remove_enabled = False
            for i in selected:
                full_path = self.file_tree.item(i, 'values')[0]
                if self.file_manager.files_tags.get(full_path, {}).get("tags"):
                    remove_enabled = True
                    break
            
            self.remove_tag_btn.config(state=tk.NORMAL if remove_enabled else tk.DISABLED)
        else:
            self.add_tag_btn.config(state=tk.DISABLED)
            self.remove_tag_btn.config(state=tk.DISABLED)
            self.edit_note_btn.config(state=tk.DISABLED)
            self.export_files_btn.config(state=tk.DISABLED)  # 确保在没有选择文件时禁用导出按钮

    def add_tags(self, event=None):
        selected_items = self.file_tree.selection()
        selected_files = [self.file_tree.item(i, 'values')[0] for i in selected_items]
        if not selected_files:
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_files_to_tag"))  # 修改
            return

        tag_window = create_modal_dialog(self, self.get_text("add_tag"), 450, 550)  # 將高度從600改為550

        # 指示標籤
        label_instruction = tb.Label(
            tag_window,
            text=self.get_text("add_tags_instruction"),  # 修改
            wraplength=400
        )
        label_instruction.pack(pady=10)

        # 新標籤輸入框
        entry_new_tag = ttk.Entry(tag_window, width=50)
        entry_new_tag.pack(pady=1)

        # 現有標籤標籤
        label_existing = tb.Label(tag_window, text=self.get_text("select_from_existing"))  # 修改
        label_existing.pack(pady=(10, 0))

        # 現有標籤篩選框
        filter_frame = tb.Frame(tag_window)
        filter_frame.pack(fill=tk.X, padx=20, pady=(10, 0))

        filter_label = tb.Label(filter_frame, text=self.get_text("filter_existing_tags"), bootstyle="info")  # 修改
        filter_label.pack(side=tk.LEFT, padx=(0, 5))

        filter_var = StringVar()
        filter_var.trace_add("write", lambda *args: filter_existing_tags())

        filter_entry = ttk.Entry(filter_frame, textvariable=filter_var, font=("Segoe UI", 14))  # 修改為12
        filter_entry.pack(fill=tk.X, expand=True)

        # 標籤列表框架
        frame_listbox = tb.Frame(tag_window)
        frame_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        scrollbar_tags = tk.Scrollbar(frame_listbox, orient=tk.VERTICAL)
        scrollbar_tags.pack(side=tk.RIGHT, fill=tk.Y)

        all_tags = self.file_manager.get_all_used_tags()
        tags_var = StringVar(value=all_tags)
        listbox_tags_existing = tk.Listbox(
            frame_listbox,
            listvariable=tags_var,
            selectmode=tk.MULTIPLE,
            yscrollcommand=scrollbar_tags.set,
            font=("Segoe UI", 14)
        )
        listbox_tags_existing.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_tags.config(command=listbox_tags_existing.yview)

        def filter_existing_tags():
            search_text = filter_var.get().strip().lower()
            filtered_tags = [tag for tag in all_tags if search_text in tag.lower()]
            listbox_tags_existing.delete(0, tk.END)
            for tag in filtered_tags:
                listbox_tags_existing.insert(tk.END, tag)

        def save_added_tags():
            # 獲取選中的現有標籤
            selected_indices = listbox_tags_existing.curselection()
            selected_tags = [listbox_tags_existing.get(i) for i in selected_indices]

            # 獲取新標籤
            new_tags = entry_new_tag.get().split(',')
            new_tags = [tag.strip() for tag in new_tags if tag.strip()]
            combined_tags = list(set(selected_tags + new_tags))

            if not combined_tags:
                CustomMessageBox.show_info(tag_window, self.get_text("info"), self.get_text("enter_select_tags"))
                return

            # 將標籤添加到所有選的檔案
            for file_path in selected_files:
                # 直接使用完整路徑新增標籤
                for tag in combined_tags:
                    self.file_manager.add_tag(file_path, tag)

            CustomMessageBox.show_info(tag_window, self.get_text("success"), self.get_text("tags_added"))
            tag_window.destroy()

            # 更新界面
            self.refresh_tag_list()
            
            # 如果當前是在未標籤列表視圖，重新載入未標籤列表
            if hasattr(self, 'showing_untagged') and self.showing_untagged:
                self.list_untagged_files()
            else:
                # 否則更新當前標籤視圖
                self.update_file_list_by_tag_selection()
            
            self.update_edit_buttons_state()

        # 保存按鈕
        btn_save = tb.Button(
            tag_window,
            text=self.get_text("save"),
            command=save_added_tags,
            bootstyle=SUCCESS
        )
        btn_save.pack(pady=20)

    def remove_tags(self, event=None):
        selected_items = self.file_tree.selection()
        selected_files = [self.file_tree.item(i, 'values')[0] for i in selected_items]
        if not selected_files:
            # 修改：使用翻譯文字
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_files"))
            return

        # 收集選擇的檔案的共同標籤
        common_tags = None
        for file_path in selected_files:
            file_tags = set(self.file_manager.files_tags.get(file_path, {}).get("tags", []))
            if common_tags is None:
                common_tags = file_tags
            else:
                common_tags = common_tags.intersection(file_tags)

        if not common_tags:
            # 修改：使用翻譯文字
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("no_common_tags"))
            return

        # 修改：使用翻譯文
        tag_window = create_modal_dialog(self, self.get_text("remove_tag"), 450, 500)

        # 指示標籤
        label_instruction = tb.Label(
            tag_window,
            text=self.get_text("remove_tags_from_files"),  # 修改：使用翻譯文字
            wraplength=400
        )
        label_instruction.pack(pady=10)

        # 標籤列表框架
        frame_listbox = tb.Frame(tag_window)
        frame_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        scrollbar_tags = tk.Scrollbar(frame_listbox, orient=tk.VERTICAL)
        scrollbar_tags.pack(side=tk.RIGHT, fill=tk.Y)

        sorted_common_tags = sorted(common_tags)
        tags_var = StringVar(value=sorted_common_tags)
        listbox_tags_current = tk.Listbox(
            frame_listbox,
            listvariable=tags_var,
            selectmode=tk.MULTIPLE,
            yscrollcommand=scrollbar_tags.set,
            font=("Segoe UI", 14)
        )
        listbox_tags_current.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_tags.config(command=listbox_tags_current.yview)

        def save_removed_tags():
            # 獲取選中的標籤
            selected_indices = listbox_tags_current.curselection()
            if not selected_indices:
                CustomMessageBox.show_info(tag_window, self.get_text("info"), self.get_text("select_tag"))
                return

            selected_tags = [listbox_tags_current.get(i) for i in selected_indices]

            # 從所有選擇的檔案中移除標籤
            for file_path in selected_files:
                for tag in selected_tags:
                    self.file_manager.remove_tag(file_path, tag)

            CustomMessageBox.show_info(tag_window, self.get_text("success"), self.get_text("tags_removed"))
            tag_window.destroy()
            
            # 更新 UI
            self.refresh_tag_list()
            self.update_file_list_by_tag_selection()
            self.update_edit_buttons_state()

        # 保存按鈕
        btn_save = tb.Button(
            tag_window,
            text=self.get_text("save"),  # 改這裡
            command=save_removed_tags,
            bootstyle=SUCCESS
        )
        btn_save.pack(pady=20)

    def edit_notes(self, event=None):
        selected_items = self.file_tree.selection()
        selected_files = [self.file_tree.item(i, 'values')[0] for i in selected_items]
        if not selected_files:
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_files"))
            return

        # 如果多個檔案選擇，提示只支援單一檔案編輯備註
        if len(selected_files) > 1:
            CustomMessageBox.show_warning(self, self.get_text("warning"), self.get_text("select_single_file"))
            return

        file_path = selected_files[0]
        current_note = self.file_manager.get_note(file_path)

        # 創建編輯備註的窗
        note_window = create_modal_dialog(self, self.get_text("edit_note"), 400, 500)  # 增加高度到500

        # 指標籤
        label_instruction = tb.Label(
            note_window,
            text=self.get_text("edit_note_for_file").format(file=os.path.basename(file_path)),
            wraplength=380
        )
        label_instruction.pack(pady=10, padx=10)

        # 備註輸入框 - 使用Text替代Entry
        note_frame = tb.Frame(note_window)
        note_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        
        self.note_text = tb.Text(note_frame, font=("Segoe UI", 14), height=12, width=50)
        self.note_text.pack(fill=tk.BOTH, expand=True)
        if current_note:
            self.note_text.insert("1.0", current_note)

        def save_note():
            note = self.note_text.get("1.0", "end-1c").strip()
            if len(note) > 500:
                CustomMessageBox.show_warning(note_window, self.get_text("warning"), self.get_text("note_too_long"))
                return
            
            # 直接使用完整路徑設置備註
            self.file_manager.set_note(file_path, note)
            
            CustomMessageBox.show_info(note_window, self.get_text("success"), self.get_text("note_updated"))
            self.update_file_list_by_tag_selection()
            self.update_edit_buttons_state()
            note_window.destroy()

        # 保按鈕
        btn_save = tb.Button(
            note_window,
            text=self.get_text("save"),  # 這裡也要修改
            command=save_note,
            bootstyle=SUCCESS
        )
        btn_save.pack(pady=10)

        self.wait_window(note_window)  # 阻塞主視窗，直到 note_window 關閉

    def edit_single_tag(self, tag):
        def save_new_tag():
            new_tag = new_tag_entry.get().strip()
            if not new_tag:
                CustomMessageBox.show_info(rename_window, self.get_text("info"), self.get_text("enter_tag_name"))
                return
            self.file_manager.rename_tag(tag, new_tag)
            CustomMessageBox.show_info(rename_window, self.get_text("success"), 
                                     self.get_text("tag_renamed").format(old=tag, new=new_tag))
            rename_window.destroy()
            self.refresh_tag_list()

        rename_window = create_modal_dialog(self, self.get_text("rename_tag_title").format(tag=tag), 300, 200)  # 增加高度到200

        tk.Label(rename_window, text=self.get_text("rename_tag_prompt").format(tag=tag), 
                font=("Segoe UI", 14)).pack(pady=10)
        new_tag_entry = ttk.Entry(rename_window, font=("Segoe UI", 14))  # 修改為14
        new_tag_entry.pack(fill=tk.X, padx=10, pady=10)
        new_tag_entry.insert(0, tag)

        save_btn = tb.Button(rename_window, text=self.get_text("save"), command=save_new_tag, bootstyle="success")
        save_btn.pack(pady=10)

        self.wait_window(rename_window)  # 阻塞主視窗，直到 rename_window 關閉

    def edit_tag_context(self, event=None):
        widget = event.widget
        if widget == self.tag_list:
            # 編輯標籤：重新命名
            selected_tags = [self.tag_list.item(i, 'text') for i in self.tag_list.selection()]
            if not selected_tags:
                CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_tag"))
                return
            tag = selected_tags[0]
            self.edit_single_tag(tag)
            return
        elif widget == self.file_tree:
            pass  # 已經有獨立的新和移除標籤功能
        else:
            return

    def show_tag_context_menu(self, event):
        """處理標籤列表的右鍵選單"""
        # 先獲取點擊的項目
        item = self.tag_list.identify_row(event.y)
        if item:
            # 清除現有選擇並選中該項目
            self.tag_list.selection_set(item)
            
            # 更新中的標籤
            self.selected_tag = self.tag_list.item(item, 'text')
            
            # 創建右鍵選單
            self.tag_context_menu = tb.Menu(self, tearoff=0)
            self.tag_context_menu.add_command(label=self.get_text("delete_tag"), command=self.delete_selected_tag)
            self.tag_context_menu.add_command(label=self.get_text("rename_tag"), command=self.rename_selected_tag)
            self.tag_context_menu.add_separator()
            self.tag_context_menu.add_command(label=self.get_text("change_color"), command=self.change_tag_color)
            
            # 顯示選單
            self.tag_context_menu.post(event.x_root, event.y_root)

    def delete_selected_tag(self):
        selected_tags = [self.tag_list.item(item, 'text') for item in self.tag_list.selection()]
        if not selected_tags:
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_tag"))
            return
        
        # 如果只選擇了一個標籤
        if len(selected_tags) == 1:
            confirm = CustomMessageBox.show_question(
                self,
                self.get_text("delete_tag_window"),
                self.get_text("confirm_delete_tag").format(tag=selected_tags[0]),
                width=400,
                height=200
            )
        else:
            # 如果選擇了多個標籤
            confirm = CustomMessageBox.show_question(
                self,
                self.get_text("delete_tag_window"),
                self.get_text("confirm_delete_multiple_tags").format(count=len(selected_tags)),
                width=400,
                height=200
            )
        
        if confirm:
            for tag in selected_tags:
                self.file_manager.delete_tag(tag)
            self.refresh_tag_list()
            # 設置未標籤標記並更新界面
            self.showing_untagged = True
            self.current_tag_selection = None
            self.list_untagged_files()
            self.update_edit_buttons_state()

    def rename_selected_tag(self):
        """重命名選中的標籤"""
        selection = self.tag_list.selection()
        if not selection:
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_tag"))
            return
        
        selected_tag = self.tag_list.item(selection[0], 'text')
        
        def save_new_tag():
            new_tag = new_tag_entry.get().strip()
            if not new_tag:
                CustomMessageBox.show_info(rename_window, self.get_text("info"), self.get_text("enter_tag_name"))
                return
            self.file_manager.rename_tag(selected_tag, new_tag)
            CustomMessageBox.show_info(rename_window, self.get_text("success"), 
                                     self.get_text("tag_renamed").format(old=selected_tag, new=new_tag))
            rename_window.destroy()
            self.refresh_tag_list()
            # 设置未标签标记并更新界面
            self.showing_untagged = True
            self.current_tag_selection = None
            self.list_untagged_files()
            self.update_edit_buttons_state()

        rename_window = create_modal_dialog(self, self.get_text("rename_tag_title").format(tag=selected_tag), 300, 200)  # 增加高度到200

        tk.Label(rename_window, text=self.get_text("rename_tag_prompt").format(tag=selected_tag), 
                font=("Segoe UI", 14)).pack(pady=10)
        new_tag_entry = ttk.Entry(rename_window, font=("Segoe UI", 14))
        new_tag_entry.pack(fill=tk.X, padx=10, pady=10)
        new_tag_entry.insert(0, selected_tag)

        save_btn = tb.Button(rename_window, text=self.get_text("save"), command=save_new_tag, bootstyle="success")
        save_btn.pack(pady=10)

        self.wait_window(rename_window)

    def get_current_selected_tag(self):
        """
        獲取當前選中標籤。
        優先使用已設置的 self.selected_tag，若未設置則從 Treeview 獲取。
        """
        if hasattr(self, 'selected_tag') and self.selected_tag:
            return self.selected_tag
        selection = self.tag_list.selection()
        if selection:
            return self.tag_list.item(selection[0], 'text')
        return None

    def show_tooltip(self, event):
        widget = event.widget
        item_id = widget.identify_row(event.y)
        if item_id:
            full_path = widget.item(item_id, 'values')[0]
            note = self.file_manager.get_note(full_path)

            if note:
                tooltip_text = f"{self.get_text('note_prefix')}\n{note}"  # 添加換行符
                # 設置有備註的 Tooltip 樣式
                self.tooltip.config(
                    text=tooltip_text,
                    foreground="white",
                    background="#007ACC",  # 藍色背景
                    font=("Segoe UI", 14, "bold"),
                    wraplength=400  # 添加自動換行
                )
            else:
                # 顯示完整路徑
                tooltip_text = full_path
                # 設置無備註的 Tooltip 樣式
                self.tooltip.config(
                    text=tooltip_text,
                    foreground="white",
                    background="#333333",  # 深灰色背景
                    font=("Segoe UI", 14, "normal"),
                    wraplength=400  # 添加自動換行
                )
            
            # 計算 Tooltip 位置
            window_x = self.winfo_rootx()
            window_y = self.winfo_rooty()

            x = event.x_root - window_x + 20
            y = event.y_root - window_y + 20

            window_width = self.winfo_width()
            window_height = self.winfo_height()

            self.tooltip.update_idletasks()
            tooltip_width = self.tooltip.winfo_reqwidth()
            tooltip_height = self.tooltip.winfo_reqheight()

            if x + tooltip_width > window_width:
                x = window_width - tooltip_width - 10
            if y + tooltip_height > window_height:
                y = window_height - tooltip_height - 10

            self.tooltip.place(x=x, y=y)
            self.tooltip.lift()
        else:
            self.hide_tooltip()

    def hide_tooltip(self, event=None):
        self.tooltip.place_forget()

    def show_context_menu(self, event):
        """處理檔案列表的右鍵選單"""
        # 先獲取點擊的項目
        item = self.file_tree.identify_row(event.y)
        if item:
            # 如果點擊項目不在當前選中項目中，則清除選擇並選中該項目
            if item not in self.file_tree.selection():
                self.file_tree.selection_set(item)
            
            # 更新選中的檔案列表
            self.selected_files = [self.file_tree.item(i, 'values')[0] for i in self.file_tree.selection()]
            # 顯示單
            self.context_menu.post(event.x_root, event.y_root)

    def open_file_location(self):
        """在檔案總管中打開檔案位置並選中檔案"""
        if hasattr(self, 'selected_files'):
            for selected_file in self.selected_files:
                try:
                    if not open_file_in_explorer(selected_file):
                        # 如果 Windows API 调用失败，回退到使用 startfile
                        os.startfile(os.path.dirname(selected_file))
                except Exception as e:
                    CustomMessageBox.show_error(self, self.get_text("error"), f"{self.get_text('unable_to_open_folder')}: {str(e)}")

    def restore_data(self):
        """還原標籤資料"""
        restore_points = self.file_manager.get_restore_points()
        if not restore_points:
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("no_restore_points"))
            return

        def select_restore_point():
            selected = restore_listbox.curselection()
            if not selected:
                CustomMessageBox.show_info(restore_window, self.get_text("info"), 
                                         self.get_text("select_restore_point"))
                return

            # 使用完整檔名來還原，但只顯示日期時間
            restore_file = os.path.join(self.file_manager.backup_dir, restore_points[selected[0]])
            restore_window.destroy()
            
            # 創建處理中的動畫對話框
            processing_dialog, progress = self.create_processing_dialog(self)
            
            def process_restore():
                try:
                    # 檢查是否已取消
                    if self.processing_cancelled:
                        logger.info("還原處理已被取消")
                        return
                    
                    # 停止所有背景操作
                    logger.info("正在停止更新管理器...")
                    self.update_manager.stop()
                    
                    logger.info("正在停止檔案監控...")
                    self.file_manager.stop_monitoring()
                    
                    # 等待確保更新管理器完全停止
                    while not self.processing_cancelled and (self.update_manager.running or self.update_manager.is_updating):
                        time.sleep(0.1)
                    
                    if self.processing_cancelled:
                        logger.info("還原處理已被取消")
                        return
                    
                    logger.info("更新管理器已停止")
                    
                    # 在主線程中清空列表
                    def clear_ui():
                        if self.processing_cancelled:
                            return
                        
                        # 清空檔案列表和標籤列表
                        self.file_tree.delete(*self.file_tree.get_children())
                        self.tag_list.delete(*self.tag_list.get_children())
                        
                        # 清空搜尋框
                        self.search_var.set("")
                        self.tag_search_var.set("")
                        
                        # 重置檔案類型選擇
                        self.file_type_var.set("ALL")
                        
                        # 清空其他相關變量
                        self.all_files = []
                        self.current_batch = 0
                        self.is_loading_cancelled = True
                        
                        logger.info("UI 已清空")
                    
                    # 在主線程中執行 UI 清理
                    self.after(0, clear_ui)
                    
                    # 等待 UI 更新完成
                    time.sleep(0.5)
                    
                    if self.processing_cancelled:
                        logger.info("還原處理已被取消")
                        return
                    
                    logger.info("開始還原資料...")
                    # 還原資料
                    self.file_manager.restore_db(restore_file)
                    
                    # 重新設置資料夾路徑以刷新資料
                    self.file_manager.set_folder_paths(self.file_manager.folder_paths)
                    
                    # 在主線程中更新 UI
                    if not self.processing_cancelled:
                        self.after(0, lambda: self.update_ui_after_processing(processing_dialog))
                    
                except Exception as e:
                    logger.error(f"還原資料時發生錯誤: {str(e)}")
                    # 如果發生錯誤，在主線程中顯示錯誤訊息
                    if not self.processing_cancelled:
                        self.after(0, lambda: self.handle_processing_error(e, processing_dialog))
            
            # 在背景線程中處理還原
            threading.Thread(target=process_restore, daemon=True).start()

        restore_window = create_modal_dialog(self, self.get_text("select_restore_point"), 400, 500)

        # 創建主框架
        main_frame = tb.Frame(restore_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 標題標籤
        title_label = tb.Label(
            main_frame,
            text=self.get_text("select_restore_point") + ":",
            font=("Segoe UI", 14)
        )
        title_label.pack(pady=(0, 10))

        # 創建列表框架
        list_frame = tb.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 添加滾動條
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 還原點列表
        restore_listbox = tk.Listbox(
            list_frame,
            font=("Segoe UI", 14),
            yscrollcommand=scrollbar.set
        )
        restore_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=restore_listbox.yview)

        # 只顯示日期時間部分
        for rp in restore_points:
            # 從檔名中提取日期時間 (格式: YYYYMMDD_HHMMSS_file_tags.json)
            date_time = rp.split('_file_tags.json')[0]  # 移除檔案名稱部分
            formatted_time = f"{date_time[:4]}/{date_time[4:6]}/{date_time[6:8]} {date_time[9:11]}:{date_time[11:13]}:{date_time[13:15]}"
            restore_listbox.insert(tk.END, formatted_time)

        # 按鈕框架
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        # 還原按鈕
        restore_btn = tb.Button(
            button_frame,
            text=self.get_text("restore"),
            command=select_restore_point,
            bootstyle="warning",
            width=15
        )
        restore_btn.pack(pady=5)

        self.wait_window(restore_window)

    def open_manage_folders_dialog(self):
        """打開管理多個資料夾的對話框"""
        dialog = create_modal_dialog(self, self.get_text("manage_folders"), 500, 400)

        # 資料夾列表框架
        folder_frame = tb.Frame(dialog)
        folder_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 資料夾列
        self.manage_folder_list = tk.Listbox(folder_frame, selectmode=tk.SINGLE, font=("Segoe UI", 14))  # 修改為14
        self.manage_folder_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        folder_scrollbar = tk.Scrollbar(folder_frame, orient=tk.VERTICAL, command=self.manage_folder_list.yview)
        self.manage_folder_list.config(yscrollcommand=folder_scrollbar.set)
        folder_scrollbar.pack(side=tk.LEFT, fill=tk.Y)

        # 填充現有資料夾
        for folder in self.file_manager.folder_paths:
            self.manage_folder_list.insert(tk.END, folder)

        # 按鈕框架
        manage_button_frame = tb.Frame(dialog)
        manage_button_frame.pack(pady=10)

        add_folder_btn = tb.Button(
            manage_button_frame,
            text=self.get_text("add_folder"),
            command=self.add_folder_to_manager,
            bootstyle="primary-outline",
            width=15
        )
        add_folder_btn.pack(side=tk.LEFT, padx=5)

        remove_folder_btn = tb.Button(
            manage_button_frame,
            text=self.get_text("remove_folder"),
            command=self.remove_selected_folder_from_manager,
            bootstyle="danger-outline",
            width=15
        )
        remove_folder_btn.pack(side=tk.LEFT, padx=5)

        # 保存按鈕
        save_btn = tb.Button(
            dialog,
            text=self.get_text("save"),
            command=lambda: self.save_folders(dialog),
            bootstyle="success",
            width=10
        )
        save_btn.pack(pady=10)

    def add_folder_to_manager(self):
        """在管理對話框中新增資料夾"""
        # 修改：使用翻譯文字
        folder_path = filedialog.askdirectory(title=self.get_text("select_folder"), parent=self)
        if folder_path:
            if folder_path not in self.file_manager.folder_paths:
                self.manage_folder_list.insert(tk.END, folder_path)
            else:
                # 修改：使用翻譯文字
                CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("folder_exists"))

    def remove_selected_folder_from_manager(self):
        """從管理對話框中移除選取的資料夾"""
        selected = self.manage_folder_list.curselection()
        if selected:
            index = selected[0]
            folder_path = self.manage_folder_list.get(index)
            # 修改：使用翻譯文字
            confirm = CustomMessageBox.show_question(
                self,
                self.get_text("remove_folder"),
                self.get_text("confirm_remove_folder"),
                width=400,
                height=200
            )
            if confirm:
                self.manage_folder_list.delete(index)

    def save_folders(self, dialog):
        """保存管理對話框中的資料夾設定"""
        folders = list(self.manage_folder_list.get(0, tk.END))
        
        # 關閉資料夾管理對話框
        dialog.destroy()
        
        # 重置取消標記
        self.processing_cancelled = False
        
        # 創建處理中的動畫對話框
        processing_dialog, progress = self.create_processing_dialog(self)
        
        def process_folders():
            try:
                # 檢查是否已取消
                if self.processing_cancelled:
                    logger.info("處理已被取消")
                    return
                
                # 停止所有背景操作
                logger.info("正在停止更新管理器...")
                self.update_manager.stop()  # 停止更新管理器
                
                logger.info("正在停止檔案監控...")
                self.file_manager.stop_monitoring()  # 停止檔案監控
                
                # 等待確保更新管理器完全停止
                while not self.processing_cancelled and (self.update_manager.running or self.update_manager.is_updating):
                    time.sleep(0.1)
                
                if self.processing_cancelled:
                    logger.info("處理已被取消")
                    return
                
                logger.info("更新管理器已停止")
                
                # 在主線程中清空列表
                def clear_ui():
                    if self.processing_cancelled:
                        return
                    
                    # 清空檔案列表和標籤列表
                    self.file_tree.delete(*self.file_tree.get_children())
                    self.tag_list.delete(*self.tag_list.get_children())
                    
                    # 清空搜尋框
                    self.search_var.set("")
                    self.tag_search_var.set("")
                    
                    # 重置檔案類型選擇
                    self.file_type_var.set("ALL")
                    
                    # 清空其他相關變量
                    self.all_files = []
                    self.current_batch = 0
                    self.is_loading_cancelled = True
                    
                    logger.info("UI 已清空")
                
                # 在主線程中執行 UI 清理
                self.after(0, clear_ui)
                
                # 等待 UI 更新完成
                time.sleep(0.5)
                
                if self.processing_cancelled:
                    logger.info("處理已被取消")
                    return
                
                logger.info("開始處理新的資料夾設定...")
                # 更新檔案管理器的資料夾路徑
                self.file_manager.folder_paths = folders
                
                # 保存配置
                self.save_config()
                
                # 重新掃描檔案
                self.file_manager.set_folder_paths(folders)
                
                # 在主線程中更新 UI
                if not self.processing_cancelled:
                    self.after(0, lambda: self.update_ui_after_processing(processing_dialog))
                
            except Exception as e:
                logger.error(f"處理資料夾時發生錯誤: {str(e)}")
                # 如果發生錯誤，在主線程中顯示錯誤訊息
                if not self.processing_cancelled:
                    self.after(0, lambda: self.handle_processing_error(e, processing_dialog))
        
        # 在背景線程中處理資料
        threading.Thread(target=process_folders, daemon=True).start()

    def update_ui_after_processing(self, processing_dialog):
        """處理完成後更新 UI"""
        try:
            logger.info("正在更新 UI...")
            # 關閉處理中的對話框
            processing_dialog.destroy()
            
            # 重新初始化更新管理器
            logger.info("重新初始化更新管理器...")
            self.update_manager = BackgroundUpdateManager(self)
            
            # 重新啟動檔案監控
            logger.info("重新啟動檔案監控...")
            self.start_file_monitoring()
            
            # 更新 UI
            logger.info("更新標籤列表...")
            self.refresh_tag_list()
            logger.info("列出未標籤檔案...")
            self.list_untagged_files()
            logger.info("更新按鈕狀態...")
            self.update_edit_buttons_state()
            logger.info("過濾檔案列表...")
            self.filter_file_list()
            
            logger.info("UI 更新完成")
            # 顯示成功訊息
            CustomMessageBox.show_info(self, self.get_text("success"), self.get_text("folders_updated"))
            
        except Exception as e:
            logger.error(f"更新 UI 時發生錯誤: {str(e)}")
            # 如果發生錯誤，顯示錯誤訊息
            self.handle_processing_error(e, processing_dialog)

    def handle_processing_error(self, error, processing_dialog):
        """處理錯誤"""
        # 關閉處理中的對話框
        processing_dialog.destroy()
        
        # 顯示錯誤訊息
        CustomMessageBox.show_error(
            self,
            self.get_text("error"),
            f"{self.get_text('folder_update_failed')}: {str(error)}"
        )

    def add_folder_dialog(self):
        """初次啟動時提示選擇資料夾"""
        self.open_manage_folders_dialog()

    def apply_file_filter(self, files):
        """根據當前篩選條更新檔案列表"""
        search_text = self.search_var.get().strip().lower()
        filtered_files = [file for file in files if search_text in file.lower()]

        self.file_tree.delete(*self.file_tree.get_children())
        self.all_files = filtered_files
        self.current_batch = 0
        self.load_files_in_batches()

        # 如果有預期選擇的檔案，嘗試選中它
        if self.intended_selection and self.intended_selection in filtered_files:
            # 延遲選操作，等文件載入完成後再選中
            self.after(100, self.select_intended_file)

        # 更新按鈕狀態
        self.update_edit_buttons_state()

    def on_window_resize(self, event):
        """當視窗大小改變時調整檔案列表欄位寬度"""
        # 只調整檔案列表的寬度，不調整標籤列表
        if event.widget == self and event.width > 0:
            self.update_treeview_column_width()

    def update_treeview_column_width(self):
        """更新檔案列表的欄位寬度"""
        # 只調整檔案列表的寬度，不調整標籤列表
        if hasattr(self, 'file_tree'):
            self.file_tree.column("#0", width=self.file_frame.winfo_width() - 20)  # 減去滾動條的寬度

    # Batch operations are already handled in FileManager

    def on_file_double_click(self, event):
        """處理檔案列表的雙擊事件"""
        selected_items = self.file_tree.selection()
        if selected_items:
            # 获取选中文件的路径
            selected_file = self.file_tree.item(selected_items[0], 'values')[0]
            try:
                # 使用系统默认程序打开文件
                os.startfile(selected_file)
            except Exception as e:
                CustomMessageBox.show_error(self, self.get_text("error"), f"{self.get_text('unable_to_open_file')}: {str(e)}")

    def start_file_monitoring(self):
        """啟動檔案監控"""
        self.file_manager.start_monitoring(self.handle_file_changes)

    def handle_file_changes(self):
        """處理檔案變更"""
        # 在後台處理文件更
        self.update_manager.queue_update("scan_files", self.file_manager.folder_paths)

    def destroy(self):
        """關閉應用程式時停止監控後台更新"""
        self.update_manager.stop()
        self.file_manager.stop_monitoring()
        super().destroy()

    def update_file_type_options(self):
        """根據所有檔案列表更新檔案類型下拉式選單的選項"""
        all_file_types = self.file_manager.get_all_file_types()
        sorted_file_types = sorted(all_file_types)
        # 使用翻譯的 "all_types" 作為第一個選項
        options = [self.get_text("all_types")] + sorted_file_types
        current_selection = self.file_type_var.get()
        self.file_type_combo['values'] = options

        # 如果當前選擇存在於選項中，保持選擇
        if current_selection in options:
            self.file_type_var.set(current_selection)
        else:
            # 預設選擇第一個選項（全部類型）
            self.file_type_var.set(options[0])

    def load_last_folders(self):
        """載入最後選擇的資料夾和檔案類型"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                    last_folders = config.get('last_folders', [])
                    last_file_type = config.get('last_file_type', self.get_text("all_types"))
                    valid_folders = [folder for folder in last_folders if os.path.exists(folder)]
                    if valid_folders:
                        # 使用後線程處理文件掃描
                        self.update_manager.queue_update("scan_files", valid_folders)
                        # 設定檔案類型選擇
                        if hasattr(self, 'file_type_combo'):
                            self.file_type_var.set(last_file_type if last_file_type in self.file_type_combo['values'] else self.get_text("all_types"))
            except Exception as e:
                print(f"載最後資料夾時發生錯誤: {str(e)}")
        else:
            # 如果沒有配置文件，也顯示空的未標籤列表
            self.after(100, self.list_untagged_files)

    def save_last_folders(self, folder_paths, file_type):
        """保存最後選擇的資料夾和檔案類型到 config.json"""
        # 如果 file_type 是 "ALL"，則使用翻譯的文字
        if file_type == "ALL":
            file_type = self.get_text("all_types")
            
        config = {
            'language': self.current_language,  # 保存語言設置
            'last_folders': folder_paths,
            'last_file_type': file_type
        }
        with open(self.config_file, 'w', encoding='utf-8') as file:
            json.dump(config, file, ensure_ascii=False, indent=4)

    def on_drop_enter(self, event):
        """當拖曳進入項目時的處理"""
        # 設置高亮樣式
        self.tag_list.tag_configure('highlight', background='#E1F5FE')

    def on_drop_leave(self, event):
        """當拖曳離開項目時的處理"""
        # 清除所有項目的高亮效果
        for item in self.tag_list.get_children():
            current_tags = list(self.tag_list.item(item)['tags'] or ())
            if 'highlight' in current_tags:
                current_tags.remove('highlight')
            self.tag_list.item(item, tags=current_tags)

    def on_drop_motion(self, event):
        """當拖曳移動時的處理"""
        # 清除所有項目的高亮效果
        for item in self.tag_list.get_children():
            self.tag_list.item(item, tags=())
            
        # 獲取當前滑鼠位置下的項目
        y = event.widget.winfo_pointery() - event.widget.winfo_rooty()
        target_item = self.tag_list.identify_row(y)
        
        if target_item:
            # 獲取該項目原有的標籤樣式（如果有的話）
            current_tags = list(self.tag_list.item(target_item)['tags'] or ())
            if 'highlight' not in current_tags:
                current_tags.append('highlight')
            
            # 設置高亮效果
            self.tag_list.tag_configure('highlight', background='#E1F5FE')
            self.tag_list.item(target_item, tags=current_tags)

    def on_drop(self, event):
        """處理檔案拖放事件"""
        # 清除高亮效果
        self.on_drop_leave(event)
        
        # 獲取拖放的檔案路徑
        data = event.data
        
        # 解析檔案路徑
        files = []
        current_file = []
        in_curly_brace = False
        i = 0
        
        while i < len(data):
            char = data[i]
            
            # 處理大括號開始
            if char == '{':
                in_curly_brace = True
                i += 1
                continue
            
            # 處理大括號結束
            elif char == '}':
                in_curly_brace = False
                if current_file:
                    files.append(''.join(current_file).strip())
                    current_file = []
                i += 1
                continue
            
            # 處理檔案分隔
            elif char == ' ' and not in_curly_brace:
                if current_file:
                    files.append(''.join(current_file).strip())
                    current_file = []
                i += 1
                continue
            
            # 收集檔案路徑字符
            current_file.append(char)
            i += 1
        
        # 處理最後一個檔案
        if current_file:
            files.append(''.join(current_file).strip())
        
        # 清理和標準化路徑
        files = [f.strip('{}').strip('"').strip() for f in files if f.strip('{}').strip('"').strip()]
        files = [os.path.normpath(f) for f in files]
        
        # 獲取拖放位置的標籤項目
        target_item = self.tag_list.identify_row(event.widget.winfo_pointery() - self.tag_list.winfo_rooty())
        if not target_item:
            CustomMessageBox.show_warning(self, self.get_text("warning"), self.get_text("drop_on_tag"))
            return
            
        target_tag = self.tag_list.item(target_item, 'text')
        
        # 確認是否要添加標籤
        confirm = CustomMessageBox.show_question(
            self,
            self.get_text("add_tag"),
            self.get_text("confirm_add_tag_to_files").format(
                tag=target_tag,
                count=len(files)
            ),
            width=400,
            height=200
        )
        
        if confirm:
            added_count = 0
            # 為每個檔案添加標籤
            for file_path in files:
                # 檢查檔案是否存在
                if os.path.exists(file_path):
                    # 檢查檔案是否在管理的資料夾中
                    is_in_managed_folder = False
                    for folder_path in self.file_manager.folder_paths:
                        if file_path.lower().startswith(folder_path.lower()):
                            is_in_managed_folder = True
                            break
                    
                    if not is_in_managed_folder:
                        # 如果檔案不在管理的資料夾中，將其所在資料夾添加到管理資料夾
                        folder_path = os.path.dirname(file_path)
                        if folder_path not in self.file_manager.folder_paths:
                            self.file_manager.folder_paths.append(folder_path)
                            self.file_manager.set_folder_paths(self.file_manager.folder_paths)
                    
                    # 添加標籤
                    self.file_manager.add_tag(file_path, target_tag)
                    added_count += 1
            
            if added_count > 0:
                # 更新界面
                self.refresh_tag_list()
                self.update_file_list_by_tag_selection()
                self.update_edit_buttons_state()
                
                # 顯示成功訊息
                CustomMessageBox.show_info(
                    self,
                    self.get_text("success"),
                    self.get_text("files_tagged_success").format(count=added_count, tag=target_tag)
                )
            else:
                CustomMessageBox.show_warning(
                    self,
                    self.get_text("warning"),
                    self.get_text("no_files_tagged")
                )

    def load_files_in_batches(self):
        """分批載入檔案到Treeview中以保持UI流暢"""
        if self.is_loading_cancelled:
            return

        if not hasattr(self, 'all_files') or self.all_files is None:
            self.all_files = []
            return

        start = self.current_batch * self.load_batch_size
        end = min(start + self.load_batch_size, len(self.all_files))
        batch = self.all_files[start:end]

        # 如果是第一批，清空現有項目
        if self.current_batch == 0:
            self.file_tree.delete(*self.file_tree.get_children())
            
        # 禁用重繪以提高性能
        self.file_tree.configure(height=1)  # 臨時減少高度以加快插入

        # 批量插入檔案
        items_to_insert = []
        for full_path in batch:
            file_name = os.path.basename(full_path)
            items_to_insert.append(('', 'end', file_name, (full_path,)))

        # 使用批量插入
        for item in items_to_insert:
            item_id = self.file_tree.insert(item[0], item[1], text=item[2], values=item[3])
            if hasattr(self, 'selected_paths') and item[3][0] in self.selected_paths:
                self.file_tree.selection_add(item_id)

        # 恢復正常高度
        self.file_tree.configure(height=20)  # 或其他適當的高度

        self.current_batch += 1
        
        # 如果還有更多檔案要載入，安排下一批
        if end < len(self.all_files) and not self.is_loading_cancelled:
            self.after(100, self.load_files_in_batches)  # 增加延遲時間到100ms以減少UI凍結

    def cancel_current_loading(self):
        """取消当前的文件加载过程"""
        self.is_loading_cancelled = True
        self.current_batch = 0

    def start_new_loading(self):
        """开始新的文件加载过程"""
        self.is_loading_cancelled = False
        self.current_batch = 0
        self.file_tree.delete(*self.file_tree.get_children())
        self.load_files_in_batches()

    def show_theme_dialog(self):
        """顯示主題選擇對話框"""
        dialog = create_modal_dialog(self, self.get_text("select_theme"), 400, 600)
        
        # 獲取所有可用的主題並確保 superhero 在第一位
        available_themes = sorted(self.style.theme_names())
        if 'superhero' in available_themes:
            available_themes.remove('superhero')
            available_themes.insert(0, 'superhero')
        
        # 創建主框架
        main_frame = tb.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 標題標籤
        title_label = tb.Label(
            main_frame,
            text=self.get_text("select_theme") + ":",
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 創建滾動框架
        scroll_frame = tb.Frame(main_frame)
        scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建畫布和滾動條
        canvas = tk.Canvas(scroll_frame)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        theme_frame = tb.Frame(canvas)
        
        # 配置滾動
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 打包滾動元件
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 創建窗口在畫布上
        canvas_frame = canvas.create_window((0, 0), window=theme_frame, anchor="nw")
        
        def get_theme_background_color(theme_name):
            """獲取主題的背景顏色"""
            theme_colors = {
                'superhero': '#2b3e50',
                'darkly': '#222222',
                'cyborg': '#060606',
                'vapor': '#0d0d2b',
                'solar': '#002b36',
                'slate': '#272b30',
                'morph': '#373a3c',
                'journal': '#f7f7f7',
                'litera': '#f8f9fa',
                'lumen': '#ffffff',
                'minty': '#fff',
                'pulse': '#fff',
                'sandstone': '#fff',
                'united': '#fff',
                'yeti': '#fff',
                'cosmo': '#fff',
                'flatly': '#fff',
                'spacelab': '#fff',
                'default': '#ffffff'
            }
            return theme_colors.get(theme_name, '#ffffff')
        
        def apply_theme(theme_name):
            # 先儲存主題設定
            self.current_theme = theme_name
            # 更新預設標籤顏色
            new_bg_color = get_theme_background_color(theme_name)
            self.file_manager.set_default_color(new_bg_color)

            # 然後切換主題
            self.style.theme_use(theme_name)
            
            # 保存當前選中的標籤
            selected_tags = [self.tag_list.item(item, 'text') for item in self.tag_list.selection()]
            
            # 更新標籤的顏色
            for item in self.tag_list.get_children():
                tag_text = self.tag_list.item(item, 'text')
                color = self.file_manager.get_tag_color(tag_text)
                tag_style = f'tag_{hash(tag_text)}'
                
                # 只有當顏色不是預設顏色時才設置標籤樣式
                if color != self.file_manager.default_color:
                    self.tag_list.tag_configure(tag_style, foreground=color)
                    self.tag_list.item(item, tags=(tag_style,))
                else:
                    # 如果是預設顏色，移除標籤樣式
                    self.tag_list.item(item, tags=())
            
            # 恢復選中的標籤
            for item in self.tag_list.get_children():
                if self.tag_list.item(item, 'text') in selected_tags:
                    self.tag_list.selection_add(item)
            
            # 最後保存配置
            self.save_config()
            
            # 設置全局字體大小
            default_font = ('TkDefaultFont', 14)
            text_font = ('TkTextFont', 14)
            fixed_font = ('TkFixedFont', 14)
            
            self.option_add('*Font', default_font)
            self.option_add('*Text.font', text_font)  # 使用option_add來設置Text部件的字體
            self.style.configure('.', font=default_font)
            self.style.configure('Treeview', font=default_font)
            self.style.configure('Treeview.Heading', font=default_font)
            self.style.configure('TCombobox', font=default_font)
            self.style.configure('TButton', font=default_font)
            self.style.configure('TLabel', font=default_font)
            self.style.configure('TEntry', font=default_font)
            self.style.configure('TCheckbutton', font=default_font)
            self.style.configure('TRadiobutton', font=default_font)
            self.style.configure('TNotebook.Tab', font=default_font)
            self.style.configure('TMenubutton', font=default_font)
            self.style.configure('TSpinbox', font=default_font)

            # 關閉對話框
            dialog.destroy()
            
            # 获取系统 DPI 缩放
            try:
                import ctypes
                user32 = ctypes.windll.user32
                user32.SetProcessDPIAware()
                dpi = user32.GetDpiForSystem()
                dpi_scale = dpi / 96.0  # 96 是标准 DPI
                if dpi_scale < 1.0:
                    dpi_scale = 1.0
            except:
                dpi_scale = 1.0

            # 根据 DPI 缩放调整字体大小和 rowheight
            base_rowheight = 30
            scaled_rowheight = int(base_rowheight * dpi_scale)
            
            tagStyle = ttk.Style()
            tagStyle.configure("Treeview", rowheight=scaled_rowheight, font=("Segoe UI", 14))
            tagStyle.configure("Treeview.Item", rowheight=scaled_rowheight, font=("Segoe UI", 14))
        
        # 獲取主題翻譯
        theme_translations = LANGUAGES[self.current_language].get("theme_names", {})
        
        # 創建主題按鈕
        for theme_name in available_themes:
            # 獲取翻譯名稱，如果沒有翻譯就使用原名
            translated_name = theme_translations.get(theme_name, theme_name)
            # 如果是當前主題，在翻譯名稱後添加"(當前)"
            if theme_name == self.current_theme:
                button_text = f"{theme_name} ({translated_name}) ★"
            else:
                button_text = f"{theme_name} ({translated_name})"
            
            theme_btn = tb.Button(
                theme_frame,
                text=button_text,
                command=lambda t=theme_name: apply_theme(t),
                bootstyle="primary-outline" if theme_name != self.current_theme else "primary",
                width=30
            )
            theme_btn.pack(pady=5)
        
        # 更新滾動區域
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        theme_frame.bind("<Configure>", configure_scroll_region)
        
        # 設置畫布大小
        def configure_canvas(event):
            canvas.itemconfig(canvas_frame, width=event.width)
        
        canvas.bind("<Configure>", configure_canvas)
        
        # 綁定滾輪事件
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # 解除綁定當對話框關閉時
        dialog.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

    def create_processing_dialog(self, parent):
        """創建處理中的動畫對話框"""
        dialog = create_modal_dialog(parent, self.get_text("processing"), 300, 150)
        
        # 設置取消標記
        self.processing_cancelled = False
        
        # 進度條
        progress = ttk.Progressbar(
            dialog,
            mode='indeterminate',
            length=200
        )
        progress.pack(pady=20)
        progress.start(10)  # 開始動畫
        
        # 處理中的文字
        label = tb.Label(
            dialog,
            text=self.get_text("processing_folders"),
            font=("Segoe UI", 14)
        )
        label.pack(pady=10)
        
        # 綁定關閉事件
        dialog.protocol("WM_DELETE_WINDOW", lambda: self.cancel_processing(dialog))
        
        return dialog, progress

    def cancel_processing(self, dialog):
        """取消處理操作"""
        logger.info("用戶取消處理操作")
        self.processing_cancelled = True
        
        # 停止所有背景操作
        logger.info("正在停止更新管理器...")
        self.update_manager.stop()
        
        logger.info("正在停止檔案監控...")
        self.file_manager.stop_monitoring()
        
        # 關閉對話框
        dialog.destroy()
        
        # 重新初始化更新管理器和檔案監控
        logger.info("重新初始化更新管理器...")
        self.update_manager = BackgroundUpdateManager(self)
        
        logger.info("重新啟動檔案監控...")
        self.start_file_monitoring()
        
        # 更新界面
        self.refresh_tag_list()
        self.filter_file_list()
        self.update_edit_buttons_state()

    def change_tag_color(self):
        """更改標籤顏色"""
        selection = self.tag_list.selection()
        if not selection:
            CustomMessageBox.show_info(self, self.get_text("info"), self.get_text("select_tag"))
            return
        
        selected_tag = self.tag_list.item(selection[0], 'text')
            
        # 創建顏色選擇對話框
        color = colorchooser.askcolor(
            title=self.get_text("select_color"),
            color=self.file_manager.get_tag_color(selected_tag)
        )
        
        if color[1]:  # color[1] 是十六進制的顏色代碼
            # 更新標籤顏色
            self.file_manager.set_tag_color(selected_tag, color[1])
            # 刷新標籤列表以顯示新顏色
            self.refresh_tag_list()

    def clear_file_list(self):
        """清空檔案列表"""
        self.file_tree.delete(*self.file_tree.get_children())
        self.file_type_var.set(self.get_text("all_types"))
        self.all_files = []

    def handle_tag_drop(self, files, tag):
        """處理標籤拖放視窗的回調"""
        # 檢查檔案是否存在
        valid_files = []
        for file_path in files:
            if os.path.exists(file_path):
                # 檢查檔案是否在管理的資料夾中
                is_in_managed_folder = False
                for folder_path in self.file_manager.folder_paths:
                    if file_path.lower().startswith(folder_path.lower()):
                        is_in_managed_folder = True
                        break
                
                if not is_in_managed_folder:
                    # 如果檔案不在管理的資料夾中，將其所在資料夾添加到管理資料夾
                    folder_path = os.path.dirname(file_path)
                    if folder_path not in self.file_manager.folder_paths:
                        self.file_manager.folder_paths.append(folder_path)
                        self.file_manager.set_folder_paths(self.file_manager.folder_paths)
                
                valid_files.append(file_path)
        
        # 為每個有效檔案添加標籤
        for file_path in valid_files:
            self.file_manager.add_tag(file_path, tag)
        
        if valid_files:
            # 更新界面
            self.refresh_tag_list()
            self.update_file_list_by_tag_selection()
            self.update_edit_buttons_state()
            
            # 使用非模態對話框顯示成功訊息
            success_window = tk.Toplevel()
            success_window.title(self.get_text("success"))
            success_window.geometry("300x100")
            success_window.resizable(False, False)
            
            # 設置窗口樣式
            success_window.transient()  # 使窗口始終顯示在最上層
            success_window.attributes('-topmost', True)
            
            # 添加消息標籤
            message = self.get_text("files_tagged_success").format(count=len(valid_files), tag=tag)
            label = ttk.Label(success_window, text=message, wraplength=280, justify='center')
            label.pack(pady=20)
            
            # 居中顯示
            self.center_window(success_window, 300, 100)
            
            # 設置自動關閉
            success_window.after(2000, success_window.destroy)  # 2秒後自動關閉

    def show_hotkey_dialog(self):
        """顯示快捷鍵設置對話框"""
        if self.hotkey_dialog is None:
            # 創建非模態對話框
            self.hotkey_dialog = tk.Toplevel(self)
            self.hotkey_dialog.title(self.get_text("edit_hotkey"))
            self.hotkey_dialog.geometry("300x200")
            
            # 設置對話框屬性
            self.hotkey_dialog.transient(self)  # 設置父窗口
            self.hotkey_dialog.resizable(False, False)  # 禁止調整大小
            
            # 創建說明標籤
            description_label = tb.Label(
                self.hotkey_dialog,
                text=self.get_text("hotkey_description"),
                wraplength=250
            )
            description_label.pack(pady=10, padx=20)
            
            # 創建輸入框架
            input_frame = tb.Frame(self.hotkey_dialog)
            input_frame.pack(fill="x", padx=20)
            
            # 獲取當前的快捷鍵設置
            current_modifier = self.hotkey_modifier if hasattr(self, 'hotkey_modifier') else "Ctrl"
            current_key = self.hotkey_key if hasattr(self, 'hotkey_key') else "1"
            
            # 創建修飾鍵選擇下拉框
            modifier_var = tk.StringVar(value=current_modifier)
            modifier_combo = tb.Combobox(
                input_frame,
                textvariable=modifier_var,
                values=["Ctrl", "Alt", "Shift"],
                state="readonly",
                width=10
            )
            modifier_combo.pack(side="left", padx=5)
            
            # 創建 + 號標籤
            plus_label = tb.Label(input_frame, text="+")
            plus_label.pack(side="left", padx=5)
            
            # 創建按鍵選擇下拉框
            key_var = tk.StringVar(value=current_key)
            key_combo = tb.Combobox(
                input_frame,
                textvariable=key_var,
                values=[str(i) for i in range(10)] + [chr(i) for i in range(65, 91)],  # 0-9 和 A-Z
                state="readonly",
                width=10
            )
            key_combo.pack(side="left", padx=5)
            
            # 創建按鈕框架
            button_frame = tb.Frame(self.hotkey_dialog)
            button_frame.pack(fill="x", padx=20, pady=20)
            
            def save_hotkey():
                modifier = modifier_var.get()
                key = key_var.get()
                # 更新 TagDropWindow 的快捷鍵
                if hasattr(self, 'tag_drop_window'):
                    self.tag_drop_window.update_hotkey(modifier, key)
                    # 保存配置到config.json
                    self.save_config()
                self.hotkey_dialog.withdraw()  # 隱藏而不是銷毀
            
            # 創建確定按鈕
            ok_btn = tb.Button(
                button_frame,
                text=self.get_text("ok"),
                command=save_hotkey,
                bootstyle="success"
            )
            ok_btn.pack(side="left", padx=5, expand=True)
            
            # 添加關閉視窗的處理
            self.hotkey_dialog.protocol("WM_DELETE_WINDOW", lambda: self.hotkey_dialog.withdraw())
            
            # 設置視窗位置為居中
            self.center_window(self.hotkey_dialog, 300, 200)
        else:
            # 如果對話框已存在，則更新顯示並顯示它
            self.hotkey_dialog.deiconify()
            self.center_window(self.hotkey_dialog, 300, 200)

    def show_settings_dialog(self):
        # 創建設定對話框
        dialog = tk.Toplevel(self)
        dialog.title(self.get_text('settings'))
        dialog.transient(self)
        dialog.grab_set()
        
        # 設置對話框大小和位置
        width = 300
        height = 490  # 增加高度以容納新按钮
        self.center_window(dialog, width, height)
        
        # 創建主框架
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建主要按鈕
        theme_btn = ttk.Button(main_frame, text=self.get_text('theme'),
                             command=lambda: [dialog.destroy(), self.show_theme_dialog()])
        theme_btn.pack(fill=tk.X, pady=10, ipady=10)
        
        language_btn = ttk.Button(main_frame, text=self.get_text('language'),
                                command=lambda: [dialog.destroy(), self.show_language_dialog()])
        language_btn.pack(fill=tk.X, pady=10, ipady=10)
        
        hotkey_btn = ttk.Button(main_frame, text=self.get_text('hotkeys'),
                               command=lambda: [dialog.destroy(), self.show_hotkey_dialog()])
        hotkey_btn.pack(fill=tk.X, pady=10, ipady=10)
        
        folder_btn = ttk.Button(main_frame, text=self.get_text('manage_folders'),
                              command=lambda: [dialog.destroy(), self.open_manage_folders_dialog()])
        folder_btn.pack(fill=tk.X, pady=10, ipady=10)
        
        export_btn = ttk.Button(main_frame, text=self.get_text('export'),
                                command=lambda: [dialog.destroy(), self.show_export_dialog()])
        export_btn.pack(fill=tk.X, pady=10, ipady=10)
        
        restore_btn = ttk.Button(main_frame, text=self.get_text('restore'),
                                command=lambda: [dialog.destroy(), self.restore_data()])
        restore_btn.pack(fill=tk.X, pady=10, ipady=10)
        

    def show_export_dialog(self):
        """顯示匯出標籤對話框"""
        dialog = create_modal_dialog(self, self.get_text('export_tags'), 400, 150)
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建格式選擇框架
        format_frame = ttk.Frame(main_frame)
        format_frame.pack(fill=tk.X, pady=10)
        
        # 格式選擇標籤
        format_label = ttk.Label(format_frame, text=self.get_text('export_format'))
        format_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # 使用 Radio Button 替代 Combobox
        format_var = tk.StringVar(value="json")
        
        radio_frame = ttk.Frame(format_frame)
        radio_frame.pack(side=tk.LEFT)
        
        json_radio = ttk.Radiobutton(
            radio_frame,
            text="JSON",
            variable=format_var,
            value="json"
        )
        json_radio.pack(side=tk.LEFT, padx=5)
        
        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # 取消按鈕
        cancel_button = ttk.Button(
            button_frame,
            text=self.get_text('cancel'),
            command=dialog.destroy
        )
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # 確認按鈕
        confirm_button = ttk.Button(
            button_frame,
            text=self.get_text('confirm'),
            command=lambda: self.export_tags(format_var.get(), dialog)
        )
        confirm_button.pack(side=tk.RIGHT)
        
        # 設置對話框位置
        self.center_window(dialog, 400, 150)
        dialog.focus_set()
        dialog.grab_set()

    def export_tags(self, format_type, dialog):
        """匯出標籤資料"""
        try:
            if format_type == "json":
                # 選擇保存位置
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json")],
                    title=self.get_text('save_as')
                )
                
                if file_path:
                    # 直接複製 file_tags.json 的內容
                    shutil.copy2(self.file_manager.db_file, file_path)
                    CustomMessageBox.show_info(
                        self,
                        self.get_text('export_successful'),
                        self.get_text('export_completed')
                    )
            dialog.destroy()
        except Exception as e:
            CustomMessageBox.show_error(
                self,
                self.get_text('export_failed'),
                str(e)
            )
            dialog.destroy()

    def show_import_dialog(self):
        """顯示匯入標籤對話框"""
        dialog = create_modal_dialog(self, self.get_text('import_tags'), 400, 200)
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建格式選擇框架
        format_frame = ttk.Frame(main_frame)
        format_frame.pack(fill=tk.X, pady=10)
        
        # 格式選擇標籤
        format_label = ttk.Label(format_frame, text=self.get_text('import_format'))
        format_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # 格式選擇下拉框
        format_var = tk.StringVar(value="json")
        format_combo = ttk.Combobox(
            format_frame,
            textvariable=format_var,
            values=["json"],
            state="readonly",
            width=15
        )
        format_combo.pack(side=tk.LEFT)
        
        # 創建匯入按鈕
        import_btn = ttk.Button(
            main_frame,
            text=self.get_text('import'),
            command=lambda: self.import_tags(format_var.get(), dialog)
        )
        import_btn.pack(side=tk.BOTTOM, pady=20)

    def import_tags(self, format_type, dialog):
        """匯入標籤資料"""
        try:
            # 選擇要匯入的檔案
            file_types = {
                'json': [('JSON files', '*.json')]
            }
            
            filename = filedialog.askopenfilename(
                title=self.get_text('select_import_file'),
                filetypes=file_types[format_type]
            )
            
            if not filename:
                return
            
            # 讀取文件內容
            with open(filename, 'r', encoding='utf-8') as f:
                imported_data = json.load(f)
            
            # 验证导入的数据格式
            if not self._validate_import_data(imported_data):
                CustomMessageBox.show_warning(
                    self,
                    self.get_text('warning'),
                    self.get_text('invalid_import_format')
                )
                return
            
            # 合并数据到当前数据库
            self._merge_import_data(imported_data)
            
            # 保存更新后的数据库
            self.file_manager.save_db()
            
            # 刷新界面（確保在保存之後進行）
            self.refresh_tag_list()
            self.update_file_list_by_tag_selection()  # 使用這個方法替代 filter_file_list
            
            # 显示成功消息
            CustomMessageBox.show_info(
                self,
                self.get_text('import_success_title'),
                self.get_text('import_success_message')
            )
            
        except json.JSONDecodeError:
            CustomMessageBox.show_error(
                self,
                self.get_text('import_error_title'),
                self.get_text('invalid_json_format')
            )
        except Exception as e:
            CustomMessageBox.show_error(
                self,
                self.get_text('import_error_title'),
                f"{self.get_text('import_error_message')}: {str(e)}"
            )
        finally:
            dialog.destroy()

    def _validate_import_data(self, data):
        """驗證匯入的資料格式是否正確"""
        # 檢查基本結構
        if not isinstance(data, dict):
            return False
        if 'files' not in data:
            return False
            
        # 檢查files结构
        if not isinstance(data['files'], dict):
            return False
            
        # 檢查每个文件的数据结构
        for file_key, file_info in data['files'].items():
            if not isinstance(file_info, dict):
                return False
            # 檢查必要的字段
            required_fields = ['tags', 'hash', 'paths', 'name']
            if not all(field in file_info for field in required_fields):
                return False
            # 檢查字段类型
            if not isinstance(file_info['tags'], list):
                return False
            if not isinstance(file_info['paths'], list):
                return False
            if not isinstance(file_info['name'], str):
                return False
            if not isinstance(file_info['hash'], str):
                return False
        
        return True

    def _merge_import_data(self, imported_data):
        """合并导入的数据到当前数据库"""
        try:
            logger.info(f"開始合併資料，匯入資料筆數: {len(imported_data['files'])}")
            
            # 記錄合併統計
            merged_count = 0
            added_count = 0
            
            for file_key, file_info in imported_data['files'].items():
                logger.debug(f"處理檔案: {file_key}")
                
                # 檢查檔案是否存在
                if file_key in self.file_manager.db_data:
                    logger.debug(f"找到現有檔案記錄: {file_key}")
                    existing_file = self.file_manager.db_data[file_key]
                    
                    # 合併標籤（去重）
                    existing_tags = set(existing_file['tags'])
                    new_tags = set(file_info['tags'])
                    merged_tags = list(existing_tags | new_tags)
                    
                    # 更新現有資料
                    existing_file['tags'] = merged_tags
                    
                    # 合併路徑（去重）
                    existing_paths = set(existing_file['paths'])
                    new_paths = set(file_info['paths'])
                    merged_paths = list(existing_paths | new_paths)
                    existing_file['paths'] = merged_paths
                    
                    logger.debug(f"合併後的標籤: {merged_tags}")
                    logger.debug(f"合併後的路徑: {merged_paths}")
                    merged_count += 1
                else:
                    logger.debug(f"新增檔案記錄: {file_key}")
                    # 如果檔案不存在，直接添加
                    self.file_manager.db_data[file_key] = {
                        'tags': file_info['tags'],
                        'paths': file_info['paths'],
                        'hash': file_info['hash'],
                        'name': file_info['name']
                    }
                    added_count += 1
                
                # 更新 files_tags
                current_file = self.file_manager.db_data[file_key]
                for path in current_file['paths']:
                    logger.debug(f"更新 files_tags: {path}")
                    self.file_manager.files_tags[path] = {
                        'tags': current_file['tags'],
                        'note': current_file.get('note', ''),  # 保留現有的note或使用空字串
                        'hash': current_file['hash'],
                        'name': current_file['name']
                    }
            
            logger.info(f"資料合併完成 - 合併: {merged_count} 筆, 新增: {added_count} 筆")
            
        except Exception as e:
            logger.error(f"合併資料時發生錯誤: {str(e)}")
            raise

    def export_selected_files(self):
        """导出选中的文件"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            CustomMessageBox.show_warning(self, self.get_text("warning"), self.get_text("no_files_selected"))
            return

        # 让用户选择导出目录
        export_dir = filedialog.askdirectory(title=self.get_text("select_export_directory"))
        if not export_dir:
            return

        try:
            # 获取选中的文件路径
            files_to_export = []
            for item in selected_items:
                file_path = self.file_tree.item(item, 'values')[0]
                if os.path.exists(file_path):
                    files_to_export.append(file_path)

            if not files_to_export:
                CustomMessageBox.show_warning(self, self.get_text("warning"), self.get_text("no_valid_files"))
                return

            # 创建进度对话框
            progress_dialog = create_modal_dialog(self, self.get_text("exporting_files"), 400, 150)
            
            # 创建进度条
            progress_var = tk.DoubleVar()
            progress_bar = tb.Progressbar(
                progress_dialog,
                variable=progress_var,
                maximum=len(files_to_export),
                bootstyle="success-striped"
            )
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # 创建进度标签
            progress_label = tb.Label(
                progress_dialog,
                text=self.get_text("preparing_export"),
                font=("Segoe UI", 12)
            )
            progress_label.pack(pady=5)

            # 创建带时间戳的子文件夹
            timestamp = datetime.now().strftime("%H%M%S")
            export_subfolder = os.path.join(export_dir, f"TagArtisan_{timestamp}")
            
            # 如果文件夹已存在（极少情况），添加随机数
            if os.path.exists(export_subfolder):
                random_num = random.randint(1, 999)
                export_subfolder = os.path.join(export_dir, f"TagArtisan_{timestamp}_{random_num:03d}")
            
            # 创建子文件夹
            os.makedirs(export_subfolder)

            # 添加取消标志
            progress_dialog.cancel_export = False
            
            def on_close():
                progress_dialog.cancel_export = True
                progress_dialog.destroy()
            
            progress_dialog.protocol("WM_DELETE_WINDOW", on_close)

            # 复制文件到目标目录
            for index, file_path in enumerate(files_to_export, 1):
                # 检查是否取消
                if progress_dialog.cancel_export:
                    break
                    
                file_name = os.path.basename(file_path)
                # 更新进度标签
                progress_label.config(text=self.get_text("exporting_file").format(
                    current=index,
                    total=len(files_to_export),
                    filename=file_name
                ))
                
                target_path = os.path.join(export_subfolder, file_name)
                
                # 如果目标文件已存在，添加数字后缀
                base_name, ext = os.path.splitext(file_name)
                counter = 1
                while os.path.exists(target_path):
                    new_name = f"{base_name}_{counter}{ext}"
                    target_path = os.path.join(export_subfolder, new_name)
                    counter += 1

                try:
                    shutil.copy2(file_path, target_path)
                    progress_var.set(index)
                    progress_dialog.update()
                except Exception as e:
                    logger.error(f"复制文件时发生错误: {str(e)}")
                    continue

            # 如果没有取消，显示成功消息
            if not progress_dialog.cancel_export:
                CustomMessageBox.show_info(
                    self,
                    self.get_text("success"),
                    self.get_text("export_success").format(count=index)
                )
            
            # 关闭进度对话框
            if progress_dialog.winfo_exists():
                progress_dialog.destroy()

            # 打开导出文件夹
            try:
                os.startfile(export_subfolder)
            except Exception as e:
                logger.error(f"打開導出文件夾失敗: {str(e)}")

        except Exception as e:
            logger.error(f"导出文件时发生错误: {str(e)}")
            CustomMessageBox.show_error(
                self,
                self.get_text("error"),
                self.get_text("export_error").format(error=str(e))
            )
            if 'progress_dialog' in locals() and progress_dialog.winfo_exists():
                progress_dialog.destroy()

    def on_window_minimize(self, event):
        """当主窗口被最小化时，关闭所有子窗口"""
        if event.widget == self and self.state() == 'iconic':
            # 关闭所有子窗口
            for window in self.winfo_children():
                if isinstance(window, (tk.Toplevel, tk.Tk)) and window.winfo_exists():
                    # 針對語言/快捷鍵/匯出視窗使用withdraw()而不是destroy()
                    if window in [self.language_dialog, self.hotkey_dialog]:
                        window.withdraw()
                    # 跳過tagdropwindow的處理
                    elif hasattr(self, 'tag_drop_window') and window == self.tag_drop_window:
                        pass
                    else:
                        window.destroy()
            # 清空子窗口列表
            self.child_windows.clear()

    def create_modal_dialog(self, parent, title, width, height):
        dialog = tk.Toplevel(parent)
        # ... existing code ...
        # 将对话框添加到子窗口列表中
        if hasattr(self, 'child_windows'):
            self.child_windows.append(dialog)
        return dialog

# 主程式入口
if __name__ == "__main__":
    app = None
    try:
        # 創建主應用程式但先不顯示
        app = Application(FileManager())
        app.withdraw()  # 隱藏主視窗
        
        # 顯示啟動畫面
        splash = SplashScreen(app)
        
        # 等待初始化完成
        splash.wait_for_initialization(app)
        
        # 主循環
        app.mainloop()
    except Exception as e:
        logger.critical("程序崩潰:", exc_info=True)
        if app and hasattr(app, 'crash_reporter'):
            app.crash_reporter.report_crash(e)
        else:
            logger.error("Unable to send crash report: Application instance not fully initialized")
