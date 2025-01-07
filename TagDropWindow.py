import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
from tkinterdnd2 import DND_FILES
import os
import ctypes
from languages import LANGUAGES

# 定義 Windows API 結構
class RECT(ctypes.Structure):
    _fields_ = [
        ('left', ctypes.c_long),
        ('top', ctypes.c_long),
        ('right', ctypes.c_long),
        ('bottom', ctypes.c_long)
    ]

class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.c_ulong),
        ('rcMonitor', RECT),
        ('rcWork', RECT),
        ('dwFlags', ctypes.c_ulong)
    ]

class CustomMessageBox(tk.Toplevel):
    def __init__(self, parent, title, message, style="info", width=400, height=200):
        super().__init__(parent)
        self.result = None

        self.title(title)
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # 信息標籤
        label = tb.Label(self, text=message, wraplength=width-20, bootstyle=style)
        label.pack(pady=20, padx=10, expand=True)

        # 按鈕框架
        button_frame = tb.Frame(self)
        button_frame.pack(pady=10)

        # 按鈕根據樣式不同而不同
        if style in ["info", "warning", "error"]:
            btn_text = "確定"
            btn_style = {
                "info": "success",
                "warning": "warning",
                "error": "danger"
            }.get(style, "info")
            btn = tb.Button(button_frame, text=btn_text, command=self.ok, bootstyle=btn_style)
            btn.pack()
        elif style == "question":
            btn_yes = tb.Button(button_frame, text="是", command=self.yes, bootstyle="success")
            btn_yes.pack(side=tk.LEFT, padx=10)
            btn_no = tb.Button(button_frame, text="否", command=self.no, bootstyle="danger")
            btn_no.pack(side=tk.RIGHT, padx=10)

        self.protocol("WM_DELETE_WINDOW", self.no if style == "question" else self.ok)

        # 將窗口居中
        self.center_window(width, height)

    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def ok(self):
        self.result = True
        self.destroy()

    def yes(self):
        self.result = True
        self.destroy()

    def no(self):
        self.result = False
        self.destroy()

    @staticmethod
    def show_info(parent, title, message, width=400, height=200):
        dialog = CustomMessageBox(parent, title, message, style="info", width=width, height=height)
        dialog.wait_window()
        return dialog.result

    @staticmethod
    def show_warning(parent, title, message, width=400, height=200):
        dialog = CustomMessageBox(parent, title, message, style="warning", width=width, height=height)
        dialog.wait_window()
        return dialog.result

    @staticmethod
    def show_error(parent, title, message, width=400, height=200):
        dialog = CustomMessageBox(parent, title, message, style="error", width=width, height=height)
        dialog.wait_window()
        return dialog.result

    @staticmethod
    def show_question(parent, title, message, width=400, height=200):
        dialog = CustomMessageBox(parent, title, message, style="question", width=width, height=height)
        dialog.wait_window()
        return dialog.result

class TagDropWindow(tk.Toplevel):
    def __init__(self, parent, tag_list, on_drop_callback):
        super().__init__(parent)
        
        # 保存回調函數和父視窗
        self.parent = parent
        self.on_drop_callback = on_drop_callback
        self.current_language = parent.current_language
        self.file_manager = parent.file_manager  # 添加對 file_manager 的引用
        
        # 初始化快捷鍵設置
        self.hotkey_modifier = "Ctrl"  # 預設修飾鍵
        self.hotkey_key = "1"  # 預設按鍵
        self.hotkey_id = 1  # 快捷鍵ID
        
        # 設置視窗屬性
        self.overrideredirect(True)  # 移除標題欄
        self.attributes('-alpha', 0.9)  # 設置透明度
        self.attributes('-topmost', True)  # 保持在最上層
        
        # 獲取螢幕資訊
        self.update_screen_info()
        
        # 設置視窗大小和位置
        width = 300
        # 使用螢幕高度減去任務欄高度
        height = self.screen_bottom - self.screen_top
        self.geometry(f"{width}x{height}+{self.screen_right-width}+{self.screen_top}")
        
        # 創建主框架
        self.main_frame = tb.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建標籤列表
        self.create_tag_list(tag_list)
        
        # 初始隱藏視窗
        self.withdraw()
        
        # 設置拖放目標
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.dnd_bind('<<DropEnter>>', self.on_drop_enter)
        self.dnd_bind('<<DropLeave>>', self.on_drop_leave)
        
        # 註冊全域快捷鍵
        self.register_hotkey()
        
        # 設置定時器檢查快捷鍵狀態
        self.check_hotkey_state()

    def get_text(self, key):
        """獲取當前語言的文字"""
        return LANGUAGES[self.current_language].get(key, LANGUAGES['en_US'][key])

    def register_hotkey(self):
        """註冊全域快捷鍵"""
        if os.name == 'nt':
            # 先取消註冊現有的快捷鍵
            self.unregister_hotkey()
            
            # 定義快捷鍵 ID
            self.hotkey_id = 1
            
            # 獲取修飾鍵的值
            modifier_values = {
                "Ctrl": 0x0002,
                "Alt": 0x0001,
                "Shift": 0x0004
            }
            modifier = modifier_values.get(self.hotkey_modifier, 0x0002)
            
            # 獲取按鍵的值
            if self.hotkey_key.isdigit():
                # 如果是數字鍵 (0-9)
                key = ord(self.hotkey_key)
            else:
                # 如果是字母鍵 (A-Z)
                key = ord(self.hotkey_key)
            
            # 註冊新的快捷鍵
            if not ctypes.windll.user32.RegisterHotKey(None, self.hotkey_id, modifier, key):
                print(f"無法註冊全域快捷鍵 {self.hotkey_modifier}+{self.hotkey_key}")

    def update_hotkey(self, modifier, key):
        """更新快捷鍵設置"""
        self.hotkey_modifier = modifier
        self.hotkey_key = key
        self.register_hotkey()

    def unregister_hotkey(self):
        """取消註冊全域快捷鍵"""
        if os.name == 'nt':
            ctypes.windll.user32.UnregisterHotKey(None, self.hotkey_id)
            
    def check_hotkey_state(self):
        """檢查快捷鍵狀態"""
        if os.name == 'nt':
            try:
                # 獲取修飾鍵的虛擬鍵碼
                modifier_vk = {
                    "Ctrl": 0x11,  # VK_CONTROL
                    "Alt": 0x12,   # VK_MENU
                    "Shift": 0x10  # VK_SHIFT
                }.get(self.hotkey_modifier)
                
                # 獲取按鍵的虛擬鍵碼
                if self.hotkey_key.isdigit():
                    key_vk = ord(self.hotkey_key)
                else:
                    key_vk = ord(self.hotkey_key)
                
                # 檢查修飾鍵狀態
                modifier_state = ctypes.windll.user32.GetAsyncKeyState(modifier_vk)
                # 檢查按鍵狀態
                key_state = ctypes.windll.user32.GetAsyncKeyState(key_vk)
                
                # 如果兩個鍵都被按下
                if (modifier_state & 0x8000) and (key_state & 0x8000):
                    if not self.winfo_viewable():
                        self.show_window()
                else:
                    # 檢查是否有子視窗（訊息框）
                    has_message_box = any(isinstance(child, CustomMessageBox) for child in self.winfo_children())
                    
                    # 只有在沒有訊息框顯示時才隱藏視窗
                    if self.winfo_viewable() and not has_message_box:
                        self.withdraw()
                        
            except Exception as e:
                print(f"檢查快捷鍵狀態時出錯: {e}")
                
        # 每 50ms 檢查一次
        self.after(50, self.check_hotkey_state)
        
    def show_window(self):
        """顯示視窗"""
        # 更新螢幕資訊
        self.update_screen_info()
        # 更新視窗高度為螢幕高度
        height = self.screen_bottom - self.screen_top
        # 更新視窗位置到當前螢幕的右上角，並設置新的高度
        self.geometry(f"{self.winfo_width()}x{height}+{self.screen_right-self.winfo_width()}+{self.screen_top}")
        # 顯示視窗
        self.deiconify()
        
    def on_drop_enter(self, event):
        """當拖曳進入項目時的處理"""
        # 設置高亮樣式
        self.tag_tree.tag_configure('highlight', background='#E1F5FE')

    def on_drop_leave(self, event):
        """當拖曳離開項目時的處理"""
        # 清除所有項目的高亮效果
        for item in self.tag_tree.get_children():
            current_tags = list(self.tag_tree.item(item)['tags'] or ())
            if 'highlight' in current_tags:
                current_tags.remove('highlight')
            self.tag_tree.item(item, tags=current_tags)

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
        y = event.widget.winfo_pointery() - event.widget.winfo_rooty()
        target_item = self.tag_tree.identify_row(y)
        if not target_item:
            CustomMessageBox.show_warning(self, self.get_text("warning"), self.get_text("drop_on_tag"))
            return
            
        target_tag = self.tag_tree.item(target_item, 'text')
        
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
            # 呼叫回調函數處理拖放
            if self.on_drop_callback:
                self.on_drop_callback(files, target_tag)
                
                # 更新標籤列表（確保與主視窗同步）
                self.update_tags(self.parent.file_manager.get_all_used_tags())
                # 更新主視窗的標籤列表
                self.parent.refresh_tag_list()

    def destroy(self):
        """銷毀視窗時取消註冊快捷鍵"""
        self.unregister_hotkey()
        super().destroy()
        
    def update_screen_info(self):
        """更新滑鼠所在螢幕的資訊"""
        if os.name == 'nt':  # Windows
            try:
                # 獲取滑鼠位置
                mouse_x = self.winfo_pointerx()
                mouse_y = self.winfo_pointery()
                
                # 獲取所有螢幕的資訊
                monitors = []
                def callback(monitor, data):
                    monitors.append(monitor)
                    return True
                
                callback_type = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
                callback_func = callback_type(callback)
                ctypes.windll.user32.EnumDisplayMonitors(None, None, callback_func, 0)
                
                # 找到滑鼠所在的螢幕
                for monitor in monitors:
                    info = MONITORINFO()
                    info.cbSize = ctypes.sizeof(MONITORINFO)
                    if ctypes.windll.user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
                        # 檢查滑鼠是否在這個螢幕內
                        if (info.rcMonitor.left <= mouse_x <= info.rcMonitor.right and
                            info.rcMonitor.top <= mouse_y <= info.rcMonitor.bottom):
                            self.screen_left = info.rcMonitor.left
                            self.screen_top = info.rcMonitor.top
                            self.screen_right = info.rcMonitor.right
                            self.screen_bottom = info.rcMonitor.bottom
                            return
                
                # 如果找不到，使用主螢幕
                self.screen_left = 0
                self.screen_top = 0
                self.screen_right = self.winfo_screenwidth()
                self.screen_bottom = self.winfo_screenheight()
            except Exception:
                # 如果出現任何錯誤，使用預設值
                self.screen_left = 0
                self.screen_top = 0
                self.screen_right = self.winfo_screenwidth()
                self.screen_bottom = self.winfo_screenheight()
        else:  # 其他作業系統
            self.screen_left = 0
            self.screen_top = 0
            self.screen_right = self.winfo_screenwidth()
            self.screen_bottom = self.winfo_screenheight()
        
    def create_tag_list(self, tag_list):
        """創建標籤列表"""
        # 創建框架包含列表和滾動條
        list_frame = tb.Frame(self.main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加滾動條
        scrollbar = tb.Scrollbar(list_frame, bootstyle="round")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 創建列表
        self.tag_tree = ttk.Treeview(
            list_frame,
            show='tree',
            selectmode='browse',
            yscrollcommand=scrollbar.set
        )
        self.tag_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置滾動條
        scrollbar.config(command=self.tag_tree.yview)
        
        # 將標籤分為兩組：有顏色的和沒有顏色的
        colored_tags = []
        normal_tags = []
        for tag in tag_list:
            if self.file_manager.get_tag_color(tag) != self.file_manager.default_color:
                colored_tags.append(tag)
            else:
                normal_tags.append(tag)
        
        # 對有顏色的標籤按修改時間排序（最新的在前）
        colored_tags.sort(key=lambda tag: self.file_manager.get_tag_color_timestamp(tag), reverse=True)
        
        # 先插入有顏色的標籤，再插入普通標籤
        for tag in colored_tags + normal_tags:
            color = self.file_manager.get_tag_color(tag)
            item = self.tag_tree.insert('', tk.END, text=tag)
            # 為每個標籤創建唯一的標籤樣式
            tag_style = f'tag_{hash(tag)}'
            # 只有當顏色不是預設顏色時才設置標籤樣式
            if color != self.file_manager.default_color:
                self.tag_tree.tag_configure(tag_style, foreground=color)
                self.tag_tree.item(item, tags=(tag_style,))
            
        # 設置拖放目標
        self.tag_tree.drop_target_register(DND_FILES)
        self.tag_tree.dnd_bind('<<Drop>>', self.on_drop)
        self.tag_tree.dnd_bind('<<DropEnter>>', self.on_drop_enter)
        self.tag_tree.dnd_bind('<<DropLeave>>', self.on_drop_leave)
        self.tag_tree.dnd_bind('<<DropPosition>>', self.on_drop_motion)

    def on_drop_motion(self, event):
        """當拖曳移動時的處理"""
        # 清除所有項目的高亮效果
        for item in self.tag_tree.get_children():
            self.tag_tree.item(item, tags=())
            
        # 獲取當前滑鼠位置下的項目
        y = event.widget.winfo_pointery() - event.widget.winfo_rooty()
        target_item = self.tag_tree.identify_row(y)
        
        if target_item:
            # 獲取該項目原有的標籤樣式（如果有的話）
            current_tags = list(self.tag_tree.item(target_item)['tags'] or ())
            if 'highlight' not in current_tags:
                current_tags.append('highlight')
            
            # 設置高亮效果
            self.tag_tree.tag_configure('highlight', background='#E1F5FE')
            self.tag_tree.item(target_item, tags=current_tags)

    def update_tags(self, tag_list):
        """更新標籤列表"""
        # 清空現有項目
        self.tag_tree.delete(*self.tag_tree.get_children())
        
        # 將標籤分為兩組：有顏色的和沒有顏色的
        colored_tags = []
        normal_tags = []
        for tag in tag_list:
            if self.file_manager.get_tag_color(tag) != self.file_manager.default_color:
                colored_tags.append(tag)
            else:
                normal_tags.append(tag)
        
        # 對有顏色的標籤按修改時間排序（最新的在前）
        colored_tags.sort(key=lambda tag: self.file_manager.get_tag_color_timestamp(tag), reverse=True)
        
        # 先插入有顏色的標籤，再插入普通標籤
        for tag in colored_tags + normal_tags:
            color = self.file_manager.get_tag_color(tag)
            item = self.tag_tree.insert('', tk.END, text=tag)
            # 為每個標籤創建唯一的標籤樣式
            tag_style = f'tag_{hash(tag)}'
            # 只有當顏色不是預設顏色時才設置標籤樣式
            if color != self.file_manager.default_color:
                self.tag_tree.tag_configure(tag_style, foreground=color)
                self.tag_tree.item(item, tags=(tag_style,)) 