import sys
import ctypes
from ctypes import wintypes
import win32gui
import win32con
import win32api
import keyboard
import pyperclip
import webbrowser
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QAction, 
                            QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton,
                            QGroupBox, QHBoxLayout, QSpinBox, QMessageBox, QWidget)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QIcon, QPixmap
from comtypes import CLSCTX_ALL
from PyQt5.QtSvg import QSvgRenderer 
from PyQt5.QtGui import QPainter, QColor
import win32process
import ctypes
import winreg
import os
import time

try:
    # 尝试直接导入UIAutomationClient来生成必要的类型库
    import comtypes.client as cc
    cc.GetModule('UIAutomationCore.dll')  # 这会生成comtypes.gen中的必要模块
    
    # 尝试直接导入生成的模块
    from comtypes.gen import UIAutomationClient
    print("成功导入UIAutomationClient")
    
except Exception as e:
    print(f"导入失败: {e}")
    # 继续使用你的备用方法

# 尝试导入pycaw，如果失败则使用ctypes作为备用方案
try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    PYWIN_PRESENT = True
except ImportError:
    PYWIN_PRESENT = False


class VolumeControl:
    """音量控制类"""
    def __init__(self):
        if PYWIN_PRESENT:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = interface.QueryInterface(IAudioEndpointVolume)
        else:
            self.volume = None
    
    def get_volume(self):
        """获取当前音量"""
        if self.volume:
            return self.volume.GetMasterVolumeLevelScalar()
        return 0.5  # 默认值
    
    def set_volume(self, level):
        """设置音量"""
        if self.volume:
            self.volume.SetMasterVolumeLevelScalar(max(0.0, min(1.0, level)), None)
        else:
            # 使用ctypes作为备用方案
            if 0 <= level <= 1:
                value = int(level * 65535)
                win32api.SendMessage(
                    win32con.HWND_BROADCAST, win32con.WM_APPCOMMAND, 0x30292,
                    win32api.MAKELONG(0, value))
    
    def increase_volume(self, step=0.05):
        """增加音量"""
        current = self.get_volume()
        self.set_volume(min(1.0, current + step))
    
    def decrease_volume(self, step=0.05):
        """降低音量"""
        current = self.get_volume()
        self.set_volume(max(0.0, current - step))


class WindowManager:
    """窗口管理类"""
    @staticmethod
    def get_foreground_window():
        """获取当前前景窗口"""
        return win32gui.GetForegroundWindow()
    
    @staticmethod
    def set_window_topmost(hwnd, topmost=True):
        """设置窗口置顶"""
        win32gui.SetWindowPos(hwnd, 
                             win32con.HWND_TOPMOST if topmost else win32con.HWND_NOTOPMOST,
                             0, 0, 0, 0,
                             win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    
    @staticmethod
    def set_window_transparency(hwnd, alpha=128):
        """设置窗口透明度"""
        # 设置窗口为分层窗口
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                              win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
        # 设置透明度
        win32gui.SetLayeredWindowAttributes(hwnd, 0, alpha, win32con.LWA_ALPHA)


class WebSearch:
    """网络搜索类"""
    def __init__(self, search_engine="https://www.baidu.com/s?ie=UTF-8&wd={}"):
        self.search_engine = search_engine
    
    def set_search_engine(self, engine_url):
        """设置搜索引擎"""
        self.search_engine = engine_url
    
    def search_text(self, text):
        """搜索文本"""
        if text.strip():
            search_url = self.search_engine.format(text)
            webbrowser.open(search_url)


class SettingsDialog(QDialog):
    """设置对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setModal(True)
        self.resize(400, 300)
        
        self.settings = QSettings("MyCompany", "WindowsUtilityTool")
    
        self.init_ui()
        self.load_settings()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 快捷键设置组
        shortcut_group = QGroupBox("快捷键设置")
        shortcut_layout = QVBoxLayout()
        
        # 窗口置顶快捷键
        topmost_layout = QHBoxLayout()
        topmost_layout.addWidget(QLabel("窗口置顶:"))
        self.topmost_edit = QLineEdit()
        topmost_layout.addWidget(self.topmost_edit)
        shortcut_layout.addLayout(topmost_layout)
        
        # 窗口透明快捷键
        transparency_layout = QHBoxLayout()
        transparency_layout.addWidget(QLabel("窗口透明:"))
        self.transparency_edit = QLineEdit()
        transparency_layout.addWidget(self.transparency_edit)
        shortcut_layout.addLayout(transparency_layout)
        
        # 音量增加快捷键
        vol_up_layout = QHBoxLayout()
        vol_up_layout.addWidget(QLabel("音量增加:"))
        self.vol_up_edit = QLineEdit()
        vol_up_layout.addWidget(self.vol_up_edit)
        shortcut_layout.addLayout(vol_up_layout)
        
        # 音量减少快捷键
        vol_down_layout = QHBoxLayout()
        vol_down_layout.addWidget(QLabel("音量减少:"))
        self.vol_down_edit = QLineEdit()
        vol_down_layout.addWidget(self.vol_down_edit)
        shortcut_layout.addLayout(vol_down_layout)
        
        # 划词搜索快捷键
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("划词搜索:"))
        self.search_edit = QLineEdit()
        search_layout.addWidget(self.search_edit)
        shortcut_layout.addLayout(search_layout)
        
        shortcut_group.setLayout(shortcut_layout)
        layout.addWidget(shortcut_group)
        
        # 搜索引擎设置组
        engine_group = QGroupBox("搜索引擎设置")
        engine_layout = QHBoxLayout()
        engine_layout.addWidget(QLabel("搜索引擎URL:"))
        self.engine_edit = QLineEdit()
        engine_layout.addWidget(self.engine_edit)
        engine_group.setLayout(engine_layout)
        layout.addWidget(engine_group)
        
        # 透明度设置组
        alpha_group = QGroupBox("透明度设置")
        alpha_layout = QHBoxLayout()
        alpha_layout.addWidget(QLabel("透明度(0-255):"))
        self.alpha_spin = QSpinBox()
        self.alpha_spin.setRange(0, 255)
        self.alpha_spin.setValue(128)
        alpha_layout.addWidget(self.alpha_spin)
        alpha_group.setLayout(alpha_layout)
        layout.addWidget(alpha_group)
        
        # 按钮
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("保存")
        self.cancel_button = QPushButton("取消")
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)
        filename = self.resource_path("./ico.ico")
        self.setWindowIcon(QIcon(filename)) 
        self.setLayout(layout)
        
        # 连接信号
        self.save_button.clicked.connect(self.save_settings)
        self.cancel_button.clicked.connect(self.reject)
    
    def resource_path(self, relative_path):
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = "./"
        ret_path = os.path.join(base_path, relative_path)
        return ret_path


    def load_settings(self):
        """加载设置"""
        self.topmost_edit.setText(self.settings.value("shortcuts/topmost", "ctrl+alt+t"))
        self.transparency_edit.setText(self.settings.value("shortcuts/transparency", "ctrl+alt+p"))
        self.vol_up_edit.setText(self.settings.value("shortcuts/vol_up", "ctrl+alt+up"))
        self.vol_down_edit.setText(self.settings.value("shortcuts/vol_down", "ctrl+alt+down"))
        self.search_edit.setText(self.settings.value("shortcuts/search", "ctrl+alt+s"))
        self.engine_edit.setText(self.settings.value("search/engine", "https://www.baidu.com/s?ie=UTF-8&wd={}"))
        self.alpha_spin.setValue(int(self.settings.value("window/alpha", 128)))
    
    def save_settings(self):
        """保存设置"""
        self.settings.setValue("shortcuts/topmost", self.topmost_edit.text())
        self.settings.setValue("shortcuts/transparency", self.transparency_edit.text())
        self.settings.setValue("shortcuts/vol_up", self.vol_up_edit.text())
        self.settings.setValue("shortcuts/vol_down", self.vol_down_edit.text())
        self.settings.setValue("shortcuts/search", self.search_edit.text())
        self.settings.setValue("search/engine", self.engine_edit.text())
        self.settings.setValue("window/alpha", self.alpha_spin.value())
        
        self.accept()
        
    def closeEvent(self, event):
        """重写关闭事件，确保不会终止整个应用程序"""
        event.accept()  # 接受关闭事件，只关闭对话框


class AboutDialog(QDialog):
    """关于对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.setModal(True)
        self.resize(300, 200)
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        filename = self.resource_path("./ico.ico")
        self.setWindowIcon(QIcon(filename))
        # 应用程序信息
        info_label = QLabel("Windows实用工具集\n\n v1.0\n\n"
                           "功能:\n"
                           "1. 窗口置顶\n  （选中窗口后按下快捷键）\n"
                           "2. 窗口半透明\n  （选中窗口后按下快捷键,如果未在退出前取消半透明，关闭半透明的窗口重开即可恢复正常）\n"
                           "3. 音量控制\n"
                           "4. 划词搜索\n  （支持不完善，只支持少量应用）\n\n"
                           "作者: Ray"
                           )
        info_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(info_label)
        
        # 确定按钮
        ok_button = QPushButton("确定")
        ok_button.clicked.connect(self.accept)
        layout.addWidget(ok_button)
        
        self.setLayout(layout)

    def resource_path(self, relative_path):
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = "./"
        ret_path = os.path.join(base_path, relative_path)
        return ret_path

    def closeEvent(self, event):
        """重写关闭事件，确保不会终止整个应用程序"""
        event.accept()  # 接受关闭事件，只关闭对话框


class WindowsUtilityTool:
    """主应用程序类"""
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.settings = QSettings("MyCompany", "WindowsUtilityTool")
        
        # 初始化组件
        self.volume_control = VolumeControl()
        self.window_manager = WindowManager()
        self.web_search = WebSearch()
        
        # 初始化系统托盘
        self.tray_icon = QSystemTrayIcon()
        self.init_tray_icon()
        
        # 注册快捷键
        self.register_hotkeys()
        
        # 存储当前置顶和透明窗口
        self.topmost_windows = set()
        self.transparent_windows = set()
    
    def init_tray_icon(self):
        """初始化系统托盘图标"""
        # 尝试从文件加载图标
        try:
            icon = QIcon("./ico.ico")
            # 检查图标是否有效
            if icon.isNull():
                raise Exception("SVG图标加载失败")
        except Exception as e:
            print(f"无法加载SVG图标文件: {e}")
            # 如果加载失败，创建一个简单的深灰色图标作为备用
            pixmap = QPixmap(16, 16)
            pixmap.fill(QColor(80, 80, 80))  # 深灰色
            icon = QIcon(pixmap)

        self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("Windows实用工具集")
        
        # 创建托盘菜单
        tray_menu = QMenu()
        
        settings_action = QAction("设置", tray_menu)
        settings_action.triggered.connect(self.show_settings)
        tray_menu.addAction(settings_action)
        
        about_action = QAction("关于", tray_menu)
        about_action.triggered.connect(self.show_about)
        tray_menu.addAction(about_action)
        
        exit_action = QAction("退出", tray_menu)
        exit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
    
    def register_hotkeys(self):
        """注册全局快捷键"""
        # 获取设置中的快捷键
        topmost_shortcut = self.settings.value("shortcuts/topmost", "ctrl+alt+t")
        transparency_shortcut = self.settings.value("shortcuts/transparency", "ctrl+alt+p")
        vol_up_shortcut = self.settings.value("shortcuts/vol_up", "ctrl+alt+up")
        vol_down_shortcut = self.settings.value("shortcuts/vol_down", "ctrl+alt+down")
        search_shortcut = self.settings.value("shortcuts/search", "ctrl+alt+s")
        
        # 注册快捷键
        keyboard.add_hotkey(topmost_shortcut, self.toggle_window_topmost)
        keyboard.add_hotkey(transparency_shortcut, self.toggle_window_transparency)
        keyboard.add_hotkey(vol_up_shortcut, self.increase_volume)
        keyboard.add_hotkey(vol_down_shortcut, self.decrease_volume)
        keyboard.add_hotkey(search_shortcut, self.search_selected_text)
    
    def toggle_window_topmost(self):
        """切换窗口置顶状态"""
        hwnd = self.window_manager.get_foreground_window()
        if hwnd in self.topmost_windows:
            self.window_manager.set_window_topmost(hwnd, False)
            self.topmost_windows.remove(hwnd)
            self.tray_icon.showMessage("窗口置顶", "已取消窗口置顶", QSystemTrayIcon.Information, 1000)
        else:
            self.window_manager.set_window_topmost(hwnd, True)
            self.topmost_windows.add(hwnd)
            self.tray_icon.showMessage("窗口置顶", "已设置窗口置顶", QSystemTrayIcon.Information, 1000)
    
    def toggle_window_transparency(self):
        """切换窗口透明状态"""
        hwnd = self.window_manager.get_foreground_window()
        alpha = int(self.settings.value("window/alpha", 128))
        
        if hwnd in self.transparent_windows:
            self.window_manager.set_window_transparency(hwnd, 255)  # 完全不透明
            self.transparent_windows.remove(hwnd)
            self.tray_icon.showMessage("窗口透明", "已取消窗口透明", QSystemTrayIcon.Information, 1000)
        else:
            self.window_manager.set_window_transparency(hwnd, alpha)
            self.transparent_windows.add(hwnd)
            self.tray_icon.showMessage("窗口透明", "已设置窗口透明", QSystemTrayIcon.Information, 1000)
    
    def increase_volume(self):
        """增加音量"""
        self.volume_control.increase_volume()
        volume = int(self.volume_control.get_volume() * 100)
        self.tray_icon.showMessage("音量控制", f"音量: {volume}%", QSystemTrayIcon.Information, 1)
    
    def decrease_volume(self):
        """降低音量"""
        self.volume_control.decrease_volume()
        volume = int(self.volume_control.get_volume() * 100)
        self.tray_icon.showMessage("音量控制", f"音量: {volume}%", QSystemTrayIcon.Information, 1)
    
    def search_selected_text(self):
        """搜索选中的文本"""
        # 使用 UI Automation API 获取选中的文本
        selected_text = self.get_selected_text_with_api()
        
        # 使用选中的文本进行搜索
        if selected_text.strip():
            search_engine = self.settings.value("search/engine", "https://www.baidu.com/s?ie=UTF-8&wd={}")
            self.web_search.set_search_engine(search_engine)
            self.web_search.search_text(selected_text)
            self.tray_icon.showMessage("划词搜索", f"正在搜索: {selected_text}", QSystemTrayIcon.Information, 1000)
        else:
            self.tray_icon.showMessage("划词搜索", "未选中任何文本", QSystemTrayIcon.Warning, 1000)

    def get_selected_text_with_uia(self):
        """使用 UI Automation API 获取选中的文本"""
        try:
            import comtypes.client as cc
            from comtypes.gen import UIAutomationClient as UIA
            
            # 初始化 UI Automation
            cc.GetModule([UIA.__file__])
            uia = cc.CreateObject(UIA.CUIAutomation)
            
            # 获取前景窗口
            foreground_window = win32gui.GetForegroundWindow()
            
            # 将窗口句柄转换为 UI Automation 元素
            element = uia.ElementFromHandle(foreground_window)
            
            # 获取文本模式
            text_pattern = element.GetCurrentPattern(UIA.UIA_TextPatternId)
            
            if text_pattern:
                # 获取文本选择
                text_selection = text_pattern.GetSelection()
                
                if text_selection and text_selection.Length > 0:
                    # 获取第一个文本范围（通常只有一个选择）
                    text_range = text_selection.GetElement(0)
                    
                    # 获取选中的文本
                    selected_text = text_range.GetText(-1)  # -1 表示获取所有文本
                    return selected_text.strip()
            
            # 如果 UI Automation 方法失败，尝试其他方法
            return self.get_selected_text_with_clipboard()
        
        except Exception as e:
            print(f"UI Automation 方法失败: {e}")
            # 如果 UI Automation 方法失败，回退到剪贴板方法
            return self.get_selected_text_with_clipboard()

    def get_selected_text_with_api(self):
        """使用 Windows API 获取选中的文本"""
        try:
            # 获取前景窗口
            hwnd = win32gui.GetForegroundWindow()
            
            # 获取窗口线程ID和进程ID
            thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
            
            # 附加到线程
            win32process.AttachThreadInput(win32api.GetCurrentThreadId(), thread_id, True)
            
            # 获取焦点控件
            focused_hwnd = win32gui.GetFocus()
            
            # 分离线程
            win32process.AttachThreadInput(win32api.GetCurrentThreadId(), thread_id, False)
            
            # 获取文本长度
            text_length = win32gui.SendMessage(focused_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            
            # 如果没有文本，返回空字符串
            if text_length == 0:
                return ""
            
            # 获取选中的文本范围
            start_selection = win32gui.SendMessage(focused_hwnd, win32con.EM_GETSEL, 0, 0)
            start = win32api.LOWORD(start_selection)
            end = win32api.HIWORD(start_selection)
            
            # 如果没有选中文本，返回空字符串
            if start == end:
                return ""
            
            # 创建缓冲区并获取选中的文本
            buffer = ctypes.create_unicode_buffer(text_length + 1)
            win32gui.SendMessage(focused_hwnd, win32con.WM_GETTEXT, text_length + 1, buffer)
            
            # 返回选中的文本部分
            return buffer.value[start:end]
        
        except Exception as e:
            print(f"获取选中文本失败: {e}")
            # 如果API方法失败，回退到剪贴板方法
            return self.get_selected_text_with_clipboard()

    def get_selected_text_with_clipboard(self):
        """使用剪贴板方法获取选中的文本（备用方法）"""
        try:
            # 保存当前剪贴板内容
            original_clipboard = pyperclip.paste()
            
            # 清空剪贴板
            pyperclip.copy('')
            
            # 模拟Ctrl+C复制选中的文本
            keyboard.send('ctrl+c')
            
            # 等待剪贴板内容变化（最多等待1秒）
            import time
            start_time = time.time()
            selected_text = ''
            
            while time.time() - start_time < 1.0:  # 最多等待1秒
                current_clipboard = pyperclip.paste()
                if current_clipboard and current_clipboard != '':
                    selected_text = current_clipboard
                    break
                time.sleep(0.05)  # 每50毫秒检查一次
            
            # 恢复剪贴板内容
            pyperclip.copy(original_clipboard)
            
            return selected_text
        
        except Exception as e:
            print(f"剪贴板方法也失败: {e}")
            return ""
        
    def show_settings(self):
        """显示设置对话框"""
        dialog = SettingsDialog()
        if dialog.exec_() == QDialog.Accepted:
            # 重新注册快捷键
            try:
                keyboard.unhook_all()
                self.register_hotkeys()
            except Exception as e:
                print(f"重新注册快捷键时出错: {e}")
    
    def show_about(self):
        """显示关于对话框"""
        dialog = AboutDialog()
        dialog.exec_()
    
    def quit_app(self):
        """退出应用程序"""
        # 恢复所有窗口的置顶和透明状态
        for hwnd in self.topmost_windows:
            try:
                self.window_manager.set_window_topmost(hwnd, False)
            except:
                pass
        
        for hwnd in self.transparent_windows:
            try:
                self.window_manager.set_window_transparency(hwnd, 255)
            except:
                pass
        
        # 退出应用程序
        self.app.quit()
    
    def run(self):
        """运行应用程序"""
        sys.exit(self.app.exec_())


if __name__ == "__main__":
    # 确保只有一个实例运行
    try:
        time.sleep(1)
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "WindowsUtilityTool")
        if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            # 先创建 QApplication 实例才能显示消息框
            app = QApplication(sys.argv)
            tray_icon = QSystemTrayIcon(QIcon("./ico.ico"))
            tray_icon.show()
            time.sleep(1)
            tray_icon.showMessage("Windows实用工具集", "程序已运行。", QSystemTrayIcon.Information, 1000)
            sys.exit(1)
    except:
        pass
    
    # 创建 QApplication 实例
    app = QApplication(sys.argv)
    
    # 然后创建主窗口
    utility_tool = WindowsUtilityTool()
    
    # 运行应用程序
    sys.exit(app.exec_())