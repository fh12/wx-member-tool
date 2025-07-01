from PyQt5.QtWidgets import (
    QMainWindow, 
    QWidget, 
    QVBoxLayout,
    QHBoxLayout,
    QPushButton, 
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QProgressBar,
    QListWidget,
    QListWidgetItem,
    QCheckBox,
    QFileDialog,
    QDialog,
    QApplication
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
import keyboard
from src.core.wechat import WeChatController
import pandas as pd
from datetime import datetime
import os
import sys
import logging
import traceback
from PyQt5.QtGui import QIcon
import uiautomation as auto

class TaskPromptDialog(QDialog):
    """任务执行提示窗口"""
    def __init__(self, parent=None):
        super().__init__(None)  # 不设置父窗口，使其成为顶级窗口
        self.main_window = parent  # 保存主窗口引用
        self.setWindowTitle("任务执行中")
        self.setFixedSize(400, 120)
        self.setWindowFlags(
            Qt.WindowStaysOnTopHint |  # 窗口置顶
            Qt.CustomizeWindowHint |   # 自定义窗口样式
            Qt.FramelessWindowHint |   # 无边框
            Qt.Tool                    # 使其成为全局工具窗口
        )
        
        # 设置窗口样式
        self.setStyleSheet("""
            QDialog {
                background-color: #000000;
                border: 1px solid #333333;
                border-radius: 4px;
            }
            QLabel {
                color: #FFFFFF;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                background-color: transparent;
            }
            QPushButton {
                background-color: #E74C3C;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-size: 13px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
            QPushButton:pressed {
                background-color: #A93226;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)
        
        # 提示文本
        self.label = QLabel("正在执行任务中，请勿移动鼠标")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        
        # 停止按钮
        button_layout = QHBoxLayout()
        self.stop_button = QPushButton("停止")
        self.stop_button.setCursor(Qt.PointingHandCursor)
        self.stop_button.clicked.connect(self.on_stop_clicked)
        button_layout.addWidget(self.stop_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # 添加定时器以处理事件
        self.process_timer = QTimer()
        self.process_timer.timeout.connect(lambda: QApplication.processEvents())
        self.process_timer.start(100)  # 每100ms处理一次事件
        
    def on_stop_clicked(self):
        """处理停止按钮点击事件"""
        print("\n=== 用户点击了停止按钮 ===")
        # 使用保存的主窗口引用
        if self.main_window and self.main_window.worker_thread and self.main_window.worker_thread.isRunning():
            print("正在停止工作线程...")
            # 先设置 WeChatController 的停止标志
            if hasattr(self.main_window.worker_thread, 'wechat'):
                print("设置 WeChatController 停止标志...")
                self.main_window.worker_thread.wechat.is_running = False
            # 停止工作线程
            self.main_window.worker_thread.stop()
            print("工作线程已停止")
            
        print("=== 开始关闭任务窗口 ===")
        self.process_timer.stop()  # 停止事件处理定时器
        self.reject()  # 关闭对话框
        
    def closeEvent(self, event):
        """窗口关闭时的处理"""
        self.process_timer.stop()  # 确保定时器被停止
        super().closeEvent(event)

    def move_to_bottom_center(self):
        """将窗口移动到屏幕底部居中"""
        # 获取主屏幕
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            x = screen_geometry.x() + (screen_geometry.width() - self.width()) // 2
            y = screen_geometry.y() + screen_geometry.height() - self.height() - 50  # 距离底部50像素
            self.move(x, y)
        
    def showEvent(self, event):
        """窗口显示时的处理"""
        super().showEvent(event)
        self.move_to_bottom_center()
        self.raise_()  # 确保窗口在最前面
        self.activateWindow()  # 激活窗口

class WorkerThread(QThread):
    """工作线程类，用于执行耗时操作"""
    progressChanged = pyqtSignal(int)  # 进度信号
    finished = pyqtSignal(object)  # 完成信号，携带结果数据
    error = pyqtSignal(str)  # 错误信号
    
    def __init__(self, task_type, wechat, **kwargs):
        super().__init__()
        self.task_type = task_type
        self.wechat = wechat
        self.kwargs = kwargs
        self._is_running = True
        
    def stop(self):
        """停止线程"""
        self._is_running = False
        
    def run(self):
        """线程执行的主要逻辑"""
        try:
            # 在线程中初始化 UI 自动化
            import uiautomation as auto
            try:
                # 使用 UIAutomationInitializerInThread 对象初始化 UI 自动化
                self.ui_automation_initializer = auto.UIAutomationInitializerInThread()
                print("线程中 UI 自动化初始化成功")
            except Exception as e:
                print(f"线程中 UI 自动化初始化失败: {e}")
            
            if self.task_type == "scan_groups":
                self._scan_groups()
            elif self.task_type == "analyze_groups":
                self._analyze_groups()
                
            # 在线程结束时释放 UI 自动化资源
            try:
                if hasattr(self, 'ui_automation_initializer'):
                    del self.ui_automation_initializer
                    print("线程中 UI 自动化资源已释放")
            except Exception as e:
                print(f"线程中释放 UI 自动化资源时出错: {e}")
                
        except Exception as e:
            self.error.emit(str(e))
            
    def _scan_groups(self):
        """扫描群聊的具体实现"""
        try:
            # 获取群聊列表
            groups = self.wechat.get_group_list(use_cache=False)
            if not self._is_running:
                return
                
            if groups:
                self.finished.emit(groups)
            else:
                self.error.emit("未找到群聊列表，请先打开微信界面")
                
        except Exception as e:
            self.error.emit(f"扫描群聊失败：{str(e)}")
            
    def _analyze_groups(self):
        """分析群聊的具体实现"""
        try:
            selected_groups = self.kwargs.get("selected_groups", [])
            
            # 确保微信窗口已初始化
            if not self.wechat.find_wechat_window():
                self.error.emit("请确保微信界面已经打开！")
                return
                
            if not self.wechat.activate_window():
                self.error.emit("无法激活微信窗口，请确保微信界面已经打开！")
                return
                
            # 获取所有选中群的成员
            all_members = {}
            for i, group_name in enumerate(selected_groups):
                if not self._is_running:
                    return
                    
                progress = int((i + 1) * 100 / len(selected_groups))
                self.progressChanged.emit(progress)
                
                members = self.wechat.get_group_members(group_name)
                if members:
                    all_members[group_name] = members
                    
            if not self._is_running:
                return
                
            if all_members:
                self.finished.emit(all_members)
            else:
                self.error.emit("未能获取任何群成员信息")
                
        except Exception as e:
            self.error.emit(f"分析失败：{str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        logging.info("开始初始化 MainWindow...")
        try:
            super().__init__()
            logging.info("MainWindow 父类初始化完成")
            
            # 设置窗口图标
            icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
                logging.info(f"设置窗口图标: {icon_path}")
            else:
                logging.warning(f"图标文件不存在: {icon_path}")
            
            self.setWindowTitle("微信群成员分析工具")
            self.setGeometry(100, 100, 1000, 600)
            logging.info("窗口基本属性设置完成")
            
            logging.info("初始化 WeChatController...")
            self.wechat = WeChatController()
            self.groups_data = {}
            self.task_dialog = None
            self.worker_thread = None
            logging.info("基本变量初始化完成")
            
            logging.info("开始初始化UI...")
            self.init_ui()
            logging.info("UI初始化完成")
            
        except Exception as e:
            logging.error(f"MainWindow 初始化失败: {str(e)}")
            logging.error(traceback.format_exc())
            QMessageBox.critical(None, "错误", f"程序初始化失败：{str(e)}")
            sys.exit(1)

    def init_ui(self):
        """初始化UI界面"""
        try:
            logging.info("开始设置UI样式...")
            # 设置整体样式
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #F5F5F5;
                }
                QLabel {
                    font-size: 14px;
                    color: #333333;
                    padding: 5px;
                }
                QLabel#user_info {
                    font-size: 13px;
                    color: #666666;
                }
                QLabel#vip_label {
                    color: #FFD700;
                    font-weight: bold;
                }
                QLabel#expire_label {
                    color: #FF6B6B;
                }
                QPushButton {
                    background-color: #4A90E2;
                    color: white;
                    border: none;
                    padding: 8px 15px;
                    border-radius: 4px;
                    font-size: 13px;
                    min-width: 80px;
                }
                QPushButton#vip_button {
                    background-color: #FFD700;
                    color: #8B4513;
                }
                QPushButton#permanent_button {
                    background-color: #FF4500;
                }
                QPushButton:hover {
                    background-color: #357ABD;
                }
                QPushButton:pressed {
                    background-color: #2E6DA4;
                }
                QListWidget, QTableWidget {
                    background-color: white;
                    border: 1px solid #DDDDDD;
                    border-radius: 4px;
                    padding: 5px;
                }
                QListWidget::item {
                    padding: 5px;
                    border-bottom: 1px solid #EEEEEE;
                }
                QListWidget::item:selected {
                    background-color: #E3F2FD;
                    color: #333333;
                }
                QTableWidget {
                    gridline-color: #EEEEEE;
                }
                QHeaderView::section {
                    background-color: #F8F9FA;
                    padding: 8px;
                    border: none;
                    border-bottom: 2px solid #DDDDDD;
                    font-weight: bold;
                    color: #333333;
                }
                QProgressBar {
                    border: 1px solid #E9ECEF;
                    border-radius: 4px;
                    background-color: #F8F9FA;
                    text-align: center;
                    color: #666666;
                    font-size: 12px;
                    padding: 1px;
                }
                QProgressBar::chunk {
                    background-color: #4CAF50;  /* 绿色 */
                    border-radius: 3px;
                    margin-right: 10px;  /* 为文字留出空间 */
                }
                QCheckBox {
                    font-size: 13px;
                    color: #333333;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                }
                QCheckBox::indicator:unchecked {
                    border: 2px solid #DDDDDD;
                    border-radius: 3px;
                    background-color: white;
                }
                QCheckBox::indicator:checked {
                    border: 2px solid #4A90E2;
                    border-radius: 3px;
                    background-color: #4A90E2;
                }
                /* 按钮容器样式 */
                QWidget#button_container {
                    background-color: white;
                    border: 1px solid #E9ECEF;
                    border-radius: 6px;
                    margin-top: 5px;
                }
            """)

            central_widget = QWidget()
            self.setCentralWidget(central_widget)
            main_layout = QVBoxLayout(central_widget)
            main_layout.setContentsMargins(20, 20, 20, 20)
            
            # 创建水平布局用于左右面板
            content_layout = QHBoxLayout()
            content_layout.setSpacing(20)
            
            # 左侧群聊列表面板
            left_panel = QWidget()
            left_panel.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-radius: 8px;
                }
                QWidget#button_container {
                    background-color: #F8F9FA;
                    border-radius: 6px;
                    border: 1px solid #E9ECEF;
                }
            """)
            left_layout = QVBoxLayout(left_panel)
            left_layout.setContentsMargins(15, 15, 15, 15)
            left_layout.setSpacing(10)
            
            # 群聊列表标题和全选框
            list_header = QWidget()
            header_layout = QHBoxLayout(list_header)
            header_layout.setContentsMargins(0, 0, 0, 0)
            group_list_label = QLabel("群聊列表")
            self.group_count_label = QLabel("(0个群聊)")  # 添加群聊数量标签
            self.group_count_label.setStyleSheet("""
                QLabel {
                    color: #666666;
                    font-size: 12px;
                    font-weight: normal;
                    padding: 5px 0px;
                }
            """)
            self.select_all_checkbox = QCheckBox("全选")
            
            self.select_all_checkbox.stateChanged.connect(self.on_select_all_changed)
            header_layout.addWidget(group_list_label)
            header_layout.addWidget(self.group_count_label)
            header_layout.addWidget(self.select_all_checkbox)
            header_layout.addStretch()
            
            self.group_list = QListWidget()
            self.group_list.setSelectionMode(QListWidget.MultiSelection)
            
            # 从缓存加载群聊列表
            cached_groups = self.wechat.get_cached_groups()
            if cached_groups:
                self.groups_data = {group["name"]: group for group in cached_groups}
                for group in cached_groups:
                    item = QListWidgetItem(group["name"])
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    self.group_list.addItem(item)
                # 更新初始群聊数量显示
                self.group_count_label.setText(f"({len(cached_groups)}个群聊)")
            
            # 添加新的事件过滤器
            self.group_list.viewport().installEventFilter(self)
            
            # 按钮容器
            button_container = QWidget()
            button_container.setObjectName("button_container")
            button_layout = QVBoxLayout(button_container)
            button_layout.setSpacing(10)
            button_layout.setContentsMargins(10, 10, 10, 10)
            
            # 扫描和分析按钮
            scan_button = QPushButton("更新群聊列表")
            analyze_button = QPushButton("分析重复成员")
            
            # 设置按钮样式
            for button in [scan_button, analyze_button]:
                button.setStyleSheet("""
                    QPushButton {
                        background-color: #4A90E2;
                        color: white;
                        border: none;
                        padding: 8px 15px;
                        border-radius: 4px;
                        font-size: 13px;
                        min-width: 80px;
                    }
                    QPushButton:hover {
                        background-color: #357ABD;
                    }
                    QPushButton:pressed {
                        background-color: #2E6DA4;
                    }
                """)
            
            # 进度条
            self.progress_bar = QProgressBar()
            self.progress_bar.setVisible(False)
            self.progress_bar.setFixedHeight(8)  # 设置进度条高度
            self.progress_bar.setTextVisible(True)  # 显示进度文字
            self.progress_bar.setFormat("  %p%")  # 设置进度文字格式，添加空格使文字向右偏移

            button_layout.addWidget(scan_button)
            button_layout.addWidget(analyze_button)
            button_layout.addWidget(self.progress_bar)

            left_layout.addWidget(list_header)
            left_layout.addWidget(self.group_list)
            left_layout.addWidget(button_container)

            # 右侧结果显示面板
            right_panel = QWidget()
            right_panel.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-radius: 8px;
                }
            """)
            right_layout = QVBoxLayout(right_panel)
            right_layout.setContentsMargins(15, 15, 15, 15)
            right_layout.setSpacing(10)
            
            # 创建标题栏容器
            title_container = QWidget()
            title_layout = QHBoxLayout(title_container)
            title_layout.setContentsMargins(0, 0, 0, 0)
            title_layout.setSpacing(10)
            
            # 添加标题和导出按钮
            result_label = QLabel("分析结果")
            export_button = QPushButton("导出结果")
            self.export_button = export_button  # 将按钮保存为类属性
            self.export_button.setStyleSheet("""
                QPushButton {
                    background-color: #4A90E2;
                    color: white;
                    border: none;
                    padding: 8px 15px;
                    border-radius: 4px;
                    font-size: 13px;
                    min-width: 80px;
                }
                QPushButton:hover {
                    background-color: #357ABD;
                }
                QPushButton:pressed {
                    background-color: #2E6DA4;
                }
            """)
            
            title_layout.addWidget(result_label)
            title_layout.addStretch()  # 添加弹性空间
            title_layout.addWidget(self.export_button)
            
            self.result_table = QTableWidget()
            self.result_table.setColumnCount(2)
            self.result_table.setHorizontalHeaderLabels(["昵称", "所在群聊"])
            self.result_table.setColumnWidth(0, 200)
            self.result_table.setColumnWidth(1, 400)
            self.result_table.setShowGrid(True)
            self.result_table.setAlternatingRowColors(True)
            self.result_table.setStyleSheet("""
                QTableWidget {
                    alternate-background-color: #F8F9FA;
                }
            """)

            right_layout.addWidget(title_container)
            right_layout.addWidget(self.result_table)

            # 添加左右面板到主布局
            content_layout.addWidget(left_panel, 1)
            content_layout.addWidget(right_panel, 2)

            # 绑定按钮事件
            scan_button.clicked.connect(self.scan_groups)
            analyze_button.clicked.connect(self.analyze_selected_groups)
            export_button.clicked.connect(self.export_results)
            
            # 设置滚动条样式
            scroll_bar_style = """
                QScrollBar:vertical {
                    border: none;
                    background: #F0F0F0;
                    width: 16px;
                    margin: 0px 0px 0px 0px;
                }
                QScrollBar::handle:vertical {
                    background: #CCCCCC;
                    min-height: 40px;
                    border-radius: 7px;
                }
                QScrollBar::handle:vertical:hover {
                    background: #BBBBBB;
                }
                QScrollBar::add-line:vertical {
                    height: 0px;
                }
                QScrollBar::sub-line:vertical {
                    height: 0px;
                }
                QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                    background: none;
                }
            """
            self.group_list.verticalScrollBar().setStyleSheet(scroll_bar_style)
            self.result_table.verticalScrollBar().setStyleSheet(scroll_bar_style)
            self.result_table.horizontalScrollBar().setStyleSheet(scroll_bar_style.replace("vertical", "horizontal").replace("width", "height"))

            # 将内容布局添加到主布局
            main_layout.addLayout(content_layout)
            
        except Exception as e:
            logging.error(f"UI初始化失败: {str(e)}")
            logging.error(traceback.format_exc())
            raise

    def show_task_dialog(self):
        """显示任务提示窗口"""
        # 如果已存在对话框，先关闭它
        if self.task_dialog:
            self.task_dialog.close()
            self.task_dialog = None
            
        # 创建新的对话框，传入self作为父窗口
        self.task_dialog = TaskPromptDialog(self)
        self.task_dialog.finished.connect(self.on_task_dialog_closed)
        self.task_dialog.show()
        QApplication.processEvents()

    def on_task_dialog_closed(self, result):
        """处理任务对话框关闭事件"""
        if result == QDialog.Rejected:  # 用户点击了停止按钮
            if self.worker_thread and self.worker_thread.isRunning():
                self.progress_bar.setVisible(False)
                QMessageBox.information(self, "提示", "任务已终止")

    def scan_groups(self):
        """扫描所有群聊"""
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # 显示任务提示窗口
            self.show_task_dialog()
            
            # 创建并启动工作线程
            self.worker_thread = WorkerThread("scan_groups", self.wechat)
            self.worker_thread.finished.connect(self.on_scan_finished)
            self.worker_thread.error.connect(self.on_worker_error)
            self.worker_thread.start()
            
        except Exception as e:
            if self.task_dialog:
                self.task_dialog.close()
            QMessageBox.critical(self, "错误", f"扫描群聊失败：{str(e)}")
            self.progress_bar.setVisible(False)

    def on_scan_finished(self, groups):
        """处理扫描完成的结果"""
        if self.task_dialog:
            self.task_dialog.close()
            
        if groups:
            # 清空并更新群聊列表
            self.group_list.clear()
            self.groups_data = {group["name"]: group for group in groups}
            
            for display_name in self.groups_data.keys():
                item = QListWidgetItem(display_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.group_list.addItem(item)
            
            # 更新群聊数量显示
            self.group_count_label.setText(f"({len(groups)}个群聊)")
            self.progress_bar.setValue(100)
            QMessageBox.information(self, "成功", f"已获取到所有微信群聊，共 {len(groups)} 个群")
            
        self.progress_bar.setVisible(False)
        self.worker_thread = None

    def analyze_selected_groups(self):
        """分析选中的群聊"""
        selected_items = self.group_list.selectedItems()
        selected_groups = [item.text() for item in selected_items]
        
        if not selected_groups:
            QMessageBox.warning(self, "警告", "请先选择要分析的群聊！")
            return
            
        # 显示任务执行窗口
        self.task_dialog = TaskPromptDialog(self)
        self.task_dialog.show()
        
        # 创建并启动工作线程
        self.worker_thread = WorkerThread("analyze_groups", self.wechat, selected_groups=selected_groups)
        self.worker_thread.finished.connect(self.on_analyze_finished)
        self.worker_thread.error.connect(self.on_worker_error)
        self.worker_thread.start()

    def on_analyze_finished(self, all_members):
        """处理分析完成的结果"""
        if self.task_dialog:
            self.task_dialog.close()
            
        if all_members:
            # 分析重复成员
            common_members = self.analyze_common_members(all_members)
            
            # 显示结果
            self.show_analysis_results(common_members)
            
            if not common_members:
                QMessageBox.information(self, "提示", "未发现重复成员！")
            else:
                QMessageBox.information(self, "成功", f"分析完成，发现 {len(common_members)} 个重复成员！")
        
        self.progress_bar.setVisible(False)
        self.worker_thread = None

    def on_worker_error(self, error_message):
        """处理工作线程的错误"""
        if self.task_dialog:
            self.task_dialog.close()
        QMessageBox.critical(self, "错误", error_message)
        self.progress_bar.setVisible(False)
        self.worker_thread = None

    def analyze_common_members(self, groups_data):
        """分析重复成员"""
        member_groups = {}
        
        # 统计每个成员在哪些群中
        for group_name, members in groups_data.items():
            for member_name, member_info in members.items():
                if member_name not in member_groups:
                    member_groups[member_name] = {
                        "groups": set()
                    }
                member_groups[member_name]["groups"].add(group_name)
        
        # 筛选出现在多个群的成员
        return {
            name: info for name, info in member_groups.items()
            if len(info["groups"]) >= 2
        }

    def show_analysis_results(self, common_members):
        """显示分析结果"""
        # 清空表格
        self.result_table.setRowCount(0)
        
        # 设置表格列数
        self.result_table.setColumnCount(3)
        self.result_table.setHorizontalHeaderLabels(["群成员", "重复出现次数", "所在群聊"])
        
        # 添加结果到表格
        total_members = len(common_members)
        
        for row, (member, info) in enumerate(common_members.items()):
            groups = list(info["groups"])  # 转换set为list
            self.result_table.insertRow(row)
            self.result_table.setItem(row, 0, QTableWidgetItem(member))
            self.result_table.setItem(row, 1, QTableWidgetItem(str(len(groups))))
            self.result_table.setItem(row, 2, QTableWidgetItem(", ".join(groups)))
        
        # 调整列宽
        self.result_table.resizeColumnsToContents()
        
        # 设置导出按钮状态
        self.export_button.setEnabled(True)
        self.export_button.setStyleSheet("""
            QPushButton {
                background-color: #4A90E2;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 13px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #357ABD;
            }
            QPushButton:pressed {
                background-color: #2E6DA4;
            }
        """)
        self.export_button.setToolTip("导出分析结果")

    def export_results(self):
        """导出分析结果到Excel文件"""
        # 检查是否有数据可以导出
        if self.result_table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "没有可导出的数据！请先进行群成员分析。")
            return
            
        try:
            # 获取默认的导出文件名（使用当前时间）
            default_filename = f"群成员分析结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            # 打开文件保存对话框
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "导出结果",
                os.path.join(os.path.expanduser("~"), "Desktop", default_filename),
                "Excel 文件 (*.xlsx);;所有文件 (*.*)"
            )
            
            if not filename:  # 用户取消了保存
                return
                
            # 如果文件名没有.xlsx后缀，添加它
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
                
            # 准备数据
            data = []
            for row in range(self.result_table.rowCount()):
                name = self.result_table.item(row, 0).text()
                repeat_count = self.result_table.item(row, 1).text()
                groups = self.result_table.item(row, 2).text()
                data.append({
                    "昵称": name,
                    "重复出现次数": repeat_count,
                    "所在群聊": groups
                })
                
            # 创建DataFrame并导出到Excel
            df = pd.DataFrame(data)
            
            # 创建Excel写入器
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 写入数据
                df.to_excel(writer, index=False, sheet_name='群成员分析结果')
                
                # 获取工作表
                worksheet = writer.sheets['群成员分析结果']
                
                # 调整列宽
                worksheet.column_dimensions['A'].width = 30  # 昵称列
                worksheet.column_dimensions['B'].width = 15  # 重复次数列
                worksheet.column_dimensions['C'].width = 50  # 群聊列
                
            QMessageBox.information(self, "成功", f"数据已成功导出到：\n{filename}")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    def eventFilter(self, source, event):
        """事件过滤器，处理列表项的点击事件"""
        if (source is self.group_list.viewport() and
            event.type() == event.MouseButtonRelease):
            item = self.group_list.itemAt(event.pos())
            if item:
                # 切换复选框状态
                item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked)
                
                # 更新全选框状态
                all_checked = True
                for i in range(self.group_list.count()):
                    if self.group_list.item(i).checkState() != Qt.Checked:
                        all_checked = False
                        break
                self.select_all_checkbox.setChecked(all_checked)
                
                return True  # 事件已处理
        
        return super().eventFilter(source, event)  # 继续传递未处理的事件

    def on_select_all_changed(self, state):
        """处理全选复选框状态改变"""
        for i in range(self.group_list.count()):
            item = self.group_list.item(i)
            item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked) 

    def update_status_label(self):
        """更新状态标签"""
        pass

    def update_vip_status(self):
        """更新会员状态显示"""
        pass

    def show_register_dialog(self, is_permanent=False):
        """显示注册对话框"""
        pass

    def show_register_dialog(self, is_permanent=False):
        """显示注册对话框"""
        pass 