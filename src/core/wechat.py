import win32gui
import win32con
import win32api
import pyautogui
import pyperclip
import time
import json
import os
from datetime import datetime
from typing import Dict, List, Optional
import uiautomation as auto
from win32com.client import Dispatch  # 修改导入方式

class WeChatController:
    def __init__(self):
        self.wechat_window = None
        self.member_list_window = None
        self.members_data = []
        self.shell = Dispatch("WScript.Shell")  # 创建 Shell 对象
        self.debug_mode = False  # 添加调试模式标志
        self.is_running = True  # 初始状态设为 True
        
        # 设置缓存文件路径
        self.cache_dir = os.path.join(os.path.expanduser("~"), "wechat_tool_cache")
        if not os.path.exists(self.cache_dir):
            os.makedirs(self.cache_dir)
        self.cache_file = os.path.join(self.cache_dir, "wechat_groups_cache.json")
        print(f"\n=== 初始化缓存 ===")
        print(f"缓存目录: {self.cache_dir}")
        print(f"缓存文件: {self.cache_file}")
        self.cached_groups = self.load_cache()
        
        # 如果有缓存数据，打印群聊数量
        if self.cached_groups and self.cached_groups.get("groups"):
            groups_count = len(self.cached_groups["groups"])
            last_update = self.cached_groups.get("last_update", "未知")
            print(f"已从缓存加载 {groups_count} 个群聊信息")
            print(f"上次更新时间: {last_update}")
            print("=== 缓存初始化完成 ===\n")
        else:
            print("未找到有效的缓存数据")
            print("=== 缓存初始化完成 ===\n")

    def load_cache(self) -> dict:
        """从缓存文件加载群聊信息"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
                print(f"已从缓存加载 {len(cache_data.get('groups', []))} 个群聊信息")
                return cache_data
            print("未找到缓存文件，将创建新的缓存")
            return {"last_update": None, "groups": {}}
        except Exception as e:
            print(f"加载缓存失败: {e}")
            return {"last_update": None, "groups": {}}

    def save_cache(self, groups_data: dict):
        """保存群聊信息到缓存文件"""
        try:
            print(f"\n=== 开始保存缓存 ===")
            print(f"缓存目录: {self.cache_dir}")
            print(f"缓存文件: {self.cache_file}")
            
            # 确保缓存目录存在
            if not os.path.exists(self.cache_dir):
                print(f"创建缓存目录: {self.cache_dir}")
                os.makedirs(self.cache_dir)
            
            # 打印要保存的数据
            print(f"要保存的数据: {json.dumps(groups_data, ensure_ascii=False, indent=2)}")
            
            # 保存数据
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(groups_data, f, ensure_ascii=False, indent=2)
            
            # 验证文件是否创建成功
            if os.path.exists(self.cache_file):
                file_size = os.path.getsize(self.cache_file)
                print(f"缓存文件创建成功，大小: {file_size} 字节")
                print(f"已将 {len(groups_data.get('groups', {}))} 个群聊信息保存到缓存")
            else:
                print("警告：缓存文件似乎未能成功创建")
            
            print("=== 缓存保存完成 ===\n")
        except Exception as e:
            print(f"保存缓存失败: {str(e)}")
            import traceback
            print(f"错误详情:\n{traceback.format_exc()}")

    def get_cached_groups(self) -> List[Dict]:
        """获取缓存的群聊列表"""
        if not self.cached_groups.get("groups"):
            print("缓存中没有群聊信息")
            return []
        
        groups = []
        for name, info in self.cached_groups["groups"].items():
            groups.append({
                "name": name,
                "member_count": info.get("member_count", "0"),
                "last_update": info.get("last_update")
            })
        return groups

    def update_group_cache(self, groups_data=None):
        """更新群聊缓存"""
        print("\n=== 开始更新群聊缓存 ===")
        
        if groups_data is None:
            groups = self.get_group_list()
            if not groups:
                print("未获取到群聊列表，请先打开微信界面")
                return False
        else:
            groups = groups_data
            
        try:
            # 转换格式以便存储
            formatted_groups = {}
            for group in groups:
                group_name = group["name"]  # 完整群名（包含成员数）
                formatted_groups[group_name] = {
                    "member_count": group["member_count"],
                    "last_update": datetime.now().isoformat(),
                    "members": {}  # 预留成员信息字段
                }
            
            print(f"处理群聊数据: {len(formatted_groups)} 个群聊")
            
            # 构建完整的缓存数据结构
            cache_data = {
                "last_update": datetime.now().isoformat(),
                "groups": formatted_groups
            }
            
            print(f"准备缓存 {len(formatted_groups)} 个群聊信息")
            
            # 保存到缓存
            self.save_cache(cache_data)
            self.cached_groups = cache_data
            
            print("=== 群聊缓存更新完成 ===\n")
            return True
        except Exception as e:
            print(f"更新缓存时发生错误: {str(e)}")
            import traceback
            print(f"错误详情:\n{traceback.format_exc()}")
            return False

    def get_group_members_with_cache(self, group_name: str, force_update: bool = False) -> Optional[Dict]:
        """获取群成员信息（支持缓存）"""
        # 检查缓存
        if not force_update and self.cached_groups.get("groups"):
            group_info = self.cached_groups["groups"].get(group_name)
            if group_info and group_info.get("members"):
                print(f"从缓存中获取群 {group_name} 的成员信息")
                return group_info["members"]
        
        # 如果没有缓存或强制更新，则获取新数据
        members = self.get_group_members(group_name)
        if members:
            # 更新缓存
            if self.cached_groups["groups"].get(group_name):
                self.cached_groups["groups"][group_name]["members"] = members
                self.cached_groups["groups"][group_name]["last_update"] = datetime.now().isoformat()
                self.save_cache(self.cached_groups["groups"])
            return members
        return None

    def find_wechat_window(self):
        """查找微信窗口"""
        try:
            print("开始查找微信窗口...")
            
            # 尝试多次查找微信窗口
            max_retries = 3
            for attempt in range(max_retries):
                print(f"尝试查找微信窗口 (第{attempt + 1}次)")
                
                # 先通过窗口句柄查找
                self.wechat_window = win32gui.FindWindow("WeChatMainWndForPC", "微信")
                if self.wechat_window and win32gui.IsWindowVisible(self.wechat_window):
                    print(f"通过窗口句柄找到微信窗口: {self.wechat_window}")
                    
                    # 再尝试通过UI自动化查找
                    try:
                        self.wechat_ui = auto.WindowControl(
                            searchDepth=1,
                            ClassName="WeChatMainWndForPC",
                            searchInterval=0.5
                        )
                        if self.wechat_ui.Exists(maxSearchSeconds=2):
                            print("UI自动化成功找到微信窗口")
                            return True
                    except Exception as e:
                        print(f"UI自动化查找失败: {e}")
                
                print("等待1秒后重试...")
                time.sleep(1)
            
            print("未找到微信窗口，请先打开微信界面")
            return False
            
        except Exception as e:
            print(f"查找微信窗口时发生错误: {e}")
            return False

    def activate_window(self):
        """激活微信窗口"""
        if not self.wechat_window:
            print("微信窗口未初始化，尝试重新查找...")
            if not self.find_wechat_window():
                print("无法找到微信窗口")
                return False
        
        try:
            max_retries = 3
            for attempt in range(max_retries):
                print(f"尝试激活微信窗口 (第{attempt + 1}次)")
                
                # 检查窗口是否最小化并还原
                placement = win32gui.GetWindowPlacement(self.wechat_window)
                if placement[1] == win32con.SW_SHOWMINIMIZED:
                    print("微信窗口已最小化，正在还原...")
                    win32gui.ShowWindow(self.wechat_window, win32con.SW_RESTORE)
                    time.sleep(0.5)
                
                # 移动窗口到固定位置
                try:
                    print("移动窗口到屏幕左上角...")
                    win32gui.MoveWindow(self.wechat_window, 0, 0, 1000, 700, True)
                    time.sleep(0.5)
                except Exception as e:
                    print(f"移动窗口失败: {e}")
                
                # 尝试多种方式激活窗口
                try:
                    # 方式1：使用SetForegroundWindow
                    win32gui.SetForegroundWindow(self.wechat_window)
                    time.sleep(0.2)
                    
                    # 方式2：使用BringWindowToTop
                    win32gui.BringWindowToTop(self.wechat_window)
                    time.sleep(0.2)
                    
                    # 方式3：使用Alt键
                    self.shell.SendKeys('%')
                    time.sleep(0.1)
                    win32gui.SetForegroundWindow(self.wechat_window)
                    
                    # 验证窗口是否真的激活
                    if win32gui.GetForegroundWindow() == self.wechat_window:
                        print("微信窗口已成功激活")
                        
                        # 获取窗口位置并点击
                        left, top, right, bottom = win32gui.GetWindowRect(self.wechat_window)
                        click_x = left + 10  # 距离左边10像素
                        click_y = top + 10   # 距离顶部10像素
                        print(f"点击窗口位置: x={click_x}, y={click_y}")
                        pyautogui.click(click_x, click_y)
                        time.sleep(0.5)
                        
                        return True
                except Exception as e:
                    print(f"激活尝试失败: {e}")
                
                print("等待0.5秒后重试...")
                time.sleep(0.5)
            
            print("无法激活微信窗口，请手动点击微信窗口")
            return False
            
        except Exception as e:
            print(f"激活窗口时发生错误: {e}")
            return False

    def find_member_list_window(self):
        """查找群成员列表窗口"""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                # 打印窗口信息，用于调试
                print(f"Window: {window_text}, Class: {class_name}")
                # 根据实际观察调整条件
                if ("群聊" in window_text or "聊天信息" in window_text) and class_name != "WeChatMainWndForPC":
                    self.member_list_window = hwnd
                    return False
            return True

        self.member_list_window = None  # 重置窗口句柄
        win32gui.EnumWindows(callback, None)
        return self.member_list_window is not None

    def get_member_list_items(self) -> List[Dict[str, str]]:
        """获取群成员列表中的所有成员信息"""
        self.members_data = []

        def enum_child_proc(hwnd, param):
            """遍历子窗口回调函数"""
            try:
                # 获取窗口类名和文本
                class_name = win32gui.GetClassName(hwnd)
                window_text = win32gui.GetWindowText(hwnd)
                
                # 打印窗口信息用于调试
                print(f"Child Window - Class: {class_name}, Text: {window_text}")
                
                # 根据实际情况调整判断条件
                if class_name == "ListBox" and window_text:  # 这里的类名需要根据实际情况调整
                    member_info = {
                        "name": window_text,
                        "wxid": "",  # TODO: 获取微信ID
                    }
                    self.members_data.append(member_info)
            except Exception as e:
                print(f"获取子窗口信息失败: {e}")

        if self.member_list_window:
            try:
                # 遍历群成员列表窗口的子窗口
                win32gui.EnumChildWindows(self.member_list_window, enum_child_proc, None)
            except Exception as e:
                print(f"遍历子窗口失败: {e}")

        return self.members_data

    def enable_debug_mode(self):
        """启用调试模式"""
        self.debug_mode = True
        print("调试模式已启用")

    def disable_debug_mode(self):
        """禁用调试模式"""
        self.debug_mode = False
        print("调试模式已禁用")

    def debug_ui_element(self, window_control, description=""):
        """交互式调试UI元素
        
        Args:
            window_control: 要调试的窗口控件
            description: 对正在调试的元素的描述
        """
        if not self.debug_mode:
            return

        print(f"\n=== 开始调试UI元素 ===")
        if description:
            print(f"正在调试: {description}")
        
        print("\n当前控件信息:")
        print(f"类型: {window_control.ControlType}")
        print(f"名称: {window_control.Name}")
        print(f"类名: {window_control.ClassName}")
        print(f"自动化ID: {window_control.AutomationId}")
        rect = window_control.BoundingRectangle
        print(f"位置: 左={rect.left}, 上={rect.top}, 右={rect.right}, 下={rect.bottom}")
        
        while True:
            print("\n请选择操作:")
            print("1 - 高亮显示元素位置")
            print("2 - 点击元素")
            print("3 - 显示子元素")
            print("4 - 显示父元素")
            print("5 - 继续执行（默认）")
            print("6 - 退出调试")
            print("Q - 退出程序")
            print("\n直接按回车将执行选项5")
            
            # 等待按键
            while True:
                try:
                    # 检查数字键1-6
                    for i in range(ord('1'), ord('7')):
                        if win32api.GetAsyncKeyState(i) & 0x8000:
                            choice = chr(i)
                            time.sleep(0.1)  # 等待按键释放
                            break
                    # 检查Q键
                    if win32api.GetAsyncKeyState(ord('Q')) & 0x8000:
                        choice = 'Q'
                        time.sleep(0.1)
                        break
                    # 检查回车键
                    if win32api.GetAsyncKeyState(0x0D) & 0x8000:  # 0x0D 是回车键的虚拟键码
                        choice = '5'  # 默认选项
                        time.sleep(0.1)
                        break
                    time.sleep(0.1)
                except:
                    continue
                
            print(f"\n选择了选项: {choice}")
            
            if choice == '1':
                # 使用pyautogui移动到元素位置
                pyautogui.moveTo(rect.xcenter(), rect.ycenter(), duration=0.5)
                print(f"已将鼠标移动到元素中心点({rect.xcenter()}, {rect.ycenter()})")
                
            elif choice == '2':
                print("确认要点击该元素吗？按Y确认，按N取消...")
                while True:
                    try:
                        if win32api.GetAsyncKeyState(ord('Y')) & 0x8000:
                            window_control.Click()
                            print("已点击元素")
                            time.sleep(0.5)
                            break
                        elif win32api.GetAsyncKeyState(ord('N')) & 0x8000:
                            print("已取消点击")
                            break
                        time.sleep(0.1)
                    except:
                        continue
                
            elif choice == '3':
                print("\n子元素列表:")
                for i, child in enumerate(window_control.GetChildren(), 1):
                    print(f"{i}. 类型={child.ControlType}, 名称={child.Name}, 类名={child.ClassName}")
                print("\n按任意键继续...")
                while True:
                    try:
                        if any(win32api.GetAsyncKeyState(i) & 0x8000 for i in range(256)):
                            break
                        time.sleep(0.1)
                    except:
                        continue
                
            elif choice == '4':
                parent = window_control.GetParentControl()
                print("\n父元素信息:")
                print(f"类型: {parent.ControlType}")
                print(f"名称: {parent.Name}")
                print(f"类名: {parent.ClassName}")
                print("\n按任意键继续...")
                while True:
                    try:
                        if any(win32api.GetAsyncKeyState(i) & 0x8000 for i in range(256)):
                            break
                        time.sleep(0.1)
                    except:
                        continue
                
            elif choice == '5':
                print("继续执行...")
                break
                
            elif choice == '6':
                print("退出调试模式")
                self.disable_debug_mode()
                break
                
            elif choice.upper() == 'Q':
                print("退出程序")
                import sys
                sys.exit(0)
                
            else:
                print("无效的选项，请重新输入")

    def get_group_list(self, use_cache=True):
        """获取群聊列表
        
        Args:
            use_cache: 是否优先使用缓存数据，默认为True
        """
        try:
            # 重置运行状态
            self.is_running = True
            
            # 如果允许使用缓存且缓存中有数据，直接返回缓存的群聊列表
            if use_cache and self.cached_groups and self.cached_groups.get("groups"):
                print("\n=== 使用缓存数据 ===")
                groups = []
                for name, info in self.cached_groups["groups"].items():
                    groups.append({
                        "name": name,
                        "member_count": info.get("member_count", "0")
                    })
                print(f"从缓存中获取到 {len(groups)} 个群聊")
                print("=== 使用缓存完成 ===\n")
                return groups
                
            print("\n=== 开始从微信获取群聊列表 ===")
            
            # 检查任务是否已终止
            if not self.is_running:
                print("任务已终止")
                self.stop_task()  # 确保关闭所有窗口
                return None
            
            # 确保能找到微信窗口
            if not self.find_wechat_window():
                print("未找到微信窗口，请确保微信已启动")
                self.stop_task()
                return None
            
            # 检查任务是否已终止
            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None
            
            # 确保能激活微信窗口
            if not self.activate_window():
                print("无法激活微信窗口")
                self.stop_task()
                return None
            
            # 检查任务是否已终止
            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None
            
            contact_manage_window = None  # 用于存储通讯录管理窗口引用
            
            try:
                max_retries = 3
                for attempt in range(max_retries):
                    # 检查任务是否已终止
                    if not self.is_running:
                        print("任务已终止")
                        self.stop_task()
                        return None
                    
                    print(f"\n尝试获取群聊列表 (第{attempt + 1}次)")
                    
                    try:
                        # 点击通讯录按钮
                        if not self.is_running:
                            self.stop_task()
                            return None
                            
                        print("点击通讯录按钮...")
                        window_rect = self.wechat_ui.BoundingRectangle
                        click_x = window_rect.left + 30
                        click_y = window_rect.top + 140
                        pyautogui.click(click_x, click_y)
                        time.sleep(1)
                        
                        # 滚动到顶部
                        print("滚动到顶部...")
                        # 移动到通讯录管理按钮的位置
                        manage_x = window_rect.left + 100
                        manage_y = window_rect.top + 100
                        pyautogui.moveTo(manage_x, manage_y)
                        time.sleep(0.2)
                        
                        # 连续滚动多次确保到达顶部
                        for _ in range(5):
                            pyautogui.scroll(1000)
                            time.sleep(0.2)
                        
                        # 点击通讯录管理按钮
                        if not self.is_running:
                            self.stop_task()
                            return None
                            
                        print("点击通讯录管理按钮...")
                        manage_x = window_rect.left + 100
                        manage_y = window_rect.top + 100
                        pyautogui.click(manage_x, manage_y)
                        time.sleep(1)
                        
                        # 等待通讯录管理窗口出现
                        if not self.is_running:
                            self.stop_task()
                            return None
                            
                        print("等待通讯录管理窗口...")
                        time.sleep(2)
                        
                        # 获取并移动通讯录管理窗口
                        print("查找通讯录管理窗口...")
                        contact_manage_window = auto.WindowControl(searchDepth=1, Name="通讯录管理")
                        if not contact_manage_window.Exists(maxSearchSeconds=3):
                            print("未找到通讯录管理窗口，请先打开微信界面")
                            continue
                        
                        if not self.is_running:
                            self.stop_task()
                            return None
                        
                        # 移动窗口到固定位置
                        try:
                            window_handle = contact_manage_window.NativeWindowHandle
                            win32gui.MoveWindow(window_handle, 0, 0, 1000, 700, True)
                            time.sleep(0.5)
                        except Exception as e:
                            print(f"移动窗口失败: {e}")
                            continue
                        
                        if not self.is_running:
                            self.stop_task()
                            return None
                        
                        # 点击群聊标签
                        print("点击群聊标签...")
                        window_rect = contact_manage_window.BoundingRectangle
                        click_x = window_rect.left + 100
                        click_y = window_rect.top + 200
                        pyautogui.click(click_x, click_y)
                        time.sleep(1)
                        
                        # 查找左侧列表
                        if not self.is_running:
                            self.stop_task()
                            return None
                            
                        print("查找群聊列表...")
                        list_view = contact_manage_window.ListControl()
                        if not list_view.Exists(maxSearchSeconds=3):
                            print("未找到群聊列表，请先打开微信界面")
                            continue
                        
                        # 收集群聊信息
                        initial_groups = []
                        
                        def collect_group_items(control):
                            """递归收集所有群聊项"""
                            if not self.is_running:  # 在收集过程中也检查终止状态
                                return False
                                
                            try:
                                rect = control.BoundingRectangle
                                if control.ControlType == 50007 and rect.left < 200:
                                    group_name = None
                                    member_count = None
                                    
                                    def find_text(ctrl):
                                        if not self.is_running:  # 在查找文本过程中也检查终止状态
                                            return False
                                            
                                        nonlocal group_name, member_count
                                        try:
                                            if ctrl.ControlType == 50020:
                                                text = ctrl.Name
                                                if text:
                                                    if "(" in text and ")" in text:
                                                        member_count = text.strip("()")
                                                        try:
                                                            int(member_count)
                                                        except:
                                                            member_count = None
                                                    elif not text.startswith("当前群聊"):
                                                        group_name = text
                                        except Exception as e:
                                            print(f"处理文本控件失败: {e}")
                                        
                                        for child in ctrl.GetChildren():
                                            if find_text(child) is False:
                                                return False
                                    
                                    if find_text(control) is False:
                                        return False
                                        
                                    if group_name and member_count:
                                        if not any(g[0] == group_name for g in initial_groups):
                                            initial_groups.append((group_name, member_count))
                                
                                for child in control.GetChildren():
                                    if collect_group_items(child) is False:
                                        return False
                                    
                            except Exception as e:
                                print(f"处理群聊项时出错: {e}")
                            return True
                        
                        # 开始收集群聊信息
                        print("\n=== 开始收集群聊信息 ===")
                        
                        # 先收集第一屏的群聊
                        if collect_group_items(list_view) is False:
                            self.stop_task()
                            return None
                            
                        last_count = len(initial_groups)
                        print(f"第一屏找到 {last_count} 个群聊")
                        
                        # 开始滚动收集
                        scroll_count = 0
                        max_no_change_count = 3
                        no_change_count = 0
                        
                        print("\n=== 开始滚动收集群聊 ===")
                        while self.is_running:  # 使用is_running作为循环条件
                            # 检查停止状态
                            if not self.is_running:
                                print("\n检测到停止信号！正在终止滚动收集...")
                                break
                            
                            # 计算滚动位置
                            scroll_x = window_rect.left + 50
                            scroll_y = window_rect.top + 300
                            
                            # 移动到滚动位置
                            print(f"移动到滚动位置: x={scroll_x}, y={scroll_y}")
                            pyautogui.moveTo(scroll_x, scroll_y)
                            time.sleep(0.5)
                            
                            # 再次检查停止状态
                            if not self.is_running:
                                print("\n检测到停止信号！正在终止滚动操作...")
                                break
                            
                            # 执行滚动
                            print(f"执行第 {scroll_count + 1} 次滚动...")
                            pyautogui.scroll(-700)
                            time.sleep(1)
                            
                            # 滚动后检查停止状态
                            if not self.is_running:
                                print("\n检测到停止信号！正在终止数据收集...")
                                break
                            
                            # 收集当前可见的群聊
                            print("收集当前页面的群聊...")
                            if collect_group_items(list_view) is False:
                                print("\n检测到停止信号或收集完成！")
                                break
                                
                            current_count = len(initial_groups)
                            new_items = current_count - last_count
                            print(f"当前共找到 {current_count} 个群聊 (新增: {new_items})")
                            
                            # 检查是否有新增
                            if new_items == 0:
                                no_change_count += 1
                                print(f"连续 {no_change_count} 次没有新增群聊")
                                if no_change_count >= max_no_change_count:
                                    print("已到达列表底部，停止滚动")
                                    break
                            else:
                                no_change_count = 0
                                
                            last_count = current_count
                            scroll_count += 1
                            
                            # 每次循环结束前检查停止状态
                            if not self.is_running:
                                print("\n检测到停止信号！正在结束本次循环...")
                                break
                            
                            time.sleep(0.5)
                        
                        # 循环结束后检查是否是因为停止信号退出
                        if not self.is_running:
                            print("\n=== 任务已被用户终止 ===")
                            break
                        
                        # 循环结束后的处理
                        if not self.is_running:
                            print("正在清理并关闭窗口...")
                            self.stop_task()
                            return None
                        
                        if initial_groups:
                            print(f"\n总共找到 {len(initial_groups)} 个群聊")
                            groups = []
                            for group_name, member_count in initial_groups:
                                groups.append({
                                    "name": f"{group_name}({member_count})",
                                    "member_count": member_count
                                })
                            
                            # 更新缓存
                            print("\n=== 更新缓存 ===")
                            self.update_group_cache(groups)
                            
                            # 关闭通讯录管理窗口
                            self.stop_task()
                            
                            return groups
                        
                        print("未找到任何群聊，重试...")
                        
                    except Exception as e:
                        print(f"获取群聊列表时出错: {e}")
                        if attempt < max_retries - 1:
                            print("等待2秒后重试...")
                            time.sleep(2)
                        continue
                
                print("\n=== 获取群聊列表失败 ===")
                return None
                
            finally:
                # 确保在任何情况下都关闭窗口
                if not self.is_running and self.wechat_window:
                    self.stop_task()
                
        except Exception as e:
            print(f"获取群聊列表失败: {e}")
            import traceback
            print(f"错误详情:\n{traceback.format_exc()}")
            self.stop_task()
            return None

    def get_group_members(self, group_name):
        """获取指定群的成员列表"""
        # 启动新任务
        if not self.start_task():
            return None

        try:
            print(f"\n=== 开始获取群 {group_name} 的成员列表 ===")
            members = {}  # 初始化成员字典
            
            # 处理群名，去掉括号和数字
            search_name = group_name
            if "(" in group_name:
                search_name = group_name.split("(")[0].strip()
            print(f"使用处理后的群名进行搜索: {search_name}")
            
            # 确保微信窗口已激活
            if not self.activate_window():
                print("无法激活微信窗口")
                self.stop_task()
                return None

            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None
            
            # 点击聊天按钮确保在主界面
            chat_btn = self.wechat_ui.ButtonControl(Name="聊天")
            if chat_btn.Exists():
                chat_btn.Click()
                time.sleep(0.2)
            
            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None
            
            # 使用 UI 自动化查找搜索框
            search_box = self.wechat_ui.EditControl(Name="搜索")
            if not search_box.Exists(maxSearchSeconds=2):
                print("未找到搜索框")
                self.stop_task()
                return None
            
            # 获取搜索框位置并输入群名
            search_rect = search_box.BoundingRectangle
            search_box.Click()
            time.sleep(0.2)
            
            # 使用剪贴板来输入群名
            original_clipboard = pyperclip.paste()  # 保存当前剪贴板内容
            try:
                pyautogui.hotkey('ctrl', 'a')  # 全选
                pyautogui.press('backspace')    # 清除
                pyperclip.copy(search_name)     # 复制群名到剪贴板
                pyautogui.hotkey('ctrl', 'v')   # 粘贴
            finally:
                time.sleep(0.2)
                pyperclip.copy(original_clipboard)  # 恢复原始剪贴板内容

            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None

            print("等待搜索结果...")
            time.sleep(0.5)
            
            # 计算并点击搜索结果的位置
            result_x = search_rect.left + 20
            result_y = search_rect.bottom + 100
            print(f"点击搜索结果位置: x={result_x}, y={result_y}")
            pyautogui.click(result_x, result_y)
            time.sleep(0.5)
            
            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None
            
            # 点击右侧的"..."设置按钮
            print("查找设置按钮...")
            more_btn = self.wechat_ui.ButtonControl(Name="聊天信息")
            if not more_btn.Exists(maxSearchSeconds=2):
                print("尝试查找备选设置按钮...")
                more_btn = self.wechat_ui.ButtonControl(Name="更多")
                if not more_btn.Exists():
                    more_btn = self.wechat_ui.ButtonControl(searchDepth=5, ClassName="Button")
            
            if not self.is_running:
                print("任务已终止")
                self.stop_task()
                return None
            
            if more_btn.Exists():
                print("点击设置按钮")
                more_btn.Click()
                time.sleep(0.5)
                
                if not self.is_running:
                    print("任务已终止")
                    self.stop_task()
                    return None
                
                # 点击"查看更多"按钮
                print("查找并点击查看更多按钮...")
                view_more_btn = self.wechat_ui.ButtonControl(Name="查看更多")
                if not view_more_btn.Exists(maxSearchSeconds=2):
                    print("尝试查找备选查看更多按钮...")
                    view_more_btn = self.wechat_ui.ButtonControl(searchDepth=5, Name="群成员")
                
                if not self.is_running:
                    print("任务已终止")
                    self.stop_task()
                    return None
                
                if view_more_btn.Exists():
                    print("点击查看更多按钮")
                    view_more_btn.Click()
                    time.sleep(1)
                    
                    if not self.is_running:
                        print("任务已终止")
                        self.stop_task()
                        return None
                    
                    # 获取主窗口的位置
                    main_rect = self.wechat_ui.BoundingRectangle
                    print(f"主窗口位置: 左={main_rect.left}, 上={main_rect.top}, 右={main_rect.right}, 下={main_rect.bottom}")
                    
                    # 初始化成员集合
                    initial_members = set()
                    processed_members = set()
                    
                    def collect_member_items(control, depth=0):
                        """递归收集所有成员项"""
                        if not self.is_running:
                            print("收集成员过程中检测到停止信号")
                            return False
                            
                        try:
                            # 每递归3层检查一次停止信号
                            if depth % 3 == 0 and not self.is_running:
                                print("递归过程中检测到停止信号")
                                return False
                            
                            rect = control.BoundingRectangle
                            if control.ControlType in [50000, 50020]:
                                name = control.Name
                                if name and name not in processed_members:
                                    # 过滤无效的成员名
                                    if (not any(x in name for x in [
                                        "群聊", "聊天", "消息", "发送", "置顶", "最小化", "最大化", "关闭",
                                        "查看更多", "群公告", "备注", "清空", "退出", "保存", "显示",
                                        ".com", ".cn", "[图片]", "[视频]", "[链接]",
                                        "播放：", "UP主：", "pdf", "weixinfile", "微信"
                                    ]) and
                                    not name.startswith(("收起", "直播", "语音", "发送", "置顶")) and
                                    not name.endswith(("M", "K", "B")) and
                                    "：" not in name
                                    ):
                                        if rect.left > main_rect.left + (main_rect.width() * 0.6):
                                            initial_members.add(name)
                                            processed_members.add(name)
                            
                            # 每处理3个子控件检查一次停止信号
                            for i, child in enumerate(control.GetChildren()):
                                if i % 3 == 0 and not self.is_running:
                                    print("处理子控件时检测到停止信号")
                                    return False
                                if collect_member_items(child, depth + 1) is False:
                                    return False
                                
                        except Exception as e:
                            print(f"处理成员项时出错: {str(e)}")
                        return True
                    
                    print("\n=== 开始收集成员信息 ===")
                    
                    # 收集当前可见的成员
                    if collect_member_items(self.wechat_ui) is False:
                        print("收集成员过程被终止")
                        self.stop_task()
                        return None
                    
                    if not self.is_running:
                        print("任务已终止")
                        self.stop_task()
                        return None
                    
                    # 处理收集到的成员信息
                    if initial_members:
                        print(f"\n总共找到 {len(initial_members)} 个成员")
                        for member_name in initial_members:
                            members[member_name] = {
                                "group": group_name
                            }
                        print(f"成功处理 {len(members)} 个成员信息")
                        return members
                    else:
                        print("未找到任何成员")
                        self.stop_task()
                        return None
                else:
                    print("未找到查看更多按钮")
                    self.stop_task()
                    return None
            else:
                print("未找到设置按钮")
                self.stop_task()
                return None
            
        except Exception as e:
            print(f"获取群成员失败: {str(e)}")
            import traceback
            print(f"错误详情:\n{traceback.format_exc()}")
            self.stop_task()
            return None
        finally:
            if not self.is_running:  # 修改条件判断
                print("任务已终止，执行清理操作")
                self.stop_task()

    def get_member_info(self, member_id):
        """获取成员信息"""
        # TODO: 实现成员信息获取逻辑
        pass 

    def stop_task(self):
        """停止当前任务"""
        print("\n=== 正在终止任务 ===")
        # 立即设置终止标志
        self.is_running = False
        
        try:
            # 先尝试使用UI自动化关闭窗口
            try:
                # 群成员列表窗口
                member_list = auto.WindowControl(searchDepth=1, Name="聊天信息")
                if member_list.Exists(maxSearchSeconds=1):
                    print("关闭群成员列表窗口...")
                    member_list.SendKeys("{ESC}")
                    time.sleep(0.2)
                
                # 通讯录管理窗口
                contact_manage = auto.WindowControl(searchDepth=1, Name="通讯录管理")
                if contact_manage.Exists(maxSearchSeconds=1):
                    print("关闭通讯录管理窗口...")
                    contact_manage.SendKeys("{ESC}")
                    time.sleep(0.2)
            except Exception as e:
                print(f"UI自动化关闭窗口时出错: {e}")
            
            # 然后使用Win32 API强制关闭
            def force_close_window(window_title):
                """强制关闭指定窗口"""
                try:
                    def callback(hwnd, windows):
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd)
                            window_class = win32gui.GetClassName(hwnd)
                            
                            # 排除主窗口和主程序窗口
                            if (window_class == "WeChatMainWndForPC" or 
                                "微信群成员分析工具" in title):
                                return True
                                
                            if window_title in title:
                                windows.append(hwnd)
                        return True
                    
                    windows = []
                    win32gui.EnumWindows(callback, windows)
                    
                    for hwnd in windows:
                        try:
                            title = win32gui.GetWindowText(hwnd)
                            # 再次确认不是主程序窗口
                            if "微信群成员分析工具" not in title:
                                print(f"尝试关闭窗口: {title}")
                                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                                time.sleep(0.2)
                        except Exception as e:
                            print(f"关闭窗口失败: {e}")
                            
                    return len(windows) > 0
                except Exception as e:
                    print(f"查找窗口失败: {e}")
                    return False
            
            # 按顺序关闭窗口
            windows_to_close = ["群成员", "聊天信息", "群聊", "通讯录管理"]
            for title in windows_to_close:
                print(f"尝试关闭 {title} 窗口...")
                if force_close_window(title):
                    print(f"已关闭 {title} 窗口")
                    time.sleep(0.3)  # 稍微增加等待时间
            
            # 最后点击微信主窗口以确保焦点回到主窗口
            if self.wechat_window and win32gui.IsWindow(self.wechat_window):
                try:
                    win32gui.SetForegroundWindow(self.wechat_window)
                    rect = win32gui.GetWindowRect(self.wechat_window)
                    pyautogui.click(rect[0] + 10, rect[1] + 10)
                except:
                    pass
            
            print("任务终止完成")
            
        except Exception as e:
            print(f"终止任务时出错: {str(e)}")
            import traceback
            print(f"错误详情:\n{traceback.format_exc()}")
        finally:
            # 最后再次确保标志被设置为False
            self.is_running = False

    def start_task(self):
        """开始新任务"""
        print("\n=== 开始新任务 ===")
        self.is_running = True
        return True 