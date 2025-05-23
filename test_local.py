import sys
import os
import logging
from datetime import datetime

def setup_test_env():
    """设置测试环境"""
    # 添加项目根目录到Python路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = current_dir  # 当前目录就是项目根目录
    if project_root not in sys.path:
        sys.path.insert(0, project_root)
    
    # 设置日志
    log_dir = os.path.join(current_dir, 'logs')
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, f'test_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return log_file

def test_imports():
    """测试所有必需模块的导入"""
    try:
        logging.info("测试导入PyQt5...")
        from PyQt5.QtWidgets import QApplication
        
        logging.info("测试导入pandas...")
        import pandas as pd
        
        logging.info("测试导入其他依赖...")
        import keyboard
        import win32gui
        import win32con
        import win32api
        import win32com.client
        import pythoncom
        import pyautogui
        import uiautomation
        import PIL
        
        logging.info("所有模块导入成功")
        return True
    except Exception as e:
        logging.error(f"模块导入失败: {str(e)}")
        return False

def test_ui():
    """测试UI界面"""
    try:
        logging.info("测试UI界面初始化...")
        import sys
        from PyQt5.QtWidgets import QApplication
        
        # 确保src目录在Python路径中
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.join(current_dir, 'src')
        if src_dir not in sys.path:
            sys.path.insert(0, src_dir)
            
        from ui.main_window import MainWindow
        from core.license import LicenseManager
        
        app = QApplication([])
        window = MainWindow()
        logging.info("UI界面初始化成功")
        
        # 不实际显示窗口，只测试创建是否成功
        return True
    except Exception as e:
        logging.error(f"UI界面测试失败: {str(e)}")
        return False

def test_license():
    """测试许可证功能"""
    try:
        logging.info("测试许可证管理...")
        # 确保src目录在Python路径中
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.join(current_dir, 'src')
        if src_dir not in sys.path:
            sys.path.insert(0, src_dir)
            
        from core.license import LicenseManager
        
        license_manager = LicenseManager()
        user_id = license_manager.get_user_id()
        logging.info(f"获取到用户ID: {user_id}")
        
        # 测试生成激活码
        license_code = license_manager.generate_license_code(user_id)
        logging.info(f"生成的激活码: {license_code}")
        
        # 测试验证激活码
        if license_manager.verify_license(license_code):
            logging.info("激活码验证成功")
            return True
        else:
            logging.error("激活码验证失败")
            return False
    except Exception as e:
        logging.error(f"许可证测试失败: {str(e)}")
        return False

def test_wechat():
    """测试微信控制功能"""
    try:
        logging.info("测试微信控制功能...")
        # 确保src目录在Python路径中
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.join(current_dir, 'src')
        if src_dir not in sys.path:
            sys.path.insert(0, src_dir)
            
        from core.wechat import WeChatController
        
        controller = WeChatController()
        # 只测试初始化，不实际操作微信
        logging.info("微信控制器初始化成功")
        return True
    except Exception as e:
        logging.error(f"微信控制功能测试失败: {str(e)}")
        return False

def main():
    """主测试函数"""
    log_file = setup_test_env()
    logging.info("=== 开始本地测试 ===")
    
    tests = [
        ("导入测试", test_imports),
        ("许可证测试", test_license),
        ("UI测试", test_ui),
        ("微信控制测试", test_wechat)
    ]
    
    all_passed = True
    for test_name, test_func in tests:
        logging.info(f"\n开始{test_name}...")
        try:
            if test_func():
                logging.info(f"{test_name}通过")
            else:
                logging.error(f"{test_name}失败")
                all_passed = False
        except Exception as e:
            logging.error(f"{test_name}发生错误: {str(e)}")
            all_passed = False
    
    if all_passed:
        logging.info("\n=== 所有测试通过 ===")
        print(f"\n测试通过！详细日志已保存到：{log_file}")
        return 0
    else:
        logging.error("\n=== 测试失败 ===")
        print(f"\n测试失败！详细日志已保存到：{log_file}")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 