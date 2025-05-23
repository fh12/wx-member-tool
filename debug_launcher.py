import sys
import os
import traceback
import logging
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMessageBox

def setup_logging():
    """设置日志记录"""
    # 确保logs目录存在
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 设置日志文件
    log_file = os.path.join(log_dir, f"debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    # 配置日志记录器
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)  # 同时输出到控制台
        ]
    )
    return log_file

def setup_python_path():
    """设置Python路径"""
    try:
        # 获取当前目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 将当前目录添加到Python路径
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)
            
        # 将src目录添加到Python路径
        src_dir = os.path.join(current_dir, 'src')
        if os.path.exists(src_dir) and src_dir not in sys.path:
            sys.path.insert(0, src_dir)
            
        logging.info(f"Python路径设置完成: {sys.path}")
        return True
    except Exception as e:
        logging.error(f"设置Python路径失败: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def test_imports():
    """测试导入所需的模块"""
    try:
        logging.info("测试导入PyQt5...")
        from PyQt5.QtWidgets import QMainWindow
        
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
        logging.error(traceback.format_exc())
        return False

def test_paths():
    """测试路径设置"""
    try:
        logging.info("检查当前目录...")
        current_dir = os.path.dirname(os.path.abspath(__file__))
        logging.info(f"当前目录: {current_dir}")
        
        logging.info("检查src目录...")
        src_dir = os.path.join(current_dir, 'src')
        if not os.path.exists(src_dir):
            logging.error(f"src目录不存在: {src_dir}")
            return False
            
        logging.info("检查资源文件...")
        icon_path = os.path.join(src_dir, 'assets', 'icon.ico')
        if not os.path.exists(icon_path):
            logging.error(f"图标文件不存在: {icon_path}")
            return False
            
        logging.info("路径检查完成")
        return True
    except Exception as e:
        logging.error(f"路径检查失败: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def test_license():
    """测试许可证功能"""
    try:
        logging.info("测试许可证管理...")
        from src.core.license import LicenseManager
        
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
        logging.error(traceback.format_exc())
        return False

def test_ui():
    """测试UI界面"""
    try:
        logging.info("测试UI界面...")
        from src.ui.main_window import MainWindow
        
        app = QApplication([])
        window = MainWindow()
        window.show()
        
        logging.info("UI界面测试成功")
        return True
    except Exception as e:
        logging.error(f"UI界面测试失败: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def main():
    """主测试函数"""
    try:
        # 设置日志
        log_file = setup_logging()
        logger = logging.getLogger(__name__)
        
        logger.info("=== 开始调试测试 ===")
        
        # 设置Python路径
        if not setup_python_path():
            logger.error("Python路径设置失败")
            return 1
        
        # 测试导入
        if not test_imports():
            logger.error("模块导入测试失败")
            return 1
            
        # 测试路径
        if not test_paths():
            logger.error("路径测试失败")
            return 1
            
        # 测试许可证
        if not test_license():
            logger.error("许可证测试失败")
            return 1
            
        # 测试UI
        if not test_ui():
            logger.error("UI测试失败")
            return 1
            
        logger.info("=== 所有测试通过 ===")
        return 0
        
    except Exception as e:
        logger.error(f"测试过程发生错误: {str(e)}")
        logger.error(traceback.format_exc())
        return 1

if __name__ == "__main__":
    sys.exit(main()) 