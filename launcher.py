import sys
import os
import traceback
import logging
from datetime import datetime

# 添加项目根目录和src目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)  # 添加当前目录到Python路径
src_dir = os.path.join(current_dir, 'src')
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)  # 添加src目录到Python路径

# 设置日志
log_dir = os.path.join(current_dir, 'logs')
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_file = os.path.join(log_dir, f'error_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def test_mode():
    """测试模式，检查程序是否可以正常运行"""
    try:
        # 记录详细的环境信息
        logging.info("=== 测试模式 ===")
        logging.info(f"Python Version: {sys.version}")
        logging.info(f"Executable Path: {sys.executable}")
        logging.info(f"Current Directory: {os.getcwd()}")
        logging.info(f"Script Directory: {current_dir}")
        logging.info(f"Python Path: {sys.path}")
        
        # 测试导入主要模块
        logging.info("测试导入主要模块...")
        from PyQt5.QtWidgets import QApplication
        
        # 测试创建应用实例
        logging.info("测试创建应用实例...")
        app = QApplication([])
        
        # 导入主窗口
        logging.info("导入主窗口模块...")
        from ui.main_window import MainWindow
        
        # 测试创建主窗口
        logging.info("测试创建主窗口...")
        window = MainWindow()
        
        logging.info("测试完成：程序可以正常运行")
        return 0
        
    except Exception as e:
        logging.error(f"测试失败：{str(e)}")
        logging.error(traceback.format_exc())
        return 1

try:
    # 检查是否在测试模式下运行
    if len(sys.argv) > 1 and sys.argv[1] == '--test':
        sys.exit(test_mode())
        
    # 记录详细的环境信息
    logging.info("=== 启动信息 ===")
    logging.info(f"Python Version: {sys.version}")
    logging.info(f"Executable Path: {sys.executable}")
    logging.info(f"Current Directory: {os.getcwd()}")
    logging.info(f"Script Directory: {current_dir}")
    logging.info("Python Path:")
    for path in sys.path:
        logging.info(f"  {path}")
    logging.info("Environment Variables:")
    for key, value in os.environ.items():
        logging.info(f"  {key}: {value}")
    
    print("正在启动程序...")
    print(f"日志文件: {log_file}")
    
    # 先导入QApplication
    logging.info("导入QApplication...")
    from PyQt5.QtWidgets import QApplication
    
    # 创建应用实例
    logging.info("创建应用实例...")
    app = QApplication(sys.argv)
    
    # 导入主窗口模块
    logging.info("正在导入主窗口模块...")
    from ui.main_window import MainWindow
    
    logging.info("创建主窗口...")
    window = MainWindow()
    
    logging.info("显示主窗口...")
    window.show()
    
    logging.info("进入事件循环...")
    sys.exit(app.exec_())
    
except Exception as e:
    # 记录详细的错误信息
    error_msg = f"程序启动失败！\n错误类型: {type(e).__name__}\n错误信息: {str(e)}"
    logging.error(error_msg)
    logging.error("详细堆栈跟踪:")
    logging.error(traceback.format_exc())
    
    # 显示错误对话框
    try:
        from PyQt5.QtWidgets import QMessageBox
        error_box = QMessageBox()
        error_box.setIcon(QMessageBox.Critical)
        error_box.setWindowTitle("错误")
        error_box.setText(f"程序启动失败！\n错误信息: {str(e)}\n\n详细错误日志已保存到: {log_file}")
        error_box.exec_()
    except:
        # 如果连错误对话框都无法显示，则使用命令行输出
        print(error_msg)
        print(f"\n详细错误日志已保存到: {log_file}")
        input("按回车键退出...") 