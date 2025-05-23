import PyInstaller.__main__
import os
import shutil

def build_exe():
    # 确保dist目录为空
    if os.path.exists("dist"):
        shutil.rmtree("dist", ignore_errors=True)
    if os.path.exists("build"):
        shutil.rmtree("build", ignore_errors=True)
        
    # PyInstaller打包参数
    params = [
        'launcher.py',  # 使用正常的启动器
        '--name=微信群成员分析工具',  # 生成的exe名称
        '--windowed',  # 使用窗口模式，不显示控制台
        '--onefile',  # 使用单文件模式
        '--icon=src/assets/icon.ico',  # 图标文件
        '--add-data=src;src',  # 添加整个src目录
        '--add-binary=src/assets/icon.ico;.',  # 添加图标文件到根目录
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=keyboard',
        '--hidden-import=win32gui',
        '--hidden-import=win32con',
        '--hidden-import=win32api',
        '--hidden-import=win32com',
        '--hidden-import=win32com.client',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
        '--hidden-import=pyautogui',
        '--hidden-import=uiautomation',
        '--hidden-import=PIL',
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=PyQt5.QtWidgets',
        '--collect-all=win32com',
        '--collect-all=pyautogui',
        '--collect-all=uiautomation',
        '--collect-all=pythoncom',
        '--collect-all=pywintypes',
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不确认覆盖
    ]
    
    # 执行打包
    PyInstaller.__main__.run(params)
    
    print("打包完成！exe文件位于dist目录中。")

if __name__ == "__main__":
    build_exe() 