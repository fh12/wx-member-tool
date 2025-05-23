import sys
import os

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.core.license import LicenseManager

def generate_license():
    """生成激活码工具"""
    print("=== 激活码生成工具 ===")
    
    while True:
        user_id = input("\n请输入用户ID (输入 'q' 退出): ").strip()
        
        if user_id.lower() == 'q':
            break
            
        if len(user_id) != 12:
            print("错误：用户ID必须是12位！")
            continue
            
        try:
            # 生成激活码
            license_manager = LicenseManager()
            license_code = license_manager.generate_license_code(user_id)
            
            print("\n=== 生成结果 ===")
            print(f"用户ID: {user_id}")
            print(f"激活码: {license_code}")
            print("=" * 20)
            
        except Exception as e:
            print(f"错误：{str(e)}")

if __name__ == "__main__":
    generate_license() 