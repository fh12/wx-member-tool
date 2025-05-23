import uuid
import hashlib
import json
import os
import sys
from datetime import datetime, timedelta

class LicenseManager:
    def __init__(self):
        # 获取应用数据目录
        self.app_data_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'WeChatGroupAnalyzer')
        if not os.path.exists(self.app_data_dir):
            os.makedirs(self.app_data_dir)
            
        self.license_file = os.path.join(self.app_data_dir, "license.json")
        # 添加固定的盐值
        self._salt = "wx_group_analyzer_2024"
        self._permanent_salt = "wx_group_analyzer_permanent_2024"  # 永久会员盐值
        
    def get_user_id(self):
        """获取用户ID（基于MAC地址）"""
        # 获取第一个有效的MAC地址
        mac = uuid.getnode()
        # 转换为十六进制字符串，去掉'0x'前缀，保证12位
        return '{:012x}'.format(mac)
        
    def generate_license_code(self, user_id, is_permanent=False):
        """根据用户ID生成激活码"""
        # 根据是否是永久会员选择盐值
        salt = self._permanent_salt if is_permanent else self._salt
        # 将用户ID和盐值组合
        salted_code = f"{user_id}{salt}"
        # 使用SHA-256进行加密
        sha256 = hashlib.sha256(salted_code.encode()).hexdigest()
        # 返回前20位作为激活码
        return sha256[:20].upper()
        
    def verify_license(self, input_code):
        """验证激活码"""
        user_id = self.get_user_id()
        # 验证年费会员激活码
        yearly_code = self.generate_license_code(user_id, is_permanent=False)
        # 验证永久会员激活码
        permanent_code = self.generate_license_code(user_id, is_permanent=True)
        
        input_code = input_code.upper()
        if input_code == permanent_code:
            # 保存永久会员信息
            self.save_license_info(user_id, input_code, is_permanent=True)
            return True, True, True  # 返回(注册成功, 是VIP, 是永久)
        elif input_code == yearly_code:
            # 保存年费会员信息
            self.save_license_info(user_id, input_code, is_permanent=False)
            return True, True, False  # 返回(注册成功, 是VIP, 不是永久)
        return False, False, False  # 返回(注册失败, 不是VIP, 不是永久)
        
    def save_license_info(self, user_id, license_code, is_permanent=False):
        """保存注册信息"""
        # 计算到期时间
        if is_permanent:
            expire_time = "永久有效"
        else:
            # 一年后的今天
            expire_time = (datetime.now() + timedelta(days=365)).strftime('%Y-%m-%d %H:%M:%S')
            
        info = {
            'user_id': user_id,
            'license_code': license_code,
            'is_vip': True,
            'is_permanent': is_permanent,
            'register_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'expire_time': expire_time
        }
        
        with open(self.license_file, 'w') as f:
            json.dump(info, f)
            
    def is_registered(self):
        """检查是否已注册"""
        if not os.path.exists(self.license_file):
            return False, False, False, None  # 返回(未注册, 不是VIP, 不是永久, 到期时间None)
            
        try:
            with open(self.license_file, 'r') as f:
                info = json.load(f)
                
            # 验证保存的注册信息是否与当前用户ID匹配
            current_user_id = self.get_user_id()
            is_permanent = info.get('is_permanent', False)
            expected_code = self.generate_license_code(current_user_id, is_permanent=is_permanent)
            
            if (info.get('user_id') == current_user_id and 
                info.get('license_code') == expected_code):
                
                # 如果不是永久会员，检查是否已过期
                if not is_permanent and info.get('expire_time') != "永久有效":
                    expire_time = datetime.strptime(info['expire_time'], '%Y-%m-%d %H:%M:%S')
                    if expire_time < datetime.now():
                        return False, False, False, None  # 会员已过期
                
                return True, True, is_permanent, info.get('expire_time')  # 返回(已注册, 是VIP, 是否永久, 到期时间)
            return False, False, False, None
        except:
            return False, False, False, None 