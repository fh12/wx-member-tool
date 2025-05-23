import os
import json
from datetime import datetime

def save_json(data, filepath):
    """保存数据到JSON文件"""
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存JSON失败: {e}")
        return False

def load_json(filepath):
    """从JSON文件加载数据"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"加载JSON失败: {e}")
        return None

def get_timestamp_filename(prefix, ext):
    """生成带时间戳的文件名"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f"{prefix}_{timestamp}.{ext}" 