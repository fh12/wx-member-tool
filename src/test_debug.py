from core.wechat import WeChatController

def main():
    # 创建控制器实例
    controller = WeChatController()
    
    # 启用调试模式
    controller.enable_debug_mode()
    
    # 开始获取群聊列表，这个过程会触发调试功能
    print("开始测试获取群聊列表...")
    groups = controller.get_group_list()
    
    # 打印获取到的群聊信息
    if groups:
        print("\n=== 最终获取到的群聊信息 ===")
        for group in groups:
            print(f"群名: {group['name']}, 成员数: {group['member_count']}")
    else:
        print("未获取到群聊信息")

if __name__ == "__main__":
    main() 