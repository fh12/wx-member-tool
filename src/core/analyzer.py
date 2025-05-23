from typing import Dict, List, Set

class GroupAnalyzer:
    def __init__(self):
        self.groups_data: Dict[str, List[str]] = {}
        self.common_members: Dict[str, Set[str]] = {}

    def analyze_common_members(self, groups: Dict[str, List[str]], min_groups: int = 2):
        """
        分析多个群的共同成员
        
        Args:
            groups: 群组数据，格式为 {群名: [成员列表]}
            min_groups: 最少出现在几个群中
        """
        self.groups_data = groups
        member_groups: Dict[str, Set[str]] = {}

        # 统计每个成员在哪些群中
        for group_name, members in groups.items():
            for member in members:
                if member not in member_groups:
                    member_groups[member] = set()
                member_groups[member].add(group_name)

        # 筛选出现在多个群的成员
        self.common_members = {
            member: groups 
            for member, groups in member_groups.items() 
            if len(groups) >= min_groups
        }

        return self.common_members

    def export_results(self, filepath: str):
        """
        导出分析结果到文件
        
        Args:
            filepath: 导出文件路径
        """
        if not self.common_members:
            return False

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("成员,所在群组\n")
                for member, groups in self.common_members.items():
                    f.write(f"{member},{','.join(groups)}\n")
            return True
        except Exception as e:
            print(f"导出失败: {e}")
            return False 