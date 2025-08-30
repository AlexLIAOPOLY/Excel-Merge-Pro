#!/usr/bin/env python3
"""
智能列识别功能演示
展示系统如何智能处理列顺序不同、空格差异、大小写差异等复杂情况
"""

import sys
import os

# 添加项目路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app_v2 import app
from models.database import db, TableGroup, TableData, TableSchema
from models.excel_processor import UniversalExcelProcessor

def demo_intelligent_recognition():
    """演示智能列识别功能"""
    with app.app_context():
        print("🎯 Excel表格智能列识别功能演示")
        print("=" * 60)
        print("📋 本演示将展示系统如何智能处理：")
        print("   🔸 列顺序不同的表格")
        print("   🔸 空格、制表符等空白字符差异")
        print("   🔸 大小写差异")
        print("   🔸 中英文混合")
        print("   🔸 特殊字符和标点符号")
        print("   🔸 复杂混合情况")
        print("=" * 60)
        
        # 清理数据库
        TableData.query.delete()
        TableSchema.query.delete()
        TableGroup.query.delete()
        db.session.commit()
        print("🧹 数据库已清理\n")
        
        # 演示案例
        test_cases = [
            {
                'name': '标准项目表格',
                'columns': ['项目编号', '项目名称', '申请部门', '预算金额', '开始日期'],
                'filename': '标准项目.xlsx'
            },
            {
                'name': '列顺序完全颠倒',
                'columns': ['开始日期', '预算金额', '申请部门', '项目名称', '项目编号'],
                'filename': '颠倒顺序项目.xlsx'
            },
            {
                'name': '列顺序部分调换',
                'columns': ['项目名称', '项目编号', '预算金额', '申请部门', '开始日期'],
                'filename': '部分调换项目.xlsx'
            },
            {
                'name': '各种空格差异',
                'columns': [' 项目编号 ', '项目  名称', '申请部门 ', ' 预算金额', '开始日期'],
                'filename': '空格差异项目.xlsx'
            },
            {
                'name': '中文全角空格',
                'columns': ['项目　编号', '项目名称', '申请部门', '预算　金额', '开始日期'],
                'filename': '全角空格项目.xlsx'
            },
            {
                'name': '制表符和换行符',
                'columns': ['项目\t编号', '项目\n名称', '申请部门', '预算金额', '开始日期'],
                'filename': '特殊空白项目.xlsx'
            },
            {
                'name': '英文大小写混合',
                'columns': ['project_id', 'PROJECT NAME', 'Department', 'BUDGET_amount', 'start_date'],
                'filename': '英文项目.xlsx'
            },
            {
                'name': '中英文标点符号差异',
                'columns': ['项目（编号）', '项目"名称"', '申请：部门', '预算-金额', '开始日期'],
                'filename': '标点符号项目.xlsx'
            },
            {
                'name': '复杂混合情况',
                'columns': [' BUDGET_amount ', '部门 Department', '  project NAME ', '项目_ID', 'START日期'],
                'filename': '复杂混合项目.xlsx'
            }
        ]
        
        print("🚀 开始演示智能列识别...\n")
        
        created_groups = []
        
        for i, case in enumerate(test_cases, 1):
            print(f"📁 [{i}] 处理: {case['name']}")
            print(f"   📋 列结构: {case['columns']}")
            print(f"   📄 文件名: {case['filename']}")
            
            # 首先尝试查找匹配的分组
            found_group, similarity = UniversalExcelProcessor.find_matching_table_group(case['columns'])
            
            if found_group:
                print(f"   ✅ 找到匹配分组: {found_group.group_name}")
                print(f"   📊 相似度: {similarity:.3f}")
                print(f"   🔗 分组ID: {found_group.id}")
            else:
                print(f"   ⚪ 未找到匹配分组，将创建新分组")
                # 创建分组
                new_group = UniversalExcelProcessor.create_table_group(case['columns'], case['filename'])
                if new_group:
                    created_groups.append(new_group)
                    print(f"   ✅ 创建新分组: {new_group.group_name}")
                    print(f"   🔗 分组ID: {new_group.id}")
                else:
                    print(f"   ❌ 创建分组失败")
            
            print()
        
        # 总结
        print("📊 演示结果总结")
        print("=" * 60)
        
        total_groups = TableGroup.query.count()
        print(f"📈 总处理表格数: {len(test_cases)}")
        print(f"📁 实际创建分组数: {total_groups}")
        print(f"🎯 智能合并成功率: {((len(test_cases) - total_groups) / len(test_cases) * 100):.1f}%")
        
        print(f"\n📋 数据库中的分组:")
        all_groups = TableGroup.query.all()
        for group in all_groups:
            print(f"   🔸 {group.group_name} (ID: {group.id})")
            print(f"      指纹: {group.schema_fingerprint}")
            print(f"      列数: {group.column_count}")
        
        print("\n🎉 演示完成！")
        
        if total_groups == 1:
            print("✨ 完美！所有表格都被智能识别为相同结构并合并到一个分组！")
        elif total_groups <= 3:
            print("👍 很好！系统成功将大部分相似表格合并！")
        else:
            print("⚠️  系统创建了较多分组，可能需要进一步优化识别算法。")
        
        print("\n🔧 支持的智能识别特性:")
        print("   ✅ 列顺序自动调整匹配")
        print("   ✅ 各种空白字符标准化")
        print("   ✅ 大小写自动统一")
        print("   ✅ 中英文混合处理")
        print("   ✅ 特殊字符和标点符号")
        print("   ✅ Unicode字符标准化")
        print("   ✅ 复杂混合情况处理")

def main():
    """主函数"""
    print("🔬 Excel表格智能列识别功能演示工具")
    print("🎯 展示系统强大的表格结构智能识别能力\n")
    
    demo_intelligent_recognition()
    
    return 0

if __name__ == "__main__":
    exit(main())
