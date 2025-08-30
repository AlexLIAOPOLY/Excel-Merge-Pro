#!/usr/bin/env python3
"""
表格分组功能全面测试
测试所有可能的场景，确保分组逻辑的健壮性
"""

import sys
import os
import time
import threading
import tempfile
import pandas as pd
import gc
import uuid
import random
import string
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from unittest.mock import patch
import multiprocessing

# 添加项目路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app_v2 import app
from models.database import db, TableGroup, TableData, TableSchema
from models.excel_processor import UniversalExcelProcessor

class GroupingTester:
    def __init__(self):
        self.app = app
        self.test_results = []
        self.test_count = 0
        self.success_count = 0
        self.failure_count = 0
        
    def log_test(self, test_name, status, message=""):
        """记录测试结果"""
        self.test_count += 1
        if status == "PASS":
            self.success_count += 1
            print(f"✅ [{self.test_count}] {test_name}: PASS {message}")
        else:
            self.failure_count += 1
            print(f"❌ [{self.test_count}] {test_name}: FAIL {message}")
        
        self.test_results.append({
            'test_name': test_name,
            'status': status,
            'message': message
        })
    
    def assert_test(self, condition, test_name, success_msg="", failure_msg=""):
        """断言测试"""
        if condition:
            self.log_test(test_name, "PASS", success_msg)
            return True
        else:
            self.log_test(test_name, "FAIL", failure_msg)
            return False
    
    def create_test_excel(self, filename, columns, data, temp_dir):
        """创建测试用的Excel文件"""
        df = pd.DataFrame(data, columns=columns)
        filepath = os.path.join(temp_dir, filename)
        df.to_excel(filepath, index=False, engine='openpyxl')
        return filepath
    
    def cleanup_database(self):
        """清理数据库"""
        with self.app.app_context():
            try:
                TableData.query.delete()
                TableSchema.query.delete()
                TableGroup.query.delete()
                db.session.commit()
                print("🧹 数据库已清理")
            except Exception as e:
                db.session.rollback()
                print(f"❌ 数据库清理失败: {e}")
    
    def test_all_scenarios(self):
        """执行所有测试场景"""
        with self.app.app_context():
            print("\n🧪 开始执行全面测试...")
            
            # 场景1：基础功能测试
            self.test_basic_functionality()
            
            # 场景2：边界条件测试
            self.test_edge_cases()
            
            # 场景3：并发测试
            self.test_concurrent_operations()
            
            # 场景4：性能测试
            self.test_performance()
            
            # 场景5：国际化测试
            self.test_internationalization()
            
            # 场景6：错误恢复测试
            self.test_error_recovery()
            
            # 场景7：真实文件处理测试
            self.test_real_file_processing()
    
    def test_basic_functionality(self):
        """基础功能测试"""
        print("\n🔬 基础功能测试")
        
        # 1. 完全相同的列结构
        columns = ['项目序号', '项目名称', '申请部门', '预算金额']
        
        # 首次创建
        group1 = UniversalExcelProcessor.create_table_group(columns, "测试文件1.xlsx")
        self.assert_test(
            group1 is not None,
            "创建第一个分组",
            f"成功创建: {group1.group_name}"
        )
        
        # 再次查找相同结构
        group2, similarity = UniversalExcelProcessor.find_matching_table_group(columns)
        self.assert_test(
            group2 is not None and group2.id == group1.id and similarity == 1.0,
            "查找相同结构分组",
            f"正确找到: {group2.group_name}, 相似度: {similarity}"
        )
        
        # 2. 空格差异测试
        space_columns = [' 项目序号', '项目名称 ', ' 申请部门 ', '预算金额']
        group3, sim3 = UniversalExcelProcessor.find_matching_table_group(space_columns)
        self.assert_test(
            group3 is not None and group3.id == group1.id,
            "空格差异处理",
            f"正确处理空格差异: {sim3:.3f}"
        )
        
        # 3. 大小写差异测试（英文）
        english_columns = ['Name', 'Age', 'Department', 'Salary']
        case_columns = ['name', 'AGE', 'Department', 'SALARY']
        
        group_eng = UniversalExcelProcessor.create_table_group(english_columns, "english.xlsx")
        group_case, sim_case = UniversalExcelProcessor.find_matching_table_group(case_columns)
        
        self.assert_test(
            group_case is not None and group_case.id == group_eng.id,
            "大小写差异处理",
            f"正确处理大小写: {sim_case:.3f}"
        )
        
        # 4. 列顺序不同测试
        original_cols = ['项目编号', '项目名称', '申请部门', '预算金额']
        reordered_cols = ['预算金额', '申请部门', '项目名称', '项目编号']  # 完全颠倒的顺序
        
        group_orig = UniversalExcelProcessor.create_table_group(original_cols, "original.xlsx")
        group_reord, sim_reord = UniversalExcelProcessor.find_matching_table_group(reordered_cols)
        
        self.assert_test(
            group_reord is not None and group_reord.id == group_orig.id and sim_reord == 1.0,
            "列顺序不同处理",
            f"正确处理顺序差异: {sim_reord:.3f}"
        )
        
        # 4.1 部分顺序不同测试
        partial_reorder = ['项目名称', '项目编号', '预算金额', '申请部门']  # 部分调换
        group_partial, sim_partial = UniversalExcelProcessor.find_matching_table_group(partial_reorder)
        
        self.assert_test(
            group_partial is not None and group_partial.id == group_orig.id and sim_partial == 1.0,
            "部分列顺序不同处理",
            f"正确处理部分顺序差异: {sim_partial:.3f}"
        )
        
        # 4.2 复杂空格和顺序混合测试
        complex_space_reorder = [' 预算金额 ', '申请部门  ', '  项目名称', '项目编号 ']
        group_complex, sim_complex = UniversalExcelProcessor.find_matching_table_group(complex_space_reorder)
        
        self.assert_test(
            group_complex is not None and group_complex.id == group_orig.id and sim_complex == 1.0,
            "复杂空格+顺序处理",
            f"正确处理复杂情况: {sim_complex:.3f}"
        )
        
        # 5. 完全不同的结构
        different_cols = ['学号', '学生姓名', '班级', '成绩']
        group_diff, sim_diff = UniversalExcelProcessor.find_matching_table_group(different_cols)
        
        self.assert_test(
            group_diff is None,
            "完全不同结构拒绝",
            f"正确拒绝不同结构: {sim_diff:.3f}"
        )
    
    def test_edge_cases(self):
        """边界条件测试"""
        print("\n🔬 边界条件测试")
        
        # 1. 空列表
        try:
            empty_result = UniversalExcelProcessor.find_matching_table_group([])
            self.assert_test(
                empty_result[0] is None,
                "空列表处理",
                "正确处理空列表"
            )
        except Exception as e:
            self.log_test("空列表处理", "FAIL", f"异常: {str(e)}")
        
        # 2. 单列
        try:
            single_result = UniversalExcelProcessor.find_matching_table_group(['单列'])
            self.assert_test(True, "单列处理", "正确处理单列")
        except Exception as e:
            self.log_test("单列处理", "FAIL", f"异常: {str(e)}")
        
        # 3. 超长列名
        try:
            long_cols = [f"超长列名{'x' * 500}_{i}" for i in range(5)]
            long_result = UniversalExcelProcessor.find_matching_table_group(long_cols)
            self.assert_test(True, "超长列名处理", "正确处理超长列名")
        except Exception as e:
            self.log_test("超长列名处理", "FAIL", f"异常: {str(e)}")
        
        # 4. 特殊字符
        try:
            special_cols = ['列!@#$%', '列^&*()', '列\\n\\t', '列"\'`']
            special_result = UniversalExcelProcessor.find_matching_table_group(special_cols)
            self.assert_test(True, "特殊字符处理", "正确处理特殊字符")
        except Exception as e:
            self.log_test("特殊字符处理", "FAIL", f"异常: {str(e)}")
        
        # 5. Unicode字符和混合编码
        try:
            unicode_cols = ['中文', '日本語', '한국어', '🔥emoji']
            unicode_result = UniversalExcelProcessor.find_matching_table_group(unicode_cols)
            self.assert_test(True, "Unicode字符处理", "正确处理Unicode")
        except Exception as e:
            self.log_test("Unicode字符处理", "FAIL", f"异常: {str(e)}")
        
        # 5.1 特殊空格字符测试
        try:
            # 测试各种类型的空格字符
            space_variants_1 = ['项目\u0020编号', '项目\u00A0名称', '申请\u2000部门', '预算\u3000金额']  # 普通空格、不换行空格、em空格、中文空格
            space_variants_2 = ['项目 编号', '项目 名称', '申请 部门', '预算 金额']  # 普通空格
            
            group_space1 = UniversalExcelProcessor.create_table_group(space_variants_1, "space1.xlsx")
            group_space2, sim_space = UniversalExcelProcessor.find_matching_table_group(space_variants_2)
            
            self.assert_test(
                group_space2 is not None and group_space2.id == group_space1.id,
                "特殊空格字符处理",
                f"正确处理各种空格: {sim_space:.3f}"
            )
        except Exception as e:
            self.log_test("特殊空格字符处理", "FAIL", f"异常: {str(e)}")
        
        # 5.2 中英文标点符号混合
        try:
            punctuation_cols1 = ['项目（编号）', '项目"名称"', '申请：部门', '预算-金额']
            punctuation_cols2 = ['项目(编号)', '项目"名称"', '申请:部门', '预算-金额']
            
            group_punct1 = UniversalExcelProcessor.create_table_group(punctuation_cols1, "punct1.xlsx")
            group_punct2, sim_punct = UniversalExcelProcessor.find_matching_table_group(punctuation_cols2)
            
            self.assert_test(
                group_punct2 is not None and group_punct2.id == group_punct1.id,
                "中英文标点符号处理",
                f"正确处理标点差异: {sim_punct:.3f}"
            )
        except Exception as e:
            self.log_test("中英文标点符号处理", "FAIL", f"异常: {str(e)}")
        
        # 6. 重复列名
        try:
            dup_cols = ['重复', '重复', '不重复', '重复']
            dup_result = UniversalExcelProcessor.find_matching_table_group(dup_cols)
            self.assert_test(True, "重复列名处理", "正确处理重复列名")
        except Exception as e:
            self.log_test("重复列名处理", "FAIL", f"异常: {str(e)}")
        
        # 7. 超多列
        try:
            many_cols = [f"列_{i}" for i in range(200)]
            many_result = UniversalExcelProcessor.find_matching_table_group(many_cols)
            self.assert_test(True, "超多列处理", f"正确处理{len(many_cols)}列")
        except Exception as e:
            self.log_test("超多列处理", "FAIL", f"异常: {str(e)}")
        
        # 8. 复杂混合测试（空格+顺序+大小写+特殊字符）
        try:
            complex_original = ['Project_ID', 'Project Name', 'Department', 'Budget Amount']
            complex_variant = [' budget amount ', 'DEPARTMENT  ', '  project name', 'project_id ']
            
            group_comp_orig = UniversalExcelProcessor.create_table_group(complex_original, "complex_orig.xlsx")
            group_comp_var, sim_comp = UniversalExcelProcessor.find_matching_table_group(complex_variant)
            
            self.assert_test(
                group_comp_var is not None and group_comp_var.id == group_comp_orig.id,
                "复杂混合情况处理",
                f"正确处理复杂混合: {sim_comp:.3f}"
            )
        except Exception as e:
            self.log_test("复杂混合情况处理", "FAIL", f"异常: {str(e)}")
    
    def test_concurrent_operations(self):
        """并发操作测试"""
        print("\n🔬 并发操作测试")
        
        # 并发创建相同结构分组
        columns = ['并发测试', '列1', '列2']
        results = []
        
        def create_group_thread(thread_id):
            with self.app.app_context():
                try:
                    group = UniversalExcelProcessor.create_table_group(columns, f"并发{thread_id}.xlsx")
                    return {'success': True, 'group_id': group.id if group else None}
                except Exception as e:
                    return {'success': False, 'error': str(e)}
        
        # 启动10个并发线程
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(create_group_thread, i) for i in range(10)]
            for future in as_completed(futures):
                results.append(future.result())
        
        successful_results = [r for r in results if r['success']]
        unique_group_ids = set(r['group_id'] for r in successful_results if r['group_id'])
        
        self.assert_test(
            len(successful_results) >= 8,
            "并发创建成功率",
            f"成功: {len(successful_results)}/10"
        )
        
        self.assert_test(
            len(unique_group_ids) == 1,
            "并发创建一致性",
            f"统一分组ID: {unique_group_ids}"
        )
    
    def test_performance(self):
        """性能测试"""
        print("\n🔬 性能测试")
        
        # 1. 大量分组创建性能
        start_time = time.time()
        created_groups = []
        
        for i in range(100):
            test_cols = [f"性能测试_{i}", f"列1_{i}", f"列2_{i}"]
            group = UniversalExcelProcessor.create_table_group(test_cols, f"性能_{i}.xlsx")
            created_groups.append(group)
        
        creation_time = time.time() - start_time
        
        self.assert_test(
            creation_time < 30.0,  # 100个分组应在30秒内创建完成
            "大量分组创建性能",
            f"创建{len(created_groups)}个分组耗时: {creation_time:.2f}秒"
        )
        
        # 2. 查找性能测试
        query_start = time.time()
        test_cols = ["性能查询测试", "列1", "列2"]
        group, similarity = UniversalExcelProcessor.find_matching_table_group(test_cols)
        query_time = time.time() - query_start
        
        self.assert_test(
            query_time < 1.0,
            "大量分组查找性能",
            f"查找耗时: {query_time:.3f}秒"
        )
        
        # 3. 重复分组清理性能
        # 手动创建重复分组
        test_fingerprint = "performance_test_fingerprint"
        for i in range(10):
            dup_group = TableGroup(
                group_name=f"性能重复_{i}",
                description=f"性能测试重复 {i}",
                schema_fingerprint=test_fingerprint,
                column_count=3
            )
            db.session.add(dup_group)
        db.session.commit()
        
        cleanup_start = time.time()
        cleaned = UniversalExcelProcessor.cleanup_duplicate_groups()
        cleanup_time = time.time() - cleanup_start
        
        self.assert_test(
            cleanup_time < 5.0 and cleaned == 9,
            "重复分组清理性能",
            f"清理{cleaned}个重复分组耗时: {cleanup_time:.3f}秒"
        )
    
    def test_internationalization(self):
        """国际化测试"""
        print("\n🔬 国际化测试")
        
        # 多语言测试用例
        test_cases = [
            (['中文列', '项目', '部门'], "中文"),
            (['English', 'Project', 'Dept'], "英文"),
            (['日本語', 'プロジェクト', '部門'], "日文"),
            (['한국어', '프로젝트', '부서'], "韩文"),
            (['Русский', 'проект', 'отдел'], "俄文"),
            (['عربي', 'مشروع', 'قسم'], "阿拉伯文"),
        ]
        
        for columns, lang in test_cases:
            try:
                group = UniversalExcelProcessor.create_table_group(columns, f"{lang}.xlsx")
                found_group, sim = UniversalExcelProcessor.find_matching_table_group(columns)
                
                self.assert_test(
                    group is not None and found_group is not None and found_group.id == group.id,
                    f"{lang}处理",
                    f"成功处理{lang}列名"
                )
            except Exception as e:
                self.log_test(f"{lang}处理", "FAIL", f"异常: {str(e)}")
        
        # 混合语言测试
        mixed_cols = ['ID', '中文名', 'English_Name', '日本語名前']
        try:
            mixed_group = UniversalExcelProcessor.create_table_group(mixed_cols, "mixed.xlsx")
            self.assert_test(
                mixed_group is not None,
                "混合语言处理",
                "成功处理混合语言列名"
            )
        except Exception as e:
            self.log_test("混合语言处理", "FAIL", f"异常: {str(e)}")
    
    def test_error_recovery(self):
        """错误恢复测试"""
        print("\n🔬 错误恢复测试")
        
        # 1. 数据库错误恢复
        try:
            with patch.object(TableGroup, 'query', side_effect=Exception("数据库连接失败")):
                try:
                    group, sim = UniversalExcelProcessor.find_matching_table_group(['错误测试'])
                    self.log_test("数据库错误处理", "FAIL", "应该抛出异常")
                except Exception as e:
                    self.assert_test(
                        "数据库连接失败" in str(e),
                        "数据库错误处理",
                        "正确捕获数据库错误"
                    )
        except Exception as e:
            self.log_test("数据库错误测试", "FAIL", f"测试异常: {str(e)}")
        
        # 2. 内存错误恢复
        try:
            with patch.object(UniversalExcelProcessor, 'generate_schema_fingerprint', side_effect=MemoryError("内存不足")):
                try:
                    group = UniversalExcelProcessor.create_table_group(['内存测试'], "memory.xlsx")
                    self.log_test("内存错误处理", "FAIL", "应该抛出内存错误")
                except MemoryError:
                    self.assert_test(True, "内存错误处理", "正确捕获内存错误")
        except Exception as e:
            self.log_test("内存错误测试", "FAIL", f"测试异常: {str(e)}")
        
        # 3. 事务回滚测试
        initial_count = TableGroup.query.count()
        try:
            with patch.object(db.session, 'commit', side_effect=Exception("提交失败")):
                try:
                    group = UniversalExcelProcessor.create_table_group(['事务测试'], "transaction.xlsx")
                    self.log_test("事务错误处理", "FAIL", "应该抛出异常")
                except Exception:
                    rollback_count = TableGroup.query.count()
                    self.assert_test(
                        rollback_count == initial_count,
                        "事务回滚正确性",
                        f"正确回滚: {initial_count} -> {rollback_count}"
                    )
        except Exception as e:
            self.log_test("事务回滚测试", "FAIL", f"测试异常: {str(e)}")
    
    def test_real_file_processing(self):
        """真实文件处理测试"""
        print("\n🔬 真实文件处理测试")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            # 1. 标准Excel文件
            standard_data = [
                ['项目1', '部门A', 100000],
                ['项目2', '部门B', 200000],
                ['项目3', '部门C', 150000']
            ]
            standard_cols = ['项目名称', '申请部门', '预算金额']
            standard_file = self.create_test_excel('标准.xlsx', standard_cols, standard_data, temp_dir)
            
            try:
                success1, msg1, count1, group_id1 = UniversalExcelProcessor.process_excel_file_with_grouping(
                    standard_file, '标准.xlsx'
                )
                
                self.assert_test(
                    success1 and count1 == 3,
                    "标准Excel处理",
                    f"成功处理{count1}条记录"
                )
            except Exception as e:
                self.log_test("标准Excel处理", "FAIL", f"异常: {str(e)}")
            
            # 2. 相似结构文件
            similar_data = [
                ['项目4', '部门D', 300000],
                ['项目5', '部门E', 250000]
            ]
            similar_cols = ['项目名称', '申请部门', '投资金额']  # 最后一列不同
            similar_file = self.create_test_excel('相似.xlsx', similar_cols, similar_data, temp_dir)
            
            try:
                success2, msg2, count2, group_id2 = UniversalExcelProcessor.process_excel_file_with_grouping(
                    similar_file, '相似.xlsx'
                )
                
                self.assert_test(
                    success2 and count2 == 2,
                    "相似Excel处理",
                    f"成功处理{count2}条记录"
                )
                
                # 验证高相似度是否合并
                if group_id1 and group_id2:
                    self.assert_test(
                        group_id1 == group_id2,
                        "相似文件自动合并",
                        f"正确合并到分组{group_id1}"
                    )
                    
            except Exception as e:
                self.log_test("相似Excel处理", "FAIL", f"异常: {str(e)}")
            
            # 3. 无效文件处理
            invalid_file = os.path.join(temp_dir, "无效.xlsx")
            with open(invalid_file, 'w') as f:
                f.write("这不是Excel文件")
            
            try:
                success3, msg3, count3, group_id3 = UniversalExcelProcessor.process_excel_file_with_grouping(
                    invalid_file, "无效.xlsx"
                )
                
                self.assert_test(
                    not success3,
                    "无效文件处理",
                    f"正确识别无效文件: {msg3}"
                )
            except Exception as e:
                self.log_test("无效文件处理", "FAIL", f"异常: {str(e)}")
    
    def run_all_tests(self):
        """运行所有测试"""
        print("🚀 Excel表格分组功能全面测试")
        print("=" * 60)
        
        start_time = time.time()
        
        # 清理数据库
        self.cleanup_database()
        
        try:
            # 执行所有测试
            self.test_all_scenarios()
            
        except Exception as e:
            print(f"❌ 测试执行异常: {str(e)}")
            import traceback
            traceback.print_exc()
        
        end_time = time.time()
        
        # 输出结果摘要
        print("\n" + "=" * 60)
        print("📊 测试结果摘要")
        print("=" * 60)
        print(f"⏱️  执行时间: {end_time - start_time:.2f} 秒")
        print(f"📋 总测试数: {self.test_count}")
        print(f"✅ 成功: {self.success_count}")
        print(f"❌ 失败: {self.failure_count}")
        print(f"📈 成功率: {(self.success_count/self.test_count*100):.1f}%" if self.test_count > 0 else "📈 成功率: 0%")
        
        if self.failure_count > 0:
            print("\n❌ 失败的测试:")
            for result in self.test_results:
                if result['status'] == 'FAIL':
                    print(f"   - {result['test_name']}: {result['message']}")
        
        print("\n" + "=" * 60)
        if self.failure_count == 0:
            print("🎉 所有测试通过！分组功能运行完美！")
        else:
            print("⚠️  存在失败测试，需要修复！")
        print("=" * 60)
        
        return self.failure_count == 0

def main():
    """主函数"""
    print("🔬 Excel表格分组功能综合测试工具")
    print("🎯 验证分组逻辑的正确性和健壮性")
    
    tester = GroupingTester()
    success = tester.run_all_tests()
    
    # 清理
    print("\n🧹 清理测试数据...")
    tester.cleanup_database()
    
    return 0 if success else 1

if __name__ == "__main__":
    exit(main())
