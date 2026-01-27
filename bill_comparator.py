#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中国电信账单差异对比工具

该工具用于对比不同月份的客户账单差异，支持Excel和CSV格式的数据文件。
"""

import os
import sys
import argparse
import pandas as pd
from datetime import datetime

class BillComparator:
    """账单差异对比类"""
    
    def __init__(self, file_path):
        """
        初始化账单对比器
        
        Args:
            file_path: 数据文件路径
        """
        self.file_path = file_path
        self.data = None
    
    def load_data(self):
        """
        加载数据文件
        
        Returns:
            bool: 是否加载成功
        """
        try:
            file_ext = os.path.splitext(self.file_path)[1].lower()
            
            if file_ext in ['.xlsx', '.xls']:
                self.data = pd.read_excel(self.file_path)
            elif file_ext == '.csv':
                self.data = pd.read_csv(self.file_path)
            else:
                print(f"不支持的文件格式: {file_ext}")
                return False
            
            # 检查必要的列是否存在
            required_columns = ['设备号码', '账务周期', '账单费用']
            for col in required_columns:
                if col not in self.data.columns:
                    print(f"数据文件缺少必要的列: {col}")
                    return False
            
            print(f"成功加载数据文件，共 {len(self.data)} 条记录")
            return True
        except Exception as e:
            print(f"加载数据失败: {str(e)}")
            return False
    
    def extract_month(self, billing_cycle):
        """
        从账务周期中提取月份
        
        Args:
            billing_cycle: 账务周期字符串，格式如 "[20240701]2024-07-01:2024-07-31"
        
        Returns:
            str: 月份字符串，格式如 "2024-07"
        """
        try:
            # 提取方括号中的日期部分
            date_part = billing_cycle.split(']')[0].strip('[')
            # 转换为年月格式
            return f"{date_part[:4]}-{date_part[4:6]}"
        except Exception:
            return ""
    
    def get_available_months(self):
        """
        获取数据中可用的月份列表
        
        Returns:
            list: 月份列表
        """
        if self.data is None:
            return []
        
        # 提取所有月份并去重
        months = self.data['账务周期'].apply(self.extract_month).unique()
        # 过滤空值并排序
        months = [month for month in months if month]
        months.sort()
        
        return months
    
    def filter_data_by_month(self, month):
        """
        根据月份过滤数据
        
        Args:
            month: 月份字符串，格式如 "2024-07"
        
        Returns:
            DataFrame: 过滤后的数据
        """
        if self.data is None:
            return pd.DataFrame()
        
        # 提取月份并过滤
        filtered_data = self.data[self.data['账务周期'].apply(self.extract_month) == month]
        return filtered_data
    
    def compare_months(self, month1, month2):
        """
        对比两个月份的账单差异
        
        Args:
            month1: 第一个月份
            month2: 第二个月份
        
        Returns:
            DataFrame: 差异对比结果
        """
        # 过滤两个月份的数据
        data1 = self.filter_data_by_month(month1)
        data2 = self.filter_data_by_month(month2)
        
        if data1.empty:
            print(f"月份 {month1} 没有数据")
            return pd.DataFrame()
        
        if data2.empty:
            print(f"月份 {month2} 没有数据")
            return pd.DataFrame()
        
        # 按照设备号码合并数据
        merged = pd.merge(
            data1[['设备号码', '账单费用']],
            data2[['设备号码', '账单费用']],
            on='设备号码',
            how='outer',
            suffixes=(f'_{month1}', f'_{month2}')
        )
        
        # 填充空值为0
        merged = merged.fillna(0)
        
        # 计算差异
        merged['差异金额'] = merged[f'账单费用_{month2}'] - merged[f'账单费用_{month1}']
        merged['差异百分比'] = ((merged[f'账单费用_{month2}'] - merged[f'账单费用_{month1}']) / 
                            merged[f'账单费用_{month1}'].replace(0, 1) * 100).round(2)
        
        return merged
    
    def save_result(self, result, output_path):
        """
        保存对比结果
        
        Args:
            result: 对比结果数据框
            output_path: 输出文件路径
        """
        try:
            file_ext = os.path.splitext(output_path)[1].lower()
            
            if file_ext in ['.xlsx', '.xls']:
                result.to_excel(output_path, index=False)
            elif file_ext == '.csv':
                result.to_csv(output_path, index=False, encoding='utf-8-sig')
            else:
                print(f"不支持的输出格式: {file_ext}")
                return False
            
            print(f"结果已保存到: {output_path}")
            return True
        except Exception as e:
            print(f"保存结果失败: {str(e)}")
            return False

def main():
    """
    主函数
    """
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='中国电信账单差异对比工具')
    parser.add_argument('file', help='数据文件路径')
    parser.add_argument('--month1', help='第一个对比月份 (格式: YYYY-MM)')
    parser.add_argument('--month2', help='第二个对比月份 (格式: YYYY-MM)')
    parser.add_argument('--output', help='输出结果文件路径')
    
    args = parser.parse_args()
    
    # 创建账单对比器
    comparator = BillComparator(args.file)
    
    # 加载数据
    if not comparator.load_data():
        return 1
    
    # 获取可用月份
    available_months = comparator.get_available_months()
    
    if not available_months:
        print("数据中没有可用的月份")
        return 1
    
    print("\n数据中可用的月份:")
    for i, month in enumerate(available_months, 1):
        print(f"{i}. {month}")
    
    # 如果没有指定月份，让用户选择
    month1 = args.month1
    month2 = args.month2
    
    if not month1 or not month2:
        print("\n请选择要对比的两个月份:")
        
        while True:
            try:
                idx1 = int(input(f"请输入第一个月份的编号 (1-{len(available_months)}): ")) - 1
                if 0 <= idx1 < len(available_months):
                    month1 = available_months[idx1]
                    break
                else:
                    print("输入编号无效，请重新输入")
            except ValueError:
                print("请输入有效的数字")
        
        while True:
            try:
                idx2 = int(input(f"请输入第二个月份的编号 (1-{len(available_months)}): ")) - 1
                if 0 <= idx2 < len(available_months):
                    month2 = available_months[idx2]
                    break
                else:
                    print("输入编号无效，请重新输入")
            except ValueError:
                print("请输入有效的数字")
    
    print(f"\n正在对比 {month1} 和 {month2} 的账单差异...")
    
    # 执行对比
    result = comparator.compare_months(month1, month2)
    
    if result.empty:
        print("对比结果为空")
        return 1
    
    # 显示结果
    print("\n对比结果:")
    print(result)
    
    # 计算汇总信息
    total_diff = result['差异金额'].sum()
    avg_diff = result['差异金额'].mean()
    max_diff_row = result.loc[result['差异金额'].abs().idxmax()]
    
    print("\n汇总信息:")
    print(f"总差异金额: {total_diff:.2f}")
    print(f"平均差异金额: {avg_diff:.2f}")
    print(f"差异最大的设备: {max_diff_row['设备号码']} (差异: {max_diff_row['差异金额']:.2f})")
    
    # 保存结果
    if args.output:
        comparator.save_result(result, args.output)
    else:
        # 默认保存路径
        default_output = f"bill_comparison_{month1}_vs_{month2}.xlsx"
        comparator.save_result(result, default_output)
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
