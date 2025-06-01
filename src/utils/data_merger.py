#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据表合并脚本
用于合并多个文件夹内的Excel(.xlsx)和CSV(.csv)数据表
"""

import os
import pandas as pd
import glob
from pathlib import Path
import sys
from datetime import datetime


class DataMerger:
    """数据合并处理类"""
    
    def __init__(self):
        self.supported_formats = ['.xlsx', '.csv']
        self.merged_data = pd.DataFrame()
        self.file_count = 0
        self.error_files = []
    
    def scan_folders(self, folder_paths):
        """
        扫描指定文件夹，查找所有支持的数据文件
        
        Args:
            folder_paths (list): 文件夹路径列表
            
        Returns:
            list: 找到的所有数据文件路径列表
        """
        all_files = []
        
        for folder_path in folder_paths:
            if not os.path.exists(folder_path):
                print(f"警告: 文件夹 '{folder_path}' 不存在，跳过...")
                continue
                
            print(f"正在扫描文件夹: {folder_path}")
            
            # 递归查找所有支持格式的文件
            for ext in self.supported_formats:
                pattern = os.path.join(folder_path, '**', f'*{ext}')
                files = glob.glob(pattern, recursive=True)
                all_files.extend(files)
                print(f"  找到 {len(files)} 个 {ext} 文件")
        
        print(f"总共找到 {len(all_files)} 个数据文件")
        return all_files
    
    def read_single_file(self, file_path):
        """
        读取单个数据文件
        
        Args:
            file_path (str): 文件路径
            
        Returns:
            pandas.DataFrame: 读取的数据，如果失败返回None
        """
        try:
            file_ext = Path(file_path).suffix.lower()
            
            if file_ext == '.xlsx':
                # 读取Excel文件
                df = pd.read_excel(file_path)
            elif file_ext == '.csv':
                # 读取CSV文件，自动检测编码
                try:
                    df = pd.read_csv(file_path, encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        df = pd.read_csv(file_path, encoding='gbk')
                    except UnicodeDecodeError:
                        df = pd.read_csv(file_path, encoding='latin-1')
            else:
                print(f"不支持的文件格式: {file_path}")
                return None
            
            # 添加文件来源信息
            df['文件来源'] = os.path.basename(file_path)
            df['文件路径'] = file_path
            
            print(f"成功读取: {file_path} (行数: {len(df)}, 列数: {len(df.columns)})")
            return df
            
        except Exception as e:
            print(f"读取文件失败: {file_path}")
            print(f"错误信息: {str(e)}")
            self.error_files.append((file_path, str(e)))
            return None
    
    def merge_data(self, file_paths):
        """
        合并多个数据文件
        
        Args:
            file_paths (list): 文件路径列表
        """
        print("\n开始合并数据...")
        all_dataframes = []
        
        for file_path in file_paths:
            df = self.read_single_file(file_path)
            if df is not None and not df.empty:
                all_dataframes.append(df)
                self.file_count += 1
        
        if all_dataframes:
            # 合并所有数据框
            self.merged_data = pd.concat(all_dataframes, ignore_index=True, sort=False)
            print(f"\n数据合并完成!")
            print(f"合并了 {self.file_count} 个文件")
            print(f"总行数: {len(self.merged_data)}")
            print(f"总列数: {len(self.merged_data.columns)}")
        else:
            print("没有有效的数据可以合并!")
    
    def save_merged_data(self, output_path):
        """
        保存合并后的数据
        
        Args:
            output_path (str): 输出文件路径
        """
        if self.merged_data.empty:
            print("没有数据可以保存!")
            return False
        
        try:
            # 根据输出文件扩展名决定保存格式
            output_ext = Path(output_path).suffix.lower()
            
            if output_ext == '.xlsx':
                self.merged_data.to_excel(output_path, index=False)
            elif output_ext == '.csv':
                self.merged_data.to_csv(output_path, index=False, encoding='utf-8-sig')
            else:
                # 默认保存为Excel格式
                output_path = output_path + '.xlsx'
                self.merged_data.to_excel(output_path, index=False)
            
            print(f"\n数据已成功保存到: {output_path}")
            return True
            
        except Exception as e:
            print(f"保存文件失败: {str(e)}")
            return False
    
    def print_summary(self):
        """打印处理摘要"""
        print("\n" + "="*50)
        print("处理摘要")
        print("="*50)
        print(f"成功处理文件数: {self.file_count}")
        print(f"合并后总行数: {len(self.merged_data)}")
        print(f"合并后总列数: {len(self.merged_data.columns)}")
        
        if self.error_files:
            print(f"处理失败文件数: {len(self.error_files)}")
            print("失败文件列表:")
            for file_path, error in self.error_files:
                print(f"  - {file_path}: {error}")
        
        if not self.merged_data.empty:
            print("\n数据概览:")
            print(self.merged_data.head())


def main():
    """主函数"""
    print("Excel/CSV数据表合并工具")
    print("="*30)
    
    # 1. 输入文件夹路径配置
    print("\n请配置输入文件夹路径:")
    print("支持以下两种方式:")
    print("1. 交互式输入")
    print("2. 直接在代码中配置")
    
    # 方式1: 交互式输入（可选）
    use_interactive = input("是否使用交互式输入？(y/n): ").lower().strip() == 'y'
    
    if use_interactive:
        folder_paths = []
        print("\n请输入要合并的文件夹路径（每行一个，输入空行结束）:")
        while True:
            folder_path = input("文件夹路径: ").strip()
            if not folder_path:
                break
            folder_paths.append(folder_path)
    else:
        # 方式2: 直接配置（推荐）
        folder_paths = [
            # 在这里配置您的文件夹路径
            "./data",           # 示例: 当前目录下的data文件夹
            "./input",          # 示例: 当前目录下的input文件夹
            # 可以添加更多文件夹路径
        ]
        print(f"使用预配置的文件夹路径: {folder_paths}")
    
    if not folder_paths:
        print("错误: 未指定任何文件夹路径!")
        return
    
    # 2. 输出文件配置
    output_file = input("请输入输出文件路径（如: merged_data.xlsx）: ").strip()
    if not output_file:
        # 使用默认输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"merged_data_{timestamp}.xlsx"
        print(f"使用默认输出文件名: {output_file}")
    
    # 3. 数据处理合并模块
    print("\n开始数据处理和合并...")
    merger = DataMerger()
    
    # 扫描文件夹
    file_paths = merger.scan_folders(folder_paths)
    
    if not file_paths:
        print("未找到任何支持的数据文件!")
        return
    
    # 合并数据
    merger.merge_data(file_paths)
    
    # 保存合并后的数据
    if merger.save_merged_data(output_file):
        merger.print_summary()
        print("\n数据合并完成!")
    else:
        print("\n数据保存失败!")


if __name__ == "__main__":
    main() 