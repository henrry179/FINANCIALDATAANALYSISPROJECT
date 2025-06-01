#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速数据合并脚本
直接修改下面的配置，然后运行即可
"""

import os
import pandas as pd
import glob
from pathlib import Path
from datetime import datetime


# ====== 配置区域 - 请根据您的需求修改 ======

# 1. 输入文件夹路径列表（请修改为您的实际路径）
INPUT_FOLDERS = [
    "/Users/mac/Downloads/WorkFiles/financedatasets0601",     # 修正为绝对路径
    "./test_data/folder2",                # 测试数据文件夹2
    # ".",                                  # 当前目录（作为示例）
    # "/Users/mac/Documents/data1",        # 您的第一个数据文件夹
    # "/Users/mac/Documents/data2",        # 您的第二个数据文件夹
    # "./local_data",                      # 相对路径示例
]

# 2. 输出文件路径（请修改为您想要的输出位置和文件名）
OUTPUT_FILE = "测试合并结果.xlsx"      # 输出文件

# 3. 支持的文件格式
SUPPORTED_FORMATS = ['.xlsx', '.csv']

# 4. 自动运行模式（设为True则跳过确认提示）
AUTO_RUN = True

# ====== 配置区域结束 ======


def merge_excel_csv_files():
    """合并多个文件夹内的Excel和CSV文件"""
    
    print("快速数据合并工具")
    print("=" * 30)
    
    # 记录处理信息
    all_files = []
    processed_files = 0
    error_files = []
    
    # 1. 扫描所有指定文件夹
    print(f"\n正在扫描 {len(INPUT_FOLDERS)} 个文件夹...")
    
    for folder_path in INPUT_FOLDERS:
        if not os.path.exists(folder_path):
            print(f"警告: 文件夹 '{folder_path}' 不存在，跳过...")
            continue
        
        print(f"扫描文件夹: {folder_path}")
        
        # 递归查找所有支持格式的文件
        for ext in SUPPORTED_FORMATS:
            pattern = os.path.join(folder_path, '**', f'*{ext}')
            files = glob.glob(pattern, recursive=True)
            all_files.extend(files)
            print(f"  找到 {len(files)} 个 {ext} 文件")
    
    print(f"\n总共找到 {len(all_files)} 个数据文件")
    
    if not all_files:
        print("未找到任何支持的数据文件!")
        return
    
    # 2. 读取并合并所有文件
    print("\n开始读取和合并数据...")
    all_dataframes = []
    
    for file_path in all_files:
        try:
            print(f"正在处理: {os.path.basename(file_path)}")
            
            # 根据文件扩展名选择读取方法
            file_ext = Path(file_path).suffix.lower()
            
            if file_ext == '.xlsx':
                df = pd.read_excel(file_path)
            elif file_ext == '.csv':
                # 尝试不同编码读取CSV文件
                try:
                    df = pd.read_csv(file_path, encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        df = pd.read_csv(file_path, encoding='gbk')
                    except UnicodeDecodeError:
                        df = pd.read_csv(file_path, encoding='latin-1')
            
            # 添加文件来源信息
            df['文件来源'] = os.path.basename(file_path)
            df['文件夹路径'] = os.path.dirname(file_path)
            
            all_dataframes.append(df)
            processed_files += 1
            print(f"  成功读取 {len(df)} 行数据")
            
        except Exception as e:
            error_msg = f"读取失败: {str(e)}"
            print(f"  {error_msg}")
            error_files.append((file_path, error_msg))
    
    # 3. 合并所有数据
    if all_dataframes:
        print(f"\n正在合并 {len(all_dataframes)} 个文件的数据...")
        merged_data = pd.concat(all_dataframes, ignore_index=True, sort=False)
        
        # 4. 保存合并后的数据
        print(f"正在保存数据到: {OUTPUT_FILE}")
        
        try:
            output_ext = Path(OUTPUT_FILE).suffix.lower()
            
            if output_ext == '.xlsx':
                merged_data.to_excel(OUTPUT_FILE, index=False)
            elif output_ext == '.csv':
                merged_data.to_csv(OUTPUT_FILE, index=False, encoding='utf-8-sig')
            else:
                # 默认保存为Excel
                output_file_with_ext = OUTPUT_FILE + '.xlsx'
                merged_data.to_excel(output_file_with_ext, index=False)
                print(f"已自动添加.xlsx扩展名: {output_file_with_ext}")
            
            print("\n✅ 数据合并完成!")
            
        except Exception as e:
            print(f"❌ 保存文件失败: {str(e)}")
            return
    
    else:
        print("❌ 没有有效的数据可以合并!")
        return
    
    # 5. 显示处理摘要
    print("\n" + "=" * 50)
    print("处理摘要")
    print("=" * 50)
    print(f"成功处理文件数: {processed_files}")
    print(f"合并后总行数: {len(merged_data)}")
    print(f"合并后总列数: {len(merged_data.columns)}")
    print(f"输出文件: {OUTPUT_FILE}")
    
    if error_files:
        print(f"\n处理失败文件数: {len(error_files)}")
        for file_path, error in error_files:
            print(f"  - {os.path.basename(file_path)}: {error}")
    
    print(f"\n数据概览（前5行）:")
    print(merged_data.head())


if __name__ == "__main__":
    # 运行前检查配置
    print("当前配置:")
    print(f"输入文件夹: {INPUT_FOLDERS}")
    print(f"输出文件: {OUTPUT_FILE}")
    print(f"支持格式: {SUPPORTED_FORMATS}")
    
    # 等待用户确认或自动运行
    if AUTO_RUN:
        print("\n自动运行模式已启用，开始合并...")
        merge_excel_csv_files()
    else:
        confirm = input(f"\n是否使用当前配置开始合并? (y/n): ").lower().strip()
        
        if confirm == 'y':
            merge_excel_csv_files()
        else:
            print("请修改脚本顶部的配置后重新运行")
            print("主要需要修改:")
            print("1. INPUT_FOLDERS - 您的数据文件夹路径")
            print("2. OUTPUT_FILE - 输出文件名和路径")
            print("3. AUTO_RUN - 设为True可跳过确认提示") 