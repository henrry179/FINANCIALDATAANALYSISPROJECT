#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简易多Sheet数据合并工具
一键将多个文件夹的数据文件合并为一个Excel文件，每个文件作为独立Sheet

使用方法：
1. 修改下面的 输入文件夹路径 列表
2. 修改 输出文件名（可选）
3. 运行脚本：python3 easy_multi_sheet_merger.py
"""

import os
import pandas as pd
import glob
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ====== 用户配置区域 - 请根据需要修改 ======

# 📁 输入文件夹路径 - 请修改为您的实际路径
INPUT_FOLDERS = [
    # 示例配置 - 请替换为您的实际路径
    "/Users/mac/Downloads/WorkFiles/financedatasets01-0601",
    "/Users/mac/Downloads/WorkFiles/financedatasets02-0601", 
    "/Users/mac/Downloads/WorkFiles/financedatasets03-0601",
    "/Users/mac/Downloads/WorkFiles/financedatasets04-0601",
    
    # 您可以添加更多文件夹路径：
    # "/您的路径/数据文件夹1",
    # "/您的路径/数据文件夹2",
    # "./相对路径/数据文件夹",
]

# 💾 输出文件名
OUTPUT_FILE = f"数据合并结果_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

# ⚙️ 高级配置
MAX_SHEETS = 200                # 最大Sheet数量（Excel限制为255个Sheet）
INCLUDE_SUMMARY = True          # 是否包含汇总信息Sheet
INCLUDE_FILE_INFO = True        # 是否在每个Sheet下方包含文件信息
SUPPORTED_FORMATS = ['.xlsx', '.csv', '.xls']  # 支持的文件格式

# ====== 配置区域结束 ======


def merge_data_to_sheets():
    """执行多Sheet数据合并"""
    
    print("🚀 简易多Sheet数据合并工具")
    print("=" * 50)
    
    # 验证配置
    print("📋 检查配置...")
    valid_folders = []
    for folder in INPUT_FOLDERS:
        if os.path.exists(folder):
            valid_folders.append(folder)
            print(f"   ✅ {folder}")
        else:
            print(f"   ❌ {folder} (文件夹不存在)")
    
    if not valid_folders:
        print("❌ 没有找到有效的输入文件夹！")
        print("请检查 INPUT_FOLDERS 配置")
        return
    
    print(f"\n📊 将处理 {len(valid_folders)} 个文件夹")
    print(f"💾 输出文件: {OUTPUT_FILE}")
    
    # 扫描文件
    print(f"\n🔍 扫描数据文件...")
    all_files = []
    
    for folder in valid_folders:
        folder_files = []
        for ext in SUPPORTED_FORMATS:
            pattern = os.path.join(folder, '**', f'*{ext}')
            files = glob.glob(pattern, recursive=True)
            folder_files.extend(files)
        
        all_files.extend(folder_files)
        print(f"   📁 {os.path.basename(folder)}: {len(folder_files)} 个文件")
    
    total_files = len(all_files)
    print(f"\n📈 总计找到 {total_files} 个数据文件")
    
    if total_files == 0:
        print("❌ 未找到任何数据文件！")
        return
    
    # 检查文件数量限制
    if total_files > MAX_SHEETS:
        print(f"⚠️  文件数量超过限制，将只处理前 {MAX_SHEETS} 个文件")
        all_files = all_files[:MAX_SHEETS]
        total_files = MAX_SHEETS
    
    # 开始合并
    print(f"\n📝 开始合并数据...")
    successful_count = 0
    error_count = 0
    total_rows = 0
    sheet_info = []
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        
        for i, file_path in enumerate(all_files, 1):
            try:
                filename = os.path.basename(file_path)
                file_ext = Path(file_path).suffix.lower()
                
                print(f"   [{i}/{total_files}] {filename}")
                
                # 读取文件
                if file_ext in ['.xlsx', '.xls']:
                    df = pd.read_excel(file_path)
                elif file_ext == '.csv':
                    # 自动检测CSV编码
                    try:
                        df = pd.read_csv(file_path, encoding='utf-8')
                    except UnicodeDecodeError:
                        try:
                            df = pd.read_csv(file_path, encoding='gbk')
                        except UnicodeDecodeError:
                            df = pd.read_csv(file_path, encoding='latin-1')
                else:
                    continue
                
                # 生成Sheet名称
                sheet_name = Path(filename).stem
                # 清理无效字符
                for char in ['\\', '/', '?', '*', '[', ']', ':']:
                    sheet_name = sheet_name.replace(char, '_')
                # 限制长度
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:28] + "..."
                
                # 确保Sheet名称唯一
                original_name = sheet_name
                counter = 1
                existing_names = [info['sheet_name'] for info in sheet_info]
                while sheet_name in existing_names:
                    sheet_name = f"{original_name[:25]}_{counter}"
                    counter += 1
                
                # 写入数据
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 添加文件信息（可选）
                if INCLUDE_FILE_INFO:
                    info_df = pd.DataFrame({
                        '文件信息': ['原始文件名', '文件路径', '处理时间', '数据行数', '数据列数'],
                        '值': [filename, file_path, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 
                              len(df), len(df.columns)]
                    })
                    info_df.to_excel(writer, sheet_name=sheet_name, index=False, 
                                   startrow=len(df)+2, startcol=0)
                
                # 记录信息
                sheet_info.append({
                    'sheet_name': sheet_name,
                    'filename': filename,
                    'rows': len(df),
                    'columns': len(df.columns),
                    'folder': os.path.basename(os.path.dirname(file_path))
                })
                
                successful_count += 1
                total_rows += len(df)
                
            except Exception as e:
                print(f"      ❌ 失败: {str(e)}")
                error_count += 1
        
        # 创建汇总Sheet
        if INCLUDE_SUMMARY and sheet_info:
            print(f"   📊 生成汇总信息...")
            
            summary_data = [
                ['数据合并汇总报告', ''],
                ['', ''],
                ['处理时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['成功处理文件', successful_count],
                ['失败文件', error_count],
                ['总Sheet数', len(sheet_info)],
                ['总数据行数', total_rows],
                ['输出文件', OUTPUT_FILE],
                ['', ''],
                ['文件清单', ''],
                ['Sheet名称', '原始文件名', '数据行数', '数据列数', '所属文件夹']
            ]
            
            for info in sheet_info:
                summary_data.append([
                    info['sheet_name'], info['filename'], 
                    info['rows'], info['columns'], info['folder']
                ])
            
            # 创建汇总DataFrame
            max_cols = max(len(row) for row in summary_data)
            summary_df = pd.DataFrame([row + [''] * (max_cols - len(row)) for row in summary_data])
            summary_df.to_excel(writer, sheet_name='📊汇总信息', index=False, header=False)
    
    # 最终报告
    print(f"\n🎉 合并完成！")
    print(f"=" * 50)
    print(f"✅ 成功处理: {successful_count} 个文件")
    if error_count > 0:
        print(f"❌ 失败文件: {error_count} 个")
    print(f"📄 生成Sheet: {len(sheet_info)} 个")
    print(f"📈 总数据行: {total_rows:,} 行")
    print(f"💾 输出文件: {OUTPUT_FILE}")
    
    try:
        file_size = os.path.getsize(OUTPUT_FILE) / 1024 / 1024
        print(f"📂 文件大小: {file_size:.1f} MB")
    except:
        pass
    
    print(f"\n💡 使用方法:")
    print(f"   1. 用Excel打开 '{OUTPUT_FILE}'")
    print(f"   2. 查看底部Sheet标签，每个原始文件都是独立的Sheet")
    print(f"   3. 查看 '📊汇总信息' Sheet了解详细统计")


def main():
    """主函数"""
    print("📊 欢迎使用简易多Sheet数据合并工具")
    print("-" * 50)
    print("本工具将多个文件夹的数据文件合并为一个Excel文件")
    print("每个原始文件保存为独立的Sheet")
    print()
    
    # 显示当前配置
    print("📋 当前配置:")
    print(f"   输入文件夹: {len(INPUT_FOLDERS)} 个")
    print(f"   输出文件: {OUTPUT_FILE}")
    print(f"   支持格式: {', '.join(SUPPORTED_FORMATS)}")
    print(f"   最大Sheet数: {MAX_SHEETS}")
    print()
    
    # 确认执行
    try:
        response = input("是否开始合并数据？(回车确认，输入n取消): ").strip().lower()
        if response in ['n', 'no', '否']:
            print("操作已取消")
            return
    except KeyboardInterrupt:
        print("\n操作已取消")
        return
    
    # 执行合并
    merge_data_to_sheets()


if __name__ == "__main__":
    main() 