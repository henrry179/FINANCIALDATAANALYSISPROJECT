#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多Sheet数据合并脚本
将多个文件夹内的Excel和CSV数据表合并成一个Excel文件，
每个原始数据文件作为独立的Sheet保存
"""

import os
import pandas as pd
import glob
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ====== 配置区域 ======

# 输入文件夹路径
INPUT_FOLDERS = [
    "/Users/mac/Downloads/WorkFiles/financedatasets01-0601",  # 金融数据集1
    "/Users/mac/Downloads/WorkFiles/financedatasets02-0601",  # 金融数据集2
    "/Users/mac/Downloads/WorkFiles/financedatasets03-0601",  # 金融数据集3
    "/Users/mac/Downloads/WorkFiles/financedatasets04-0601",  # 金融数据集4
    # 可以继续添加更多文件夹
]

# 输出文件配置
OUTPUT_FILE = f"完整金融数据合并_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

# 支持的文件格式
SUPPORTED_FORMATS = ['.xlsx', '.csv', '.xls']

# 功能配置
MAX_SHEETS = 200        # 最大Sheet数量限制（增加到200）
MAX_SHEET_NAME_LEN = 31 # Excel Sheet名称最大长度
ENABLE_SUMMARY = True   # 是否生成汇总Sheet
AUTO_RUN = True

# ====== 配置区域结束 ======


class MultiSheetMerger:
    """多Sheet数据合并器"""
    
    def __init__(self):
        self.processed_files = 0
        self.error_files = []
        self.sheet_info = []
        self.total_rows = 0
        
    def clean_sheet_name(self, filename):
        """清理并生成有效的Sheet名称"""
        # 移除文件扩展名
        name = Path(filename).stem
        
        # 替换Excel中不允许的字符
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # 限制长度
        if len(name) > MAX_SHEET_NAME_LEN:
            name = name[:MAX_SHEET_NAME_LEN-3] + "..."
        
        return name
    
    def scan_and_collect_files(self):
        """扫描所有文件夹并收集数据文件"""
        print("🔍 扫描文件夹，收集数据文件...")
        print("-" * 50)
        
        all_files = []
        folder_stats = {}
        
        for folder_path in INPUT_FOLDERS:
            if not os.path.exists(folder_path):
                print(f"⚠️  文件夹不存在，跳过: {folder_path}")
                continue
            
            print(f"📁 扫描: {folder_path}")
            folder_files = []
            
            # 递归查找所有支持格式的文件
            for ext in SUPPORTED_FORMATS:
                pattern = os.path.join(folder_path, '**', f'*{ext}')
                files = glob.glob(pattern, recursive=True)
                folder_files.extend(files)
            
            all_files.extend(folder_files)
            folder_stats[folder_path] = len(folder_files)
            print(f"   📊 找到 {len(folder_files)} 个数据文件")
        
        print(f"\n📈 扫描结果汇总:")
        for folder, count in folder_stats.items():
            print(f"   {os.path.basename(folder)}: {count} 个文件")
        
        print(f"   总计: {len(all_files)} 个数据文件")
        
        # 检查是否超出限制
        if len(all_files) > MAX_SHEETS:
            print(f"⚠️  警告: 文件数量({len(all_files)})超过最大Sheet限制({MAX_SHEETS})")
            print(f"   将只处理前 {MAX_SHEETS} 个文件")
            all_files = all_files[:MAX_SHEETS]
        
        return all_files
    
    def read_single_file(self, file_path):
        """读取单个数据文件"""
        try:
            file_ext = Path(file_path).suffix.lower()
            filename = os.path.basename(file_path)
            
            print(f"📖 正在读取: {filename}")
            
            # 根据文件类型选择读取方法
            if file_ext in ['.xlsx', '.xls']:
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
            else:
                print(f"   ❌ 不支持的文件格式: {file_ext}")
                return None, None, None
            
            # 生成Sheet名称
            sheet_name = self.clean_sheet_name(filename)
            
            # 确保Sheet名称唯一
            original_name = sheet_name
            counter = 1
            existing_names = [info['sheet_name'] for info in self.sheet_info]
            while sheet_name in existing_names:
                if len(original_name) > MAX_SHEET_NAME_LEN - 4:
                    sheet_name = original_name[:MAX_SHEET_NAME_LEN-4] + f"_{counter}"
                else:
                    sheet_name = f"{original_name}_{counter}"
                counter += 1
            
            print(f"   ✅ 成功读取: {len(df)} 行, {len(df.columns)} 列 → Sheet: {sheet_name}")
            
            return df, sheet_name, filename
            
        except Exception as e:
            error_msg = f"读取失败: {str(e)}"
            print(f"   ❌ {error_msg}")
            self.error_files.append((file_path, error_msg))
            return None, None, None
    
    def merge_to_multiple_sheets(self, file_paths):
        """将多个文件合并到多个Sheet中"""
        print(f"\n📝 开始合并数据到多个Sheet...")
        print("-" * 50)
        
        # 创建ExcelWriter对象
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            
            # 处理每个文件
            for i, file_path in enumerate(file_paths, 1):
                print(f"[{i}/{len(file_paths)}] ", end="")
                
                df, sheet_name, filename = self.read_single_file(file_path)
                
                if df is not None and sheet_name is not None:
                    try:
                        # 添加元数据信息
                        df_with_meta = df.copy()
                        
                        # 在数据末尾添加元信息（可选）
                        meta_info = pd.DataFrame({
                            '元信息': ['原始文件名', '文件路径', '处理时间', '数据行数', '数据列数'],
                            '值': [
                                filename,
                                file_path,
                                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                len(df),
                                len(df.columns)
                            ]
                        })
                        
                        # 写入Sheet
                        df_with_meta.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                        
                        # 在数据下方添加元信息
                        meta_info.to_excel(writer, sheet_name=sheet_name, index=False, 
                                         startrow=len(df_with_meta)+2, startcol=0)
                        
                        # 记录Sheet信息
                        sheet_info = {
                            'sheet_name': sheet_name,
                            'original_file': filename,
                            'file_path': file_path,
                            'rows': len(df),
                            'columns': len(df.columns),
                            'folder': os.path.dirname(file_path)
                        }
                        self.sheet_info.append(sheet_info)
                        self.total_rows += len(df)
                        self.processed_files += 1
                        
                    except Exception as e:
                        print(f"   ❌ 写入Sheet失败: {str(e)}")
                        self.error_files.append((file_path, f"写入失败: {str(e)}"))
            
            # 生成汇总Sheet
            if ENABLE_SUMMARY and self.sheet_info:
                self.create_summary_sheet(writer)
        
        print(f"\n✅ 数据合并完成!")
        print(f"   输出文件: {OUTPUT_FILE}")
    
    def create_summary_sheet(self, writer):
        """创建汇总信息Sheet"""
        print("📋 正在生成汇总Sheet...")
        
        try:
            # 创建汇总数据
            summary_data = []
            
            # 添加总体统计
            summary_data.extend([
                ['合并统计', ''],
                ['处理时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['成功处理文件数', self.processed_files],
                ['失败文件数', len(self.error_files)],
                ['总Sheet数', len(self.sheet_info)],
                ['总数据行数', self.total_rows],
                ['输出文件', OUTPUT_FILE],
                ['', ''],
            ])
            
            # 添加文件夹统计
            folder_stats = {}
            for info in self.sheet_info:
                folder = os.path.basename(info['folder'])
                if folder not in folder_stats:
                    folder_stats[folder] = {'count': 0, 'rows': 0}
                folder_stats[folder]['count'] += 1
                folder_stats[folder]['rows'] += info['rows']
            
            summary_data.extend([
                ['文件夹统计', ''],
                ['文件夹名称', '文件数量', '数据行数'],
            ])
            
            for folder, stats in folder_stats.items():
                summary_data.append([folder, stats['count'], stats['rows']])
            
            summary_data.append(['', '', ''])
            
            # 添加详细文件列表
            summary_data.extend([
                ['详细文件列表', '', '', '', ''],
                ['Sheet名称', '原始文件名', '数据行数', '数据列数', '所属文件夹'],
            ])
            
            for info in self.sheet_info:
                summary_data.append([
                    info['sheet_name'],
                    info['original_file'],
                    info['rows'],
                    info['columns'],
                    os.path.basename(info['folder'])
                ])
            
            # 添加错误文件列表
            if self.error_files:
                summary_data.extend([
                    ['', '', '', '', ''],
                    ['处理失败的文件', ''],
                    ['文件路径', '错误信息'],
                ])
                
                for file_path, error in self.error_files:
                    summary_data.append([os.path.basename(file_path), error])
            
            # 创建汇总DataFrame
            max_cols = max(len(row) for row in summary_data) if summary_data else 5
            summary_df = pd.DataFrame([row + [''] * (max_cols - len(row)) for row in summary_data])
            
            # 写入汇总Sheet
            summary_df.to_excel(writer, sheet_name='📊汇总信息', index=False, header=False)
            
            print("   ✅ 汇总Sheet创建完成")
            
        except Exception as e:
            print(f"   ❌ 创建汇总Sheet失败: {str(e)}")
    
    def analyze_file_types(self, file_paths):
        """分析文件类型分布"""
        print(f"\n📊 文件类型分析")
        print("-" * 40)
        
        type_stats = {}
        folder_stats = {}
        
        for file_path in file_paths:
            # 文件类型统计
            file_ext = Path(file_path).suffix.lower()
            type_stats[file_ext] = type_stats.get(file_ext, 0) + 1
            
            # 文件夹统计
            folder = os.path.basename(os.path.dirname(file_path))
            folder_stats[folder] = folder_stats.get(folder, 0) + 1
        
        print("文件格式分布:")
        for ext, count in type_stats.items():
            print(f"   {ext}: {count} 个文件")
        
        print("\n文件夹分布:")
        for folder, count in folder_stats.items():
            print(f"   {folder}: {count} 个文件")
    
    def print_final_summary(self):
        """打印最终处理摘要"""
        print(f"\n" + "=" * 60)
        print("📊 多Sheet合并处理摘要")
        print("=" * 60)
        print(f"✅ 成功处理文件数: {self.processed_files}")
        print(f"📄 生成Sheet数: {len(self.sheet_info)}")
        print(f"📈 总数据行数: {self.total_rows:,}")
        print(f"💾 输出文件: {OUTPUT_FILE}")
        
        if self.error_files:
            print(f"\n❌ 处理失败文件数: {len(self.error_files)}")
            print("失败文件:")
            for file_path, error in self.error_files[:5]:  # 只显示前5个
                print(f"   - {os.path.basename(file_path)}: {error}")
            if len(self.error_files) > 5:
                print(f"   ... 还有 {len(self.error_files) - 5} 个失败文件")
        
        # 显示文件大小
        try:
            file_size = os.path.getsize(OUTPUT_FILE) / 1024 / 1024
            print(f"📂 文件大小: {file_size:.1f} MB")
        except:
            pass
        
        print(f"\n💡 使用建议:")
        print(f"   - 使用Excel打开 '{OUTPUT_FILE}' 查看所有数据表")
        print(f"   - 查看 '📊汇总信息' Sheet了解详细统计")
        print(f"   - 每个原始文件都保存为独立的Sheet")
        if len(self.sheet_info) > 20:
            print(f"   - 文件包含大量Sheet({len(self.sheet_info)}个)，建议按需查看")


def main():
    """主函数"""
    print("📊 多Sheet数据合并工具")
    print("=" * 40)
    print("将多个文件夹的数据文件合并为一个Excel文件，每个原始文件作为独立Sheet")
    
    # 显示当前配置
    print(f"\n⚙️ 当前配置:")
    print(f"📁 输入文件夹数量: {len(INPUT_FOLDERS)}")
    for i, folder in enumerate(INPUT_FOLDERS, 1):
        exists = "✅" if os.path.exists(folder) else "❌"
        print(f"   {i}. {exists} {folder}")
    print(f"💾 输出文件: {OUTPUT_FILE}")
    print(f"📊 支持格式: {', '.join(SUPPORTED_FORMATS)}")
    print(f"📄 最大Sheet数: {MAX_SHEETS}")
    
    if not AUTO_RUN:
        confirm = input(f"\n是否开始合并? (y/n): ").lower().strip()
        if confirm != 'y':
            print("操作已取消")
            return
    
    # 开始处理
    merger = MultiSheetMerger()
    
    # 1. 扫描文件
    file_paths = merger.scan_and_collect_files()
    
    if not file_paths:
        print("❌ 未找到任何支持的数据文件!")
        return
    
    # 2. 分析文件类型
    merger.analyze_file_types(file_paths)
    
    # 3. 执行合并
    merger.merge_to_multiple_sheets(file_paths)
    
    # 4. 显示最终摘要
    merger.print_final_summary()
    
    print(f"\n🎉 多Sheet合并完成!")


if __name__ == "__main__":
    main() 