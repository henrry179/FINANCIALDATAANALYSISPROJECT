#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多Sheet数据分析工具
专门分析合并后的Excel文件，提供全面的数据洞察和分析报告
"""

import os
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 尝试导入可视化库
try:
    import matplotlib.pyplot as plt
    import seaborn as sns
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    HAS_PLOTTING = True
except ImportError:
    print("📊 提示: 安装 matplotlib 和 seaborn 可获得可视化功能")
    HAS_PLOTTING = False

# ====== 配置区域 ======

# 输入文件配置
EXCEL_FILES = [
    "数据合并结果_20250601_1703.xlsx",  # 最新合并文件
    "完整金融数据合并_20250601_1658.xlsx",  # 完整金融数据
    "多表合并数据_20250601_1658.xlsx",  # 多表合并数据
]

# 输出配置
ANALYSIS_OUTPUT_DIR = "分析结果"
GENERATE_CHARTS = True
SAVE_DETAILED_REPORTS = True

# ====== 配置区域结束 ======


class MultiSheetAnalyzer:
    """多Sheet数据分析器"""
    
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.sheets_data = {}
        self.summary_info = {}
        self.analysis_results = {}
        
        # 创建输出目录
        os.makedirs(ANALYSIS_OUTPUT_DIR, exist_ok=True)
        
    def load_excel_data(self):
        """加载Excel文件的所有Sheet数据"""
        print(f"📖 正在加载Excel文件: {self.excel_file}")
        
        try:
            # 获取所有Sheet名称
            excel_file = pd.ExcelFile(self.excel_file)
            sheet_names = excel_file.sheet_names
            
            print(f"   📊 发现 {len(sheet_names)} 个Sheet")
            
            # 读取每个Sheet
            for i, sheet_name in enumerate(sheet_names, 1):
                try:
                    print(f"   [{i}/{len(sheet_names)}] 读取Sheet: {sheet_name}")
                    
                    df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                    
                    # 过滤掉元信息行（如果存在）
                    if '元信息' in df.columns or '文件信息' in df.columns:
                        # 找到元信息开始的位置
                        meta_start = None
                        for idx, row in df.iterrows():
                            if '原始文件名' in str(row.values):
                                meta_start = idx
                                break
                        
                        if meta_start is not None:
                            df = df.iloc[:meta_start]  # 只保留数据部分
                    
                    self.sheets_data[sheet_name] = df
                    print(f"      ✅ 成功读取: {len(df)} 行, {len(df.columns)} 列")
                    
                except Exception as e:
                    print(f"      ❌ 读取失败: {str(e)}")
                    continue
            
            print(f"✅ 成功加载 {len(self.sheets_data)} 个Sheet的数据")
            return True
            
        except Exception as e:
            print(f"❌ 加载Excel文件失败: {str(e)}")
            return False
    
    def analyze_data_overview(self):
        """数据概览分析"""
        print(f"\n📊 数据概览分析")
        print("-" * 50)
        
        total_rows = 0
        total_cols = 0
        sheet_info = []
        
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':  # 跳过汇总Sheet
                continue
                
            rows = len(df)
            cols = len(df.columns)
            total_rows += rows
            
            sheet_info.append({
                'sheet_name': sheet_name,
                'rows': rows,
                'columns': cols,
                'memory_usage_mb': df.memory_usage(deep=True).sum() / 1024 / 1024,
                'non_null_ratio': df.count().sum() / (rows * cols) if rows * cols > 0 else 0
            })
        
        # 排序：按行数排序
        sheet_info.sort(key=lambda x: x['rows'], reverse=True)
        
        print(f"数据集总览:")
        print(f"   📄 总Sheet数: {len(sheet_info)}")
        print(f"   📈 总数据行数: {total_rows:,}")
        print(f"   💾 总内存占用: {sum(info['memory_usage_mb'] for info in sheet_info):.1f} MB")
        
        print(f"\n📋 各Sheet详细信息:")
        print(f"{'Sheet名称':<30} {'行数':<10} {'列数':<8} {'数据完整度':<10}")
        print("-" * 65)
        
        for info in sheet_info[:10]:  # 显示前10个最大的Sheet
            completeness = f"{info['non_null_ratio']*100:.1f}%"
            print(f"{info['sheet_name']:<30} {info['rows']:<10,} {info['columns']:<8} {completeness:<10}")
        
        if len(sheet_info) > 10:
            print(f"... 还有 {len(sheet_info) - 10} 个Sheet")
        
        self.analysis_results['overview'] = {
            'total_sheets': len(sheet_info),
            'total_rows': total_rows,
            'sheet_info': sheet_info
        }
        
        return sheet_info
    
    def analyze_data_types(self):
        """数据类型分析"""
        print(f"\n🔍 数据类型分析")
        print("-" * 50)
        
        all_numeric_cols = set()
        all_text_cols = set()
        all_date_cols = set()
        common_columns = None
        
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':
                continue
            
            # 分析列类型
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            text_cols = df.select_dtypes(include=['object']).columns.tolist()
            date_cols = df.select_dtypes(include=['datetime']).columns.tolist()
            
            # 尝试识别可能的日期列
            potential_date_cols = []
            for col in text_cols:
                if any(keyword in col.lower() for keyword in ['date', '日期', 'time', '时间', '年', '月', '日']):
                    potential_date_cols.append(col)
            
            all_numeric_cols.update(numeric_cols)
            all_text_cols.update(text_cols)
            all_date_cols.update(date_cols + potential_date_cols)
            
            # 找出共同列
            if common_columns is None:
                common_columns = set(df.columns)
            else:
                common_columns = common_columns.intersection(set(df.columns))
        
        print(f"数据类型统计:")
        print(f"   🔢 数值型列: {len(all_numeric_cols)} 个")
        print(f"   📝 文本型列: {len(all_text_cols)} 个")
        print(f"   📅 日期型列: {len(all_date_cols)} 个")
        print(f"   🔗 共同列: {len(common_columns)} 个")
        
        if len(common_columns) > 0:
            print(f"\n共同列名:")
            for col in sorted(list(common_columns))[:10]:
                print(f"   - {col}")
            if len(common_columns) > 10:
                print(f"   ... 还有 {len(common_columns) - 10} 个")
        
        self.analysis_results['data_types'] = {
            'numeric_columns': list(all_numeric_cols),
            'text_columns': list(all_text_cols),
            'date_columns': list(all_date_cols),
            'common_columns': list(common_columns)
        }
        
        return all_numeric_cols, all_text_cols, all_date_cols, common_columns
    
    def analyze_data_quality(self):
        """数据质量分析"""
        print(f"\n🎯 数据质量分析")
        print("-" * 50)
        
        quality_report = []
        
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':
                continue
            
            # 计算各种质量指标
            total_cells = len(df) * len(df.columns)
            null_cells = df.isnull().sum().sum()
            duplicate_rows = df.duplicated().sum()
            
            # 数值列的统计
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            numeric_stats = {}
            
            if len(numeric_cols) > 0:
                for col in numeric_cols:
                    stats = {
                        'mean': df[col].mean(),
                        'std': df[col].std(),
                        'min': df[col].min(),
                        'max': df[col].max(),
                        'zeros': (df[col] == 0).sum(),
                        'negatives': (df[col] < 0).sum()
                    }
                    numeric_stats[col] = stats
            
            quality_info = {
                'sheet_name': sheet_name,
                'total_cells': total_cells,
                'null_cells': null_cells,
                'null_percentage': (null_cells / total_cells * 100) if total_cells > 0 else 0,
                'duplicate_rows': duplicate_rows,
                'duplicate_percentage': (duplicate_rows / len(df) * 100) if len(df) > 0 else 0,
                'numeric_stats': numeric_stats
            }
            
            quality_report.append(quality_info)
        
        # 显示质量报告
        print(f"数据质量报告:")
        print(f"{'Sheet名称':<30} {'缺失率':<10} {'重复率':<10} {'数值列数':<10}")
        print("-" * 70)
        
        for info in quality_report[:10]:
            null_pct = f"{info['null_percentage']:.1f}%"
            dup_pct = f"{info['duplicate_percentage']:.1f}%"
            num_cols = len(info['numeric_stats'])
            print(f"{info['sheet_name']:<30} {null_pct:<10} {dup_pct:<10} {num_cols:<10}")
        
        self.analysis_results['quality'] = quality_report
        return quality_report
    
    def analyze_numerical_data(self):
        """数值数据分析"""
        print(f"\n📈 数值数据分析")
        print("-" * 50)
        
        all_numeric_data = []
        numeric_summary = {}
        
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':
                continue
            
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            if len(numeric_cols) > 0:
                for col in numeric_cols:
                    values = df[col].dropna()
                    if len(values) > 0:
                        all_numeric_data.extend(values)
                        
                        if col not in numeric_summary:
                            numeric_summary[col] = {
                                'total_values': 0,
                                'sum': 0,
                                'min': float('inf'),
                                'max': float('-inf'),
                                'sheets': []
                            }
                        
                        numeric_summary[col]['total_values'] += len(values)
                        numeric_summary[col]['sum'] += values.sum()
                        numeric_summary[col]['min'] = min(numeric_summary[col]['min'], values.min())
                        numeric_summary[col]['max'] = max(numeric_summary[col]['max'], values.max())
                        numeric_summary[col]['sheets'].append(sheet_name)
        
        # 显示数值分析结果
        if all_numeric_data:
            print(f"数值数据概述:")
            print(f"   📊 总数值个数: {len(all_numeric_data):,}")
            print(f"   📈 数值范围: {min(all_numeric_data):.2f} ~ {max(all_numeric_data):.2f}")
            print(f"   📊 平均值: {np.mean(all_numeric_data):.2f}")
            print(f"   📊 中位数: {np.median(all_numeric_data):.2f}")
            
            print(f"\n主要数值列分析:")
            print(f"{'列名':<25} {'总值数':<10} {'最小值':<12} {'最大值':<12} {'平均值':<12}")
            print("-" * 80)
            
            # 按总值数排序，显示前10个
            sorted_cols = sorted(numeric_summary.items(), 
                               key=lambda x: x[1]['total_values'], reverse=True)
            
            for col_name, stats in sorted_cols[:10]:
                avg_val = stats['sum'] / stats['total_values']
                print(f"{col_name:<25} {stats['total_values']:<10,} {stats['min']:<12.2f} {stats['max']:<12.2f} {avg_val:<12.2f}")
        
        self.analysis_results['numerical'] = {
            'total_values': len(all_numeric_data),
            'column_summary': numeric_summary
        }
        
        return numeric_summary
    
    def find_patterns_and_insights(self):
        """寻找数据模式和洞察"""
        print(f"\n🔍 数据模式和洞察分析")
        print("-" * 50)
        
        insights = []
        
        # 1. 寻找高相关性的数值列
        correlations = []
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':
                continue
                
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) >= 2:
                try:
                    corr_matrix = df[numeric_cols].corr()
                    # 找出高相关性的列对
                    for i in range(len(numeric_cols)):
                        for j in range(i+1, len(numeric_cols)):
                            corr_val = corr_matrix.iloc[i, j]
                            if abs(corr_val) > 0.7 and not np.isnan(corr_val):
                                correlations.append({
                                    'sheet': sheet_name,
                                    'col1': numeric_cols[i],
                                    'col2': numeric_cols[j],
                                    'correlation': corr_val
                                })
                except:
                    continue
        
        if correlations:
            print(f"发现高相关性列对:")
            correlations.sort(key=lambda x: abs(x['correlation']), reverse=True)
            for corr in correlations[:5]:
                print(f"   📊 {corr['sheet']}: {corr['col1']} ↔ {corr['col2']} (相关性: {corr['correlation']:.3f})")
        
        # 2. 寻找异常值
        outliers = []
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':
                continue
                
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            for col in numeric_cols:
                values = df[col].dropna()
                if len(values) > 10:  # 足够的数据点
                    Q1 = values.quantile(0.25)
                    Q3 = values.quantile(0.75)
                    IQR = Q3 - Q1
                    lower_bound = Q1 - 1.5 * IQR
                    upper_bound = Q3 + 1.5 * IQR
                    
                    outlier_count = ((values < lower_bound) | (values > upper_bound)).sum()
                    if outlier_count > 0:
                        outliers.append({
                            'sheet': sheet_name,
                            'column': col,
                            'outlier_count': outlier_count,
                            'outlier_percentage': outlier_count / len(values) * 100
                        })
        
        if outliers:
            print(f"\n发现数据异常值:")
            outliers.sort(key=lambda x: x['outlier_percentage'], reverse=True)
            for outlier in outliers[:5]:
                print(f"   ⚠️  {outlier['sheet']}.{outlier['column']}: {outlier['outlier_count']} 个异常值 ({outlier['outlier_percentage']:.1f}%)")
        
        # 3. 数据分布分析
        distribution_insights = []
        for sheet_name, df in self.sheets_data.items():
            if sheet_name == '📊汇总信息':
                continue
                
            # 检查是否有明显的数据模式
            if len(df) > 100:  # 足够的数据进行分析
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                for col in numeric_cols[:3]:  # 只分析前3个数值列
                    values = df[col].dropna()
                    if len(values) > 50:
                        # 检查数据分布
                        skewness = values.skew()
                        if abs(skewness) > 1:
                            distribution_insights.append({
                                'sheet': sheet_name,
                                'column': col,
                                'skewness': skewness,
                                'distribution': 'right_skewed' if skewness > 1 else 'left_skewed'
                            })
        
        if distribution_insights:
            print(f"\n数据分布特征:")
            for insight in distribution_insights[:5]:
                dist_type = "右偏" if insight['distribution'] == 'right_skewed' else "左偏"
                print(f"   📊 {insight['sheet']}.{insight['column']}: {dist_type}分布 (偏度: {insight['skewness']:.2f})")
        
        insights.extend([
            {'type': 'correlations', 'data': correlations},
            {'type': 'outliers', 'data': outliers},
            {'type': 'distributions', 'data': distribution_insights}
        ])
        
        self.analysis_results['insights'] = insights
        return insights
    
    def generate_summary_report(self):
        """生成分析总结报告"""
        print(f"\n📋 生成分析总结报告")
        print("-" * 50)
        
        report_file = os.path.join(ANALYSIS_OUTPUT_DIR, f"数据分析报告_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")
        
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(f"多Sheet数据分析报告\n")
            f.write(f"=" * 50 + "\n")
            f.write(f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"分析文件: {self.excel_file}\n\n")
            
            # 数据概览
            if 'overview' in self.analysis_results:
                overview = self.analysis_results['overview']
                f.write(f"数据概览:\n")
                f.write(f"  总Sheet数: {overview['total_sheets']}\n")
                f.write(f"  总行数: {overview['total_rows']:,}\n\n")
                
                f.write(f"主要Sheet信息:\n")
                for info in overview['sheet_info'][:10]:
                    f.write(f"  - {info['sheet_name']}: {info['rows']:,} 行, {info['columns']} 列\n")
                f.write("\n")
            
            # 数据质量
            if 'quality' in self.analysis_results:
                f.write(f"数据质量摘要:\n")
                quality_data = self.analysis_results['quality']
                avg_null_rate = np.mean([q['null_percentage'] for q in quality_data])
                avg_dup_rate = np.mean([q['duplicate_percentage'] for q in quality_data])
                f.write(f"  平均缺失率: {avg_null_rate:.1f}%\n")
                f.write(f"  平均重复率: {avg_dup_rate:.1f}%\n\n")
            
            # 数值分析
            if 'numerical' in self.analysis_results:
                numerical = self.analysis_results['numerical']
                f.write(f"数值数据分析:\n")
                f.write(f"  总数值个数: {numerical['total_values']:,}\n")
                f.write(f"  主要数值列数: {len(numerical['column_summary'])}\n\n")
            
            # 关键洞察
            if 'insights' in self.analysis_results:
                f.write(f"关键洞察:\n")
                insights = self.analysis_results['insights']
                for insight_group in insights:
                    if insight_group['type'] == 'correlations' and insight_group['data']:
                        f.write(f"  - 发现 {len(insight_group['data'])} 个高相关性列对\n")
                    elif insight_group['type'] == 'outliers' and insight_group['data']:
                        f.write(f"  - 发现 {len(insight_group['data'])} 个包含异常值的列\n")
                f.write("\n")
            
            # 建议
            f.write(f"分析建议:\n")
            f.write(f"  1. 建议重点关注数据量最大的几个Sheet\n")
            f.write(f"  2. 对于缺失率高的数据，建议进行数据清洗\n")
            f.write(f"  3. 对于高相关性的列，可以考虑降维或特征选择\n")
            f.write(f"  4. 对于包含异常值的列，建议进一步调查数据来源\n")
        
        print(f"✅ 分析报告已保存到: {report_file}")
        return report_file
    
    def run_full_analysis(self):
        """运行完整分析流程"""
        print(f"🚀 开始多Sheet数据分析")
        print(f"分析文件: {self.excel_file}")
        print("=" * 60)
        
        # 1. 加载数据
        if not self.load_excel_data():
            return False
        
        # 2. 数据概览
        self.analyze_data_overview()
        
        # 3. 数据类型分析
        self.analyze_data_types()
        
        # 4. 数据质量分析
        self.analyze_data_quality()
        
        # 5. 数值数据分析
        self.analyze_numerical_data()
        
        # 6. 模式和洞察分析
        self.find_patterns_and_insights()
        
        # 7. 生成报告
        report_file = self.generate_summary_report()
        
        print(f"\n🎉 分析完成!")
        print(f"📋 详细报告: {report_file}")
        
        return True


def select_file_to_analyze():
    """选择要分析的文件"""
    print(f"📁 检测到以下Excel文件:")
    
    available_files = []
    for i, filename in enumerate(EXCEL_FILES, 1):
        if os.path.exists(filename):
            file_size = os.path.getsize(filename) / 1024 / 1024
            available_files.append(filename)
            print(f"   {i}. {filename} ({file_size:.1f} MB)")
        else:
            print(f"   {i}. {filename} (文件不存在)")
    
    if not available_files:
        print("❌ 没有找到可分析的Excel文件!")
        return None
    
    try:
        print(f"\n请选择要分析的文件 (1-{len(available_files)})，或直接回车分析第一个文件:")
        choice = input().strip()
        
        if not choice:
            return available_files[0]
        
        choice_idx = int(choice) - 1
        if 0 <= choice_idx < len(available_files):
            return available_files[choice_idx]
        else:
            print("无效选择，将分析第一个文件")
            return available_files[0]
            
    except (ValueError, KeyboardInterrupt):
        print("将分析第一个文件")
        return available_files[0]


def main():
    """主函数"""
    print("📊 多Sheet数据分析工具")
    print("=" * 40)
    print("专门分析合并后的Excel文件，提供全面的数据洞察")
    print()
    
    # 选择分析文件
    excel_file = select_file_to_analyze()
    if not excel_file:
        return
    
    print(f"\n🎯 将分析文件: {excel_file}")
    
    try:
        confirm = input("按回车开始分析，输入n取消: ").strip().lower()
        if confirm in ['n', 'no', '否']:
            print("分析已取消")
            return
    except KeyboardInterrupt:
        print("\n分析已取消")
        return
    
    # 创建分析器并运行分析
    analyzer = MultiSheetAnalyzer(excel_file)
    success = analyzer.run_full_analysis()
    
    if success:
        print(f"\n💡 分析结果已保存到 '{ANALYSIS_OUTPUT_DIR}' 文件夹")
        print(f"   - 详细分析报告")
        if HAS_PLOTTING and GENERATE_CHARTS:
            print(f"   - 数据可视化图表")
    else:
        print(f"\n❌ 分析过程中出现错误")


if __name__ == "__main__":
    main() 