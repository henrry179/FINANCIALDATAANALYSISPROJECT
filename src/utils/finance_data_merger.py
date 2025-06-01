#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金融数据合并与智能分析脚本
专门用于合并金融相关的Excel和CSV数据表，并提供多维度智能分析
"""

import os
import pandas as pd
import numpy as np
import glob
from pathlib import Path
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# 尝试导入可视化和统计分析库
try:
    import matplotlib.pyplot as plt
    import seaborn as sns
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    HAS_PLOTTING = True
except ImportError:
    print("📊 提示: 安装 matplotlib 和 seaborn 可获得可视化功能")
    HAS_PLOTTING = False

try:
    from scipy import stats
    from sklearn.preprocessing import StandardScaler
    from sklearn.decomposition import PCA
    HAS_ADVANCED_STATS = True
except ImportError:
    print("📈 提示: 安装 scipy 和 scikit-learn 可获得高级统计分析功能")
    HAS_ADVANCED_STATS = False


# ====== 金融数据配置区域 ======

# 金融数据文件夹路径
FINANCE_DATA_FOLDERS = [
    "/Users/mac/Downloads/WorkFiles/financedatasets01-0601",  # 金融数据集
    "/Users/mac/Downloads/WorkFiles/financedatasets02-0601",  # 金融数据集
    "/Users/mac/Downloads/WorkFiles/financedatasets03-0601",  # 金融数据集
    "/Users/mac/Downloads/WorkFiles/financedatasets04-0601",  # 金融数据集

    # 可以添加更多金融数据文件夹路径
    # "/Users/mac/Documents/金融数据2024",
    # "/Users/mac/Downloads/股票数据",
]

# 输出文件配置
OUTPUT_FILE = f"金融数据汇总_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
ANALYSIS_REPORT = f"金融数据分析报告_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"

# 支持的文件格式
SUPPORTED_FORMATS = ['.xlsx', '.csv', '.xls']

# 功能开关
AUTO_RUN = True
ENABLE_ANALYSIS = True  # 是否启用数据分析功能
SAVE_PLOTS = True       # 是否保存图表
DETAILED_REPORT = True  # 是否生成详细报告

# ====== 配置区域结束 ======


class FinanceDataAnalyzer:
    """金融数据智能分析器"""
    
    def __init__(self, data):
        self.data = data
        self.numeric_cols = []
        self.date_cols = []
        self.analysis_results = {}
        self._prepare_data()
    
    def _prepare_data(self):
        """数据预处理和类型识别"""
        print("🔍 正在分析数据结构...")
        
        # 识别数值列
        for col in self.data.columns:
            # 确保列名是字符串类型
            col_name = str(col)
            if self.data[col].dtype in ['int64', 'float64']:
                if not col_name.startswith(('数据来源', '数据文件夹', '处理时间')):
                    self.numeric_cols.append(col)
        
        # 识别日期列
        for col in self.data.columns:
            col_name = str(col)
            if '日期' in col_name or '时间' in col_name or 'date' in col_name.lower():
                try:
                    pd.to_datetime(self.data[col])
                    self.date_cols.append(col)
                except:
                    pass
        
        print(f"   📊 识别到 {len(self.numeric_cols)} 个数值列")
        print(f"   📅 识别到 {len(self.date_cols)} 个日期列")
    
    def basic_statistics(self):
        """基础统计分析"""
        print("\n📈 1. 基础统计分析")
        print("-" * 40)
        
        if not self.numeric_cols:
            print("   ⚠️  未发现数值型数据，跳过统计分析")
            return {}
        
        # 描述性统计
        numeric_data = self.data[self.numeric_cols]
        desc_stats = numeric_data.describe()
        
        # 计算额外统计指标
        additional_stats = {}
        for col in self.numeric_cols[:10]:  # 限制分析前10列，避免输出过长
            col_data = numeric_data[col].dropna()
            if len(col_data) > 0:
                additional_stats[col] = {
                    '缺失值数量': self.data[col].isnull().sum(),
                    '缺失值比例': f"{self.data[col].isnull().mean()*100:.2f}%",
                    '唯一值数量': self.data[col].nunique(),
                    '偏度': f"{col_data.skew():.3f}",
                    '峰度': f"{col_data.kurtosis():.3f}",
                }
        
        # 显示关键统计信息
        print("   📊 数据概览:")
        print(f"      总行数: {len(self.data):,}")
        print(f"      总列数: {len(self.data.columns)}")
        print(f"      数值列数: {len(self.numeric_cols)}")
        
        print("\n   📋 主要数值列统计（前5列）:")
        for col in self.numeric_cols[:5]:
            if col in additional_stats:
                stats = additional_stats[col]
                print(f"      {col}:")
                print(f"         均值: {numeric_data[col].mean():.3f}")
                print(f"         标准差: {numeric_data[col].std():.3f}")
                print(f"         缺失值: {stats['缺失值数量']} ({stats['缺失值比例']})")
        
        self.analysis_results['basic_stats'] = {
            'desc_stats': desc_stats,
            'additional_stats': additional_stats
        }
        
        return self.analysis_results['basic_stats']
    
    def trend_analysis(self):
        """趋势分析"""
        print("\n📈 2. 趋势分析")
        print("-" * 40)
        
        if not self.date_cols or not self.numeric_cols:
            print("   ⚠️  缺少日期或数值数据，跳过趋势分析")
            return {}
        
        trends = {}
        
        # 分析时间序列趋势
        for date_col in self.date_cols[:2]:  # 限制分析前2个日期列
            try:
                # 转换日期列
                self.data[date_col] = pd.to_datetime(self.data[date_col])
                
                # 按日期排序
                temp_data = self.data.sort_values(date_col)
                
                print(f"   📅 分析时间序列: {date_col}")
                print(f"      时间范围: {temp_data[date_col].min()} 至 {temp_data[date_col].max()}")
                
                # 分析数值列的趋势
                for num_col in self.numeric_cols[:3]:  # 限制分析前3个数值列
                    if num_col in temp_data.columns:
                        valid_data = temp_data[[date_col, num_col]].dropna()
                        if len(valid_data) > 10:
                            # 计算趋势
                            x = np.arange(len(valid_data))
                            y = valid_data[num_col].values
                            
                            if HAS_ADVANCED_STATS:
                                slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
                                trend_direction = "上升" if slope > 0 else "下降"
                                trend_strength = abs(r_value)
                                
                                trends[f"{date_col}_{num_col}"] = {
                                    'slope': slope,
                                    'r_squared': r_value**2,
                                    'direction': trend_direction,
                                    'strength': trend_strength,
                                    'significance': 'significant' if p_value < 0.05 else 'not_significant'
                                }
                                
                                print(f"      {num_col}: {trend_direction}趋势 (R²={r_value**2:.3f})")
            except Exception as e:
                print(f"   ⚠️  {date_col} 趋势分析失败: {str(e)}")
        
        self.analysis_results['trends'] = trends
        return trends
    
    def correlation_analysis(self):
        """相关性分析"""
        print("\n🔗 3. 相关性分析")
        print("-" * 40)
        
        if len(self.numeric_cols) < 2:
            print("   ⚠️  数值列少于2个，跳过相关性分析")
            return {}
        
        # 计算相关性矩阵
        numeric_data = self.data[self.numeric_cols[:10]].dropna()  # 限制前10列，删除缺失值
        
        if len(numeric_data) < 10:
            print("   ⚠️  有效数据不足，跳过相关性分析")
            return {}
        
        corr_matrix = numeric_data.corr()
        
        # 找出强相关关系
        strong_correlations = []
        for i in range(len(corr_matrix.columns)):
            for j in range(i+1, len(corr_matrix.columns)):
                corr_value = corr_matrix.iloc[i, j]
                if abs(corr_value) > 0.7:  # 强相关阈值
                    strong_correlations.append({
                        'var1': corr_matrix.columns[i],
                        'var2': corr_matrix.columns[j],
                        'correlation': corr_value,
                        'strength': '强正相关' if corr_value > 0.7 else '强负相关'
                    })
        
        print(f"   🔍 发现 {len(strong_correlations)} 对强相关变量:")
        for corr in strong_correlations[:5]:  # 显示前5对
            print(f"      {corr['var1']} ↔ {corr['var2']}: {corr['correlation']:.3f} ({corr['strength']})")
        
        self.analysis_results['correlations'] = {
            'matrix': corr_matrix,
            'strong_correlations': strong_correlations
        }
        
        # 保存相关性热图
        if HAS_PLOTTING and SAVE_PLOTS:
            try:
                plt.figure(figsize=(12, 10))
                sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, 
                           fmt='.2f', square=True)
                plt.title('金融数据相关性热图')
                plt.tight_layout()
                plt.savefig('金融数据相关性热图.png', dpi=300, bbox_inches='tight')
                plt.close()
                print("   📊 相关性热图已保存为 '金融数据相关性热图.png'")
            except Exception as e:
                print(f"   ⚠️  保存热图失败: {str(e)}")
        
        return self.analysis_results['correlations']
    
    def risk_analysis(self):
        """风险分析"""
        print("\n⚠️  4. 风险分析")
        print("-" * 40)
        
        if not self.numeric_cols:
            print("   ⚠️  无数值数据，跳过风险分析")
            return {}
        
        risk_metrics = {}
        
        # 计算各类风险指标
        for col in self.numeric_cols[:5]:  # 限制前5列
            col_data = self.data[col].dropna()
            if len(col_data) > 10:
                # 波动性分析
                volatility = col_data.std() / abs(col_data.mean()) if col_data.mean() != 0 else float('inf')
                
                # VaR计算 (Value at Risk)
                var_95 = np.percentile(col_data, 5)
                var_99 = np.percentile(col_data, 1)
                
                # 极值分析
                q1 = col_data.quantile(0.25)
                q3 = col_data.quantile(0.75)
                iqr = q3 - q1
                outliers = col_data[(col_data < q1 - 1.5*iqr) | (col_data > q3 + 1.5*iqr)]
                
                risk_metrics[col] = {
                    'volatility': volatility,
                    'var_95': var_95,
                    'var_99': var_99,
                    'outliers_count': len(outliers),
                    'outliers_ratio': len(outliers) / len(col_data) * 100,
                    'max_drawdown': (col_data.max() - col_data.min()) / col_data.max() if col_data.max() != 0 else 0
                }
                
                risk_level = "高" if volatility > 0.5 else "中" if volatility > 0.2 else "低"
                print(f"   📊 {col}:")
                print(f"      波动性: {volatility:.3f} (风险等级: {risk_level})")
                print(f"      VaR(95%): {var_95:.3f}")
                print(f"      异常值比例: {len(outliers) / len(col_data) * 100:.2f}%")
        
        self.analysis_results['risk_analysis'] = risk_metrics
        return risk_metrics
    
    def market_insights(self):
        """市场洞察分析"""
        print("\n💡 5. 市场洞察分析")
        print("-" * 40)
        
        insights = []
        
        # 数据完整性分析
        missing_ratio = self.data.isnull().sum().sum() / (len(self.data) * len(self.data.columns))
        if missing_ratio > 0.1:
            insights.append(f"⚠️  数据缺失率较高 ({missing_ratio*100:.1f}%)，建议关注数据质量")
        
        # 数值分布分析
        if self.numeric_cols:
            for col in self.numeric_cols[:3]:
                col_data = self.data[col].dropna()
                if len(col_data) > 100:
                    skewness = col_data.skew()
                    if abs(skewness) > 1:
                        skew_type = "右偏" if skewness > 0 else "左偏"
                        insights.append(f"📈 {col} 呈现{skew_type}分布，可能存在极端值影响")
        
        # 时间序列分析
        if self.date_cols and 'trends' in self.analysis_results:
            trend_count = len(self.analysis_results['trends'])
            if trend_count > 0:
                insights.append(f"📅 检测到 {trend_count} 个时间序列趋势，建议关注时间效应")
        
        # 相关性洞察
        if 'correlations' in self.analysis_results:
            strong_corr_count = len(self.analysis_results['correlations']['strong_correlations'])
            if strong_corr_count > 5:
                insights.append(f"🔗 发现 {strong_corr_count} 对强相关变量，存在多重共线性风险")
        
        # 风险评估
        if 'risk_analysis' in self.analysis_results:
            high_risk_vars = sum(1 for metrics in self.analysis_results['risk_analysis'].values() 
                               if metrics['volatility'] > 0.5)
            if high_risk_vars > 0:
                insights.append(f"⚠️  发现 {high_risk_vars} 个高风险变量，建议加强风险管控")
        
        # 数据规模洞察
        data_size_mb = self.data.memory_usage(deep=True).sum() / 1024 / 1024
        if data_size_mb > 100:
            insights.append(f"💾 数据量较大 ({data_size_mb:.1f}MB)，建议考虑分批处理或优化存储")
        
        print("   🔍 关键洞察:")
        for i, insight in enumerate(insights, 1):
            print(f"      {i}. {insight}")
        
        if not insights:
            insights.append("✅ 数据质量良好，未发现明显异常")
            print("   ✅ 数据质量良好，未发现明显异常")
        
        self.analysis_results['insights'] = insights
        return insights
    
    def generate_comprehensive_report(self):
        """生成综合分析报告"""
        print("\n📝 6. 生成综合分析报告")
        print("-" * 40)
        
        report_lines = [
            "=" * 80,
            f"金融数据智能分析报告",
            f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 80,
            "",
            "📊 数据概览",
            "-" * 40,
            f"数据行数: {len(self.data):,}",
            f"数据列数: {len(self.data.columns)}",
            f"数值列数: {len(self.numeric_cols)}",
            f"日期列数: {len(self.date_cols)}",
            "",
        ]
        
        # 添加基础统计
        if 'basic_stats' in self.analysis_results:
            report_lines.extend([
                "📈 基础统计摘要",
                "-" * 40,
            ])
            for col in self.numeric_cols[:5]:
                if col in self.data.columns:
                    mean_val = self.data[col].mean()
                    std_val = self.data[col].std()
                    report_lines.append(f"{col}: 均值={mean_val:.3f}, 标准差={std_val:.3f}")
            report_lines.append("")
        
        # 添加相关性分析
        if 'correlations' in self.analysis_results:
            strong_corrs = self.analysis_results['correlations']['strong_correlations']
            report_lines.extend([
                "🔗 强相关关系",
                "-" * 40,
            ])
            for corr in strong_corrs[:10]:
                report_lines.append(f"{corr['var1']} ↔ {corr['var2']}: {corr['correlation']:.3f}")
            report_lines.append("")
        
        # 添加风险分析
        if 'risk_analysis' in self.analysis_results:
            report_lines.extend([
                "⚠️ 风险评估",
                "-" * 40,
            ])
            for var, metrics in list(self.analysis_results['risk_analysis'].items())[:5]:
                risk_level = "高" if metrics['volatility'] > 0.5 else "中" if metrics['volatility'] > 0.2 else "低"
                report_lines.append(f"{var}: 波动性={metrics['volatility']:.3f} (风险等级: {risk_level})")
            report_lines.append("")
        
        # 添加市场洞察
        if 'insights' in self.analysis_results:
            report_lines.extend([
                "💡 关键洞察",
                "-" * 40,
            ])
            for insight in self.analysis_results['insights']:
                report_lines.append(f"• {insight}")
            report_lines.append("")
        
        # 添加建议
        report_lines.extend([
            "🎯 分析建议",
            "-" * 40,
            "1. 定期监控高风险变量的波动情况",
            "2. 关注强相关变量间的关系变化",
            "3. 建立预警机制识别异常数据",
            "4. 考虑使用机器学习模型进行预测分析",
            "5. 持续收集和更新数据以提高分析准确性",
            "",
            "=" * 80,
            "报告结束"
        ])
        
        # 保存报告
        try:
            with open(ANALYSIS_REPORT, 'w', encoding='utf-8') as f:
                f.write('\n'.join(report_lines))
            print(f"   ✅ 分析报告已保存为: {ANALYSIS_REPORT}")
        except Exception as e:
            print(f"   ❌ 保存报告失败: {str(e)}")
        
        return '\n'.join(report_lines)
    
    def run_full_analysis(self):
        """运行完整分析流程"""
        print("\n🤖 启动金融数据智能分析系统")
        print("=" * 50)
        
        # 执行各项分析
        self.basic_statistics()
        self.trend_analysis()
        self.correlation_analysis()
        self.risk_analysis()
        self.market_insights()
        
        if DETAILED_REPORT:
            self.generate_comprehensive_report()
        
        print("\n✅ 智能分析完成!")
        return self.analysis_results


def merge_finance_data():
    """合并金融数据文件"""
    
    print("🏦 金融数据合并工具")
    print("=" * 40)
    
    # 记录处理信息
    all_files = []
    processed_files = 0
    error_files = []
    
    # 1. 扫描所有指定文件夹
    print(f"\n正在扫描 {len(FINANCE_DATA_FOLDERS)} 个金融数据文件夹...")
    
    for folder_path in FINANCE_DATA_FOLDERS:
        if not os.path.exists(folder_path):
            print(f"⚠️  警告: 文件夹 '{folder_path}' 不存在，跳过...")
            continue
        
        print(f"📁 扫描文件夹: {folder_path}")
        
        # 递归查找所有支持格式的文件
        for ext in SUPPORTED_FORMATS:
            pattern = os.path.join(folder_path, '**', f'*{ext}')
            files = glob.glob(pattern, recursive=True)
            all_files.extend(files)
            print(f"   📊 找到 {len(files)} 个 {ext} 文件")
    
    print(f"\n📈 总共找到 {len(all_files)} 个金融数据文件")
    
    if not all_files:
        print("❌ 未找到任何支持的金融数据文件!")
        return None
    
    # 2. 按文件类型分类显示
    print("\n📋 文件类型分析:")
    file_types = {}
    for file_path in all_files:
        filename = os.path.basename(file_path).lower()
        
        # 简单分类
        if any(keyword in filename for keyword in ['股票', '股指', 'a股', '上证', '深证']):
            category = "📈 股票数据"
        elif any(keyword in filename for keyword in ['汇率', '外汇', '人民币']):
            category = "💱 汇率数据"  
        elif any(keyword in filename for keyword in ['利率', 'shibor', 'lpr', '拆借']):
            category = "💰 利率数据"
        elif any(keyword in filename for keyword in ['银行', '工商', '建设', '交通']):
            category = "🏦 银行数据"
        elif any(keyword in filename for keyword in ['货币', 'm2', '供应量']):
            category = "💸 货币数据"
        else:
            category = "📊 其他数据"
            
        if category not in file_types:
            file_types[category] = 0
        file_types[category] += 1
    
    for category, count in file_types.items():
        print(f"   {category}: {count} 个文件")
    
    # 3. 读取并合并所有文件
    print(f"\n🔄 开始读取和合并数据...")
    all_dataframes = []
    
    for i, file_path in enumerate(all_files, 1):
        try:
            filename = os.path.basename(file_path)
            print(f"📖 [{i}/{len(all_files)}] 正在处理: {filename}")
            
            # 根据文件扩展名选择读取方法
            file_ext = Path(file_path).suffix.lower()
            
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
            
            # 添加文件来源信息
            df['数据来源文件'] = filename
            df['数据文件夹'] = os.path.dirname(file_path)
            df['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            all_dataframes.append(df)
            processed_files += 1
            print(f"   ✅ 成功读取 {len(df)} 行数据，{len(df.columns)} 列")
            
        except Exception as e:
            error_msg = f"读取失败: {str(e)}"
            print(f"   ❌ {error_msg}")
            error_files.append((file_path, error_msg))
    
    # 4. 合并所有数据
    if all_dataframes:
        print(f"\n🔗 正在合并 {len(all_dataframes)} 个文件的数据...")
        merged_data = pd.concat(all_dataframes, ignore_index=True, sort=False)
        
        # 5. 保存合并后的数据
        print(f"💾 正在保存数据到: {OUTPUT_FILE}")
        
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
            
            print("\n🎉 金融数据合并完成!")
            
        except Exception as e:
            print(f"❌ 保存文件失败: {str(e)}")
            return None
    
    else:
        print("❌ 没有有效的数据可以合并!")
        return None
    
    # 6. 显示详细处理摘要
    print("\n" + "=" * 60)
    print("📊 金融数据处理摘要")
    print("=" * 60)
    print(f"✅ 成功处理文件数: {processed_files}")
    print(f"📈 合并后总行数: {len(merged_data):,}")
    print(f"📋 合并后总列数: {len(merged_data.columns)}")
    print(f"💾 输出文件: {OUTPUT_FILE}")
    print(f"📂 文件大小: {os.path.getsize(OUTPUT_FILE) / 1024:.1f} KB")
    
    if error_files:
        print(f"\n❌ 处理失败文件数: {len(error_files)}")
        for file_path, error in error_files:
            print(f"   - {os.path.basename(file_path)}: {error}")
    
    print(f"\n📋 数据概览（前3行）:")
    pd.set_option('display.max_columns', 10)
    pd.set_option('display.width', 1000)
    print(merged_data.head(3))
    
    print(f"\n💡 提示: 您可以使用Excel或其他数据分析工具打开 '{OUTPUT_FILE}' 查看完整数据")
    
    return merged_data


if __name__ == "__main__":
    # 运行前检查配置
    print("当前配置:")
    print(f"📁 金融数据文件夹: {FINANCE_DATA_FOLDERS}")
    print(f"💾 输出文件: {OUTPUT_FILE}")
    print(f"📊 支持格式: {SUPPORTED_FORMATS}")
    print(f"🔬 智能分析: {'启用' if ENABLE_ANALYSIS else '禁用'}")
    
    # 等待用户确认或自动运行
    if AUTO_RUN:
        print("\n🚀 自动运行模式已启用，开始合并...")
        
        # 1. 数据合并
        merged_data = merge_finance_data()
        
        # 2. 智能分析
        if merged_data is not None and ENABLE_ANALYSIS:
            analyzer = FinanceDataAnalyzer(merged_data)
            analysis_results = analyzer.run_full_analysis()
            
            print(f"\n📋 分析完成！生成了以下文件:")
            print(f"   📊 数据文件: {OUTPUT_FILE}")
            if DETAILED_REPORT:
                print(f"   📝 分析报告: {ANALYSIS_REPORT}")
            if HAS_PLOTTING and SAVE_PLOTS:
                print(f"   📈 可视化图表: 金融数据相关性热图.png")
        
    else:
        confirm = input(f"\n是否使用当前配置开始合并金融数据? (y/n): ").lower().strip()
        
        if confirm == 'y':
            merged_data = merge_finance_data()
            
            if merged_data is not None and ENABLE_ANALYSIS:
                run_analysis = input("是否运行智能分析? (y/n): ").lower().strip() == 'y'
                if run_analysis:
                    analyzer = FinanceDataAnalyzer(merged_data)
                    analyzer.run_full_analysis()
        else:
            print("请修改脚本顶部的配置后重新运行")
            print("主要需要修改:")
            print("1. FINANCE_DATA_FOLDERS - 您的金融数据文件夹路径")
            print("2. OUTPUT_FILE - 输出文件名和路径")
            print("3. ENABLE_ANALYSIS - 是否启用智能分析功能") 