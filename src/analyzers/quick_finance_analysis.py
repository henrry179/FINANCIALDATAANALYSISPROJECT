#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
大规模金融数据快速分析脚本
专门处理大型合并数据集的快速智能分析
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 配置
DATA_FILE = "金融数据汇总_20250601_1650.xlsx"  # 输入数据文件
SAMPLE_SIZE = 10000  # 采样大小，用于快速分析

def quick_analysis():
    """快速分析大规模金融数据"""
    
    print("🚀 启动大规模金融数据快速分析")
    print("=" * 50)
    
    # 1. 数据加载
    print("📖 正在加载数据...")
    try:
        # 先读取小样本来了解数据结构
        sample_data = pd.read_excel(DATA_FILE, nrows=100)
        print(f"   ✅ 成功加载样本数据: {len(sample_data)} 行")
        
        # 读取完整数据
        print("   📊 正在加载完整数据集...")
        full_data = pd.read_excel(DATA_FILE)
        print(f"   ✅ 成功加载完整数据: {len(full_data):,} 行, {len(full_data.columns)} 列")
        
    except Exception as e:
        print(f"   ❌ 数据加载失败: {str(e)}")
        return
    
    # 2. 数据概览
    print(f"\n📊 数据集概览")
    print("-" * 40)
    print(f"总行数: {len(full_data):,}")
    print(f"总列数: {len(full_data.columns)}")
    
    # 内存使用情况
    memory_usage = full_data.memory_usage(deep=True).sum() / 1024 / 1024
    print(f"内存占用: {memory_usage:.1f} MB")
    
    # 3. 识别数据类型
    print(f"\n🔍 数据类型分析")
    print("-" * 40)
    
    numeric_cols = []
    date_cols = []
    text_cols = []
    
    for col in full_data.columns:
        col_name = str(col)
        if full_data[col].dtype in ['int64', 'float64']:
            if not col_name.startswith(('数据来源', '数据文件夹', '处理时间')):
                numeric_cols.append(col)
        elif '日期' in col_name or '时间' in col_name or 'date' in col_name.lower():
            date_cols.append(col)
        else:
            text_cols.append(col)
    
    print(f"数值列数量: {len(numeric_cols)}")
    print(f"日期列数量: {len(date_cols)}")
    print(f"文本列数量: {len(text_cols)}")
    
    # 4. 缺失值分析
    print(f"\n⚠️ 数据质量分析")
    print("-" * 40)
    
    missing_data = full_data.isnull().sum()
    missing_ratio = missing_data / len(full_data) * 100
    
    # 找出缺失值最多的列
    high_missing = missing_ratio[missing_ratio > 50].sort_values(ascending=False)
    print(f"缺失值超过50%的列数: {len(high_missing)}")
    
    if len(high_missing) > 0:
        print("缺失值最高的前5列:")
        for col in high_missing.head(5).index:
            print(f"   {str(col)[:50]}...: {missing_ratio[col]:.1f}%")
    
    # 5. 采样分析（针对数值数据）
    if len(numeric_cols) > 0:
        print(f"\n📈 数值数据分析（采样分析）")
        print("-" * 40)
        
        # 随机采样
        if len(full_data) > SAMPLE_SIZE:
            sample_data = full_data.sample(n=SAMPLE_SIZE, random_state=42)
            print(f"采用随机采样: {SAMPLE_SIZE:,} 行")
        else:
            sample_data = full_data
            print(f"使用全量数据: {len(sample_data):,} 行")
        
        # 分析前10个数值列
        numeric_sample = sample_data[numeric_cols[:10]]
        
        print(f"\n主要数值列统计:")
        for col in numeric_cols[:5]:
            if col in numeric_sample.columns:
                col_data = numeric_sample[col].dropna()
                if len(col_data) > 0:
                    print(f"   {str(col)[:30]}:")
                    print(f"      均值: {col_data.mean():.3f}")
                    print(f"      中位数: {col_data.median():.3f}")
                    print(f"      标准差: {col_data.std():.3f}")
                    print(f"      范围: [{col_data.min():.3f}, {col_data.max():.3f}]")
    
    # 6. 数据来源分析
    print(f"\n📁 数据来源分析")
    print("-" * 40)
    
    if '数据来源文件' in full_data.columns:
        source_counts = full_data['数据来源文件'].value_counts()
        print(f"数据来源文件数量: {len(source_counts)}")
        print(f"平均每文件行数: {len(full_data) / len(source_counts):.1f}")
        
        print("数据量最大的前5个文件:")
        for file_name, count in source_counts.head(5).items():
            print(f"   {str(file_name)[:50]}...: {count:,} 行")
    
    # 7. 时间跨度分析
    if len(date_cols) > 0:
        print(f"\n📅 时间数据分析")
        print("-" * 40)
        
        for date_col in date_cols[:3]:  # 分析前3个日期列
            try:
                date_data = pd.to_datetime(full_data[date_col], errors='coerce').dropna()
                if len(date_data) > 0:
                    print(f"{str(date_col)[:30]}:")
                    print(f"   时间跨度: {date_data.min()} 至 {date_data.max()}")
                    print(f"   有效日期数: {len(date_data):,}")
                    time_span = (date_data.max() - date_data.min()).days
                    print(f"   跨度天数: {time_span} 天")
            except Exception as e:
                print(f"   {str(date_col)[:30]}: 日期解析失败")
    
    # 8. 智能洞察
    print(f"\n💡 智能洞察")
    print("-" * 40)
    
    insights = []
    
    # 数据规模洞察
    if len(full_data) > 50000:
        insights.append("🎯 超大规模数据集，建议考虑分批处理或使用分布式计算")
    
    # 列数洞察
    if len(full_data.columns) > 500:
        insights.append("📊 列数量极多，建议进行特征选择和降维分析")
    
    # 缺失值洞察
    overall_missing = full_data.isnull().sum().sum() / (len(full_data) * len(full_data.columns))
    if overall_missing > 0.3:
        insights.append(f"⚠️ 整体缺失率高达 {overall_missing*100:.1f}%，建议数据清洗")
    
    # 数据类型洞察
    if len(numeric_cols) / len(full_data.columns) > 0.8:
        insights.append("📈 主要为数值型数据，适合进行量化分析和机器学习")
    
    # 数据来源洞察
    if '数据来源文件' in full_data.columns:
        unique_sources = full_data['数据来源文件'].nunique()
        if unique_sources > 100:
            insights.append(f"📁 数据来源多样化（{unique_sources}个文件），信息丰富度高")
    
    # 时间维度洞察
    if len(date_cols) > 0:
        insights.append("📅 包含时间维度，可进行时序分析和趋势预测")
    
    if not insights:
        insights.append("✅ 数据结构良好，可进行深度分析")
    
    for i, insight in enumerate(insights, 1):
        print(f"   {i}. {insight}")
    
    # 9. 分析建议
    print(f"\n🎯 分析建议")
    print("-" * 40)
    recommendations = [
        "1. 优先处理高缺失率列，考虑删除或插值",
        "2. 对数值变量进行标准化处理",
        "3. 利用时间维度进行趋势分析",
        "4. 考虑按数据来源进行分组分析",
        "5. 使用采样技术进行快速探索性分析",
        "6. 建立数据字典记录列含义",
        "7. 考虑使用降维技术处理高维数据"
    ]
    
    for rec in recommendations:
        print(f"   {rec}")
    
    # 10. 生成快速报告
    print(f"\n📝 生成分析报告")
    print("-" * 40)
    
    report_name = f"快速分析报告_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
    
    report_lines = [
        "=" * 60,
        "大规模金融数据快速分析报告",
        f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "=" * 60,
        "",
        "📊 数据概览",
        f"总行数: {len(full_data):,}",
        f"总列数: {len(full_data.columns)}",
        f"内存占用: {memory_usage:.1f} MB",
        "",
        "🔍 数据类型分布",
        f"数值列: {len(numeric_cols)}",
        f"日期列: {len(date_cols)}",
        f"文本列: {len(text_cols)}",
        "",
        "⚠️ 数据质量",
        f"整体缺失率: {overall_missing*100:.1f}%",
        f"高缺失列数: {len(high_missing)}",
        "",
        "💡 关键洞察",
    ]
    
    for insight in insights:
        report_lines.append(f"• {insight}")
    
    report_lines.extend([
        "",
        "🎯 分析建议",
    ])
    
    for rec in recommendations:
        report_lines.append(f"• {rec}")
    
    report_lines.extend([
        "",
        "=" * 60,
        "报告结束"
    ])
    
    # 保存报告
    try:
        with open(report_name, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_lines))
        print(f"   ✅ 快速分析报告已保存: {report_name}")
    except Exception as e:
        print(f"   ❌ 报告保存失败: {str(e)}")
    
    print(f"\n🎉 快速分析完成！")
    print(f"📊 数据文件: {DATA_FILE}")
    print(f"📝 分析报告: {report_name}")

if __name__ == "__main__":
    quick_analysis() 