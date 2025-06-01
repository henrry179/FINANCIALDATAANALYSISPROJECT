#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金融数据深度洞察分析工具
基于多Sheet数据分析结果，提供专门的金融数据深度洞察
"""

import os
import pandas as pd
import numpy as np
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
    HAS_PLOTTING = False

# ====== 配置区域 ======

# 输入文件
EXCEL_FILE = "../../data/数据合并结果_20250601_1703.xlsx"
OUTPUT_DIR = "金融洞察分析"

# ====== 配置区域结束 ======


class FinanceInsightsAnalyzer:
    """金融数据深度洞察分析器"""
    
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.insights = {}
        
        # 创建输出目录
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
    def analyze_stock_data(self):
        """分析股票相关数据"""
        print(f"\n📈 股票数据深度分析")
        print("-" * 50)
        
        # 读取关键股票数据Sheet
        stock_sheets = [
            "工商银行与沪深300指数",
            "构建投资组合的五只股票数据（2016-2018）",
            "沪深300指数（2016-2018）",
            "国内A股主要股指的日收盘数据（2014-2018）",
            "东方航空股票价格（2014-2018）"
        ]
        
        stock_analysis = {}
        
        for sheet_name in stock_sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                
                # 过滤元信息
                if '元信息' in df.columns or '文件信息' in df.columns:
                    meta_start = None
                    for idx, row in df.iterrows():
                        if '原始文件名' in str(row.values):
                            meta_start = idx
                            break
                    if meta_start is not None:
                        df = df.iloc[:meta_start]
                
                # 分析数值列
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                
                analysis = {
                    'sheet_name': sheet_name,
                    'total_records': len(df),
                    'numeric_columns': len(numeric_cols),
                    'time_span': self._estimate_time_span(df),
                    'price_volatility': self._calculate_volatility(df, numeric_cols),
                    'trend_analysis': self._analyze_trends(df, numeric_cols)
                }
                
                stock_analysis[sheet_name] = analysis
                print(f"   📊 {sheet_name}: {len(df)} 条记录, {len(numeric_cols)} 个数值指标")
                
            except Exception as e:
                print(f"   ❌ 分析失败: {sheet_name} - {str(e)}")
                continue
        
        self.insights['stock_analysis'] = stock_analysis
        return stock_analysis
    
    def analyze_bond_data(self):
        """分析债券相关数据"""
        print(f"\n💰 债券数据深度分析")
        print("-" * 50)
        
        # 读取债券相关Sheet
        bond_sheets = [
            "债券存量规模与GDP（2010-2018年）",
            "国内债券市场按照交易场所分类（2018年末）",
            "2020年末按照债券品种划分的债券余额情况",
            "债券存量规模与GDP（2010-2020年）",
            "2020年末存量债券的市场分布情况"
        ]
        
        bond_analysis = {}
        
        for sheet_name in bond_sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                
                # 过滤元信息
                if '元信息' in df.columns or '文件信息' in df.columns:
                    meta_start = None
                    for idx, row in df.iterrows():
                        if '原始文件名' in str(row.values):
                            meta_start = idx
                            break
                    if meta_start is not None:
                        df = df.iloc[:meta_start]
                
                # 分析债券市场规模和结构
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                
                analysis = {
                    'sheet_name': sheet_name,
                    'total_records': len(df),
                    'market_structure': self._analyze_market_structure(df),
                    'growth_trends': self._analyze_growth_trends(df, numeric_cols)
                }
                
                bond_analysis[sheet_name] = analysis
                print(f"   💳 {sheet_name}: {len(df)} 条记录")
                
            except Exception as e:
                print(f"   ❌ 分析失败: {sheet_name} - {str(e)}")
                continue
        
        self.insights['bond_analysis'] = bond_analysis
        return bond_analysis
    
    def analyze_interest_rate_data(self):
        """分析利率相关数据"""
        print(f"\n📊 利率数据深度分析")
        print("-" * 50)
        
        # 读取利率相关Sheet
        rate_sheets = [
            "贷款基础利率（LPR）数据",
            "银行间回购定盘利率（2018年）",
            "Shibor利率（2018年）",
            "银行间同业拆借利率（2018年）"
        ]
        
        rate_analysis = {}
        
        for sheet_name in rate_sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                
                # 过滤元信息
                if '元信息' in df.columns or '文件信息' in df.columns:
                    meta_start = None
                    for idx, row in df.iterrows():
                        if '原始文件名' in str(row.values):
                            meta_start = idx
                            break
                    if meta_start is not None:
                        df = df.iloc[:meta_start]
                
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                
                if len(numeric_cols) > 0:
                    # 计算利率统计指标
                    rate_stats = {}
                    for col in numeric_cols:
                        values = df[col].dropna()
                        if len(values) > 0:
                            rate_stats[col] = {
                                'mean': values.mean(),
                                'std': values.std(),
                                'min': values.min(),
                                'max': values.max(),
                                'latest': values.iloc[-1] if len(values) > 0 else None
                            }
                
                analysis = {
                    'sheet_name': sheet_name,
                    'total_records': len(df),
                    'rate_types': len(numeric_cols),
                    'rate_statistics': rate_stats,
                    'volatility_analysis': self._analyze_rate_volatility(df, numeric_cols)
                }
                
                rate_analysis[sheet_name] = analysis
                print(f"   📈 {sheet_name}: {len(df)} 条记录, {len(numeric_cols)} 种利率")
                
            except Exception as e:
                print(f"   ❌ 分析失败: {sheet_name} - {str(e)}")
                continue
        
        self.insights['rate_analysis'] = rate_analysis
        return rate_analysis
    
    def analyze_fund_data(self):
        """分析基金相关数据"""
        print(f"\n🏦 基金数据深度分析")
        print("-" * 50)
        
        # 读取基金相关Sheet
        fund_sheets = [
            "四只开放式股票型基金的净值（2016-2018年）",
            "国内4只开放式股票型基金净值数据（2018-2020）"
        ]
        
        fund_analysis = {}
        
        for sheet_name in fund_sheets:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                
                # 过滤元信息
                if '元信息' in df.columns or '文件信息' in df.columns:
                    meta_start = None
                    for idx, row in df.iterrows():
                        if '原始文件名' in str(row.values):
                            meta_start = idx
                            break
                    if meta_start is not None:
                        df = df.iloc[:meta_start]
                
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                
                # 计算基金表现指标
                fund_performance = {}
                for col in numeric_cols:
                    values = df[col].dropna()
                    if len(values) > 1:
                        # 计算收益率
                        returns = values.pct_change().dropna()
                        fund_performance[col] = {
                            'total_return': (values.iloc[-1] / values.iloc[0] - 1) * 100,
                            'volatility': returns.std() * np.sqrt(252) * 100,  # 年化波动率
                            'max_drawdown': self._calculate_max_drawdown(values),
                            'sharpe_ratio': self._calculate_sharpe_ratio(returns)
                        }
                
                analysis = {
                    'sheet_name': sheet_name,
                    'total_records': len(df),
                    'fund_count': len(numeric_cols),
                    'performance_metrics': fund_performance
                }
                
                fund_analysis[sheet_name] = analysis
                print(f"   🏦 {sheet_name}: {len(df)} 条记录, {len(numeric_cols)} 只基金")
                
            except Exception as e:
                print(f"   ❌ 分析失败: {sheet_name} - {str(e)}")
                continue
        
        self.insights['fund_analysis'] = fund_analysis
        return fund_analysis
    
    def _estimate_time_span(self, df):
        """估算数据时间跨度"""
        # 寻找可能的日期列
        date_cols = []
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in ['date', '日期', 'time', '时间']):
                date_cols.append(col)
        
        if date_cols:
            try:
                date_col = date_cols[0]
                dates = pd.to_datetime(df[date_col], errors='coerce').dropna()
                if len(dates) > 0:
                    return f"{dates.min().strftime('%Y-%m')} 至 {dates.max().strftime('%Y-%m')}"
            except:
                pass
        
        return f"约 {len(df)} 个数据点"
    
    def _calculate_volatility(self, df, numeric_cols):
        """计算价格波动率"""
        volatilities = {}
        for col in numeric_cols:
            values = df[col].dropna()
            if len(values) > 1:
                returns = values.pct_change().dropna()
                if len(returns) > 0:
                    volatilities[col] = returns.std() * 100
        return volatilities
    
    def _analyze_trends(self, df, numeric_cols):
        """分析趋势"""
        trends = {}
        for col in numeric_cols:
            values = df[col].dropna()
            if len(values) > 10:
                # 简单线性趋势分析
                x = np.arange(len(values))
                slope = np.polyfit(x, values, 1)[0]
                trends[col] = {
                    'slope': slope,
                    'direction': 'upward' if slope > 0 else 'downward',
                    'strength': abs(slope)
                }
        return trends
    
    def _analyze_market_structure(self, df):
        """分析市场结构"""
        # 寻找市场份额或规模相关的列
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            # 计算各项占比
            total_col = None
            for col in numeric_cols:
                if '总' in str(col) or '合计' in str(col):
                    total_col = col
                    break
            
            if total_col:
                total = df[total_col].sum()
                structure = {}
                for col in numeric_cols:
                    if col != total_col:
                        structure[col] = (df[col].sum() / total * 100) if total > 0 else 0
                return structure
        return {}
    
    def _analyze_growth_trends(self, df, numeric_cols):
        """分析增长趋势"""
        growth_trends = {}
        for col in numeric_cols:
            values = df[col].dropna()
            if len(values) > 1:
                # 计算年度增长率
                if len(values) >= 2:
                    growth_rate = (values.iloc[-1] / values.iloc[0]) ** (1/len(values)) - 1
                    growth_trends[col] = growth_rate * 100
        return growth_trends
    
    def _analyze_rate_volatility(self, df, numeric_cols):
        """分析利率波动性"""
        volatility_analysis = {}
        for col in numeric_cols:
            values = df[col].dropna()
            if len(values) > 1:
                volatility_analysis[col] = {
                    'std': values.std(),
                    'coefficient_of_variation': values.std() / values.mean() if values.mean() != 0 else 0,
                    'range': values.max() - values.min()
                }
        return volatility_analysis
    
    def _calculate_max_drawdown(self, values):
        """计算最大回撤"""
        peak = values.expanding().max()
        drawdown = (values - peak) / peak
        return drawdown.min() * 100
    
    def _calculate_sharpe_ratio(self, returns, risk_free_rate=0.02):
        """计算夏普比率"""
        excess_returns = returns - risk_free_rate / 252
        if excess_returns.std() != 0:
            return excess_returns.mean() / excess_returns.std() * np.sqrt(252)
        return 0
    
    def generate_insights_summary(self):
        """生成洞察总结"""
        print(f"\n🎯 生成金融洞察总结")
        print("-" * 50)
        
        summary_file = os.path.join(OUTPUT_DIR, f"金融洞察报告_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")
        
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(f"金融数据深度洞察报告\n")
            f.write(f"=" * 50 + "\n")
            f.write(f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"数据来源: {self.excel_file}\n\n")
            
            # 股票分析总结
            if 'stock_analysis' in self.insights:
                f.write(f"股票市场分析:\n")
                stock_data = self.insights['stock_analysis']
                f.write(f"  分析Sheet数: {len(stock_data)}\n")
                
                total_records = sum(data['total_records'] for data in stock_data.values())
                f.write(f"  总数据记录: {total_records:,}\n")
                
                # 找出数据量最大的股票数据
                largest_dataset = max(stock_data.values(), key=lambda x: x['total_records'])
                f.write(f"  最大数据集: {largest_dataset['sheet_name']} ({largest_dataset['total_records']:,} 条记录)\n\n")
            
            # 债券分析总结
            if 'bond_analysis' in self.insights:
                f.write(f"债券市场分析:\n")
                bond_data = self.insights['bond_analysis']
                f.write(f"  分析Sheet数: {len(bond_data)}\n")
                f.write(f"  涵盖债券市场结构、规模和增长趋势\n\n")
            
            # 利率分析总结
            if 'rate_analysis' in self.insights:
                f.write(f"利率市场分析:\n")
                rate_data = self.insights['rate_analysis']
                f.write(f"  分析Sheet数: {len(rate_data)}\n")
                
                # 统计利率品种
                total_rate_types = sum(data['rate_types'] for data in rate_data.values())
                f.write(f"  利率品种总数: {total_rate_types}\n")
                
                # LPR数据特别分析
                if '贷款基础利率（LPR）数据' in [data['sheet_name'] for data in rate_data.values()]:
                    f.write(f"  包含LPR历史数据，支持利率政策分析\n\n")
            
            # 基金分析总结
            if 'fund_analysis' in self.insights:
                f.write(f"基金市场分析:\n")
                fund_data = self.insights['fund_analysis']
                f.write(f"  分析Sheet数: {len(fund_data)}\n")
                
                total_funds = sum(data['fund_count'] for data in fund_data.values())
                f.write(f"  基金总数: {total_funds}\n")
                f.write(f"  包含收益率、波动率、夏普比率等关键指标\n\n")
            
            # 关键洞察
            f.write(f"关键洞察:\n")
            f.write(f"  1. 数据集涵盖股票、债券、利率、基金四大金融市场\n")
            f.write(f"  2. 时间跨度从2010年至2020年，具有良好的历史覆盖\n")
            f.write(f"  3. 包含沪深300、上证180等主要股指数据\n")
            f.write(f"  4. 利率数据包含LPR、Shibor、银行间拆借等关键品种\n")
            f.write(f"  5. 适合进行多资产配置、风险管理和政策影响分析\n\n")
            
            # 应用建议
            f.write(f"应用建议:\n")
            f.write(f"  1. 股票分析: 可进行指数跟踪、个股表现和行业分析\n")
            f.write(f"  2. 债券分析: 支持债券市场结构和利率环境研究\n")
            f.write(f"  3. 基金分析: 适合基金业绩评估和风险收益分析\n")
            f.write(f"  4. 宏观分析: 可研究货币政策对各类资产的影响\n")
            f.write(f"  5. 投资组合: 支持多资产配置和风险管理决策\n")
        
        print(f"✅ 金融洞察报告已保存到: {summary_file}")
        return summary_file
    
    def run_full_analysis(self):
        """运行完整的金融洞察分析"""
        print(f"🚀 开始金融数据深度洞察分析")
        print(f"分析文件: {self.excel_file}")
        print("=" * 60)
        
        try:
            # 1. 股票数据分析
            self.analyze_stock_data()
            
            # 2. 债券数据分析
            self.analyze_bond_data()
            
            # 3. 利率数据分析
            self.analyze_interest_rate_data()
            
            # 4. 基金数据分析
            self.analyze_fund_data()
            
            # 5. 生成洞察总结
            report_file = self.generate_insights_summary()
            
            print(f"\n🎉 金融洞察分析完成!")
            print(f"📋 详细报告: {report_file}")
            
            return True
            
        except Exception as e:
            print(f"❌ 分析过程中出现错误: {str(e)}")
            return False


def main():
    """主函数"""
    print("📊 金融数据深度洞察分析工具")
    print("=" * 40)
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ 文件不存在: {EXCEL_FILE}")
        return
    
    # 创建分析器并运行分析
    analyzer = FinanceInsightsAnalyzer(EXCEL_FILE)
    success = analyzer.run_full_analysis()
    
    if success:
        print(f"\n💡 金融洞察分析结果已保存到 '{OUTPUT_DIR}' 文件夹")
    else:
        print(f"\n❌ 分析过程中出现错误")


if __name__ == "__main__":
    main() 