#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高级数据深度分析工具
提供专业的统计分析、趋势分析、相关性分析和异常检测
"""

import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# 统计分析库
from scipy import stats
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans

# 可视化库
try:
    import plotly.graph_objects as go
    import plotly.express as px
    import plotly.figure_factory as ff
    from plotly.subplots import make_subplots
    import plotly.offline as pyo
    HAS_PLOTLY = True
except ImportError:
    print("⚠️ 请安装plotly: pip install plotly")
    HAS_PLOTLY = False

try:
    import seaborn as sns
    import matplotlib.pyplot as plt
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# ====== 配置区域 ======

EXCEL_FILE = "../../data/数据合并结果_20250601_1703.xlsx"
OUTPUT_DIR = "../../output/analysis_results"
CHARTS_DIR = os.path.join(OUTPUT_DIR, "图表")

# ====== 配置区域结束 ======


class AdvancedDataAnalyzer:
    """高级数据深度分析器"""
    
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.data_cache = {}
        self.analysis_results = {}
        self.charts = []
        
        # 创建输出目录
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(CHARTS_DIR, exist_ok=True)
        
    def load_key_datasets(self):
        """加载关键数据集"""
        print("📖 加载关键数据集...")
        
        # 定义要深度分析的关键Sheet
        key_sheets = {
            'stock_index': '沪深300指数（2016-2018）',
            'stock_portfolio': '构建投资组合的五只股票数据（2016-2018）',
            'fund_data': '四只开放式股票型基金的净值（2016-2018年）',
            'lpr_rates': '贷款基础利率（LPR）数据',
            'shibor_rates': 'Shibor利率（2018年）',
            'bond_gdp': '债券存量规模与GDP（2010-2020年）',
            'stock_indices': '国内A股主要股指的日收盘数据（2014-2018）'
        }
        
        for key, sheet_name in key_sheets.items():
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
                
                self.data_cache[key] = df
                print(f"   ✅ {sheet_name}: {len(df)} 行, {len(df.columns)} 列")
                
            except Exception as e:
                print(f"   ❌ 加载失败: {sheet_name} - {str(e)}")
        
        print(f"✅ 成功加载 {len(self.data_cache)} 个关键数据集")
        return True
    
    def analyze_time_series_trends(self):
        """时间序列趋势分析"""
        print(f"\n📈 时间序列趋势深度分析")
        print("-" * 50)
        
        trend_analysis = {}
        
        # 分析沪深300指数趋势
        if 'stock_index' in self.data_cache:
            df = self.data_cache['stock_index']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                values = df[col].dropna()
                if len(values) > 30:  # 至少30个数据点
                    # 计算趋势指标
                    x = np.arange(len(values))
                    slope, intercept, r_value, p_value, std_err = stats.linregress(x, values)
                    
                    # 计算移动平均
                    ma_20 = values.rolling(window=20).mean()
                    ma_50 = values.rolling(window=50).mean() if len(values) > 50 else None
                    
                    # 计算波动率
                    returns = values.pct_change().dropna()
                    volatility = returns.std() * np.sqrt(252)  # 年化波动率
                    
                    trend_analysis[f'沪深300_{col}'] = {
                        'slope': slope,
                        'r_squared': r_value**2,
                        'p_value': p_value,
                        'volatility': volatility,
                        'trend_strength': abs(slope) * r_value**2,
                        'direction': 'upward' if slope > 0 else 'downward'
                    }
        
        # 分析利率趋势
        if 'lpr_rates' in self.data_cache:
            df = self.data_cache['lpr_rates']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                values = df[col].dropna()
                if len(values) > 10:
                    # 利率变化分析
                    rate_changes = values.diff().dropna()
                    
                    trend_analysis[f'LPR_{col}'] = {
                        'mean_rate': values.mean(),
                        'rate_range': values.max() - values.min(),
                        'volatility': values.std(),
                        'recent_trend': values.tail(10).mean() - values.head(10).mean(),
                        'change_frequency': (rate_changes != 0).sum()
                    }
        
        self.analysis_results['trends'] = trend_analysis
        
        # 打印关键发现
        print("关键趋势发现:")
        for metric, data in trend_analysis.items():
            if 'slope' in data:
                trend = "上升" if data['direction'] == 'upward' else "下降"
                strength = "强" if data['trend_strength'] > 0.5 else "弱"
                print(f"   📊 {metric}: {trend}趋势，强度{strength} (R²={data['r_squared']:.3f})")
        
        return trend_analysis
    
    def analyze_correlations(self):
        """相关性深度分析"""
        print(f"\n🔗 相关性深度分析")
        print("-" * 50)
        
        correlation_results = {}
        
        # 股票投资组合相关性分析
        if 'stock_portfolio' in self.data_cache:
            df = self.data_cache['stock_portfolio']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            if len(numeric_cols) >= 2:
                # 计算相关性矩阵
                corr_matrix = df[numeric_cols].corr()
                
                # 找出高相关性对
                high_corr_pairs = []
                for i in range(len(corr_matrix.columns)):
                    for j in range(i+1, len(corr_matrix.columns)):
                        corr_val = corr_matrix.iloc[i, j]
                        if abs(corr_val) > 0.7:
                            high_corr_pairs.append({
                                'pair': f"{corr_matrix.columns[i]} - {corr_matrix.columns[j]}",
                                'correlation': corr_val,
                                'strength': 'strong' if abs(corr_val) > 0.8 else 'moderate'
                            })
                
                correlation_results['portfolio_correlations'] = {
                    'matrix': corr_matrix,
                    'high_correlations': high_corr_pairs
                }
        
        # 基金间相关性分析
        if 'fund_data' in self.data_cache:
            df = self.data_cache['fund_data']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            if len(numeric_cols) >= 2:
                # 计算收益率相关性
                returns_df = df[numeric_cols].pct_change().dropna()
                returns_corr = returns_df.corr()
                
                correlation_results['fund_correlations'] = returns_corr
        
        self.analysis_results['correlations'] = correlation_results
        
        # 打印相关性发现
        if 'portfolio_correlations' in correlation_results:
            high_corr = correlation_results['portfolio_correlations']['high_correlations']
            print(f"发现 {len(high_corr)} 个高相关性股票对:")
            for pair_info in high_corr[:5]:
                print(f"   🔗 {pair_info['pair']}: {pair_info['correlation']:.3f} ({pair_info['strength']})")
        
        return correlation_results
    
    def analyze_risk_metrics(self):
        """风险指标分析"""
        print(f"\n⚠️ 风险指标深度分析")
        print("-" * 50)
        
        risk_analysis = {}
        
        # 股票风险分析
        if 'stock_portfolio' in self.data_cache:
            df = self.data_cache['stock_portfolio']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            stock_risks = {}
            for col in numeric_cols:
                values = df[col].dropna()
                if len(values) > 30:
                    returns = values.pct_change().dropna()
                    
                    # 计算风险指标
                    stock_risks[col] = {
                        'volatility': returns.std() * np.sqrt(252),  # 年化波动率
                        'var_95': np.percentile(returns, 5),  # 95% VaR
                        'var_99': np.percentile(returns, 1),  # 99% VaR
                        'max_drawdown': self._calculate_max_drawdown(values),
                        'sharpe_ratio': self._calculate_sharpe_ratio(returns),
                        'skewness': returns.skew(),
                        'kurtosis': returns.kurtosis()
                    }
            
            risk_analysis['stock_risks'] = stock_risks
        
        # 基金风险分析
        if 'fund_data' in self.data_cache:
            df = self.data_cache['fund_data']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            fund_risks = {}
            for col in numeric_cols:
                values = df[col].dropna()
                if len(values) > 30:
                    returns = values.pct_change().dropna()
                    
                    # 找一个基准列（第一个数值列）
                    benchmark_col = None
                    for benchmark_candidate in numeric_cols:
                        if benchmark_candidate != col:
                            benchmark_values = df[benchmark_candidate].dropna()
                            if len(benchmark_values) > 30:
                                benchmark_col = benchmark_candidate
                                break
                    
                    beta_value = None
                    if benchmark_col:
                        benchmark_returns = df[benchmark_col].pct_change().dropna()
                        # 确保长度匹配
                        min_length = min(len(returns), len(benchmark_returns))
                        if min_length > 10:
                            beta_value = self._calculate_beta(
                                returns.iloc[:min_length], 
                                benchmark_returns.iloc[:min_length]
                            )
                    
                    fund_risks[col] = {
                        'volatility': returns.std() * np.sqrt(252),
                        'tracking_error': returns.std() * np.sqrt(252),
                        'max_drawdown': self._calculate_max_drawdown(values),
                        'calmar_ratio': self._calculate_calmar_ratio(values, returns),
                        'beta': beta_value
                    }
            
            risk_analysis['fund_risks'] = fund_risks
        
        self.analysis_results['risks'] = risk_analysis
        
        # 打印风险发现
        print("关键风险指标:")
        if 'stock_risks' in risk_analysis:
            for stock, metrics in list(risk_analysis['stock_risks'].items())[:3]:
                vol = metrics['volatility'] * 100
                sr = metrics['sharpe_ratio']
                print(f"   📊 {stock}: 波动率 {vol:.1f}%, 夏普比率 {sr:.2f}")
        
        return risk_analysis
    
    def detect_anomalies(self):
        """异常值检测"""
        print(f"\n🔍 异常值检测分析")
        print("-" * 50)
        
        anomaly_results = {}
        
        # 对主要数据集进行异常检测
        for dataset_name, df in self.data_cache.items():
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            dataset_anomalies = {}
            
            for col in numeric_cols:
                values = df[col].dropna()
                if len(values) > 10:
                    # Z-score异常检测
                    z_scores = np.abs(stats.zscore(values))
                    z_anomalies = (z_scores > 3).sum()
                    
                    # IQR异常检测
                    Q1 = values.quantile(0.25)
                    Q3 = values.quantile(0.75)
                    IQR = Q3 - Q1
                    lower_bound = Q1 - 1.5 * IQR
                    upper_bound = Q3 + 1.5 * IQR
                    iqr_anomalies = ((values < lower_bound) | (values > upper_bound)).sum()
                    
                    dataset_anomalies[col] = {
                        'z_score_anomalies': z_anomalies,
                        'iqr_anomalies': iqr_anomalies,
                        'anomaly_percentage': (iqr_anomalies / len(values)) * 100
                    }
            
            anomaly_results[dataset_name] = dataset_anomalies
        
        self.analysis_results['anomalies'] = anomaly_results
        
        # 打印异常检测结果
        print("异常值检测结果:")
        for dataset, anomalies in anomaly_results.items():
            total_anomalies = sum(info['iqr_anomalies'] for info in anomalies.values())
            if total_anomalies > 0:
                print(f"   ⚠️ {dataset}: 发现 {total_anomalies} 个异常值")
        
        return anomaly_results
    
    def perform_clustering_analysis(self):
        """聚类分析"""
        print(f"\n🎯 聚类分析")
        print("-" * 50)
        
        clustering_results = {}
        
        # 对股票投资组合进行聚类
        if 'stock_portfolio' in self.data_cache:
            df = self.data_cache['stock_portfolio']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            if len(numeric_cols) >= 2:
                # 准备数据
                data = df[numeric_cols].dropna()
                
                if len(data) > 10:
                    # 标准化
                    scaler = StandardScaler()
                    scaled_data = scaler.fit_transform(data)
                    
                    # K-means聚类
                    kmeans = KMeans(n_clusters=3, random_state=42)
                    clusters = kmeans.fit_predict(scaled_data)
                    
                    # 分析聚类结果
                    cluster_analysis = {}
                    for i in range(3):
                        cluster_mask = clusters == i
                        cluster_data = data[cluster_mask]
                        cluster_analysis[f'cluster_{i}'] = {
                            'size': cluster_mask.sum(),
                            'mean_values': cluster_data.mean().to_dict(),
                            'characteristics': self._analyze_cluster_characteristics(cluster_data)
                        }
                    
                    clustering_results['stock_clusters'] = cluster_analysis
        
        self.analysis_results['clustering'] = clustering_results
        
        # 打印聚类结果
        if 'stock_clusters' in clustering_results:
            print("股票聚类分析结果:")
            for cluster_id, info in clustering_results['stock_clusters'].items():
                print(f"   📊 {cluster_id}: {info['size']} 个数据点")
        
        return clustering_results
    
    def _calculate_max_drawdown(self, values):
        """计算最大回撤"""
        peak = values.expanding().max()
        drawdown = (values - peak) / peak
        return drawdown.min()
    
    def _calculate_sharpe_ratio(self, returns, risk_free_rate=0.02):
        """计算夏普比率"""
        excess_returns = returns - risk_free_rate / 252
        if excess_returns.std() == 0:
            return 0
        return excess_returns.mean() / excess_returns.std() * np.sqrt(252)
    
    def _calculate_calmar_ratio(self, values, returns):
        """计算卡玛比率"""
        annual_return = (values.iloc[-1] / values.iloc[0]) ** (252 / len(values)) - 1
        max_dd = abs(self._calculate_max_drawdown(values))
        return annual_return / max_dd if max_dd != 0 else 0
    
    def _calculate_beta(self, stock_returns, market_returns):
        """计算Beta系数"""
        if len(stock_returns) == len(market_returns):
            covariance = np.cov(stock_returns, market_returns)[0][1]
            market_variance = np.var(market_returns)
            return covariance / market_variance if market_variance != 0 else 0
        return None
    
    def _analyze_cluster_characteristics(self, cluster_data):
        """分析聚类特征"""
        characteristics = {}
        for col in cluster_data.columns:
            values = cluster_data[col]
            characteristics[col] = {
                'mean': values.mean(),
                'std': values.std(),
                'volatility_level': 'high' if values.std() > values.mean() * 0.1 else 'low'
            }
        return characteristics
    
    def generate_insights_summary(self):
        """生成深度洞察总结"""
        print(f"\n🎯 生成深度分析洞察")
        print("-" * 50)
        
        insights = []
        
        # 趋势洞察
        if 'trends' in self.analysis_results:
            trends = self.analysis_results['trends']
            strong_trends = [k for k, v in trends.items() if v.get('trend_strength', 0) > 0.5]
            if strong_trends:
                insights.append(f"发现 {len(strong_trends)} 个强趋势指标")
        
        # 相关性洞察
        if 'correlations' in self.analysis_results:
            corr = self.analysis_results['correlations']
            if 'portfolio_correlations' in corr:
                high_corr_count = len(corr['portfolio_correlations']['high_correlations'])
                if high_corr_count > 0:
                    insights.append(f"投资组合中发现 {high_corr_count} 对高相关性资产")
        
        # 风险洞察
        if 'risks' in self.analysis_results:
            risks = self.analysis_results['risks']
            if 'stock_risks' in risks:
                high_risk_assets = [k for k, v in risks['stock_risks'].items() if v.get('volatility', 0) > 0.3]
                insights.append(f"识别出 {len(high_risk_assets)} 个高风险资产")
        
        # 异常值洞察
        if 'anomalies' in self.analysis_results:
            anomalies = self.analysis_results['anomalies']
            total_anomalies = sum(
                sum(col_info['iqr_anomalies'] for col_info in dataset.values())
                for dataset in anomalies.values()
            )
            if total_anomalies > 0:
                insights.append(f"检测到 {total_anomalies} 个潜在异常值")
        
        print("关键洞察:")
        for insight in insights:
            print(f"   💡 {insight}")
        
        # 保存洞察报告
        insights_file = os.path.join(OUTPUT_DIR, f"深度分析洞察_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")
        with open(insights_file, 'w', encoding='utf-8') as f:
            f.write("深度数据分析洞察报告\n")
            f.write("=" * 50 + "\n")
            f.write(f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            for category, results in self.analysis_results.items():
                f.write(f"{category.upper()} 分析结果:\n")
                f.write(f"{str(results)}\n\n")
        
        print(f"✅ 深度分析洞察已保存到: {insights_file}")
        return insights
    
    def run_advanced_analysis(self):
        """运行高级分析"""
        print("🚀 开始高级数据深度分析")
        print("=" * 60)
        
        # 1. 加载数据
        if not self.load_key_datasets():
            return False
        
        # 2. 时间序列趋势分析
        self.analyze_time_series_trends()
        
        # 3. 相关性分析
        self.analyze_correlations()
        
        # 4. 风险指标分析
        self.analyze_risk_metrics()
        
        # 5. 异常值检测
        self.detect_anomalies()
        
        # 6. 聚类分析
        self.perform_clustering_analysis()
        
        # 7. 生成洞察总结
        insights = self.generate_insights_summary()
        
        print(f"\n🎉 高级分析完成!")
        print(f"📁 结果保存在: {OUTPUT_DIR}")
        
        return True


def main():
    """主函数"""
    print("📊 高级数据深度分析工具")
    print("=" * 40)
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ 文件不存在: {EXCEL_FILE}")
        return
    
    analyzer = AdvancedDataAnalyzer(EXCEL_FILE)
    analyzer.run_advanced_analysis()


if __name__ == "__main__":
    main() 