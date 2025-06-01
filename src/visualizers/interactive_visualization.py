#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交互式数据可视化网页生成器
生成包含多种交互式图表的HTML网页
"""

import os
import pandas as pd
import numpy as np
from datetime import datetime
import json
import warnings
warnings.filterwarnings('ignore')

# 可视化库
try:
    import plotly.graph_objects as go
    import plotly.express as px
    import plotly.figure_factory as ff
    from plotly.subplots import make_subplots
    import plotly.offline as pyo
    from plotly.graph_objs import *
    HAS_PLOTLY = True
except ImportError:
    print("❌ 请安装plotly: pip install plotly")
    HAS_PLOTLY = False

# ====== 配置区域 ======

EXCEL_FILE = "data/数据合并结果_20250601_1703.xlsx"
OUTPUT_DIR = "../../output/visualizations"
HTML_FILE = "金融数据交互分析仪表板.html"

# ====== 配置区域结束 ======


class InteractiveVisualization:
    """交互式可视化生成器"""
    
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.data_cache = {}
        self.figures = {}
        
        # 创建输出目录
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
    def load_visualization_data(self):
        """加载用于可视化的数据（扩展数据表）"""
        print("📊 加载可视化数据...")
        viz_sheets = {
            'stock_index': '沪深300指数（2016-2018）',
            'stock_portfolio': '构建投资组合的五只股票数据（2016-2018）',
            'fund_performance': '四只开放式股票型基金的净值（2016-2018年）',
            'shibor_rates': 'Shibor利率（2018年）',
            'bond_market': '债券存量规模与GDP（2010-2020年）',
            'stock_major_indices': '国内A股主要股指的日收盘数据（2014-2018）',
            'bank_rates': '银行间同业拆借利率（2018年）',
            'macro_economy': '宏观经济指标',
            'market_cap': '市值分布',
            'industry': '行业分布',
            'asset_allocation': '资产配置',
            'cashflow': '现金流量'
        }
        for key, sheet_name in viz_sheets.items():
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                if '元信息' in df.columns or '文件信息' in df.columns:
                    meta_start = None
                    for idx, row in df.iterrows():
                        if '原始文件名' in str(row.values):
                            meta_start = idx
                            break
                    if meta_start is not None:
                        df = df.iloc[:meta_start]
                self.data_cache[key] = df
                print(f"   ✅ {sheet_name}: {len(df)} 行数据")
            except Exception as e:
                print(f"   ⚠️ 跳过: {sheet_name} - {str(e)}")
        print(f"✅ 成功加载 {len(self.data_cache)} 个数据集")
        return True
    
    def create_stock_index_chart(self):
        """创建股票指数趋势图"""
        if 'stock_index' not in self.data_cache:
            return None
        
        df = self.data_cache['stock_index']
        
        # 创建子图
        fig = make_subplots(
            rows=2, cols=1,
            subplot_titles=('沪深300指数价格趋势', '成交量趋势'),
            vertical_spacing=0.1,
            shared_xaxes=True
        )
        
        # 寻找价格和成交量列
        price_cols = []
        volume_cols = []
        
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['价格', 'price', '收盘', 'close']):
                price_cols.append(col)
            elif any(keyword in col_lower for keyword in ['成交量', 'volume', '交易量']):
                volume_cols.append(col)
        
        # 绘制价格趋势
        for i, col in enumerate(price_cols[:3]):  # 最多显示3个价格序列
            values = df[col].dropna()
            fig.add_trace(
                go.Scatter(
                    x=list(range(len(values))),
                    y=values,
                    mode='lines',
                    name=f'{col}',
                    line=dict(width=2),
                    hovertemplate=f'{col}: %{{y:.2f}}<br>日期: %{{x}}<extra></extra>'
                ),
                row=1, col=1
            )
        
        # 绘制成交量
        for col in volume_cols[:1]:  # 只显示一个成交量序列
            values = df[col].dropna()
            fig.add_trace(
                go.Bar(
                    x=list(range(len(values))),
                    y=values,
                    name=f'{col}',
                    marker_color='lightblue',
                    opacity=0.7
                ),
                row=2, col=1
            )
        
        fig.update_layout(
            title={
                'text': '沪深300指数分析仪表板',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 20}
            },
            height=600,
            showlegend=True,
            template='plotly_white'
        )
        
        fig.update_xaxes(title_text="时间序列", row=2, col=1)
        fig.update_yaxes(title_text="价格", row=1, col=1)
        fig.update_yaxes(title_text="成交量", row=2, col=1)
        
        self.figures['stock_index'] = fig
        return fig
    
    def create_portfolio_performance_chart(self):
        """创建投资组合表现图"""
        if 'stock_portfolio' not in self.data_cache:
            return None
        
        df = self.data_cache['stock_portfolio']
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) == 0:
            return None
        
        # 创建多个子图
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=('投资组合价格走势', '收益率分布', '累计收益率', '风险收益散点图'),
            specs=[[{"secondary_y": False}, {"secondary_y": False}],
                   [{"secondary_y": False}, {"secondary_y": False}]]
        )
        
        # 1. 价格走势图
        for i, col in enumerate(numeric_cols[:4]):  # 最多4只股票
            values = df[col].dropna()
            # 标准化处理，以便比较
            normalized_values = (values / values.iloc[0]) * 100
            
            fig.add_trace(
                go.Scatter(
                    x=list(range(len(normalized_values))),
                    y=normalized_values,
                    mode='lines',
                    name=f'{col}',
                    line=dict(width=2),
                    hovertemplate=f'{col}<br>标准化价格: %{{y:.2f}}<extra></extra>'
                ),
                row=1, col=1
            )
        
        # 2. 收益率分布
        returns_data = []
        for col in numeric_cols[:4]:
            values = df[col].dropna()
            returns = values.pct_change().dropna() * 100
            returns_data.extend(returns.tolist())
        
        if returns_data:
            fig.add_trace(
                go.Histogram(
                    x=returns_data,
                    nbinsx=30,
                    name='收益率分布',
                    marker_color='lightgreen',
                    opacity=0.7
                ),
                row=1, col=2
            )
        
        # 3. 累计收益率
        for col in numeric_cols[:4]:
            values = df[col].dropna()
            returns = values.pct_change().fillna(0)
            cumulative_returns = (1 + returns).cumprod() - 1
            
            fig.add_trace(
                go.Scatter(
                    x=list(range(len(cumulative_returns))),
                    y=cumulative_returns * 100,
                    mode='lines',
                    name=f'{col} 累计收益',
                    line=dict(width=2)
                ),
                row=2, col=1
            )
        
        # 4. 风险收益散点图
        risk_return_data = []
        for col in numeric_cols[:4]:
            values = df[col].dropna()
            returns = values.pct_change().dropna()
            avg_return = returns.mean() * 252 * 100  # 年化收益率
            volatility = returns.std() * np.sqrt(252) * 100  # 年化波动率
            risk_return_data.append((volatility, avg_return, col))
        
        if risk_return_data:
            volatilities, avg_returns, names = zip(*risk_return_data)
            fig.add_trace(
                go.Scatter(
                    x=volatilities,
                    y=avg_returns,
                    mode='markers+text',
                    text=names,
                    textposition="top center",
                    marker=dict(size=12, color='red'),
                    name='风险收益',
                    hovertemplate='波动率: %{x:.2f}%<br>年化收益率: %{y:.2f}%<extra></extra>'
                ),
                row=2, col=2
            )
        
        fig.update_layout(
            title={
                'text': '投资组合综合分析',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            height=800,
            showlegend=True,
            template='plotly_white'
        )
        
        self.figures['portfolio'] = fig
        return fig
    
    def create_fund_comparison_chart(self):
        """创建基金对比图"""
        if 'fund_performance' not in self.data_cache:
            return None
        
        df = self.data_cache['fund_performance']
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) == 0:
            return None
        
        # 创建雷达图比较基金表现
        categories = ['收益率', '夏普比率', '最大回撤', '波动率', '稳定性']
        
        fig = go.Figure()
        
        colors = ['red', 'blue', 'green', 'orange']
        
        for i, col in enumerate(numeric_cols[:4]):
            values = df[col].dropna()
            if len(values) > 30:
                returns = values.pct_change().dropna()
                
                # 计算指标
                annual_return = (values.iloc[-1] / values.iloc[0]) ** (252 / len(values)) - 1
                sharpe_ratio = self._calculate_sharpe_ratio(returns)
                max_drawdown = abs(self._calculate_max_drawdown(values))
                volatility = returns.std() * np.sqrt(252)
                stability = 1 / (returns.std() + 0.001)  # 稳定性指标
                
                # 标准化指标 (0-100)
                metrics = [
                    max(0, min(100, annual_return * 100 + 50)),  # 收益率
                    max(0, min(100, (sharpe_ratio + 2) * 25)),   # 夏普比率
                    max(0, min(100, (1 - max_drawdown) * 100)),  # 最大回撤 (反向)
                    max(0, min(100, (1 - volatility) * 100)),    # 波动率 (反向)
                    max(0, min(100, stability * 20))             # 稳定性
                ]
                
                fig.add_trace(go.Scatterpolar(
                    r=metrics,
                    theta=categories,
                    fill='toself',
                    name=col,
                    line_color=colors[i % len(colors)]
                ))
        
        fig.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 100]
                )),
            title={
                'text': '基金表现雷达图',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            height=500,
            template='plotly_white'
        )
        
        self.figures['fund_radar'] = fig
        return fig
    
    def create_interest_rates_chart(self):
        """创建利率走势图"""
        if 'shibor_rates' not in self.data_cache:
            return None
        
        df = self.data_cache['shibor_rates']
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) == 0:
            return None
        
        fig = go.Figure()
        
        # 添加不同期限的利率曲线
        colors = px.colors.qualitative.Set1
        
        for i, col in enumerate(numeric_cols):
            values = df[col].dropna()
            fig.add_trace(go.Scatter(
                x=list(range(len(values))),
                y=values,
                mode='lines',
                name=col,
                line=dict(width=2, color=colors[i % len(colors)]),
                hovertemplate=f'{col}: %{{y:.4f}}%<br>日期: %{{x}}<extra></extra>'
            ))
        
        fig.update_layout(
            title={
                'text': 'Shibor利率走势分析',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            xaxis_title='时间序列',
            yaxis_title='利率 (%)',
            height=500,
            template='plotly_white',
            hovermode='x unified'
        )
        
        self.figures['interest_rates'] = fig
        return fig
    
    def create_correlation_heatmap(self):
        """创建相关性热力图"""
        if 'stock_portfolio' not in self.data_cache:
            return None
        
        df = self.data_cache['stock_portfolio']
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) < 2:
            return None
        
        # 计算相关性矩阵
        correlation_matrix = df[numeric_cols].corr()
        
        # 创建热力图
        fig = go.Figure(data=go.Heatmap(
            z=correlation_matrix.values,
            x=correlation_matrix.columns,
            y=correlation_matrix.columns,
            colorscale='RdBu',
            zmid=0,
            text=np.round(correlation_matrix.values, 3),
            texttemplate="%{text}",
            textfont={"size": 10},
            hovertemplate='%{x} vs %{y}<br>相关系数: %{z:.3f}<extra></extra>'
        ))
        
        fig.update_layout(
            title={
                'text': '投资组合相关性分析',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            height=500,
            template='plotly_white'
        )
        
        self.figures['correlation'] = fig
        return fig
    
    def create_bond_market_chart(self):
        """创建债券市场分析图"""
        if 'bond_market' not in self.data_cache:
            return None
        
        df = self.data_cache['bond_market']
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) == 0:
            return None
        
        # 创建双Y轴图表
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        # 添加债券规模
        for col in numeric_cols:
            if '债券' in str(col) or 'bond' in str(col).lower():
                values = df[col].dropna()
                fig.add_trace(
                    go.Scatter(
                        x=list(range(len(values))),
                        y=values,
                        mode='lines+markers',
                        name=col,
                        line=dict(width=3),
                        marker=dict(size=6)
                    ),
                    secondary_y=False,
                )
                break
        
        # 添加GDP数据
        for col in numeric_cols:
            if 'GDP' in str(col) or 'gdp' in str(col).lower():
                values = df[col].dropna()
                fig.add_trace(
                    go.Scatter(
                        x=list(range(len(values))),
                        y=values,
                        mode='lines+markers',
                        name=col,
                        line=dict(width=3, dash='dash'),
                        marker=dict(size=6)
                    ),
                    secondary_y=True,
                )
                break
        
        fig.update_xaxes(title_text="年份")
        fig.update_yaxes(title_text="债券存量规模", secondary_y=False)
        fig.update_yaxes(title_text="GDP", secondary_y=True)
        
        fig.update_layout(
            title={
                'text': '债券市场规模与GDP关系',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            height=500,
            template='plotly_white'
        )
        
        self.figures['bond_market'] = fig
        return fig
    
    def create_candlestick_chart(self):
        """创建K线图（蜡烛图）"""
        if 'stock_portfolio' not in self.data_cache:
            return None
        df = self.data_cache['stock_portfolio']
        if not all(col in df.columns for col in ['日期', 'open', 'high', 'low', 'close']):
            return None
        fig = go.Figure(data=[go.Candlestick(
            x=df['日期'],
            open=df['open'],
            high=df['high'],
            low=df['low'],
            close=df['close'],
            name='K线图'
        )])
        fig.update_layout(title='股票K线图', xaxis_title='日期', yaxis_title='价格', template='plotly_white')
        self.figures['candlestick'] = fig
        return fig

    def create_return_histogram(self):
        """创建收益率分布直方图"""
        if 'stock_portfolio' not in self.data_cache:
            return None
        df = self.data_cache['stock_portfolio']
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            return None
        returns = df[numeric_cols[0]].pct_change().dropna()
        fig = go.Figure(data=[go.Histogram(x=returns, nbinsx=50)])
        fig.update_layout(title='收益率分布直方图', xaxis_title='收益率', yaxis_title='频数', template='plotly_white')
        self.figures['return_histogram'] = fig
        return fig
    
    def _calculate_sharpe_ratio(self, returns, risk_free_rate=0.02):
        """计算夏普比率"""
        excess_returns = returns - risk_free_rate / 252
        if excess_returns.std() == 0:
            return 0
        return excess_returns.mean() / excess_returns.std() * np.sqrt(252)
    
    def _calculate_max_drawdown(self, values):
        """计算最大回撤"""
        peak = values.expanding().max()
        drawdown = (values - peak) / peak
        return drawdown.min()
    
    def create_macro_economic_chart(self):
        """创建宏观经济指标趋势图"""
        if 'macro_economy' not in self.data_cache:
            return None
        df = self.data_cache['macro_economy']
        fig = go.Figure()
        for col in df.columns[1:]:
            fig.add_trace(go.Scatter(x=df[df.columns[0]], y=df[col], mode='lines+markers', name=col))
        fig.update_layout(title='宏观经济指标趋势', xaxis_title='时间', yaxis_title='指标值', template='plotly_white')
        self.figures['macro_economy'] = fig
        return fig

    def create_market_cap_pie(self):
        """创建市值分布饼图"""
        if 'market_cap' not in self.data_cache:
            return None
        df = self.data_cache['market_cap']
        values = df.groupby('市值类型')['市值'].sum()
        fig = go.Figure(data=[go.Pie(labels=values.index, values=values.values, hole=0.4)])
        fig.update_layout(title='不同市值类型占比', template='plotly_white')
        self.figures['market_cap_pie'] = fig
        return fig

    def create_industry_bar(self):
        """创建行业分布柱状图"""
        if 'industry' not in self.data_cache:
            return None
        df = self.data_cache['industry']
        counts = df['行业'].value_counts()
        fig = go.Figure(data=[go.Bar(x=counts.index, y=counts.values)])
        fig.update_layout(title='行业分布', xaxis_title='行业', yaxis_title='公司数量', template='plotly_white')
        self.figures['industry_bar'] = fig
        return fig

    def create_asset_allocation_pie(self):
        """创建资产配置饼图"""
        if 'asset_allocation' not in self.data_cache:
            return None
        df = self.data_cache['asset_allocation']
        fig = go.Figure(data=[go.Pie(labels=df['资产类别'], values=df['市值'], hole=0.4)])
        fig.update_layout(title='资产配置结构', template='plotly_white')
        self.figures['asset_allocation_pie'] = fig
        return fig

    def create_monte_carlo_simulation(self, n_sim=100, n_days=252):
        """创建投资组合蒙特卡洛模拟图"""
        if 'stock_portfolio' not in self.data_cache:
            return None
        df = self.data_cache['stock_portfolio']
        returns = df.select_dtypes(include=[np.number]).pct_change().dropna().mean(axis=1)
        last_price = df.select_dtypes(include=[np.number]).iloc[-1].mean()
        simulations = []
        for _ in range(n_sim):
            price = [last_price]
            for _ in range(n_days):
                price.append(price[-1] * (1 + np.random.choice(returns)))
            simulations.append(price)
        fig = go.Figure()
        for sim in simulations:
            fig.add_trace(go.Scatter(y=sim, mode='lines', line=dict(width=1), opacity=0.3, showlegend=False))
        fig.update_layout(title='投资组合蒙特卡洛模拟', xaxis_title='未来交易日', yaxis_title='价格', template='plotly_white')
        self.figures['monte_carlo'] = fig
        return fig

    def create_cashflow_forecast(self):
        """创建现金流量预测图"""
        if 'cashflow' not in self.data_cache:
            return None
        df = self.data_cache['cashflow']
        if '日期' not in df.columns or '现金流' not in df.columns:
            return None
        from sklearn.linear_model import LinearRegression
        X = np.arange(len(df)).reshape(-1, 1)
        y = df['现金流'].values
        model = LinearRegression().fit(X, y)
        y_pred = model.predict(X)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df['日期'], y=y, mode='lines+markers', name='历史现金流'))
        fig.add_trace(go.Scatter(x=df['日期'], y=y_pred, mode='lines', name='预测现金流'))
        fig.update_layout(title='现金流量预测', xaxis_title='日期', yaxis_title='现金流', template='plotly_white')
        self.figures['cashflow_forecast'] = fig
        return fig

    def generate_html_dashboard(self):
        """生成HTML仪表板（集成新图表）"""
        print("🌐 生成交互式HTML仪表板...")
        charts_created = 0
        if self.create_stock_index_chart():
            charts_created += 1
        if self.create_portfolio_performance_chart():
            charts_created += 1
        if self.create_fund_comparison_chart():
            charts_created += 1
        if self.create_interest_rates_chart():
            charts_created += 1
        if self.create_correlation_heatmap():
            charts_created += 1
        if self.create_bond_market_chart():
            charts_created += 1
        if self.create_candlestick_chart():
            charts_created += 1
        if self.create_return_histogram():
            charts_created += 1
        if self.create_macro_economic_chart():
            charts_created += 1
        if self.create_market_cap_pie():
            charts_created += 1
        if self.create_industry_bar():
            charts_created += 1
        if self.create_asset_allocation_pie():
            charts_created += 1
        if self.create_monte_carlo_simulation():
            charts_created += 1
        if self.create_cashflow_forecast():
            charts_created += 1
        print(f"   ✅ 成功创建 {charts_created} 个交互式图表")
        html_content = self._generate_html_template()
        html_path = os.path.join(OUTPUT_DIR, HTML_FILE)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"✅ 交互式仪表板已生成: {html_path}")
        return html_path
    
    def _generate_chart_divs(self):
        """分区分组生成图表容器DIV，带标题和说明"""
        chart_descriptions = {
            'stock_index': '展示沪深300等主要指数的价格与成交量趋势。',
            'portfolio': '投资组合的价格、收益、风险等多维度综合分析。',
            'candlestick': '典型股票K线图，反映价格波动区间。',
            'return_histogram': '收益率分布，衡量风险特征。',
            'fund_radar': '多只基金的收益、风险、回撤等指标雷达对比。',
            'interest_rates': 'Shibor等利率走势分析。',
            'correlation': '资产间相关性热力图。',
            'bond_market': '债券市场规模与GDP关系。',
            'market_cap_pie': '大/中/小市值资产占比。',
            'industry_bar': '各行业上市公司数量分布。',
            'asset_allocation_pie': '资产类别配置结构。',
            'macro_economy': 'GDP、CPI等宏观经济指标趋势。',
            'monte_carlo': '投资组合未来价格蒙特卡洛模拟。',
            'cashflow_forecast': '历史与预测现金流量对比。'
        }
        sections = [
            ("市场表现", [
                ('stock_index', '股票指数分析', 'full-width'),
                ('portfolio', '投资组合分析', 'full-width'),
                ('candlestick', 'K线图', ''),
                ('return_histogram', '收益率分布', ''),
                ('fund_radar', '基金表现对比', ''),
                ('interest_rates', '利率走势', ''),
                ('correlation', '相关性分析', ''),
                ('bond_market', '债券市场', '')
            ]),
            ("资产配置", [
                ('market_cap_pie', '市值分布', ''),
                ('industry_bar', '行业分布', ''),
                ('asset_allocation_pie', '资产配置', '')
            ]),
            ("宏观经济", [
                ('macro_economy', '宏观经济指标', '')
            ]),
            ("风险与预测", [
                ('monte_carlo', '蒙特卡洛模拟', ''),
                ('cashflow_forecast', '现金流量预测', '')
            ])
        ]
        divs = []
        for section_title, charts in sections:
            divs.append(f'<a id="{section_title}" style="display:block;position:relative;top:-80px;"></a>')
            divs.append(f'<h2 style="margin-top:40px;">{section_title}</h2>')
            for chart_id, title, css_class in charts:
                if chart_id in self.figures:
                    class_attr = f'class="chart-container {css_class}"' if css_class else 'class="chart-container"'
                    desc = chart_descriptions.get(chart_id, '')
                    divs.append(f'<div {class_attr}><h3>{title}</h3><div id="{chart_id}" style="height: 100%;"></div>'
                                f'<p class="chart-desc" style="color:#888;font-size:0.95em;margin-top:8px;">{desc}</p></div>')
        return '\n        '.join(divs)
    
    def _generate_chart_scripts(self):
        """生成图表JavaScript代码"""
        scripts = ["window.charts = {};"]
        
        for chart_id, fig in self.figures.items():
            chart_json = fig.to_json()
            scripts.append(f"""
        // {chart_id} 图表
        var {chart_id}_data = {chart_json};
        Plotly.newPlot('{chart_id}', {chart_id}_data.data, {chart_id}_data.layout, {{
            responsive: true,
            displayModeBar: true,
            modeBarButtonsToRemove: ['pan2d', 'lasso2d', 'select2d'],
            displaylogo: false
        }});
        window.charts['{chart_id}'] = '{chart_id}';
            """)
        
        return '\n'.join(scripts)
    
    def _generate_html_template(self):
        """生成HTML模板，含导航和美观CSS"""
        nav_links = [
            ('市场表现', '市场表现'),
            ('资产配置', '资产配置'),
            ('宏观经济', '宏观经济'),
            ('风险与预测', '风险与预测')
        ]
        nav_html = '<nav class="dashboard-nav" style="text-align:center;margin-bottom:20px;">' + \
            ' | '.join([f'<a href="#{sec}" style="margin:0 12px;font-weight:bold;color:#667eea;">{name}</a>' for sec, name in nav_links]) + '</nav>'
        html_template = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>金融数据交互分析仪表板</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .header {{
            text-align: center;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px 0 10px 0;
            margin-bottom: 10px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .dashboard-nav a {{
            text-decoration: none;
            color: #667eea;
            font-size: 1.1em;
            transition: color 0.2s;
        }}
        .dashboard-nav a:hover {{
            color: #764ba2;
        }}
        .dashboard-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(600px, 1fr));
            gap: 20px;
            max-width: 1400px;
            margin: 0 auto;
        }}
        .chart-container {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            margin-bottom: 18px;
        }}
        .full-width {{
            grid-column: 1 / -1;
        }}
        .stats-panel {{
            background: white;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
        }}
        .stat-item {{
            display: inline-block;
            margin: 10px 24px;
            text-align: center;
            min-width: 120px;
        }}
        .stat-value {{
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
        }}
        .stat-label {{
            color: #666;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            color: #666;
            background: white;
            border-radius: 10px;
        }}
        .chart-desc {{
            color: #888;
            font-size: 0.95em;
            margin-top: 8px;
        }}
        @media (max-width: 900px) {{
            .dashboard-grid {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>金融数据交互分析仪表板</h1>
        <p>多维度金融数据可视化分析平台</p>
        <p>生成时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M')}</p>
    </div>
    {nav_html}
    <div class="stats-panel">
        <div class="stat-item">
            <div class="stat-value">{len(self.data_cache)}</div>
            <div class="stat-label">数据集</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{len(self.figures)}</div>
            <div class="stat-label">交互图表</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">200+</div>
            <div class="stat-label">原始数据表</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">20K+</div>
            <div class="stat-label">数据记录</div>
        </div>
    </div>
    <div class="dashboard-grid">
        {self._generate_chart_divs()}
    </div>
    <div class="footer">
        <h3>使用说明</h3>
        <p>📊 所有图表支持缩放、平移、悬停查看详情</p>
        <p>🔍 点击图例可以显示/隐藏数据系列</p>
        <p>💾 右上角工具栏可以下载图表为PNG格式</p>
        <p>🔄 双击图表可以重置缩放</p>
        <br>
        <p>© 2025 金融数据分析平台 | 基于Plotly.js技术</p>
    </div>
    <script>
        {self._generate_chart_scripts()}
        window.addEventListener('load', function() {{
            console.log('金融数据仪表板已加载完成');
            window.addEventListener('resize', function() {{
                Object.keys(window.charts).forEach(function(chartId) {{
                    Plotly.Plots.resize(chartId);
                }});
            }});
        }});
    </script>
</body>
</html>
        """
        return html_template
    
    def run_visualization(self):
        """运行可视化生成"""
        print("🎨 开始生成交互式可视化")
        print("=" * 50)
        
        if not HAS_PLOTLY:
            print("❌ 缺少plotly库，无法生成交互式图表")
            return False
        
        # 1. 加载数据
        if not self.load_visualization_data():
            return False
        
        # 2. 生成HTML仪表板
        html_path = self.generate_html_dashboard()
        
        print(f"\n🎉 交互式可视化生成完成!")
        print(f"📁 文件位置: {html_path}")
        print(f"🌐 打开方式: 双击HTML文件或在浏览器中打开")
        
        return True


def main():
    """主函数"""
    print("🎨 交互式数据可视化生成器")
    print("=" * 40)
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ 文件不存在: {EXCEL_FILE}")
        return
    
    visualizer = InteractiveVisualization(EXCEL_FILE)
    visualizer.run_visualization()


if __name__ == "__main__":
    main() 