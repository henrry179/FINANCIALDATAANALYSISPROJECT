import pandas as pd
import numpy as np
from datetime import datetime
from pandas import ExcelWriter

# 宏观经济指标
macro = pd.DataFrame({
    '时间': pd.date_range('2018-01-01', periods=24, freq='Q'),
    '全球GDP': np.random.uniform(70, 100, 24).round(2),
    '全球CPI': np.random.uniform(1.5, 4, 24).round(2),
    '全球失业率': np.random.uniform(4, 7, 24).round(2)
})

# 市值分布
market_cap = pd.DataFrame({
    '市值类型': ['大盘', '中盘', '小盘'],
    '市值': [60000, 25000, 15000]
})

# 行业分布
industry = pd.DataFrame({
    '行业': ['科技', '金融', '能源', '消费', '医疗', '工业', '原材料', '公用事业'],
    '公司数量': [120, 80, 60, 90, 70, 50, 40, 30]
})

# 资产配置
asset_allocation = pd.DataFrame({
    '资产类别': ['股票', '债券', '外汇', '期货', '现金'],
    '市值': [50000, 20000, 15000, 10000, 5000]
})

# 现金流量
cashflow = pd.DataFrame({
    '日期': pd.date_range('2022-01', periods=24, freq='M'),
    '现金流': np.random.uniform(-2000, 3000, 24).round(2)
})

with ExcelWriter('data/虚拟金融多市场数据集示例.xlsx') as writer:
    macro.to_excel(writer, sheet_name='宏观经济指标', index=False)
    market_cap.to_excel(writer, sheet_name='市值分布', index=False)
    industry.to_excel(writer, sheet_name='行业分布', index=False)
    asset_allocation.to_excel(writer, sheet_name='资产配置', index=False)
    cashflow.to_excel(writer, sheet_name='现金流量', index=False)

print('已生成虚拟金融多市场数据集示例.xlsx') 