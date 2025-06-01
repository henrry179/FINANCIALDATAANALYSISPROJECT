from src.data_loader import load_data, preprocess_data
from src.feature_engineering import add_features
from src.analysis.risk_analysis import calculate_volatility, calculate_sharpe_ratio
from src.utils.helpers import plot_timeseries

# 示例：加载数据
# df = load_data('data/your_data.csv')
# df = preprocess_data(df)
# df = add_features(df)
# print('年化波动率:', calculate_volatility(df))
# print('夏普比率:', calculate_sharpe_ratio(df))
# plot_timeseries(df, 'close') 