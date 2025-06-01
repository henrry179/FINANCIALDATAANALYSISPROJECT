def calculate_volatility(df):
    """计算年化波动率"""
    return df['return'].std() * (252 ** 0.5)

def calculate_sharpe_ratio(df, risk_free_rate=0.03):
    """计算夏普比率"""
    excess_return = df['return'] - risk_free_rate / 252
    return excess_return.mean() / excess_return.std() * (252 ** 0.5) 