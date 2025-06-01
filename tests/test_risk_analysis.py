import pandas as pd
from src.analysis.risk_analysis import calculate_volatility, calculate_sharpe_ratio

def test_calculate_volatility():
    df = pd.DataFrame({'return': [0.01, 0.02, -0.01, 0.03]})
    vol = calculate_volatility(df)
    assert vol > 0

def test_calculate_sharpe_ratio():
    df = pd.DataFrame({'return': [0.01, 0.02, -0.01, 0.03]})
    sharpe = calculate_sharpe_ratio(df)
    assert isinstance(sharpe, float) 