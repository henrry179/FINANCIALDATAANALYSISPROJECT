import pandas as pd
from src.feature_engineering import add_features

def test_add_features():
    df = pd.DataFrame({'close': [1, 2, 3, 4]})
    df = add_features(df)
    assert 'return' in df.columns 