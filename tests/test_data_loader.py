import pandas as pd
from src.data_loader import load_data, preprocess_data

def test_load_data_csv():
    # 假设有一个测试用csv文件 test.csv
    df = load_data('data/测试数据.csv')
    assert isinstance(df, pd.DataFrame)

def test_preprocess_data():
    df = pd.DataFrame({'a': [1, None, 3]})
    clean_df = preprocess_data(df)
    assert clean_df.isnull().sum().sum() == 0 