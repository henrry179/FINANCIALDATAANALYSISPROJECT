import pandas as pd

def load_data(filepath):
    """加载CSV或Excel数据文件"""
    if filepath.endswith('.csv'):
        return pd.read_csv(filepath)
    elif filepath.endswith('.xlsx') or filepath.endswith('.xls'):
        return pd.read_excel(filepath)
    else:
        raise ValueError('Unsupported file format')

def preprocess_data(df):
    """基础数据清洗：去除缺失值"""
    return df.dropna() 