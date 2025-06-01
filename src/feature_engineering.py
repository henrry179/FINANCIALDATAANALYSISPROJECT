def add_features(df):
    """为数据集添加常用金融特征"""
    if 'close' in df.columns:
        df['return'] = df['close'].pct_change()
    return df 