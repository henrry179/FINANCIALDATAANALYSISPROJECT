import os
from src.data_loader import load_data, preprocess_data
from src.feature_engineering import add_features
from src.analysis.risk_analysis import calculate_volatility, calculate_sharpe_ratio

DATA_DIR = 'data/'
OUTPUT_FILE = 'output/batch_analysis_results.csv'

results = []

for filename in os.listdir(DATA_DIR):
    if filename.endswith('.csv') or filename.endswith('.xlsx') or filename.endswith('.xls'):
        filepath = os.path.join(DATA_DIR, filename)
        try:
            df = load_data(filepath)
            df = preprocess_data(df)
            df = add_features(df)
            vol = calculate_volatility(df)
            sharpe = calculate_sharpe_ratio(df)
            results.append({'file': filename, 'volatility': vol, 'sharpe_ratio': sharpe})
        except Exception as e:
            results.append({'file': filename, 'error': str(e)})

import pandas as pd
pd.DataFrame(results).to_csv(OUTPUT_FILE, index=False)
print(f'批量分析完成，结果已保存到 {OUTPUT_FILE}') 