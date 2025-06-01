#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建测试数据文件
用于演示数据合并脚本的功能
"""

import pandas as pd
import os

def create_test_data():
    """创建测试数据文件"""
    
    # 创建测试文件夹
    os.makedirs("test_data", exist_ok=True)
    os.makedirs("test_data/folder1", exist_ok=True)
    os.makedirs("test_data/folder2", exist_ok=True)
    
    # 创建测试数据1 - 销售数据
    sales_data = {
        '日期': ['2024-01-01', '2024-01-02', '2024-01-03'],
        '产品': ['产品A', '产品B', '产品C'],
        '销量': [100, 150, 200],
        '价格': [10.5, 15.8, 25.0],
        '销售额': [1050, 2370, 5000]
    }
    df1 = pd.DataFrame(sales_data)
    df1.to_csv("test_data/folder1/销售数据.csv", index=False, encoding='utf-8-sig')
    
    # 创建测试数据2 - 库存数据  
    inventory_data = {
        '产品编号': ['A001', 'B002', 'C003', 'D004'],
        '产品名称': ['笔记本', '鼠标', '键盘', '显示器'],
        '库存量': [50, 100, 75, 25],
        '采购价格': [3000, 50, 200, 1500]
    }
    df2 = pd.DataFrame(inventory_data)
    df2.to_excel("test_data/folder1/库存数据.xlsx", index=False)
    
    # 创建测试数据3 - 员工数据
    employee_data = {
        '员工ID': ['E001', 'E002', 'E003'],
        '姓名': ['张三', '李四', '王五'],
        '部门': ['销售部', '技术部', '财务部'],
        '工资': [8000, 12000, 9000]
    }
    df3 = pd.DataFrame(employee_data)
    df3.to_csv("test_data/folder2/员工数据.csv", index=False, encoding='utf-8-sig')
    
    # 创建测试数据4 - 财务数据
    finance_data = {
        '月份': ['2024-01', '2024-02', '2024-03'],
        '收入': [100000, 120000, 95000],
        '支出': [80000, 90000, 85000],
        '利润': [20000, 30000, 10000]
    }
    df4 = pd.DataFrame(finance_data)
    df4.to_excel("test_data/folder2/财务数据.xlsx", index=False)
    
    print("✅ 测试数据创建完成!")
    print("\n创建的文件:")
    print("├── test_data/")
    print("│   ├── folder1/")
    print("│   │   ├── 销售数据.csv")
    print("│   │   └── 库存数据.xlsx") 
    print("│   └── folder2/")
    print("│       ├── 员工数据.csv")
    print("│       └── 财务数据.xlsx")
    
    return True

if __name__ == "__main__":
    create_test_data() 