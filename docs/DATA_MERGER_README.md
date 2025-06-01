# Excel/CSV数据表合并工具

这是一个Python脚本，用于合并多个文件夹内的Excel(.xlsx)和CSV(.csv)数据表。

## 功能特点

- ✅ 支持Excel(.xlsx)和CSV(.csv)格式
- ✅ 递归遍历多个文件夹
- ✅ 自动处理CSV文件编码问题（UTF-8, GBK, Latin-1）
- ✅ 为每行数据添加文件来源信息
- ✅ 详细的处理日志和错误报告
- ✅ 支持交互式输入和预配置两种模式

## 安装依赖

```bash
pip install -r requirements.txt
```

或者单独安装：

```bash
pip install pandas openpyxl
```

## 使用方法

### 方法1: 直接运行脚本（推荐）

1. 编辑 `data_merger.py` 文件中的文件夹路径配置：

```python
folder_paths = [
    "./data",           # 您的数据文件夹路径1
    "./input",          # 您的数据文件夹路径2
    "/path/to/folder3", # 可以添加更多文件夹路径
]
```

2. 运行脚本：

```bash
python data_merger.py
```

3. 按提示选择是否使用交互式输入，然后输入输出文件名

### 方法2: 交互式输入

运行脚本后选择交互式输入模式，逐个输入要合并的文件夹路径。

## 输出结果

合并后的数据表将包含以下额外列：
- `文件来源`: 原始文件名
- `文件路径`: 原始文件的完整路径

## 文件夹结构示例

```
项目目录/
├── data_merger.py          # 主脚本
├── requirements.txt        # 依赖文件
├── data/                  # 数据文件夹1
│   ├── file1.xlsx
│   ├── file2.csv
│   └── subfolder/
│       └── file3.xlsx
├── input/                 # 数据文件夹2
│   ├── data1.csv
│   └── data2.xlsx
└── merged_data_20240304_143022.xlsx  # 输出文件
```

## 错误处理

脚本会自动处理以下情况：
- 文件夹不存在
- 文件读取失败
- 编码问题
- 空文件或损坏文件

所有错误信息会在处理摘要中显示。

## 注意事项

1. 确保所有要合并的文件具有相似的数据结构
2. 不同文件的列名可能不同，pandas会自动处理缺失列
3. 大量数据文件可能需要较长处理时间
4. 确保有足够的磁盘空间保存合并后的文件

## 自定义配置

### 修改支持的文件格式

在 `DataMerger` 类中修改 `supported_formats` 列表：

```python
self.supported_formats = ['.xlsx', '.csv', '.xls']  # 添加.xls支持
```

### 修改输出格式

脚本根据输出文件扩展名自动判断格式：
- `.xlsx` → Excel格式
- `.csv` → CSV格式（UTF-8编码）
- 其他 → 默认Excel格式

## 故障排除

### 常见问题

1. **导入错误**: 确保已安装pandas和openpyxl
   ```bash
   pip install pandas openpyxl
   ```

2. **文件路径错误**: 使用绝对路径或确保相对路径正确

3. **内存不足**: 处理大量数据时可能需要更多内存

4. **编码问题**: 脚本已自动处理常见编码，如仍有问题请检查文件编码

### 获取帮助

如有问题，请检查：
1. 文件路径是否正确
2. 文件是否有读取权限
3. Python版本是否兼容（建议Python 3.7+） 