# -- coding: utf-8 --
# @Time : 2025-03-10 13:42
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : 1234567.py
# @Software: PyCharm

import pandas as pd
import numpy as np

# 读取Excel文件
file_path = r"C:\Users\13560\Desktop\测试用.xlsx"  # 替换为你的Excel文件路径
df = pd.read_excel(file_path)

# 如果C列不存在，则创建C列并初始化为空
if 'C' not in df.columns:
    df['C'] = None

# 定义需要匹配的关键词
keywords = ['代发月工资', '工资', '小时工工资', '货款', '劳务服务费', '@KD付款', '劳务派遣', '劳务费', '．']

# 遍历A列，从第2行开始
for i in range(1, len(df)):
    cell_value = df.iloc[i, 0]
    # 检查是否为NaN
    if pd.isna(cell_value):
        cell_value = ""
    else:
        cell_value = str(cell_value)
    # 检查A列的值是否包含关键词列表中的任何一个关键词，或者是否为纯数字，或者是否为空
    if any(keyword in cell_value for keyword in keywords) or cell_value.isdigit() or not cell_value.strip():
        df.at[df.index[i], 'C'] = str(df.iloc[i, 1]) + "劳务服务费"  # 在C列输出B列内容加上"劳务服务费"

# 保存修改后的Excel文件
output_file_path = r"C:\Users\13560\Desktop\测试用.xlsx"  # 替换为你想保存的文件路径
df.to_excel(output_file_path, index=False)

print("处理完成，文件已保存为:", output_file_path)