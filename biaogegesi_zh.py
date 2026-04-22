# -- coding: utf-8 --
# @Time : 2025-03-12 17:03
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : biaogegesi_zh.py
# @Software: PyCharm

import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox


def select_directory(title):
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    directory = filedialog.askdirectory(title=title)
    return directory


def convert_and_copy_files(source_dir, target_dir):
    # 支持的扩展名
    extensions = ['.xlsx', '.xlsm', '.xlsb', '.xls', '.csv']

    # 遍历源目录中的所有文件
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            file_path = os.path.join(root, file)
            file_name, file_ext = os.path.splitext(file)

            # 如果文件扩展名在支持的列表中
            if file_ext.lower() in extensions:
                # 读取文件
                if file_ext.lower() == '.csv':
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_excel(file_path)

                # 构建目标文件路径
                target_file_name = file_name + '.xlsx'
                target_file_path = os.path.join(target_dir, target_file_name)

                # 如果目标文件已经存在，则在文件名后加上_1
                counter = 1
                while os.path.exists(target_file_path):
                    target_file_name = f"{file_name}_{counter}.xlsx"
                    target_file_path = os.path.join(target_dir, target_file_name)
                    counter += 1

                # 保存为xlsx格式
                df.to_excel(target_file_path, index=False)
                print(f"Converted and copied: {file} -> {target_file_name}")


if __name__ == "__main__":
    # 选择源目录和目标目录
    source_directory = select_directory("选择要读取的文件夹")
    target_directory = select_directory("选择要保存的文件夹")

    if source_directory and target_directory:
        convert_and_copy_files(source_directory, target_directory)
        print("转换和复制完成！")

        # 显示处理完成的消息框
        root = Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showinfo("处理完成", "文件转换和复制已完成!")
    else:
        print("未选择目录，程序退出。")