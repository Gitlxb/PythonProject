# -- coding: utf-8 --
# @Time : 2025-03-12 14:15
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : N层级文件夹_复制粘贴.py
# @Software: PyCharm

import os
import shutil
from tkinter import Tk, filedialog, messagebox


def choose_directory(title):
    """选择文件夹路径"""
    root = Tk()
    root.withdraw()  # 隐藏Tkinter主窗口
    directory = filedialog.askdirectory(title=title)
    return directory


def copy_files(src_dir, dst_dir):
    """递归提取文件到目标文件夹，不保留文件夹结构"""
    for root_dir, _, files in os.walk(src_dir):
        for filename in files:
            src_file = os.path.join(root_dir, filename)
            dst_file = os.path.join(dst_dir, filename)

            # 处理同名文件
            base, ext = os.path.splitext(filename)
            counter = 1
            while os.path.exists(dst_file):
                dst_file = os.path.join(dst_dir, f"{base}_{counter}{ext}")
                counter += 1

            shutil.copy2(src_file, dst_file)
            print(f"Copied: {src_file} -> {dst_file}")


def show_completion_message():
    """显示处理完成的消息框"""
    root = Tk()
    root.withdraw()  # 隐藏Tkinter主窗口
    messagebox.showinfo("处理完成", "文件复制已完成。")


def main_zfzwj():
    # 选择源文件夹
    src_dir = choose_directory("选择源文件夹")
    if not src_dir:
        print("未选择源文件夹。")
        return

    # 选择目标文件夹
    dst_dir = choose_directory("选择目标文件夹")
    if not dst_dir:
        print("未选择目标文件夹。")
        return

    # 复制文件
    copy_files(src_dir, dst_dir)
    print("文件复制完成。")

    # 显示处理完成的提示
    show_completion_message()


# if __name__ == "__main__":
#     main_zfzwj()