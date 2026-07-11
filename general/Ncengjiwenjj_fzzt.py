# -- coding: utf-8 --
# @Time : 2025-03-12 14:15
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : N层级文件夹_复制粘贴.py
# @Software: PyCharm

import os
import shutil
import tkinter as tk
from tkinter import Tk, filedialog, messagebox


def choose_directory(title, parent=None):
    """
    选择文件夹路径
    :param title: 对话框标题
    :param parent: 父窗口，如果提供则使用其作为对话框父窗口
    """
    if parent and (isinstance(parent, tk.Tk) or isinstance(parent, tk.Toplevel)):
        # 如果有父窗口，创建一个临时的Toplevel用于文件对话框
        temp = tk.Toplevel(parent)
        temp.withdraw()
        directory = filedialog.askdirectory(title=title, parent=temp)
        temp.destroy()
    else:
        # 否则使用独立的Tk实例
        root = tk.Tk()
        root.withdraw()  # 隐藏Tkinter主窗口
        directory = filedialog.askdirectory(title=title)
        root.destroy()
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


def show_completion_message(parent=None):
    """
    显示处理完成的消息框
    :param parent: 父窗口，如果提供则使用其作为消息框的父窗口
    """
    if parent and (isinstance(parent, tk.Tk) or isinstance(parent, tk.Toplevel)):
        # 如果有父窗口，使用其作为消息框的父窗口
        messagebox.showinfo("处理完成", "文件复制已完成。", parent=parent)
    else:
        # 否则创建临时窗口
        root = tk.Tk()
        root.withdraw()  # 隐藏Tkinter主窗口
        messagebox.showinfo("处理完成", "文件复制已完成。")
        root.destroy()


def main_zfzwj(parent=None):
    """
    主函数：从N层级文件夹中提取文件
    :param parent: 父窗口，如果提供则使用其作为对话框的父窗口
    """
    # 选择源文件夹
    src_dir = choose_directory("选择源文件夹", parent)
    if not src_dir:
        print("未选择源文件夹。")
        return

    # 选择目标文件夹
    dst_dir = choose_directory("选择目标文件夹", parent)
    if not dst_dir:
        print("未选择目标文件夹。")
        return

    # 复制文件
    copy_files(src_dir, dst_dir)
    print("文件复制完成。")

    # 显示处理完成的提示
    show_completion_message(parent)


# if __name__ == "__main__":
#     main_zfzwj()