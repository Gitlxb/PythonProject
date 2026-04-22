# -- coding: utf-8 --
# @Time : 2025-10-22 14:35
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : txm_zyy.py
# @Software: PyCharm

import sys
import subprocess
import os

def check_and_install():
    """检查并自动安装依赖"""
    try:
        import barcode
        from barcode.writer import ImageWriter
        return True
    except ImportError:
        print("正在自动安装依赖库...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "python-barcode==0.14.0", "pillow"])
            return True
        except:
            return False

if not check_and_install():
    print("无法自动安装依赖，请手动执行：")
    print("pip install python-barcode==0.14.0 pillow")
    sys.exit(1)

# --- 修复 Pillow 兼容性问题 ---
from PIL import ImageFont

# 保存原始构造函数
original_init = ImageFont.FreeTypeFont.__init__

def patched_init(self, *args, **kwargs):
    original_init(self, *args, **kwargs)
    if not hasattr(self, 'getsize'):
        self.getsize = lambda text: self.getbbox(text)[:2]  # 返回 width, height

# 打补丁
ImageFont.FreeTypeModel = ImageFont.FreeTypeFont  # 防止重复打补丁
ImageFont.FreeTypeFont.__init__ = patched_init
# ---------------------------------

# 以下为正式代码
from barcode import get_barcode_class
from barcode.writer import ImageWriter

def generate_barcode(data, filename="barcode"):
    try:
        Code128 = get_barcode_class('code128')
        code = Code128(data, writer=ImageWriter())
        full_path = code.save(filename)
        print(f"条形码已生成: {os.path.abspath(full_path)}.png")
        return full_path
    except Exception as e:
        print(f"生成失败: {str(e)}")
        return None

if __name__ == "__main__":
    print("GS1条形码生成器（输入内容包含括号即可自动识别为GS1格式）")
    while True:
        data = input("\n请输入内容（留空退出）: ").strip()
        if not data:
            break
        output = input("保存文件名（默认barcode）: ").strip() or "barcode"
        generate_barcode(data, output)