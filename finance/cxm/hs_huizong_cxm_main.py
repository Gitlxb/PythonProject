# -*- coding: utf-8 -*-
"""
Excel汇总处理工具 - 启动入口
"""

import tkinter as tk
from finance.cxm.hs_gui import SummaryProcessorApp


def main():
    root = tk.Tk()
    app = SummaryProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
