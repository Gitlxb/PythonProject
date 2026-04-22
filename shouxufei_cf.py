# -- coding: utf-8 --
# @Time : 2025-03-11 9:43
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : 手续费与拆分.py
# @Software: PyCharm

def shouxufei_cf_xzy():
    """拆分_拆分成多个表格文件功能"""
    try:
        import os
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill
        from tkinter import Tk, filedialog, messagebox
        from tkinter import ttk
        import queue
        import threading
        import sys

        # 创建 Tkinter 根窗口
        root = Tk()
        root.withdraw()

        # 选择输入文件
        input_file = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not input_file:
            print("未选择输入文件，程序退出。")
            return

        # 设置默认保存路径为桌面
        desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
        output_file = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=desktop_path
        )
        if not output_file:
            print("未选择输出文件，程序退出。")
            return

        # 创建进度条窗口
        progress_window = Tk()
        progress_window.title("处理进度")
        progress_window.geometry("400x150")

        # 使窗口居中
        progress_window.update_idletasks()
        width = progress_window.winfo_width()
        height = progress_window.winfo_height()
        x = (progress_window.winfo_screenwidth() // 2) - (width // 2)
        y = (progress_window.winfo_screenheight() // 2) - (height // 2)
        progress_window.geometry(f'{width}x{height}+{x}+{y}')

        # 添加标签
        label = ttk.Label(progress_window, text="正在处理数据，请稍候...", font=("Arial", 12))
        label.pack(pady=20)

        # 添加进度条
        progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
        progress_bar.pack(pady=10)

        # 添加进度百分比标签
        progress_label = ttk.Label(progress_window, text="0%", font=("Arial", 10))
        progress_label.pack(pady=5)

        # 创建队列用于线程间通信
        progress_queue = queue.Queue()
        def update_progress_from_queue():
            """从队列更新进度条"""
            try:
                while True:
                    try:
                        progress, message = progress_queue.get_nowait()
                    except queue.Empty:
                        break
        
                    progress_bar['value'] = progress
                    progress_label.config(text=f"{progress:.0f}%")
                    if message:
                        label.config(text=message)
        
                    if progress >= 100:
                        progress_window.after(1000, show_completion_dialog)
            finally:
                progress_window.after(100, update_progress_from_queue)

        def show_completion_dialog():
            """显示完成对话框"""
            try:
                progress_window.destroy()
                messagebox.showinfo("完成", "处理完成，文件已保存！")
                root.quit()
                root.destroy()
                sys.exit(0)
            except:
                try:
                    root.quit()
                    root.destroy()
                except:
                    pass
                sys.exit(0)

        def show_error_dialog(error_msg):
            """显示错误对话框"""
            try:
                progress_window.destroy()
                messagebox.showerror("错误", error_msg)
                root.quit()
                root.destroy()
                sys.exit(0)
            except:
                try:
                    root.quit()
                    root.destroy()
                except:
                    pass
                sys.exit(0)

        # 处理函数
        def process_data():
            """处理 Excel 数据的核心函数"""
            try:
                progress_queue.put((5, "正在读取 Excel 文件..."))
                df = pd.read_excel(input_file)

                progress_queue.put((10, "正在处理列映射..."))
                # 定义列字母到索引的映射
                col_map = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'L': 11}
                col_name = {letter: df.columns[idx] for letter, idx in col_map.items()}

                # 对 L 列去除所有空格
                if col_name['L'] in df.columns:
                    df[col_name['L']] = df[col_name['L']].astype(str).str.replace(' ', '')

                progress_queue.put((15, "正在应用条件规则..."))

                # 使用向量化操作优化条件判断（性能提升关键）
                total_rows = len(df)
                
                # 条件 1: H 列带有"招商银行"且 C 列没有"伟明"，E 列为 0
                mask1 = df[col_name['H']].astype(str).str.contains('招商银行', na=False) & \
                        ~df[col_name['C']].astype(str).str.contains('伟明', na=False)
                df.loc[mask1, col_name['E']] = 0

                # 条件 2: I 列出现"驻厂"，E 列为 0
                mask2 = df[col_name['I']].astype(str).str.contains('驻厂', na=False)
                df.loc[mask2, col_name['E']] = 0

                # 条件 3: C 列有"伟明"且 H 列有"中信银行"，E 列为 0
                mask3 = df[col_name['C']].astype(str).str.contains('伟明', na=False) & \
                        df[col_name['H']].astype(str).str.contains('中信银行', na=False)
                df.loc[mask3, col_name['E']] = 0

                # 条件 4: C 列有"伟明"、H 列有"招商银行"且 I 列有"垫付"，E 列为 0
                mask4 = df[col_name['C']].astype(str).str.contains('伟明', na=False) & \
                        df[col_name['H']].astype(str).str.contains('招商银行', na=False) & \
                        df[col_name['I']].astype(str).str.contains('垫付', na=False)
                df.loc[mask4, col_name['E']] = 0

                # 条件 5: C 列有"瑞立"且 H 列有"招商银行"或"邮政"，E 列为 0
                mask5 = df[col_name['C']].astype(str).str.contains('瑞立', na=False) & \
                        (df[col_name['H']].astype(str).str.contains('招商银行', na=False) | \
                         df[col_name['H']].astype(str).str.contains('邮政', na=False))
                df.loc[mask5, col_name['E']] = 0

                # 条件 6: C 列有"中世"，E 列为 0
                mask6 = df[col_name['C']].astype(str).str.contains('中世', na=False)
                df.loc[mask6, col_name['E']] = 0

                progress_queue.put((35, "正在计算 E 列数值..."))

                # 对 E 列未填充的单元格，根据 D 列值输出
                mask_empty_e = df[col_name['E']].isna()
                df.loc[mask_empty_e & (df[col_name['D']] <= 500), col_name['E']] = 5
                df.loc[mask_empty_e & (df[col_name['D']] > 500), col_name['E']] = \
                    df.loc[mask_empty_e, col_name['D']] * 0.01

                progress_queue.put((45, "正在计算 F 列数值..."))
                # 计算 F 列 = D 列 - E 列
                df[col_name['F']] = df[col_name['D']] - df[col_name['E']]

                progress_queue.put((50, "正在保存文件..."))
                # 保存修改后的 Excel 文件
                df.to_excel(output_file, index=False, engine='openpyxl')

                # 使用 openpyxl 添加公式到 F 列
                wb_calc = load_workbook(output_file)
                ws_calc = wb_calc.active
                last_data_row = len(df) + 1

                # 添加公式 =D2-E2, =D3-E3, 等
                for row in range(2, last_data_row + 1):
                    ws_calc[f'F{row}'] = f'=D{row}-E{row}'

                wb_calc.save(output_file)
                
                progress_queue.put((60, "正在拆分工作表..."))

                # 根据 G 列的内容创建新的工作表
                wb_temp = load_workbook(output_file, data_only=True)
                ws_temp = wb_temp.active

                # 从 Excel 中读取 F 列的数值
                df_for_split = df.copy()
                for idx in range(len(df_for_split)):
                    row_num = idx + 2
                    f_value = ws_temp[f'F{row_num}'].value
                    if f_value is not None:
                        try:
                            df_for_split.at[idx, col_name['F']] = float(f_value)
                        except (ValueError, TypeError):
                            d_val = df_for_split.at[idx, col_name['D']]
                            e_val = df_for_split.at[idx, col_name['E']]
                            if pd.notna(d_val) and pd.notna(e_val):
                                try:
                                    df_for_split.at[idx, col_name['F']] = float(d_val) - float(e_val)
                                except (ValueError, TypeError):
                                    df_for_split.at[idx, col_name['F']] = 0
                            else:
                                df_for_split.at[idx, col_name['F']] = 0
                    else:
                        d_val = df_for_split.at[idx, col_name['D']]
                        e_val = df_for_split.at[idx, col_name['E']]
                        if pd.notna(d_val) and pd.notna(e_val):
                            try:
                                df_for_split.at[idx, col_name['F']] = float(d_val) - float(e_val)
                            except (ValueError, TypeError):
                                df_for_split.at[idx, col_name['F']] = 0
                        else:
                            df_for_split.at[idx, col_name['F']] = 0

                wb_temp.close()
                                
                with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    unique_values = df_for_split[df_for_split.columns[col_map['G']]].unique()

                    for i, value in enumerate(unique_values):
                        if pd.isna(value):
                            continue
                        filtered_df = df_for_split[df_for_split[df_for_split.columns[col_map['G']]] == value]
                        filtered_df.to_excel(writer, sheet_name=str(value), index=False)

                        progress = 60 + (i / len(unique_values) * 10)
                        progress_queue.put((progress, f"正在拆分工作表：{value}"))

                progress_queue.put((70, "正在加载工作簿..."))

                # 加载修改后的 Excel 文件
                wb = load_workbook(output_file)

                # 定义黄色填充
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                # 定义颜色填充样式
                duplicate_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                progress_queue.put((75, "正在处理工作表格式..."))

                # 遍历每个工作表
                total_sheets = len(wb.sheetnames)
                for sheet_idx, sheet_name in enumerate(wb.sheetnames):
                    ws = wb[sheet_name]

                    # 找到 C 列最后一个有内容的行，添加合计行
                    last_row = ws.max_row
                    total_row = None
                    for row in range(1, last_row + 2):
                        if ws[f'C{row}'].value is None:
                            ws[f'C{row}'] = '合计'
                            ws[f'D{row}'] = f'=SUM(D2:D{row - 1})'
                            ws[f'E{row}'] = f'=SUM(E2:E{row - 1})'
                            ws[f'F{row}'] = f'=SUM(F2:F{row - 1})'
                            ws[f'E{row}'].fill = yellow_fill
                            ws[f'F{row}'].fill = yellow_fill
                            total_row = row
                            break

                    # 标记特殊条件的行（合并循环优化）
                    for row in range(1, ws.max_row + 1):
                        # 驻厂标记
                        if ws[f'I{row}'].value and '驻厂' in str(ws[f'I{row}'].value):
                            for col in ['D', 'E', 'F', 'G', 'H', 'I']:
                                ws[f'{col}{row}'].fill = yellow_fill
                        
                        # 招商银行标记
                        if ws[f'H{row}'].value and '招商银行' in str(ws[f'H{row}'].value):
                            for col in ['D', 'E', 'F', 'G', 'H']:
                                ws[f'{col}{row}'].fill = yellow_fill
                        
                        # 伟明 + 中信银行标记
                        c_value = ws[f'C{row}'].value
                        h_value = ws[f'H{row}'].value
                        if c_value and '伟明' in str(c_value) and h_value and '中信银行' in str(h_value):
                            for col in ['D', 'E', 'F', 'G', 'H']:
                                ws[f'{col}{row}'].fill = yellow_fill
                        
                        # 瑞立 + 招商银行/邮政标记
                        if c_value and '瑞立' in str(c_value) and h_value and ('招商银行' in str(h_value) or '邮政' in str(h_value)):
                            for col in ['D', 'E', 'F', 'G', 'H']:
                                ws[f'{col}{row}'].fill = yellow_fill
                        
                        # 中世标记
                        if c_value and '中世' in str(c_value):
                            for col in ['D', 'E', 'F', 'G', 'H']:
                                ws[f'{col}{row}'].fill = yellow_fill

                    # 标记 J 列和 K 列的重复值
                    j_values = [ws[f'J{row}'].value for row in range(1, ws.max_row + 1)]
                    k_values = [ws[f'K{row}'].value for row in range(1, ws.max_row + 1)]
                    j_duplicates = set([x for x in j_values if j_values.count(x) > 1])
                    k_duplicates = set([x for x in k_values if k_values.count(x) > 1])

                    for row in range(1, ws.max_row + 1):
                        if ws[f'J{row}'].value in j_duplicates:
                            ws[f'J{row}'].fill = duplicate_fill
                        if ws[f'K{row}'].value in k_duplicates:
                            ws[f'K{row}'].fill = duplicate_fill

                    # 添加账号、户名、金额、汇款备注信息
                    start_row = total_row + 5
                    ws[f'C{start_row}'] = '账号'
                    ws[f'D{start_row}'] = '户名'
                    ws[f'E{start_row}'] = '金额'
                    ws[f'F{start_row}'] = '汇款备注'

                    data_row = start_row + 1
                    amount_sum = 0

                    for row_idx in range(2, total_row):
                        bank_card = ws[f'L{row_idx}'].value
                        account_name = ws[f'K{row_idx}'].value

                        # 获取 F 列的实际金额
                        f_cell = ws[f'F{row_idx}']
                        if f_cell.value is not None:
                            if isinstance(f_cell.value, str) and f_cell.value.startswith('='):
                                d_value = ws[f'D{row_idx}'].value
                                e_value = ws[f'E{row_idx}'].value
                                if d_value is not None and e_value is not None:
                                    try:
                                        actual_amount = float(d_value) - float(e_value)
                                    except (ValueError, TypeError):
                                        actual_amount = 0
                                else:
                                    actual_amount = 0
                            else:
                                try:
                                    actual_amount = float(f_cell.value)
                                except (ValueError, TypeError):
                                    actual_amount = 0
                        else:
                            actual_amount = 0

                        name = ws[f'J{row_idx}'].value

                        ws[f'C{data_row}'] = bank_card
                        ws[f'D{data_row}'] = account_name
                        ws[f'E{data_row}'] = actual_amount
                        ws[f'F{data_row}'] = name

                        if actual_amount is not None:
                            try:
                                amount_sum += float(actual_amount)
                            except (ValueError, TypeError):
                                pass

                        data_row += 1

                    # 添加金额合计行
                    amount_total_row = data_row
                    ws[f'D{amount_total_row}'] = '金额合计'
                    ws[f'E{amount_total_row}'] = amount_sum
                    ws[f'E{amount_total_row}'].fill = yellow_fill

                    # 更新进度
                    progress = 75 + ((sheet_idx + 1) / total_sheets * 20)
                    progress_queue.put((progress, f"正在处理工作表：{sheet_name}"))

                progress_queue.put((95, "正在保存最终文件..."))
                wb.save(output_file)
                progress_queue.put((100, "处理完成！"))

            except Exception as e:
                import traceback
                error_msg = f"拆分成多个表格文件时出错：{str(e)}\n{traceback.format_exc()}"
                print(error_msg)
                progress_queue.put((0, f"处理出错：{str(e)}"))
                progress_window.after(0, lambda: show_error_dialog(error_msg))

        # 启动队列检查
        progress_window.after(100, update_progress_from_queue)

        # 在新线程中处理数据
        thread = threading.Thread(target=process_data)
        thread.daemon = True
        thread.start()

        # 启动进度窗口主循环
        progress_window.mainloop()

    except Exception as e:
        import traceback
        error_msg = f"程序初始化时出错：{str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        try:
            messagebox.showerror("错误", error_msg)
        except:
            pass
        finally:
            import sys
            try:
                root.quit()
                root.destroy()
            except:
                pass
            sys.exit(0)


if __name__ == "__main__":
    shouxufei_cf_xzy()