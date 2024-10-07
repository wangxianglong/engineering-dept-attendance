import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
import calendar
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment,PatternFill,Border, Side,Font
from openpyxl.worksheet.dimensions import ColumnDimension,DimensionHolder
from openpyxl.utils import get_column_letter

root = tk.Tk()
root.geometry("600x300+50+50")
root.title("工程部考勤记录生成器")

select_path = tk.StringVar()


def select_file():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    select_path.set(selected_file_path)

def generate_excel():
    selected_month = month_combo.get().replace('月','')
    selected_year = year_combo.get()
    # print(get_days_of_month(int(selected_year), int(selected_month)))
    workbook = Workbook()
    sheet = workbook.active

# 创建实线边框的边界定义
    border = Border(left=Side(border_style='thin', color='000000'),  # 左边界，实线，颜色为黑色
                right=Side(border_style='thin', color='000000'),  # 右边界
                top=Side(border_style='thin', color='000000'),  # 上边界
                bottom=Side(border_style='thin', color='000000'))  # 下边界

    sheet.merge_cells('A1:B1')
    sheet.cell(row=1, column=1).value = f"工程部{selected_year}年{selected_month}月加班调休明细表"

    sheet.merge_cells('A3:A4')
    sheet.cell(row=3, column=1).value = '星期'
    sheet.cell(row=3, column=1).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=3, column=1).border = border

    sheet.cell(row=3, column=2).value = '姓名'
    sheet.cell(row=3, column=2).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=3, column=2).border = border

    sheet.cell(row=4, column=2).value = '日期'
    sheet.cell(row=4, column=2).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=4, column=2).border = border

    df = pd.read_excel(select_path.get())
    print(df)
     
    # 生成日期数据
    for column_index in range(1,32):
        date_iloc = df.iloc[1, column_index]  #取日期数据
        if type(date_iloc).__name__ == "int":
            row_index = column_index + 4
            
            # print(date_iloc)
            sheet.cell(row=row_index, column=2).value = f'{selected_month}月{date_iloc}日'
            sheet.cell(row=row_index, column=2).alignment = Alignment(horizontal="center",vertical="center")
            sheet.cell(row=row_index, column=2).border = border

            day_iloc = df.iloc[2, column_index]  #取星期几数据
            sheet.cell(row=row_index, column=1).value = day_iloc
            sheet.cell(row=row_index, column=1).alignment = Alignment(horizontal="center",vertical="center")
            sheet.cell(row=row_index, column=1).border = border

            if day_iloc == '六' or day_iloc == '日':
                for col in range(1,sheet.max_column + 1):
                    sheet.cell(row=row_index, column=col).fill = PatternFill(start_color='FFFF00', fill_type='solid')
     
    #生成最后几行汇总数据的第一列内容
    row_index = row_index + 1
    sheet.cell(row=row_index, column=1).value = f'{selected_month}月合计'
    sheet.cell(row=row_index, column=1).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=row_index, column=1).border = border
    sheet.cell(row=row_index, column=1).font = Font(size=10, bold=True)
    sheet.merge_cells(f'A{row_index}:B{row_index}')

    row_index = row_index + 1
    sheet.cell(row=row_index, column=1).value = f'{int(selected_month) - 1}月剩余加班小时'
    sheet.cell(row=row_index, column=1).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=row_index, column=1).border = border
    sheet.cell(row=row_index, column=1).font = Font(size=8, bold=True)
    sheet.merge_cells(f'A{row_index}:B{row_index}')

    row_index = row_index + 1
    sheet.cell(row=row_index, column=1).value = f'{selected_month}月剩余加班小时数'
    sheet.cell(row=row_index, column=1).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=row_index, column=1).border = border
    sheet.cell(row=row_index, column=1).font = Font(size=8, bold=True)
    sheet.merge_cells(f'A{row_index}:B{row_index}')

    row_index = row_index + 1
    sheet.cell(row=row_index, column=1).value = "签名"
    sheet.cell(row=row_index, column=1).alignment = Alignment(horizontal="center",vertical="center")
    sheet.cell(row=row_index, column=1).border = border
    sheet.cell(row=row_index, column=1).font = Font(size=11, bold=True)
    sheet.merge_cells(f'A{row_index}:B{row_index}')
    sheet.row_dimensions[row_index].height = 35

    row_index = row_index + 1
    sheet.cell(row=row_index, column=1).value = "制表人："
    sheet.cell(row=row_index, column=1).alignment = Alignment(horizontal="center",vertical="bottom")
    sheet.merge_cells(f'A{row_index}:B{row_index}')

    sheet.cell(row=row_index, column=8).value = "项目负责人："
    sheet.cell(row=row_index, column=8).alignment = Alignment(horizontal="left",vertical="bottom")
    sheet.row_dimensions[row_index].height = 31.1

    #生成员工数据
    name_row_index = 3
    name_column_index = 2
  
    while True:
        name_iloc = df.iloc[name_row_index, 0]
        if name_iloc == "排班说明":
            break
        else: # 一个员工的数据
            sheet.cell(row=3, column=name_column_index + 1).value = name_iloc
            sheet.cell(row=3, column=name_column_index + 1).alignment = Alignment(horizontal="center",vertical="center")
            sheet.cell(row=3, column=name_column_index + 1).border = border
            sheet.cell(row=3, column=name_column_index + 1).font = Font(size=11)
            sheet.merge_cells(start_row=3, start_column=name_column_index + 1, end_row=3, end_column=name_column_index + 4)

            cell1 = sheet.cell(row=4, column=name_column_index + 1)
            cell1.value = '平时加班'
            cell1.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            cell1.border = border
            cell1.font = Font(size=8)
            sheet.column_dimensions[get_column_letter(cell1.column)].width = 4
            set_cell_border(row_index,sheet,name_column_index + 1,border)

            cell2 = sheet.cell(row=4, column=name_column_index + 2)
            cell2.value = '法定加班'
            cell2.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            cell2.border = border
            cell2.font = Font(size=8)
            sheet.column_dimensions[get_column_letter(cell2.column)].width = 4
            set_cell_border(row_index,sheet,name_column_index + 2,border)

            cell3 = sheet.cell(row=4, column=name_column_index + 3)
            cell3.value = '周末加班'
            cell3.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            cell3.border = border
            cell3.font = Font(size=8)
            sheet.column_dimensions[get_column_letter(cell3.column)].width = 4
            set_cell_border(row_index,sheet,name_column_index + 3,border)

            cell4 = sheet.cell(row=4, column=name_column_index + 4)
            cell4.value = '调休'
            cell4.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            cell4.border = border
            cell4.font = Font(size=8)
            sheet.column_dimensions[get_column_letter(cell4.column)].width = 4
            set_cell_border(row_index,sheet,name_column_index + 4,border)

            sheet.cell(row=row_index - 1, column=name_column_index + 1).border = border
            sheet.merge_cells(start_row=row_index - 1, start_column=name_column_index + 1, end_row=row_index - 1, end_column=name_column_index + 4)

            for date_col_index in range(1,32):
                result_iloc = df.iloc[name_row_index, date_col_index]
                if result_iloc == "调休":
                    name_value = df.iloc[name_row_index, 0]
                    date_value = df.iloc[1, date_col_index]
                    # day_value = df.iloc[2, date_col_index]
                    for cell in sheet[2]:
                        if cell.value == name_value:
                            sheet.cell(row=4 + date_value, column=cell.column + 3).value = 8
                                
                if result_iloc == "加班":
                    name_value = df.iloc[name_row_index, 0]
                    date_value = df.iloc[1, date_col_index]
                    day_value = df.iloc[2, date_col_index] 
            
                    for cell in sheet[2]:
                        if cell.value == name_value:
                            if day_value == "六" or day_value == "日":
                                sheet.cell(row=4 + date_value, column=cell.column + 2).value = 8
                            else:
                                sheet.cell(row=4 + date_value, column=cell.column).value = 8

            name_column_index = name_column_index + 4
            name_row_index = name_row_index + 1

    workbook.save("工程部加班调休明细表.xlsx")

def set_cell_border(row_index,sheet,column_index,border):
    for border_row in range(5,row_index - 1):
        cell_border = sheet.cell(row=border_row, column=column_index)
        cell_border.alignment = Alignment(horizontal="center",vertical="center")
        cell_border.border = border
        cell_border.font = Font(size=9, bold=True)
        cell0 = sheet.cell(row=border_row, column=1)
        if cell0.value == "六" or cell0.value == "日":
            cell_border.fill = PatternFill(start_color='FFFF00', fill_type='solid')
            

def select_date_range():
    print(f"Selected month: {month_combo.get()}")

def get_days_of_month(year, month):
    # 获取指定年月的日历
    cal = calendar.monthcalendar(year, month)
    days_of_month = []
    for week in cal:
        for day in week:
            if day > 0:
                # 将日期转换为字符串格式
                day_str = str(day)
                # 获取星期几
                weekday = datetime(year, month, day).weekday()
                # 将星期几转换为星期几的名称
                weekday_name = calendar.day_name[weekday]
                # 打包成字典
                day_info = {
                    'day': day_str,
                    'weekday': weekday_name
                }
                days_of_month.append(day_info)
    return days_of_month

if __name__ == '__main__':

    years = list(range(2024,2055))
    year_combo = ttk.Combobox(root, values=years,width=5)
    year_combo.current(0)  # 设置默认选择为"一月"
    year_combo.configure(state="readonly")
    year_combo.grid(column=0, row=0)

    months = ["1月", "2月", "3月", "4月", "5月", "6月",
          "7月", "8月", "9月", "10月", "11月", "12月"]
    month_combo = ttk.Combobox(root, values=months,width=5)
    month_combo.current(0)  # 设置默认选择为"一月"
    month_combo.configure(state="readonly")
    month_combo.grid(column=1, row=0)
    
    # select_button = tk.Button(root, text="选择年月", command=select_date_range)
    # select_button.grid(column=2, row=0)

    entry = tk.Entry(root, textvariable=select_path,width=45)
    entry.grid(column=0, row=1)
    entry.configure(state="readonly")
    tk.Button(root, text="选择排班表", command=select_file).grid(row=1, column=2)

    button = tk.Button(root, text="1、生成加班调休明细表",command=generate_excel)
    button.grid(row=5, column=1, sticky="EWNS",pady=50)  # 使Button在row=1, column=1的位置，sticky选项使其在水平和垂直方向上扩展

    root.mainloop()

