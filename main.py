import tkinter as tk
from tkinter import filedialog,ttk,messagebox
import pandas as pd
import calendar
from datetime import datetime, timedelta
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment,PatternFill,Border, Side,Font
from openpyxl.worksheet.dimensions import ColumnDimension,DimensionHolder
from openpyxl.utils import get_column_letter
import re

root = tk.Tk()
root.geometry("580x300+50+50") # widthxheight+x+y
root.title("工程部/客服部/保安考勤记录生成器")
root.resizable(False,False)

select_path = tk.StringVar() #本月排班表本地路径
select_path_lastmonth = tk.StringVar() # 上个月加班调休明细表本地路径
select_path_dayoff = tk.StringVar() # 本月调休表本地路径
select_path_ot = tk.StringVar() # 本月加班表本地路径

def select_file():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    select_path.set(selected_file_path)
    if len(select_path.get()) > 0 and len(select_path_lastmonth.get()) > 0 and len(select_path_dayoff.get()) > 0 and len(select_path_ot.get()) > 0:
        button1.config(state=tk.ACTIVE)
    messagebox.showwarning("本月排班表文件是", f'{selected_file_path}\n\n如果文件不正确可重新选择')

def select_file_lastmonth():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    select_path_lastmonth.set(selected_file_path)
    if len(select_path.get()) > 0 and len(select_path_lastmonth.get()) > 0 and len(select_path_dayoff.get()) > 0 and len(select_path_ot.get()) > 0:
        button1.config(state=tk.ACTIVE)
    messagebox.showwarning("上个月加班调休明细表文件是", f'{selected_file_path}\n\n如果文件不正确可重新选择')

    # print(get_remaining_hours("李琦琛"))

def select_file_dayoff():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    select_path_dayoff.set(selected_file_path)
    # data_frame = pd.read_excel(select_path_dayoff.get())
    # result = get_dayoff_data("蔡晓曼",data_frame)
    # print(result)
    if len(select_path.get()) > 0 and len(select_path_lastmonth.get()) > 0 and len(select_path_dayoff.get()) > 0 and len(select_path_ot.get()) > 0:
        button1.config(state=tk.ACTIVE)
    messagebox.showwarning("本月调休记录表文件是", f'{selected_file_path}\n\n如果文件不正确可重新选择')

def select_file_ot():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    select_path_ot.set(selected_file_path)
    if len(select_path.get()) > 0 and len(select_path_lastmonth.get()) > 0 and len(select_path_dayoff.get()) > 0 and len(select_path_ot.get()) > 0:
        button1.config(state=tk.ACTIVE)
    messagebox.showwarning("本月加班记录表文件是", f'{selected_file_path}\n\n如果文件不正确可重新选择')

# 根据姓名获取调休记录表的数据
def get_dayoff_data(name) -> dict:
    data_frame = pd.read_excel(select_path_dayoff.get())
    dayoff_dict = dict()
    for row_index in range(data_frame.shape[0]):
        if data_frame.iloc[row_index, 0] == name:
            for j in range(20,len(data_frame.columns)):
               dayoff_data = data_frame.iloc[row_index, j]
               if not pd.isna(dayoff_data):
                   dayoff_data = re.sub(r'[\u4e00-\u9fff]', '', dayoff_data)
                   dayoff_dict[j - 19] = dayoff_data.replace("/","")
            break
    return dayoff_dict

# 根据姓名获取加班记录表的数据
def get_ot_data(name) -> list:
    data_frame = pd.read_excel(select_path_ot.get())
    ot_list = list()
    for row_index in range(data_frame.shape[0]):
        if data_frame.iloc[row_index, 0] == name:
            
            each_date = list()

            date = data_frame.iloc[row_index, 4] # 加班日期
            data_arr = date.split("-")
            each_date.append(int(data_arr[2]))

            hour = data_frame.iloc[row_index, 8] # 加班时长
            each_date.append(float(hour))

            ot_type = data_frame.iloc[row_index, 21] # 加班类型
            each_date.append(ot_type)

            ot_list.append(each_date)

    return ot_list


# 获取上个月加班调休明细表的剩余加班小时数
def get_remaining_hours(name) -> list:
    data_frame = pd.read_excel(select_path_lastmonth.get())
     
    # 初始化行列位置
    row = 1
    column = None
    result_list = ["","","",""]

    # 遍历DataFrame寻找内容
    for j in range(len(data_frame.columns)):
        if data_frame.iloc[row, j] == name:
            column = j
            # print(f'内容找到在行：{row+1}，列：{j+1}')
            break

    # 如果需要获取具体的单元格数据
    if column is not None:
        row_idx = data_frame.shape[0] - 4
   
        for list_index in range(0,4):
            if not pd.isna(data_frame.iloc[row_idx, column + list_index]):
                last_cell_data = float(data_frame.iloc[row_idx, column + list_index])
                if not pd.isna(data_frame.iloc[row_idx + 1, column]):
                    cur_cell_data = float(data_frame.iloc[row_idx + 1, column + list_index])
                    result_list[list_index] = str(cur_cell_data + last_cell_data)
                else:
                    result_list[list_index] = str(last_cell_data)
            else:
                if not pd.isna(data_frame.iloc[row_idx + 1, column + list_index]):
                    cur_cell_data = float(data_frame.iloc[row_idx + 1, column + list_index])
                    result_list[list_index] = str(cur_cell_data)
                else:
                    result_list[list_index] = ''
    # print(result_list)
    return result_list

# 生成调休加班明细表
def generate_excel():
    try:
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

        df = pd.read_excel(select_path.get()) # 排班表的数据
        # print(df)
        
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
                        sheet.cell(row=row_index, column=col).fill = PatternFill(start_color='FFFF00', fill_type='solid') # 把当前行的背景色设置为黄色
        
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
            name_iloc = df.iloc[name_row_index, 0] # 循环获取员工姓名
            if name_iloc == "排班说明": # 表示已经读取完员工姓名
                break
            else: # 一个员工的数据
                sheet.cell(row=3, column=name_column_index + 1).value = name_iloc
                sheet.cell(row=3, column=name_column_index + 1).alignment = Alignment(horizontal="center",vertical="center")
                sheet.cell(row=3, column=name_column_index + 1).border = border
                sheet.cell(row=3, column=name_column_index + 1).font = Font(size=11)
                # 合并“姓名”的单元格
                sheet.merge_cells(start_row=3, start_column=name_column_index + 1, end_row=3, end_column=name_column_index + 4)

                last_month_hours = get_remaining_hours(name_iloc) # 获取上个月的剩余加班小时数

                cell1 = sheet.cell(row=4, column=name_column_index + 1)
                cell1.value = '平时加班'
                cell1.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                cell1.border = border
                cell1.font = Font(size=8)
                sheet.column_dimensions[get_column_letter(cell1.column)].width = 4.2
                set_cell_border(row_index,sheet,name_column_index + 1,border)
                cell_total_1 = sheet.cell(row=row_index - 4, column=name_column_index + 1) 
                cell_total_1.value = f'=SUM({cell1.column_letter}5:{cell1.column_letter}{row_index - 5})'
                # 写入上个月的平时加班小时数
                ordinary_ot_cell = sheet.cell(row=row_index - 3, column=name_column_index + 1)
                ordinary_ot_cell.value = last_month_hours[0]
    

                cell2 = sheet.cell(row=4, column=name_column_index + 2)
                cell2.value = '法定加班'
                cell2.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                cell2.border = border
                cell2.font = Font(size=8)
                sheet.column_dimensions[get_column_letter(cell2.column)].width = 4.2
                set_cell_border(row_index,sheet,name_column_index + 2,border)
                cell_total_2 = sheet.cell(row=row_index - 4, column=name_column_index + 2) 
                cell_total_2.value = f'=SUM({cell2.column_letter}5:{cell2.column_letter}{row_index - 5})'
                # 写入上个月的法定加班小时数
                statutory_ot_cell = sheet.cell(row=row_index - 3, column=name_column_index + 2)
                statutory_ot_cell.value = last_month_hours[1]
    

                cell3 = sheet.cell(row=4, column=name_column_index + 3)
                cell3.value = '周末加班'
                cell3.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                cell3.border = border
                cell3.font = Font(size=8)
                sheet.column_dimensions[get_column_letter(cell3.column)].width = 4.2
                set_cell_border(row_index,sheet,name_column_index + 3,border)
                cell_total_3 = sheet.cell(row=row_index - 4, column=name_column_index + 3) 
                cell_total_3.value = f'=SUM({cell3.column_letter}5:{cell3.column_letter}{row_index - 5})'
                # 写入上个月的周末加班小时数
                weekend_ot_cell = sheet.cell(row=row_index - 3, column=name_column_index + 3)
                weekend_ot_cell.value = last_month_hours[2]


                cell4 = sheet.cell(row=4, column=name_column_index + 4)
                cell4.value = '调休'
                cell4.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
                cell4.border = border
                cell4.font = Font(size=8)
                sheet.column_dimensions[get_column_letter(cell4.column)].width = 4.2
                set_cell_border(row_index,sheet,name_column_index + 4,border)
                cell_total_4 = sheet.cell(row=row_index - 4, column=name_column_index + 4) 
                cell_total_4.value = f'=SUM({cell4.column_letter}5:{cell4.column_letter}{row_index - 5})'
            
                # 合并“签名”的单元格
                sheet.cell(row=row_index - 1, column=name_column_index + 1).border = border
                sheet.merge_cells(start_row=row_index - 1, start_column=name_column_index + 1, end_row=row_index - 1, end_column=name_column_index + 4)

                # 填入每个人的调休时间
                dayoff_dict = get_dayoff_data(name_iloc)
                if len(dayoff_dict) > 0:
                    for key,val in dayoff_dict.items():
                        sheet.cell(row=4 + key, column=name_column_index + 4).value = float(val)

                # 填入每个人的加班时间
                ot_list = get_ot_data(name_iloc)
                if len(ot_list) > 0: # 该员工有加班
                    for each_ot_data in ot_list:
                        if each_ot_data[2] == '节假日': # 法定加班
                            sheet.cell(row=4 + each_ot_data[0], column=name_column_index + 2).value = each_ot_data[1]
                        if each_ot_data[2] == '公休日': # 周末加班
                            sheet.cell(row=4 + each_ot_data[0], column=name_column_index + 3).value = each_ot_data[1]
                        if each_ot_data[2] == '工作日': # 平时加班
                            sheet.cell(row=4 + each_ot_data[0], column=name_column_index + 1).value = each_ot_data[1]

                '''
                rest_hours = 0 #一个员工当月的调休小时数
                for date_col_index in range(1,32):
                    result_iloc = df.iloc[name_row_index, date_col_index]
                    if result_iloc == "调休":
                        name_value = df.iloc[name_row_index, 0]
                        date_value = df.iloc[1, date_col_index]
                        # day_value = df.iloc[2, date_col_index]
                        for cell in sheet[3]:
                            if cell.value == name_value:
                                sheet.cell(row=4 + date_value, column=cell.column + 3).value = 8
                                rest_hours = rest_hours + 8
                                    
                    if result_iloc == "加班":
                        name_value = df.iloc[name_row_index, 0]
                        date_value = df.iloc[1, date_col_index]
                        day_value = df.iloc[2, date_col_index] 
                
                        for cell in sheet[3]:
                            if cell.value == name_value:
                                if day_value == "六" or day_value == "日":
                                    sheet.cell(row=4 + date_value, column=cell.column + 2).value = 8
                                else:
                                    sheet.cell(row=4 + date_value, column=cell.column).value = 8
                '''
    
                name_column_index = name_column_index + 4
                name_row_index = name_row_index + 1

 
        recalculate_left_hours(sheet)
   

        selected_type = type_combo.get()
        content = f"{selected_type}{selected_year}年{selected_month}月加班调休明细表"
        sheet.cell(row=1, column=1).value = content
        sheet.cell(row=1, column=1).alignment = Alignment(horizontal="center",vertical="center")
        sheet.cell(row=1, column=1).font = Font(size=18)
        sheet.merge_cells(start_row=1, start_column=1, end_row=2, end_column=sheet.max_column)
        global file_name
        file_name = f"{content}.xlsx"
        workbook.save(file_name)
        
        messagebox.showinfo("提示", f"生成文件【{file_name}】成功")
     
    except Exception as e:
        print(e)
        messagebox.showerror("错误", "生成文件失败，请检查选择的文件内容是否正确!原因：" + repr(e))

    finally:
        workbook.close()
       
# 根据调休小时数计算剩余加班小时数
def recalculate_left_hours(write_sheet):
   
    # 读Excel文件用来取数据
    # workbook = load_workbook(file_name,read_only=True,data_only=True)
    # sheet = workbook.active
    # workbook.close()

    # 读Excel文件用来写数据
    # write_workbook = load_workbook(file_name)
    # write_sheet = write_workbook.active

    # print(sheet.max_row)
    # print(sheet.max_column)
    # 把本月剩余的加班小时数抄过来
    continue_count = 0
    for col_iter_index in range(3,write_sheet.max_column + 1):
        continue_count = continue_count + 1
        if continue_count % 4 == 0: # 调休数据直接跳过
            continue
        current_total_hours = write_sheet.cell(row = write_sheet.max_row - 4,column = col_iter_index).value
        current_total_hours = current_total_hours.replace("=SUM(","").replace(")","")
        # print(f'********{current_total_hours}******')

        sum_value = sum(cell.value for row in write_sheet[current_total_hours] for cell in row if cell.value is not None)
        if sum_value is not None:
            write_sheet.cell(row = write_sheet.max_row - 2,column = col_iter_index).value = sum_value
    
    # 计算每个员工的剩余加班小时数
    employees_count = (write_sheet.max_column - 3) / 4 # 员工数量
    start_col_index = 3 # 从第三列开始
    for _ in range(0,int(employees_count)):
        hours_data = list() # 员工的小时数数据
        for j in range(0,4):
            earch_hour = write_sheet.cell(row = write_sheet.max_row - 3,column = start_col_index + j).value
            hours_data.append(earch_hour if earch_hour is not None else "")
            if j == 3:
                rest_hour = write_sheet.cell(row = write_sheet.max_row - 4,column = start_col_index + j).value
                rest_hour = rest_hour.replace("=SUM(","").replace(")","")
                total_hours_data = sum(cell.value for row in write_sheet[rest_hour] for cell in row if cell.value is not None)
                hours_data.append(str(total_hours_data) if total_hours_data is not None else "")
        
        # 根据调休时间计算一个员工的剩余加班时间
        hours_data = cal_remaining_hours(0,hours_data)
        
        # employee_name = write_sheet.cell(row = 3,column = start_col_index).value
        # print(f"{employee_name}+++原始+++{hours_data}")

        if float(hours_data[4]) > 0: # 上个月剩余的加班小时数不够扣调休小时数
            for l in range(0,4):
                earch_hour = write_sheet.cell(row = write_sheet.max_row - 4,column = start_col_index + l).value
                earch_hour = earch_hour.replace("=SUM(","").replace(")","")
                total_earch_hour = sum(cell.value for row in write_sheet[earch_hour] for cell in row if cell.value is not None)
                hours_data[l] = str(total_earch_hour) if total_earch_hour is not None else ""
            # print(f"{employee_name}+++修改前+++{hours_data}")    
            hours_data = cal_remaining_hours(0,hours_data) # 用本月的加班小时数扣调休小时数
            # print(f"{employee_name}+++修改后+++{hours_data}")
            for n in range(0,3):
                write_sheet.cell(row = write_sheet.max_row - 2,column = start_col_index + n).value = str(hours_data[n])
        else: # 上个月剩余的加班小时数够扣调休小时数,直接更新本月剩余加班小时数
            for m in range(0,3):
                write_sheet.cell(row = write_sheet.max_row - 3,column = start_col_index + m).value = str(hours_data[m])
        

        start_col_index = start_col_index + 4

    # write_workbook.save(file_name)
    # messagebox.showinfo("提示", "计算剩余小时数成功")
    for col_iter_index in range(3,write_sheet.max_column + 1):
        current_last_hour = write_sheet.cell(row = write_sheet.max_row - 2,column = col_iter_index).value
        if current_last_hour is not None and float(current_last_hour) == 0:
            write_sheet.cell(row = write_sheet.max_row - 2,column = col_iter_index).value = ""
  

'''
 递归用剩余加班小时数减去调休小时数

last_remaining_hours的前4个元素是上个月或者本月的加班小时数，第5个元素是调休小时数
'''
def cal_remaining_hours(index,last_remaining_hours) -> list:
    if index <= 3:
        if last_remaining_hours[index] != "":
            if float(last_remaining_hours[index]) >= float(last_remaining_hours[4]):
                last_remaining_hours[index] = str(float(last_remaining_hours[index]) - float(last_remaining_hours[4]))
                last_remaining_hours[4] = '0'
            else:
                last_remaining_hours[4] = str(float(last_remaining_hours[4]) - float(last_remaining_hours[index]))
                last_remaining_hours[index] = '0'
                cal_remaining_hours(index + 1,last_remaining_hours)
             
        else:
            cal_remaining_hours(index + 1,last_remaining_hours)

    return last_remaining_hours

# 每个员工具体每一天的单元格设置边框
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

def button1_click():
    generate_excel()
    recalculate_left_hours()


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
    
    types = ["工程部", "客服部", "保安"]
    type_combo = ttk.Combobox(root, values=types,width=5)
    type_combo.current(0)  
    type_combo.configure(state="readonly")
    type_combo.grid(column=2, row=0)

    entry = tk.Entry(root, textvariable=select_path,width=45)
    entry.grid(column=0, row=1)
    entry.configure(state="readonly")
    tk.Button(root, text="选择本月排班表", command=select_file).grid(row=1, column=1)

    entry = tk.Entry(root, textvariable=select_path_lastmonth,width=45)
    entry.grid(column=0, row=2)
    entry.configure(state="readonly")
    tk.Button(root, text="选择上个月加班调休明细表", command=select_file_lastmonth).grid(row=2, column=1)

    entry = tk.Entry(root, textvariable=select_path_dayoff,width=45)
    entry.grid(column=0, row=3)
    entry.configure(state="readonly")
    tk.Button(root, text="选择本月调休记录表", command=select_file_dayoff).grid(row=3, column=1)

    entry = tk.Entry(root, textvariable=select_path_ot,width=45)
    entry.grid(column=0, row=4)
    entry.configure(state="readonly")
    tk.Button(root, text="选择本月加班记录表", command=select_file_ot).grid(row=4, column=1)

    description = '''
    1、选择正确的年/月/部门。
    2、根据按钮提示选择正确的Excel文件。
    3、点击“1、生成加班调休明细表”按钮，然后打开生成的明细表Excel文件（如：工程部2024年9月加班调休明细表.xlsx）,文件的数据即是正确的数据。
    '''
    text = tk.Text(root, font=("Helvetica", 10), fg="blue",width=50,height=7)# 设定文本内容、字体、字号、字体颜色 
    text.grid(row=5, column=0, sticky="EWNS",pady=20) # sticky选项使其在水平和垂直方向上扩展
    text.insert("insert",description)
    text['state'] = 'disabled'

    button1 = tk.Button(root, text="1、生成加班调休明细表",command=generate_excel)
    button1.grid(row=5, column=1, pady=20)  # 使Button在row=1, column=1的位置
    button1.config(state=tk.DISABLED)
    
    # button2 = tk.Button(root, text="2、计算剩余加班小时数",command=recalculate_left_hours)
    # button2.grid(row=6, column=1, pady=10)
    # button2.config(state=tk.DISABLED)

    root.mainloop()
 

