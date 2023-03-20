import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re
import xlrd

# 创建Tkinter窗口
window = tk.Tk()
window.title('出港航班信息统计程序')
window.geometry('300x200')

# 创建选择文件按钮
def choose_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        process_flight_data(file_path)
        

button = tk.Button(window, text='选择当日航班数据表', command=choose_file)
button.pack()

# 创建数据处理函数
def process_flight_data(file_path):
    # 读取Excel表格数据
    
    df = pd.read_excel(file_path)

    # 筛选航班号中不包含'CA'和'ZH'的行
    df_filtered = df[~df['出港航班号'].str.contains('CA|ZH|SC|CZ')]

    # 删除包含'CA'和'ZH'的行
    df_filtered.drop(df_filtered[df_filtered['出港航班号'].str.contains('CA|ZH|SC|CZ')].index, inplace=True)

    # 筛选航班属性为国内或国际，航班性质为正班、补班、加班的航班
    df_filtered = df_filtered[(df_filtered['属性'].str.contains('国内|国际|地区')) & (df_filtered['任务'].str.contains('正班|补班|加班|旅包'))]

    #只选择出港航班
    #df_filtered = df_filtered[(df_filtered['进出'].isin(['出港']))]
    df_filtered = df_filtered[df_filtered['出港航班号'] != '-']


    #提取前两位
    df_filtered['航司'] = df_filtered['出港航班号'].str[:2]
    
    #-----------------------------春秋航班提取配载数据------------------------------------#
    for index , row in df_filtered.iterrows():
        
        if pd.isna(row['配平备注']):
            continue  # 如果配平备注
        
        if '9C' in row['航司']:
        # 获取配平备注前三位
            passenger_str = row['配平备注'][:3]
            
            # 判断前三位是否全为数字
            if passenger_str.isdigit():
                
                passenger_num = int(passenger_str)
                
            else:
                # 获取配平备注前两位
                passenger_num = int(row['配平备注'][:2])
            # 更新旅客人数列
            df_filtered.loc[index, '旅客人数'] = round(passenger_num)
            
    #---------------------机位备注------------------------#
    df_filtered['机位备注'] = ''
    
    for index, row in df_filtered.iterrows():
        
        if pd.isna(row['机位']):
            continue  # 如果配平备注
        
        park_str = row['机位']
        
        if park_str[:3].isdigit():
            park_num = int(park_str[:3])
            if park_num < 300 or (park_num > 362 and park_num < 400):
                df_filtered.loc[index, '机位备注'] = 2
            elif park_num > 500 and park_num <= 590:
                df_filtered.loc[index, '机位备注'] = 5
            elif park_num >= 301 and park_num <= 362:
                df_filtered.loc[index, '机位备注'] = 1
            else:
                df_filtered.loc[index, '机位备注'] = 0
                
        if '国际' in row['属性']:
            df_filtered.loc[index, '机位备注'] = 3
         
        
            
        if pd.isna(row['登机口']):
            continue  # 如果配平备注
        gate_str = row['登机口']
            
        if isinstance(gate_str, str) and gate_str[:3].isdigit():
            
            gate_num = int(gate_str[:3])
                
            
            
        elif isinstance(gate_str, str) and gate_str[:2].isdigit():
                
            gate_num = int(gate_str[:2])
            
            
        
                
        if gate_num != park_num:
            
            if park_num - 300 != gate_num:
                        
                df_filtered.loc[index, '机位备注'] = 2
                
        if gate_num > 10 and gate_num <= 13:
                df_filtered.loc[index, '机位备注'] = 3
                

    # 对航班按预计起飞时间进行排序
    df_filtered = df_filtered.sort_values(by='计起')



    new_df = df_filtered[['任务', '属性', '航司', '机位','登机口','计起','配平备注','机位备注','旅客人数','机位备注']]

    # 让用户选择保存文件路径和文件名
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("XLSX files", "*.xlsx"), ("All files", "*.*")))
    if save_path:
        # 保存数据表到指定路径下
        new_df.to_excel(save_path, index=False)
        print(f'已成功将数据保存至{save_path}')

        # 显示处理后的数据
        text = tk.Text(window)
        text.insert(tk.END, df_filtered.to_string())
        text.pack()

# 运行Tkinter窗口
window.mainloop()