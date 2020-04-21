#!/usr/bin/env python3
import pandas as pd
import os
import xlsxwriter
from openpyxl import load_workbook

total_xls_path = './excel/total.xlsx'
files_dir = './data/'
lst_files = os.listdir(files_dir)


def write_xls(file_path):
    # headers for df
    headers1 = ['f', 'alpha_75', 'Y1_75', 'R1_75']
    headers2 = ['alpha_80', 'Y1_80', 'R1_80']
    headers3 = ['alpha_85', 'Y1_85', 'R1_85']
    headers4 = ['alpha_90', 'Y1_90', 'R1_90']
    headers5 = ['alpha_95', 'Y1_95', 'R1_95']
    headers = headers1+headers2+headers3+headers4+headers5
    writer = pd.ExcelWriter(file_path)
    workbook = writer.book

    for file in lst_files:
        file_name = files_dir+file
        sheet_name = file[:30]  # ограниченя экселя((

        data1 = pd.read_csv(file_name, encoding='windows-1251', skiprows=35,
                            nrows=16, header=None, decimal=',', sep=";", usecols=[0, 5, 6, 7],names=headers1)
        data2 = pd.read_csv(file_name, encoding='windows-1251', skiprows=56,
                            nrows=16, header=None, decimal=',', sep=";", usecols=[5, 6, 7],names=headers2)
        data3 = pd.read_csv(file_name, encoding='windows-1251', skiprows=77,
                            nrows=16, header=None, decimal=',', sep=";", usecols=[5, 6, 7],names=headers3)
        data4 = pd.read_csv(file_name, encoding='windows-1251', skiprows=98,
                            nrows=16, header=None, decimal=',', sep=";", usecols=[5, 6, 7],names=headers4)
        data5 = pd.read_csv(file_name, encoding='windows-1251', skiprows=119,
                            nrows=16, header=None, decimal=',', sep=";", usecols=[5, 6, 7],names=headers5)
        data = pd.concat([data1, data2, data3, data4, data5], axis=1)
        data.index=data['f']
        data.to_excel(writer, sheet_name=sheet_name,
                      header=headers, index=False)
        worksheet = writer.sheets[sheet_name]
        # make charts
        chart1 = workbook.add_chart({'type': 'line'})
        chart2 = workbook.add_chart({'type': 'line'})
        chart3 = workbook.add_chart({'type': 'line'})

        for i in range(2, 18, 3):
            chart1.add_series({
                'name':       [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, 16, 0],
                'values':     [sheet_name, 1, i, 16, i],
                'line':   {'width': 1.25}
            })
            chart2.add_series({
                'name':       [sheet_name, 0, i-1],
                'categories': [sheet_name, 1, 0, 16, 0],
                'values':     [sheet_name, 1, i-1, 16, i-1],
            })
            chart3.add_series({
                'name':       [sheet_name, 0, i+1],
                'categories': [sheet_name, 1, 0, 16, 0],
                'values':     [sheet_name, 1, i+1, 16, i+1],
            })

        chart1.set_title({'name': 'Y','name_font': {'size': 14, 'bold': False}})
        chart1.set_x_axis({
            'name': 'f, Гц',
            'name_font': {'size': 14, 'bold': False},
             })
        chart1.set_style(2)

        chart2.set_title({'name': 'alpha','name_font': {'size': 14, 'bold': False}})
        chart2.set_x_axis({
            'name': 'f, Гц',
            'name_font': {'size': 14, 'bold': False},
             })
        chart2.set_style(2)
        chart3.set_title({'name': 'R','name_font': {'size': 14, 'bold': False}})
        chart3.set_x_axis({
            'name': 'f, Гц',
            'name_font': {'size': 14, 'bold': False},
             })
        chart3.set_style(2)
        # Insert the chart into the worksheet
        worksheet.insert_chart('A19', chart1, {'x_offset': 0, 'y_offset': 0})
        chart1.set_size({'width': 448, 'height': 300})#320x 226
        worksheet.insert_chart('A34', chart2, {'x_offset': 0, 'y_offset': 0})
        chart2.set_size({'width': 448, 'height': 300})#320x 226
        worksheet.insert_chart('A49', chart3, {'x_offset': 0, 'y_offset': 0})
        chart3.set_size({'width': 448, 'height': 300})#320x 226
    # сортировка листов
    workbook.worksheets_objs.sort(key=lambda x: x.name)
    writer.save()
    writer.close()

def find_resonance(file_path):
    alpha=['alpha_75','alpha_80','alpha_85','alpha_90','alpha_95']
    xl=pd.ExcelFile(file_path)
    sheets=xl.sheet_names
    for sheet in sheets:
        print(sheet)
        df=pd.read_excel(open(file_path, 'rb'),
              sheet_name=sheet,index_col=0)  
        alpha_max=[]
        f_alpha_max=[]
        for a in [alpha]:
            alpha_max.append(df[a].max())
            f_alpha_max.append(df[a].idxmax())
        print(alpha_max)
        print(f_alpha_max)
    wb = load_workbook(file_path)
    wb.create_sheet('alpha')
    wb.save(file_path)    

# TODO: максимум альфа, пересение У
write_xls(total_xls_path)
find_resonance(total_xls_path)
print('Done !')
