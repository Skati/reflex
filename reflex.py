#!/usr/bin/env python3
# -*- coding: utf-8 -*
import os
import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
import re
from scipy.optimize import fsolve,root,brentq
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import BPoly, CubicSpline
import warnings


total_xls_path = './excel/total.xlsx'
file_alpha='./excel/alpha.xlsx'
files_dir = './data/'
lst_files = os.listdir(files_dir)

if not os.path.exists('./excel'):
    os.makedirs('./excel')
if not os.path.exists('./data'):
    os.makedirs('./data')

def write_xls(file_path):
    # headers for df
    headers1 = ['f', 'alpha_75', 'Y1_75', 'R1_75']
    headers2 = ['alpha_80', 'Y1_80', 'R1_80']
    headers3 = ['alpha_85', 'Y1_85', 'R1_85']
    headers4 = ['alpha_90', 'Y1_90', 'R1_90']
    headers5 = ['alpha_95', 'Y1_95', 'R1_95']
    headers = headers1+headers2+headers3+headers4+headers5
    writer = pd.ExcelWriter(file_path)# https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    workbook = writer.book

    for file in lst_files:
        file_name = files_dir+file
        sheet_name = file[:30].split('.')[0]  # ограниченя экселя((

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
'''
old version
def find_resonance(file_path,file_alpha):
    alphas=['alpha_75','alpha_80','alpha_85','alpha_90','alpha_95']
    y1=['Y1_75','Y1_80','Y1_85','Y1_90','Y1_95']
    xl_file=pd.ExcelFile(file_path)
    sheets=xl_file.sheet_names
    #dfs = {sheet_name: xl_file.parse(sheet_name)for sheet_name in xl_file.sheet_names}
    wb=xlsxwriter.Workbook(file_alpha)
    arr_sheets=list(dict.fromkeys([re.sub(r'\d+', '',sheets[i]) for i in range(len(sheets))]))#name of worksheets
    ws2=wb.add_worksheet('Total_alpha_max')
    ws3=wb.add_worksheet('Total_y_min')
    tot_row=0
    tot_r=1
    
    for a in arr_sheets:#проход по листам
        reg=r'{0}\d+'.format(a)
        matches=re.findall(reg,''.join(sheets))
        ws=wb.add_worksheet(a)
        ws2.write(tot_row+1,0,a)
        for i in range(len(matches)):
            ws2.write(0,i+1,i)
            ws3.write(0,i+1,i)
        ws3.write(tot_row+1,0,a)
        merge_format = wb.add_format({'align': 'center',})
        merge_index=1
        sheet_index=0
        tot_c=1
        print(a)
        print('***************')
        for match in matches:
            print(match)
            ws.merge_range(0,merge_index,0,merge_index+2,match,merge_format) #запись имен файлов со слиянием
            df=pd.read_excel(file_path,sheet_name=match,index_col=0)
            
            df_y=df.iloc[:,[1,4,7,10,13]] #срезы по y,alpha,r
            df_alpha=df.iloc[:,[0,3,6,9,12]]
            df_r=df.iloc[:,[2,5,8,11,14]]
            x_coords=np.array(df_alpha['alpha_75'].keys().tolist())
            y_coords=np.array(df_alpha['alpha_75'].tolist())

            cubic=CubicSpline(x_coords,y_coords)
            bpoly = BPoly.from_derivatives(x_coords,y_coords[:,np.newaxis])
            akima=Akima1DInterpolator(x_coords,y_coords)

            xnew = np.arange(300, 700, 0.1)   
            ynew=cubic(xnew)
            max_index = np.argmax(ynew)
            max_value = ynew[max_index]
            max_x = xnew[max_index]
            
            
            root1 = optimize.fsolve(bpoly, 350.0) #350-стартовая частота
            root2=optimize.fsolve(cubic, 350.0) 
            plt.figure()
            plt.title('alpha')
            plt.plot(x_coords, y_coords, 'o', xnew, ynew,'g',xnew,bpoly(xnew),'r',xnew,akima(xnew))
            plt.legend(['data', 'cubic spline','bpoly','akima'], loc = 'best')
            #plt.show()
                        
            ws.write(1,merge_index,'f(alpha_max)')
            ws.write(1,merge_index+1,'alpha_max')
            ws.write(1,merge_index+2,'R(alpha_max')
            ws.write(13,merge_index,'f(y0)')
            ws.write(13,merge_index+1,'alpha_max')
            ws.write(13,merge_index+2,'R(alpha_max')
            ws.write(8,sheet_index+1,matches[sheet_index])
            #Запись имен листов
            ws.write(20,sheet_index+1,matches[sheet_index])
            i=2#номер строки для записи
            df_y_abs=df_y.abs()
            f_change_a=[]#список резонансных частот по альфа
            f_change_y=[]#список резонансных частот по y
            list_db=[75,80,85,90,95]
            for alpha in alphas:
                #по максимумам alpha
                alpha_max=df_alpha[alpha].max()#max alpha
                alpha_idxmax=df_alpha[alpha].idxmax()#f(max(alpha))
                alpha_x=df_alpha.index.get_loc(alpha_idxmax)#indexis
                alpha_y=df_alpha.columns.get_loc(alpha)
                r_alpha=df_r.iloc[alpha_x,alpha_y]#r(max(alpha))
                ws.write(i,0, alpha[6:]+'dB')
                ws.write(i+12,0, alpha[6:]+'dB')
                ws.write(i,merge_index,alpha_idxmax)   
                ws.write(i,merge_index+1,alpha_max)
                ws.write(i,merge_index+2,r_alpha)    
                f_change_a.append(alpha_idxmax)
                i+=1
            
            for f in f_change_a:
                if f>f_change_a[0]:
                    report='Частота повысилась'
                    symbol=u'\u2191'
                    break
                elif f<f_change_a[0]:
                    report='Частота понизилась'
                    symbol=u'\u2193'
                    break
                else:
                    report='Частота не изменилась'
                    symbol=u'\u2192'
                    continue
    
                return f,report,symbol

            ws.write(9,sheet_index+1,list_db[f_change_a.index(f)])       
            ws.write(10,sheet_index+1,report)  
            ws.write(11,sheet_index+1,symbol) 
            ws2.write(tot_r,tot_c,list_db[f_change_a.index(f)])
            
            
                 #-------------------------по пересечению нуля
            j=14#№ строки заголовков
            for y in y1:
                y_idxmin=df_y_abs[y].idxmin()
                y_x=df_y_abs.index.get_loc(y_idxmin)
                y_y=df_y.columns.get_loc(y)
                
                alpha_y=df_alpha.iloc[y_x,y_y]
                r_y=df_r.iloc[y_x,y_y]

                ws.write(j,merge_index,y_idxmin)   
                ws.write(j,merge_index+1,alpha_y)
                ws.write(j,merge_index+2,r_y)  
                j+=1 
                f_change_y.append(y_idxmin)
            
            for f_y in f_change_y:
                if f_y>f_change_y[0]:
                    report='Частота повысилась'
                    symbol=u'\u2191'
                    break
                elif f_y<f_change_y[0]:
                    report='Частота понизилась'
                    symbol=u'\u2193'
                    break
                else:
                    report='Частота не изменилась'
                    symbol=u'\u2192'
                    continue
                return f_y,report,symbol
            ws.write(21,sheet_index+1,list_db[f_change_y.index(f_y)])       
            ws.write(22,sheet_index+1,report)  
            ws.write(23,sheet_index+1,symbol)
            ws3.write(tot_r,tot_c,list_db[f_change_y.index(f_y)]) 
            merge_index+=3
            sheet_index+=1
            tot_c+=1
        tot_r+=1    
        print('------------------------------------------')    
        tot_row+=1
    wb.close()
'''
def find_resonance(file_path,file_alpha):
    xl_file=pd.ExcelFile(file_path)
    sheets=xl_file.sheet_names
    #dfs = {sheet_name: xl_file.parse(sheet_name)for sheet_name in xl_file.sheet_names}
    wb=xlsxwriter.Workbook(file_alpha)
    arr_sheets=list(dict.fromkeys([re.sub(r'\d+', '',sheets[i]) for i in range(len(sheets))]))#name of worksheets
    ws1=wb.add_worksheet('Total_alpha_max')
    ws2=wb.add_worksheet('Total_R_min')
    ws3=wb.add_worksheet('Total_Y0')
    total_index_row=1
    root_index_row=1
    for arr_sheet in arr_sheets:#проход по фио
        reg=r'{0}\d+'.format(arr_sheet)
        experiments=re.findall(reg,''.join(sheets))#list of experiments by fullname
        ws0=wb.add_worksheet(arr_sheet)#добавление листов по фио
        ws1.write(total_index_row,0,arr_sheet)#запись в листы тотал названий фио
        ws2.write(total_index_row,0,arr_sheet)
        ws3.write(total_index_row,0,arr_sheet)
        merge_format = wb.add_format({'align': 'center',})
        merge_title_col=1#для записи титульников
        data_index_col=1# для записи данных в общую таблицу
        root_index_col=1      
        for experiment in experiments:
            df=pd.read_excel(file_path,sheet_name=experiment,index_col=0)
            #срезы по y,alpha,r
            df_y=df.iloc[:,[1,4,7,10,13]] 
            df_alpha=df.iloc[:,[0,3,6,9,12]]
            df_r=df.iloc[:,[2,5,8,11,14]]

            alpha_names=df_alpha.keys()

            print(experiment)
            print('-------')
            ws0.merge_range(0,merge_title_col,0,merge_title_col+4,experiment,merge_format)#запись имен файлов со слиянием
            ws0.write(1,merge_title_col,'f0')
            ws0.write(1,merge_title_col+1,'alpha_max')
            ws0.write(1,merge_title_col+2,'f(alpha_max)')
            ws0.write(1,merge_title_col+3,'Rmin')
            ws0.write(1,merge_title_col+4,'f(Rmin)')
            ws0.write(8,root_index_col,experiment)
            ws0.write(14,root_index_col,experiment)
            ws0.write(20,root_index_col,experiment)
            #-----------------alpha-----------------------
            pressure_title_index_row=2
            for alpha in alpha_names:
                ws0.write(pressure_title_index_row,0, alpha[6:]+'dB')
                pressure_title_index_row+=1
            data_index_row=2
            
            max_alpha_values=[]
            max_alpha_frequencis=[]
            roots_y=[]
            min_r_values=[]
            min_r_frequencis=[]
            for a_index in range(len(alpha_names)):
                
                #y_coords=np.array(df_alpha[df.columns[a_index]].tolist())             
                x_coords=np.array(df_alpha[df_alpha.columns[a_index]].keys().tolist())
                y_coords_alpha=np.array(df_alpha[df_alpha.columns[a_index]].tolist())
                y_coords_y=np.array(df_y[df_y.columns[a_index]].tolist())
                y_coords_r=np.array(df_r[df_r.columns[a_index]].tolist())
                
                cubic_alpha=CubicSpline(x_coords,y_coords_alpha)
                #cubic_y=CubicSpline(x_coords,y_coords_y)
                cubic_r=CubicSpline(x_coords,y_coords_r)

                xnew = np.arange(300, 675, 0.1)
                ynew_alpha=cubic_alpha(xnew)
                #ynew_y=cubic_y(xnew)
                ynew_r=cubic_r(xnew)
                bpoly = BPoly.from_derivatives(x_coords,y_coords_y[:,np.newaxis],extrapolate=None)

                max_index_alpha = np.argmax(ynew_alpha)
                max_value_alpha = ynew_alpha[max_index_alpha]
                max_x_alpha = xnew[max_index_alpha]
                
                min_index_r = np.argmin(ynew_r)
                min_value_r = ynew_r[min_index_r]
                min_x_r = xnew[min_index_r]
                #root_y = brentq(bpoly, 350, max_x_alpha+50)
                try:
                    root_y = brentq(bpoly, 350, max_x_alpha+50)# Нормуль!
                except ValueError:
                    root_y=0
                ws0.write(data_index_row,data_index_col,root_y)
                ws0.write(data_index_row,data_index_col+1,max_value_alpha)
                ws0.write(data_index_row,data_index_col+2,max_x_alpha)
                ws0.write(data_index_row,data_index_col+3,min_value_r)
                ws0.write(data_index_row,data_index_col+4,min_x_r)             
                '''
                fig, ax1 = plt.subplots()
                ax2 = ax1.twinx()

                ax1.axhline(y=0, color='r', linestyle='-')   
                ax2.axhline(y=max_value_alpha,color='r',linestyle='-')            
                ax1.plot(x_coords, y_coords_y, 'o', xnew, bpoly(xnew),'-')
                ax2.plot(x_coords,y_coords_alpha,'o',xnew,ynew_alpha,'-')
                ax1.plot(root_y, 0, marker='o', markersize=4, color="red")
                ax1.annotate('zero = (%.1f, %.1f)'%(root_y, 0),xy=(root_y, 0),xytext=(root_y, 0))                                                                         
                ax2.plot(max_x_alpha,max_value_alpha,marker='o', markersize=4, color="red")
                ax2.annotate('max_alpha = (%.1f, %.1f)'%(max_x_alpha,max_value_alpha),xy=(max_x_alpha,max_value_alpha),xytext=(max_x_alpha,max_value_alpha))   
                #ax2.annotate('display = (%.1f, %.1f)'%(max_x_alpha,max_value_alpha)) 
                ax1.legend(['zero','data_y','spline_y','y0'], loc = 'best')
                ax2.legend(['zero','data_alpha','spline_alpha','alpha_max'], loc = 'best')
                fig.tight_layout() 
                plt.show()'''
                max_alpha_values.append('%0.2f' %max_value_alpha)
                max_alpha_frequencis.append('%0.1f' % max_x_alpha)
                roots_y.append('%0.1f' % root_y)
                min_r_values.append('%0.2f' %min_value_r)
                min_r_frequencis.append('%0.1f' %min_x_r)
                data_index_row+=1
            print(max_alpha_values)
            print(max_alpha_frequencis)
            print(roots_y)
            print(min_r_values)
            print(min_r_frequencis)#TODO вернуть 1 частоту если не изменилась
            for f_y in roots_y:
                if float(f_y)>=float(roots_y[0])+5:
                    report='Частота повысилась'
                    symbol=u'\u2191'
                    break
                elif float(f_y)+5<=float(roots_y[0]):
                    report='Частота понизилась'
                    symbol=u'\u2193'
                    break
                else:
                    report='Частота не изменилась'
                    symbol=u'\u2192'
                    continue
                return f_y,report,symbol
            print(type(f_y))
            if f_y=='0.0':
                ws0.write(9,root_index_col,'-')       
                ws0.write(10,root_index_col,'-')  
                ws0.write(11,root_index_col,'-')
                ws0.write(12,root_index_col,'-')
                ws3.write(root_index_row,root_index_col,'-')
            else:
                ws0.write(9,root_index_col,f_y)       
                ws0.write(10,root_index_col,report)  
                ws0.write(11,root_index_col,symbol)
                ws0.write(12,root_index_col,alpha_names[roots_y.index(f_y)][6:])
                ws3.write(root_index_row,root_index_col,alpha_names[roots_y.index(f_y)][6:])

            for f_alpha in max_alpha_frequencis:
                if float(f_alpha)>float(max_alpha_frequencis[0])+5:
                    report='Частота повысилась'
                    symbol=u'\u2191'
                    break
                elif float(f_alpha)+5<float(max_alpha_frequencis[0]):
                    report='Частота понизилась'
                    symbol=u'\u2193'
                    break
                else:
                    report='Частота не изменилась'
                    symbol=u'\u2192'
                    continue
                return f_alpha,report,symbol
            ws0.write(15,root_index_col,f_alpha)       
            ws0.write(16,root_index_col,report)  
            ws0.write(17,root_index_col,symbol)
            ws0.write(18,root_index_col,alpha_names[max_alpha_frequencis.index(f_alpha)][6:])
            ws1.write(root_index_row,root_index_col,alpha_names[max_alpha_frequencis.index(f_alpha)][6:])
            for f_r in min_r_frequencis:
                if float(f_r)>float(min_r_frequencis[0])+5:
                    report='Частота повысилась'
                    symbol=u'\u2191'
                    break
                elif float(f_r)+5<float(min_r_frequencis[0]):
                    report='Частота понизилась'
                    symbol=u'\u2193'
                    break
                else:
                    report='Частота не изменилась'
                    symbol=u'\u2192'
                    continue
                return f_r,report,symbol
            ws0.write(21,root_index_col,f_r)       
            ws0.write(22,root_index_col,report)  
            ws0.write(23,root_index_col,symbol)
            ws0.write(24,root_index_col,alpha_names[min_r_frequencis.index(f_r)][6:])
            ws2.write(root_index_row,root_index_col,alpha_names[min_r_frequencis.index(f_r)][6:])
            root_index_col+=1
            #root_index_row+=1
            data_index_col+=5
            merge_title_col+=5            
        root_index_row+=1
        print('*********************************************************')    
        total_index_row+=1
    wb.close()

# TODO: максимум альфа, пересение У
write_xls(total_xls_path)
find_resonance(total_xls_path,file_alpha)
print('Done !')
