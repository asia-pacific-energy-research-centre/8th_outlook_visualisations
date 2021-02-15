import pandas as pd
import numpy as np
import os
import xlsxwriter

# https://xlsxwriter.readthedocs.io/example_chart_line.html

# Macro data file is saved in: 8th_outlook_visualisations\data\5_Macro\Macro data.xlsx

# read in the Excel file and assign each sheet to a dictionary key:value pair
_path = "../../../data/5_Macro/Macro data.xlsx"
_dict = pd.read_excel(_path,sheet_name=None,index_col=0)

# assign each sheet to a dataframe
df_GDP = _dict['GDP'].drop(columns=['Notes','Unit'])
df_Pop = _dict['Population'].drop(columns=['UN Scenario','Unit'])
df_GDPpercap = _dict['GDP per capita'].drop(columns=['UN Scenario','Unit'])

workbook = xlsxwriter.Workbook('../../../results/Macro charts.xlsx')
bold = workbook.add_format({'bold': 1})

for i,row in df_GDP.iterrows():
    worksheet = workbook.add_worksheet(name=i)
    # == GDP chart ==
    worksheet.write_column('A1',['Year','GDP','Population','GDP per capita'],bold)
    worksheet.write_row('B1',row.index,bold)
    worksheet.write_row('B2',row.values)
    _chart = workbook.add_chart({'type':'line'})
    _chart.add_series({
        'categories': i+'!$B$1:$CD$1',
        'values': i+'!$B$2:$CD$2'
    })
    _chart.set_legend({'none': True})
    _chart.set_title({'name': 'GDP'})
    _chart.set_size({
        'width': 500,
        'height': 300
        })
    _chart.set_chartarea({
            'border': {'none': True}
        })
    worksheet.insert_chart('B7',_chart)

    _chart.set_x_axis({
        #'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
print('wrote GDP')

for i,row in df_Pop.iterrows():
    # == Population chart ==
    worksheet = workbook.get_worksheet_by_name(i)   # access sheets created in the GDP loop
    worksheet.write_row('B3',row.values)
    _chart = workbook.add_chart({'type':'line'})
    _chart.add_series({                             # add the data
        'categories': i+'!$B$1:$CD$1',
        'values': i+'!$B$3:$CD$3'
    })
    _chart.set_legend({'none': True})               # style
    _chart.set_title({'name': 'Population'})
    _chart.set_size({
        'width': 500,
        'height': 300
        })
    _chart.set_chartarea({
            'border': {'none': True}
        })
    worksheet.insert_chart('K7',_chart)

    _chart.set_x_axis({
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
print('wrote Population')

# add GDP per capita here

workbook.close()    # close the workbook