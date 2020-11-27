# EGEDA TPES plots for each economy

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
import xlsxwriter
import pandas.io.formats.excel

# Import the recently created data frame that joins OSeMOSYS results to EGEDA historical 

EGEDA_years = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA.csv')

# Define unique values for economy, fuels, and items columns

Economy_codes = EGEDA_years.economy.unique()
Fuels = EGEDA_years.fuel_code.unique()
Items = EGEDA_years.item_code_new.unique()

# Colours for charting (to be amended later)

colours = pd.read_excel('./data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

# Define month and year to create folder for saving charts/tables

month_year = pd.to_datetime('today').strftime('%B_%Y')

# Subsets for impending df builds

First_level_fuels = list(Fuels[[0, 9, 17, 24, 45, 49, 50, 51, 60, 76, 77, 78, 79]])
Required_fuels = list(Fuels[[0, 9, 17, 24, 45, 49, 50, 51, 61, 62, 63, 64, 65, 66, 68, 69, 70, 75, 76, 77, 78, 79]])

# Aggregate's of different fuels for TPES/prod charting
Coal_fuels = list(Fuels[[0, 9]])

Oil_fuels = list(Fuels[[17, 24]])

Other_fuels = list(Fuels[[66, 69, 75]])

Renewables_fuels = list(Fuels[[49, 51, 61, 62, 63, 64, 65, 68, 70]])

tpes_items = list(Items[[0, 1, 2, 3, 4, 5, 6, 7, 8]])

Prod_items = tpes_items[:3]

Petroleum_fuels = list(Fuels[[24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 29, 40, 41, 42, 43, 44]])

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written

# Define column chart years
col_chart_years = ['2000', '2010', '2017', '2020', '2030', '2040', '2050']

TPES_agg_fuels = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']
TPES_agg_trade = ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']

# Total Primary Energy Supply fuel breakdown for each economy

########### Build TPES dataframes for each economy providing various breakdowns (by fuel, TPES component, etc)  

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # First data frame: TPES by fuels (and also fourth and sixth dataframe with slight tweaks)
    tpes_df = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                          (EGEDA_years['item_code_new'] == '6_total_primary_energy_supply') &
                          (EGEDA_years['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]
    
    #nrows1 = tpes_df.shape[0]
    #ncols1 = tpes_df.shape[1]

    coal = tpes_df[tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '6_total_primary_energy_supply')
    
    oil = tpes_df[tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '6_total_primary_energy_supply')
    
    renewables = tpes_df[tpes_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '6_total_primary_energy_supply')
    
    others = tpes_df[tpes_df['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '6_total_primary_energy_supply')
    
    tpes_df1 = tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(tpes_df.loc[:, '2000':])].reset_index(drop = True)

    tpes_df1.loc[tpes_df1['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    tpes_df1.loc[tpes_df1['fuel_code'] == '7_nuclear', 'fuel_code'] = 'Nuclear'

    tpes_df1 = tpes_df1[tpes_df1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    nrows4 = tpes_df1.shape[0]
    ncols4 = tpes_df1.shape[1]

    tpes_df2 = tpes_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows6 = tpes_df2.shape[0]
    ncols6 = tpes_df2.shape[1]
    
    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    prod_df = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                          (EGEDA_years['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]
    
    #nrows2 = prod_df.shape[0]
    #ncols2 = prod_df.shape[1]

    coal = prod_df[prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = prod_df[prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = prod_df[prod_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '1_indigenous_production')
    
    others = prod_df[prod_df['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '1_indigenous_production')
    
    prod_df1 = prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(prod_df.loc[:, '2000':])].reset_index(drop = True)

    prod_df1.loc[prod_df1['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    prod_df1.loc[prod_df1['fuel_code'] == '7_nuclear', 'fuel_code'] = 'Nuclear'

    prod_df1 = prod_df1[prod_df1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    nrows5 = prod_df1.shape[0]
    ncols5 = prod_df1.shape[1]

    prod_df2 = prod_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows7 = prod_df2.shape[0]
    ncols7 = prod_df2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    tpes_comp_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                           (EGEDA_years['item_code_new'].isin(tpes_items)) &
                           (EGEDA_years['fuel_code'] == '12_total')]
    
    net_trade = tpes_comp_df1[tpes_comp_df1['item_code_new'].isin(['2_imports', 
                                                                     '3_exports'])].groupby(['economy', 
                                                                                             'fuel_code']).sum().assign(fuel_code = '12_total',
                                                                                                                        item_code_new = 'Net trade')
                           
    bunkers = tpes_comp_df1[tpes_comp_df1['item_code_new'].isin(['4_1_international_marine_bunkers', 
                                                                 '4_2_international_aviation_bunkers'])].groupby(['economy', 
                                                                                                                  'fuel_code']).sum().assign(fuel_code = '12_total',
                                                                                                                                             item_code_new = 'Bunkers')
    
    tpes_comp_df1 = tpes_comp_df1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    tpes_comp_df1.loc[tpes_comp_df1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    tpes_comp_df1.loc[tpes_comp_df1['item_code_new'] == '5_stock_changes', 'item_code_new'] = 'Stock changes'
    
    tpes_comp_df1 = tpes_comp_df1.loc[tpes_comp_df1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True)
    
    nrows3 = tpes_comp_df1.shape[0]
    ncols3 = tpes_comp_df1.shape[1]

    # Imports/exports data frame

    imports_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                              (EGEDA_years['item_code_new'] == '2_imports') & 
                              (EGEDA_years['fuel_code'].isin(Required_fuels))]

    coal = imports_df1[imports_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '2_imports')
    
    # oil = imports_df1[imports_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
    #                                                                                                     item_code_new = '2_imports')
    
    renewables = imports_df1[imports_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '2_imports')
    
    others = imports_df1[imports_df1['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '2_imports')
    
    imports_df1 = imports_df1.append([coal, oil, renewables, others]).reset_index(drop = True)

    imports_df1.loc[imports_df1['fuel_code'] == '3_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    imports_df1.loc[imports_df1['fuel_code'] == '4_petroleum_products', 'fuel_code'] = 'Petroleum products'
    imports_df1.loc[imports_df1['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    imports_df1.loc[imports_df1['fuel_code'] == '7_nuclear', 'fuel_code'] = 'Nuclear'

    imports_df1 = imports_df1[imports_df1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(imports_df1.loc[:, '2000':])]

    nrows8 = imports_df1.shape[0]
    ncols8 = imports_df1.shape[1] 

    imports_df2 = imports_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows12 = imports_df2.shape[0]
    ncols12 = imports_df2.shape[1]                             

    exports_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                              (EGEDA_years['item_code_new'] == '3_exports') & 
                              (EGEDA_years['fuel_code'].isin(Required_fuels))].copy()

    # Change export values to positive rather than negative

    exports_df1[list(exports_df1.columns[3:])] = exports_df1[list(exports_df1.columns[3:])].apply(lambda x: x * -1)

    coal = exports_df1[exports_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '3_exports')
    
    # oil = exports_df1[exports_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
    #                                                                                                     item_code_new = '3_exports')
    
    renewables = exports_df1[exports_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '3_exports')
    
    others = exports_df1[exports_df1['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '3_exports')
    
    exports_df1 = exports_df1.append([coal, oil, renewables, others]).reset_index(drop = True)

    exports_df1.loc[exports_df1['fuel_code'] == '3_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    exports_df1.loc[exports_df1['fuel_code'] == '4_petroleum_products', 'fuel_code'] = 'Petroleum products'
    exports_df1.loc[exports_df1['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    exports_df1.loc[exports_df1['fuel_code'] == '7_nuclear', 'fuel_code'] = 'Nuclear'

    exports_df1 = exports_df1[exports_df1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(exports_df1.loc[:, '2000':])]

    nrows9 = exports_df1.shape[0]
    ncols9 = exports_df1.shape[1]

    exports_df2 = exports_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows13 = exports_df2.shape[0]
    ncols13 = exports_df2.shape[1] 

    # Bunkers dataframe

    bunkers_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                              (EGEDA_years['item_code_new'] == '4_1_international_marine_bunkers') & 
                              (EGEDA_years['fuel_code'].isin(['4_5_gas_diesel_oil', '4_6_fuel_oil']))]

    bunkers_df1 = bunkers_df1[['fuel_code', 'item_code_new'] + list(bunkers_df1.loc[:, '2000':])]

    bunkers_df1.loc[bunkers_df1['fuel_code'] == '4_5_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    bunkers_df1.loc[bunkers_df1['fuel_code'] == '4_6_fuel_oil', 'fuel_code'] = 'Fuel oil'

    nrows10 = bunkers_df1.shape[0]
    ncols10 = bunkers_df1.shape[1]

    bunkers_df2 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                              (EGEDA_years['item_code_new'] == '4_2_international_aviation_bunkers') & 
                              (EGEDA_years['fuel_code'].isin(['4_3_jet_fuel', '4_1_2_aviation_gasoline']))]

    bunkers_df2 = bunkers_df2[['fuel_code', 'item_code_new'] + list(bunkers_df2.loc[:, '2000':])]

    bunkers_df2.loc[bunkers_df2['fuel_code'] == '4_1_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'
    bunkers_df2.loc[bunkers_df2['fuel_code'] == '4_3_jet_fuel', 'fuel_code'] = 'Jet fuel'

    nrows11 = bunkers_df2.shape[0]
    ncols11 = bunkers_df2.shape[1]
    
    # Define directory
    script_dir = './results/' + month_year + '/TPES/'
    results_dir = os.path.join(script_dir, 'economy_breakdown/', economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
    
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_tpes.xlsx', engine = 'xlsxwriter')
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    tpes_df1.to_excel(writer, sheet_name = economy + '_TPES', index = False, startrow = chart_height)
    tpes_df2.to_excel(writer, sheet_name = economy + '_TPES', index = False, startrow = chart_height + nrows4 + 3)
    prod_df1.to_excel(writer, sheet_name = economy + '_prod', index = False, startrow = chart_height)
    prod_df2.to_excel(writer, sheet_name = economy + '_prod', index = False, startrow = chart_height + nrows5 + 3)
    tpes_comp_df1.to_excel(writer, sheet_name = economy + '_TPES_components_I', index = False, startrow = chart_height)
    imports_df1.to_excel(writer, sheet_name = economy + '_TPES_components_I', index = False, startrow = chart_height + nrows3 + 3)
    imports_df2.to_excel(writer, sheet_name = economy + '_TPES_components_I', index = False, startrow = chart_height + nrows3 + nrows8 + 6)
    exports_df1.to_excel(writer, sheet_name = economy + '_TPES_components_I', index = False, startrow = chart_height + nrows3 + nrows8 + nrows12 + 9)
    exports_df2.to_excel(writer, sheet_name = economy + '_TPES_components_I', index = False, startrow = chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12)
    bunkers_df1.to_excel(writer, sheet_name = economy + '_TPES_components_II', index = False, startrow = chart_height)
    bunkers_df2.to_excel(writer, sheet_name = economy + '_TPES_components_II', index = False, startrow = chart_height + nrows10 + 3)

    #ImEx_df1.to_excel(writer, sheet_name = economy + '_TPES_components', index = False, startrow = chart_height + nrows3 + 3)
    
    # Access the workbook
    workbook = writer.book
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
    
    # Access the sheet created using writer above
    worksheet1 = writer.sheets[economy + '_TPES']
    
    # Apply comma format and header format to relevant data rows
    worksheet1.set_column(2, ncols4 + 1, None, comma_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height + nrows4 + 3, None, header_format)
    worksheet1.write(0, 0, economy + ' TPES fuel', cell_format1)

    ################################################################### CHARTS ###################################################################

    # Create a TPES chart
    tpes_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    tpes_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    tpes_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    tpes_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    tpes_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    tpes_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    tpes_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows4):
        tpes_chart2.add_series({
            'name':       [economy + '_TPES', chart_height + i + 1, 0],
            'categories': [economy + '_TPES', chart_height, 2, chart_height, ncols4 - 1],
            'values':     [economy + '_TPES', chart_height + i + 1, 2, chart_height + i + 1, ncols4 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet1.insert_chart('B3', tpes_chart2)

    ######## same chart as above but line

    # TPES line chart
    tpes_chart4 = workbook.add_chart({'type': 'line'})
    tpes_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    tpes_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    tpes_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    tpes_chart4.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    tpes_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    tpes_chart4.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows4):
        tpes_chart4.add_series({
            'name':       [economy + '_TPES', chart_height + i + 1, 0],
            'categories': [economy + '_TPES', chart_height, 2, chart_height, ncols4 - 1],
            'values':     [economy + '_TPES', chart_height + i + 1, 2, chart_height + i + 1, ncols4 - 1],
            'line':       {'color': colours_hex[i], 
                           'width': 1}
        })    
        
    worksheet1.insert_chart('R3', tpes_chart4)

    ###################### Create another TPES chart showing proportional share #################################

    # Create a TPES chart
    tpes_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    tpes_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    tpes_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    tpes_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    tpes_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    tpes_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    tpes_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in TPES_agg_fuels:
        i = tpes_df2[tpes_df2['fuel_code'] == component].index[0]
        tpes_chart3.add_series({
            'name':       [economy + '_TPES', chart_height + nrows4 + i + 4, 0],
            'categories': [economy + '_TPES', chart_height + nrows4 + 3, 2, chart_height + nrows4 + 3, ncols6 - 1],
            'values':     [economy + '_TPES', chart_height + nrows4 + i + 4, 2, chart_height + nrows4 + i + 4, ncols6 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet1.insert_chart('J3', tpes_chart3)

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    worksheet2 = writer.sheets[economy + '_prod']
    
    # Apply comma format and header format to relevant data rows
    worksheet2.set_column(2, ncols5 + 1, None, comma_format)
    worksheet2.set_row(chart_height, None, header_format)
    worksheet2.set_row(chart_height + nrows5 + 3, None, header_format)
    worksheet2.write(0, 0, economy + ' prod fuel', cell_format1)

    ###################### Create another PRODUCTION chart with only 6 categories #################################

    # Create a PROD chart
    prod_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    prod_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    prod_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    prod_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    prod_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    prod_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    prod_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows5):
        prod_chart2.add_series({
            'name':       [economy + '_prod', chart_height + i + 1, 0],
            'categories': [economy + '_prod', chart_height, 2, chart_height, ncols5 - 1],
            'values':     [economy + '_prod', chart_height + i + 1, 2, chart_height + i + 1, ncols5 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet2.insert_chart('B3', prod_chart2)

    ############ Same as above but with a line ###########

    # Create a PROD chart
    prod_chart2 = workbook.add_chart({'type': 'line'})
    prod_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    prod_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    prod_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    prod_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    prod_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    prod_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows5):
        prod_chart2.add_series({
            'name':       [economy + '_prod', chart_height + i + 1, 0],
            'categories': [economy + '_prod', chart_height, 2, chart_height, ncols5 - 1],
            'values':     [economy + '_prod', chart_height + i + 1, 2, chart_height + i + 1, ncols5 - 1],
            'line':       {'color': colours_hex[i],
                           'width': 1} 
        })    
        
    worksheet2.insert_chart('R3', prod_chart2)

    ###################### Create another PRODUCTION chart showing proportional share #################################

    # Create a production chart
    prod_chart3 = workbook.add_chart({'type': 'column', 
                                      'subtype': 'percent_stacked'})
    prod_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    prod_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    prod_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    prod_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    prod_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    prod_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in TPES_agg_fuels:
        i = prod_df2[prod_df2['fuel_code'] == component].index[0]
        prod_chart3.add_series({
            'name':       [economy + '_prod', chart_height + nrows5 + i + 4, 0],
            'categories': [economy + '_prod', chart_height + nrows5 + 3, 2, chart_height + nrows5 + 3, ncols7 - 1],
            'values':     [economy + '_prod', chart_height + nrows5 + i + 4, 2, chart_height + nrows5 + i + 4, ncols7 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet2.insert_chart('J3', prod_chart3)
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    worksheet3 = writer.sheets[economy + '_TPES_components_I']
    
    # Apply comma format and header format to relevant data rows
    worksheet3.set_column(2, ncols8 + 1, None, comma_format)
    worksheet3.set_row(chart_height, None, header_format)
    worksheet3.set_row(chart_height + nrows3 + 3, None, header_format)
    worksheet3.set_row(chart_height + nrows3 + nrows8 + 6, None, header_format)
    worksheet3.set_row(chart_height + nrows3 + nrows8 + nrows12 + 9, None, header_format)
    worksheet3.set_row(chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12, None, header_format)
    worksheet3.write(0, 0, economy + ' TPES components I', cell_format1)
    
    # Create a TPES components chart
    tpes_comp_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    tpes_comp_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    tpes_comp_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    tpes_comp_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    tpes_comp_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'PJ',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    tpes_comp_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    tpes_comp_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Net trade', 'Bunkers', 'Stock changes']:
        i = tpes_comp_df1[tpes_comp_df1['item_code_new'] == component].index[0]
        tpes_comp_chart1.add_series({
            'name':       [economy + '_TPES_components_I', chart_height + i + 1, 1],
            'categories': [economy + '_TPES_components_I', chart_height, 2, chart_height, ncols3 - 1],
            'values':     [economy + '_TPES_components_I', chart_height + i + 1, 2, chart_height + i + 1, ncols3 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    worksheet3.insert_chart('B3', tpes_comp_chart1)

    # IMPORTS: Create a line chart subset by fuel
    
    imports_line = workbook.add_chart({'type': 'line'})
    imports_line.set_size({
        'width': 500,
        'height': 300
    })
    
    imports_line.set_chartarea({
        'border': {'none': True}
    })
    
    imports_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    imports_line.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Imports (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    imports_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    imports_line.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for fuel in ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Other fuels']:
        i = imports_df1[imports_df1['fuel_code'] == fuel].index[0]
        imports_line.add_series({
            'name':       [economy + '_TPES_components_I', chart_height + nrows3 + i + 4, 0],
            'categories': [economy + '_TPES_components_I', chart_height + nrows3 + 3, 2, chart_height + nrows3 + 3, ncols8 - 1],
            'values':     [economy + '_TPES_components_I', chart_height + nrows3 + i + 4, 2, chart_height + nrows3 + i + 4, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    worksheet3.insert_chart('J3', imports_line)

    # Create a imports by fuel column
    imports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    imports_column.set_size({
        'width': 500,
        'height': 300
    })
    
    imports_column.set_chartarea({
        'border': {'none': True}
    })
    
    imports_column.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    imports_column.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Imports by fuel (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    imports_column.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    imports_column.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows12):
        imports_column.add_series({
            'name':       [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + i + 7, 0],
            'categories': [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + 6, 2, chart_height + nrows3 + nrows8 + 6, ncols12 - 1],
            'values':     [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + i + 7, 2, chart_height + nrows3 + nrows8 + i + 7, ncols12 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    worksheet3.insert_chart('R3', imports_column)

    # EXPORTS: Create a line chart subset by fuel
    
    exports_line = workbook.add_chart({'type': 'line'})
    exports_line.set_size({
        'width': 500,
        'height': 300
    })
    
    exports_line.set_chartarea({
        'border': {'none': True}
    })
    
    exports_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    exports_line.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Exports (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    exports_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    exports_line.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for fuel in ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Other fuels']:
        i = exports_df1[exports_df1['fuel_code'] == fuel].index[0]
        exports_line.add_series({
            'name':       [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + nrows12 + i + 10, 0],
            'categories': [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + nrows12 + 9, 2, chart_height + nrows3 + nrows8 + nrows12 + 9, ncols8 - 1],
            'values':     [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + nrows12 + i + 10, 2, chart_height + nrows3 + nrows8 + nrows12 + i + 10, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    worksheet3.insert_chart('Z3', exports_line)

    # Create a imports by fuel column
    exports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    exports_column.set_size({
        'width': 500,
        'height': 300
    })
    
    exports_column.set_chartarea({
        'border': {'none': True}
    })
    
    exports_column.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    exports_column.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Exports by fuel (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    exports_column.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    exports_column.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows13):
        exports_column.add_series({
            'name':       [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + nrows12 + nrows9 + i + 13, 0],
            'categories': [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12, 2, chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12, ncols13 - 1],
            'values':     [economy + '_TPES_components_I', chart_height + nrows3 + nrows8 + nrows12 + nrows9 + i + 13, 2, chart_height + nrows3 + nrows8 + nrows12 + nrows9 + i + 13, ncols13 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    worksheet3.insert_chart('AH3', exports_column)

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    worksheet4 = writer.sheets[economy + '_TPES_components_II']
    
    # Apply comma format and header format to relevant data rows
    worksheet4.set_column(2, ncols10 + 1, None, comma_format)
    worksheet4.set_row(chart_height, None, header_format)
    worksheet4.set_row(chart_height + nrows10 + 3, None, header_format)
    worksheet4.write(0, 0, economy + ' TPES components II', cell_format1)
    
    # MARINE BUNKER: Create a line chart subset by fuel
    
    marine_line = workbook.add_chart({'type': 'line'})
    marine_line.set_size({
        'width': 500,
        'height': 300
    })
    
    marine_line.set_chartarea({
        'border': {'none': True}
    })
    
    marine_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    marine_line.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Marine bunkers (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    marine_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    marine_line.set_title({
        'none': True
    }) 

    # Configure the series of the chart from the dataframe data.
    for i in range(nrows10):
        marine_line.add_series({
            'name':       [economy + '_TPES_components_II', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_components_II', chart_height, 2, chart_height, ncols10 - 1],
            'values':     [economy + '_TPES_components_II', chart_height + i + 1, 2, chart_height + i + 1, ncols10 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    worksheet4.insert_chart('B3', marine_line)

    # AVIATION BUNKER: Create a line chart subset by fuel
    
    aviation_line = workbook.add_chart({'type': 'line'})
    aviation_line.set_size({
        'width': 500,
        'height': 300
    })
    
    aviation_line.set_chartarea({
        'border': {'none': True}
    })
    
    aviation_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    aviation_line.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Aviation bunkers (PJ)',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    aviation_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    aviation_line.set_title({
        'none': True
    }) 

    # Configure the series of the chart from the dataframe data.
    for i in range(nrows11):
        aviation_line.add_series({
            'name':       [economy + '_TPES_components_II', chart_height + nrows10 + i + 4, 0],
            'categories': [economy + '_TPES_components_II', chart_height + nrows10 + 3, 2, chart_height + nrows10 + 3, ncols11 - 1],
            'values':     [economy + '_TPES_components_II', chart_height + nrows10 + i + 4, 2, chart_height + nrows10 + i + 4, ncols11 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    worksheet4.insert_chart('J3', aviation_line)
    
    writer.save()

print('Bling blang blaow, you have some TPES charts now')


