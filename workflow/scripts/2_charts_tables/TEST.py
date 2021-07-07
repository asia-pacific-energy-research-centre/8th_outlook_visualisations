# Bolt projected emissions to historical

# Import relevant packages

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
import xlsxwriter
import pandas.io.formats.excel
import glob
import re

# Import the recently created EMISSIONS data frame that joins OSeMOSYS results to EGEDA historical 

EGEDA_emissions_reference = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_emissions_2018_reference.csv')
EGEDA_emissions_netzero = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_emissions_2018_netzero.csv')

# Define unique values for economy, fuels, and items columns

Economy_codes = EGEDA_emissions_reference.economy.unique()
Fuels = EGEDA_emissions_reference.fuel_code.unique()
Items = EGEDA_emissions_reference.item_code_new.unique()

# Colours for charting (to be amended later)

colours = pd.read_excel('./data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

# Subsets for impending emissions df builds

First_level_emiss = ['1_coal', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products',
                     '8_gas', '16_others', '17_electricity', '18_heat', '19_total']

Required_emiss = ['1_coal', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products',
                  '8_gas', '16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources',
                  '17_electricity', '18_heat', '19_total']

Coal_emiss = ['1_coal', '2_coal_products', '3_peat', '4_peat_products']

Oil_emiss = ['6_crude_oil_and_ngl', '7_petroleum_products']

Heat_others_emiss = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources', '18_heat']

# Emissions sectors (DEMANDS)

Sectors_emiss = ['9_x_power', '10_losses_and_own_use', 
               '14_industry_sector', '15_transport_sector', '16_1_commercial_and_public_services', '16_2_residential',
               '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others', '17_nonenergy_use']

Buildings_emiss = ['16_1_commercial_and_public_services', '16_2_residential']

Ag_emiss = ['16_3_agriculture', '16_4_fishing']

# FED aggregate fuels

Emissions_agg_fuels = ['Coal', 'Oil', 'Gas', 'Electricity', 'Heat & others']

Emissions_agg_sectors = ['Power', 'Own use', 'Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-specified']

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # REFERENCE DATA FRAMES
    # First data frame construction: Emissions by fuels
    ref_emiss_1 = EGEDA_emissions_reference[(EGEDA_emissions_reference['economy'] == economy) & 
                          (EGEDA_emissions_reference['item_code_new'].isin(['13_x_dem_pow_own'])) &
                          (EGEDA_emissions_reference['fuel_code'].isin(Required_emiss))].loc[:, 'fuel_code':].reset_index(drop = True)

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = ref_emiss_1[ref_emiss_1['fuel_code'].isin(Coal_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_x_dem_pow_own')
    
    oil = ref_emiss_1[ref_emiss_1['fuel_code'].isin(Oil_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_x_dem_pow_own')
    
    heat_others = ref_emiss_1[ref_emiss_1['fuel_code'].isin(Heat_others_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_x_dem_pow_own')

    # EMISSIONS fuel data frame 1 (data frame 6)

    ref_emiss_fuel_1 = ref_emiss_1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(ref_emiss_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_emiss_fuel_1.loc[ref_emiss_fuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_emiss_fuel_1.loc[ref_emiss_fuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_emiss_fuel_1 = ref_emiss_fuel_1[ref_emiss_fuel_1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()

    ref_emiss_fuel_1_rows = ref_emiss_fuel_1.shape[0]
    ref_emiss_fuel_1_cols = ref_emiss_fuel_1.shape[1]

    ref_emiss_fuel_2 = ref_emiss_fuel_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_emiss_fuel_2_rows = ref_emiss_fuel_2.shape[0]
    ref_emiss_fuel_2_cols = ref_emiss_fuel_2.shape[1]

    # Second data frame construction: FED by sectors
    ref_emiss_2 = EGEDA_emissions_reference[(EGEDA_emissions_reference['economy'] == economy) &
                               (EGEDA_emissions_reference['item_code_new'].isin(Sectors_emiss)) &
                               (EGEDA_emissions_reference['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    ref_emiss_2 = ref_emiss_2[['fuel_code', 'item_code_new'] + list(ref_emiss_2.loc[:,'2000':'2050'])]
    
    ref_emiss_2_rows = ref_emiss_2.shape[0]
    ref_emiss_2_cols = ref_emiss_2.shape[1]

    # Now build aggregate sector variables
    
    buildings = ref_emiss_2[ref_emiss_2['item_code_new'].isin(Buildings_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = ref_emiss_2[ref_emiss_2['item_code_new'].isin(Ag_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    ref_emiss_sector_1 = ref_emiss_2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(ref_emiss_2.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '9_x_power', 'item_code_new'] = 'Power'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own use'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_emiss_sector_1 = ref_emiss_sector_1[ref_emiss_sector_1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()
    ref_emiss_sector_1 = ref_emiss_sector_1[['fuel_code', 'item_code_new'] + list(ref_emiss_sector_1.loc[:, '2000':'2050'])]

    ref_emiss_sector_1_rows = ref_emiss_sector_1.shape[0]
    ref_emiss_sector_1_cols = ref_emiss_sector_1.shape[1]

    ref_emiss_sector_2 = ref_emiss_sector_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_emiss_sector_2_rows = ref_emiss_sector_2.shape[0]
    ref_emiss_sector_2_cols = ref_emiss_sector_2.shape[1]

    ##################################################################################################################################
    # NET ZERO DATA FRAMES
    # First data frame construction: Emissions by fuels
    netz_emiss_1 = EGEDA_emissions_netzero[(EGEDA_emissions_netzero['economy'] == economy) & 
                          (EGEDA_emissions_netzero['item_code_new'].isin(['13_x_dem_pow_own'])) &
                          (EGEDA_emissions_netzero['fuel_code'].isin(Required_emiss))].loc[:, 'fuel_code':].reset_index(drop = True)

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = netz_emiss_1[netz_emiss_1['fuel_code'].isin(Coal_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_x_dem_pow_own')
    
    oil = netz_emiss_1[netz_emiss_1['fuel_code'].isin(Oil_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_x_dem_pow_own')
    
    heat_others = netz_emiss_1[netz_emiss_1['fuel_code'].isin(Heat_others_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_x_dem_pow_own')

    # EMISSIONS fuel data frame 1 (data frame 6)

    netz_emiss_fuel_1 = netz_emiss_1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(netz_emiss_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_emiss_fuel_1.loc[netz_emiss_fuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_emiss_fuel_1.loc[netz_emiss_fuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_emiss_fuel_1 = netz_emiss_fuel_1[netz_emiss_fuel_1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()

    netz_emiss_fuel_1_rows = netz_emiss_fuel_1.shape[0]
    netz_emiss_fuel_1_cols = netz_emiss_fuel_1.shape[1]

    netz_emiss_fuel_2 = netz_emiss_fuel_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_emiss_fuel_2_rows = netz_emiss_fuel_2.shape[0]
    netz_emiss_fuel_2_cols = netz_emiss_fuel_2.shape[1]

    # Second data frame construction: FED by sectors
    netz_emiss_2 = EGEDA_emissions_netzero[(EGEDA_emissions_netzero['economy'] == economy) &
                               (EGEDA_emissions_netzero['item_code_new'].isin(Sectors_emiss)) &
                               (EGEDA_emissions_netzero['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    netz_emiss_2 = netz_emiss_2[['fuel_code', 'item_code_new'] + list(netz_emiss_2.loc[:,'2000':'2050'])]
    
    netz_emiss_2_rows = netz_emiss_2.shape[0]
    netz_emiss_2_cols = netz_emiss_2.shape[1]

    # Now build aggregate sector variables
    
    buildings = netz_emiss_2[netz_emiss_2['item_code_new'].isin(Buildings_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = netz_emiss_2[netz_emiss_2['item_code_new'].isin(Ag_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    netz_emiss_sector_1 = netz_emiss_2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(netz_emiss_2.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '9_x_power', 'item_code_new'] = 'Power'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own use'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_emiss_sector_1 = netz_emiss_sector_1[netz_emiss_sector_1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()
    netz_emiss_sector_1 = netz_emiss_sector_1[['fuel_code', 'item_code_new'] + list(netz_emiss_sector_1.loc[:, '2000':'2050'])]

    netz_emiss_sector_1_rows = netz_emiss_sector_1.shape[0]
    netz_emiss_sector_1_cols = netz_emiss_sector_1.shape[1]

    netz_emiss_sector_2 = netz_emiss_sector_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_emiss_sector_2_rows = netz_emiss_sector_2.shape[0]
    netz_emiss_sector_2_cols = netz_emiss_sector_2.shape[1]

    # Define directory
    script_dir = './results/'
    results_dir = os.path.join(script_dir,economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_emissions.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Insert the various dataframes into different sheets of the workbook    
    ref_emiss_fuel_1.to_excel(writer, sheet_name = economy + '_Emiss_fuel', index = False, startrow = chart_height)
    netz_emiss_fuel_1.to_excel(writer, sheet_name = economy + '_Emiss_fuel', index = False, startrow = (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6)
    ref_emiss_fuel_2.to_excel(writer, sheet_name = economy + '_Emiss_fuel', index = False, startrow = chart_height + ref_emiss_fuel_1_rows + 3)
    netz_emiss_fuel_2.to_excel(writer, sheet_name = economy + '_Emiss_fuel', index = False, startrow = (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9)
    ref_emiss_sector_1.to_excel(writer, sheet_name = economy + '_Emiss_sector', index = False, startrow = chart_height)
    netz_emiss_sector_1.to_excel(writer, sheet_name = economy + '_Emiss_sector', index = False, startrow = (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6)
    ref_emiss_sector_2.to_excel(writer, sheet_name = economy + '_Emiss_sector', index = False, startrow = chart_height + ref_emiss_sector_1_rows + 3)
    netz_emiss_sector_2.to_excel(writer, sheet_name = economy + '_Emiss_sector', index = False, startrow = (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9)

    # Access the workbook and first sheet with data from df1
    both_worksheet34 = writer.sheets[economy + '_Emiss_fuel']
    
        
    # Apply comma format and header format to relevant data rows
    both_worksheet34.set_column(1, ref_emiss_fuel_1_cols + 1, None, space_format)
    both_worksheet34.set_row(chart_height, None, header_format)
    both_worksheet34.set_row(chart_height + ref_emiss_fuel_1_rows + 3, None, header_format)
    both_worksheet34.set_row((2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, None, header_format)
    both_worksheet34.set_row((2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9, None, header_format)
    both_worksheet34.write(0, 0, economy + ' emissions by fuel reference scenario', cell_format1)
    both_worksheet34.write(chart_height + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 0, economy + ' emissions by fuel net-zero scenario', cell_format1)

    ################################################################### CHARTS ###################################################################

    # Create a FED area chart
    ref_em_fuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_em_fuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_fuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_fuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_em_fuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_fuel_1_rows):
        ref_em_fuel_chart1.add_series({
            'name':       [economy + '_Emiss_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_Emiss_fuel', chart_height, 2, chart_height, ref_emiss_fuel_1_cols - 1],
            'values':     [economy + '_Emiss_fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_fuel_1_cols - 1],
            'fill':       {'color': ref_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    both_worksheet34.insert_chart('B3', ref_em_fuel_chart1)

    ###################### Create another EMISSIONS chart showing proportional share #################################

    # Create a another chart
    ref_em_fuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_em_fuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_fuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_fuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'CO2 proportion',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_em_fuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in Emissions_agg_fuels:
        i = ref_emiss_fuel_2[ref_emiss_fuel_2['fuel_code'] == component].index[0]
        ref_em_fuel_chart2.add_series({
            'name':       [economy + '_Emiss_fuel', chart_height + ref_emiss_fuel_1_rows + i + 4, 0],
            'categories': [economy + '_Emiss_fuel', chart_height + ref_emiss_fuel_1_rows + 3, 2, chart_height + ref_emiss_fuel_1_rows + 3, ref_emiss_fuel_2_cols - 1],
            'values':     [economy + '_Emiss_fuel', chart_height + ref_emiss_fuel_1_rows + i + 4, 2, chart_height + ref_emiss_fuel_1_rows + i + 4, ref_emiss_fuel_2_cols - 1],
            'fill':       {'color': ref_emiss_fuel_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    both_worksheet34.insert_chart('J3', ref_em_fuel_chart2)

    # Create a Emissions line chart with higher level aggregation
    ref_em_fuel_chart3 = workbook.add_chart({'type': 'line'})
    ref_em_fuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_fuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_fuel_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_fuel_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_em_fuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_fuel_1_rows):
        ref_em_fuel_chart3.add_series({
            'name':       [economy + '_Emiss_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_Emiss_fuel', chart_height, 2, chart_height, ref_emiss_fuel_1_cols - 1],
            'values':     [economy + '_Emiss_fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_fuel_1_cols - 1],
            'line':       {'color': ref_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25}
        })    
        
    both_worksheet34.insert_chart('R3', ref_em_fuel_chart3)


    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    both_worksheet35 = writer.sheets[economy + '_Emiss_sector']
        
    # Apply comma format and header format to relevant data rows
    both_worksheet35.set_column(1, ref_emiss_2_cols + 1, None, space_format)
    both_worksheet35.set_row(chart_height, None, header_format)
    both_worksheet35.set_row(chart_height + ref_emiss_sector_1_rows + 3, None, header_format)
    both_worksheet35.set_row((2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, None, header_format)
    both_worksheet35.set_row((2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9, None, header_format)
    both_worksheet35.write(0, 0, economy + ' emissions by demand sector reference scenario', cell_format1)
    both_worksheet35.write(chart_height + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, 0, economy + ' emissions by demand sector net-zero scenario', cell_format1)
    
    # Create an EMISSIONS sector line chart

    ref_em_sector_chart1 = workbook.add_chart({'type': 'line'})
    ref_em_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_sector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_em_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_sector_1_rows):
        ref_em_sector_chart1.add_series({
            'name':       [economy + '_Emiss_sector', chart_height + i + 1, 1],
            'categories': [economy + '_Emiss_sector', chart_height, 2, chart_height, ref_emiss_sector_1_cols - 1],
            'values':     [economy + '_Emiss_sector', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_sector_1_cols - 1],
            'line':       {'color': ref_emiss_sector_1['item_code_new'].map(colours_dict).loc[i], 
                           'width': 1.25}
        })    
        
    both_worksheet35.insert_chart('R3', ref_em_sector_chart1)

    # Create a EMISSIONS sector area chart

    ref_em_sector_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_em_sector_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_sector_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_sector_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_em_sector_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ref_emiss_sector_1_rows):
        ref_em_sector_chart2.add_series({
            'name':       [economy + '_Emiss_sector', chart_height + i + 1, 1],
            'categories': [economy + '_Emiss_sector', chart_height, 2, chart_height, ref_emiss_sector_1_cols - 1],
            'values':     [economy + '_Emiss_sector', chart_height + i + 1, 2, chart_height + i + 1, ref_emiss_sector_1_cols - 1],
            'fill':       {'color': ref_emiss_sector_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    both_worksheet35.insert_chart('B3', ref_em_sector_chart2)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    ref_em_sector_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_em_sector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_em_sector_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_em_sector_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Proportion of CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    ref_em_sector_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_em_sector_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in Emissions_agg_sectors:
        i = ref_emiss_sector_2[ref_emiss_sector_2['item_code_new'] == component].index[0]
        ref_em_sector_chart3.add_series({
            'name':       [economy + '_Emiss_sector', chart_height + ref_emiss_sector_1_rows + i + 4, 1],
            'categories': [economy + '_Emiss_sector', chart_height + ref_emiss_sector_1_rows + 3, 2, chart_height + ref_emiss_sector_1_rows + 3, ref_emiss_sector_2_cols - 1],
            'values':     [economy + '_Emiss_sector', chart_height + ref_emiss_sector_1_rows + i + 4, 2, chart_height + ref_emiss_sector_1_rows + i + 4, ref_emiss_sector_2_cols - 1],
            'fill':       {'color': ref_emiss_sector_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    both_worksheet35.insert_chart('J3', ref_em_sector_chart3)

    #############################################################################################################################
    # NET ZERO CHARTS
    # Access the workbook and first sheet with data from df1
    # both_worksheet34 = writer.sheets[economy + '_Emiss_fuel']
        
    # # Apply comma format and header format to relevant data rows
    # both_worksheet34.set_column(1, netz_emiss_fuel_1_cols + 1, None, space_format)
    # both_worksheet34.set_row(chart_height, None, header_format)
    # both_worksheet34.set_row(chart_height, None, header_format)
    # both_worksheet34.set_row(chart_height + netz_emiss_fuel_1_rows + 3, None, header_format)
    # both_worksheet34.write(0, 0, economy + ' emissions by fuel net zero scenario', cell_format1)

    ################################################################### CHARTS ###################################################################

    # Create a FED area chart
    netz_em_fuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_em_fuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_fuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_fuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_em_fuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_fuel_1_rows):
        netz_em_fuel_chart1.add_series({
            'name':       [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 0],
            'categories': [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 2,\
                (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, netz_emiss_fuel_1_cols - 1],
            'values':     [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, netz_emiss_fuel_1_cols - 1],
            'fill':       {'color': netz_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    both_worksheet34.insert_chart('B37', netz_em_fuel_chart1)

    ###################### Create another EMISSIONS chart showing proportional share #################################

    # Create a another chart
    netz_em_fuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_em_fuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_fuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_fuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'CO2 proportion',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_em_fuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in Emissions_agg_fuels:
        i = netz_emiss_fuel_2[netz_emiss_fuel_2['fuel_code'] == component].index[0]
        netz_em_fuel_chart2.add_series({
            'name':       [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + i + 10, 0],
            'categories': [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9, 2,\
                (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + 9, netz_emiss_fuel_2_cols - 1],
            'values':     [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + i + 10, 2,\
                (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + netz_emiss_fuel_1_rows + i + 10, netz_emiss_fuel_2_cols - 1],
            'fill':       {'color': netz_emiss_fuel_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    both_worksheet34.insert_chart('J37', netz_em_fuel_chart2)

    # Create a Emissions line chart with higher level aggregation
    netz_em_fuel_chart3 = workbook.add_chart({'type': 'line'})
    netz_em_fuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_fuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_fuel_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_fuel_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_em_fuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_fuel_1_rows):
        netz_em_fuel_chart3.add_series({
            'name':       [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 0],
            'categories': [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, 2,\
                (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + 6, netz_emiss_fuel_1_cols - 1],
            'values':     [economy + '_Emiss_fuel', (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_emiss_fuel_1_rows + ref_emiss_fuel_2_rows + i + 7, netz_emiss_fuel_1_cols - 1],
            'line':       {'color': netz_emiss_fuel_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25}
        })    
        
    both_worksheet34.insert_chart('R37', netz_em_fuel_chart3)


    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # # Access the workbook and second sheet with data from df2
    # both_worksheet35 = writer.sheets[economy + '_Emiss_sector_netz']
        
    # # Apply comma format and header format to relevant data rows
    # both_worksheet35.set_column(1, netz_emiss_2_cols + 1, None, space_format)
    # both_worksheet35.set_row(chart_height, None, header_format)
    # both_worksheet35.set_row(chart_height + netz_emiss_2_rows + 3, None, header_format)
    # both_worksheet35.set_row(chart_height + netz_emiss_2_rows + netz_emiss_sector_1_rows + 6, None, header_format)
    # both_worksheet35.write(0, 0, economy + ' emissions by demand sector net-zero scenario', cell_format1)
    
    # Create an EMISSIONS sector line chart

    netz_em_sector_chart1 = workbook.add_chart({'type': 'line'})
    netz_em_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_sector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_em_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_sector_1_rows):
        netz_em_sector_chart1.add_series({
            'name':       [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 1],
            'categories': [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, 2,\
                (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, netz_emiss_sector_1_cols - 1],
            'values':     [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, netz_emiss_sector_1_cols - 1],
            'line':       {'color': netz_emiss_sector_1['item_code_new'].map(colours_dict).loc[i], 
                           'width': 1.25}
        })    
        
    both_worksheet35.insert_chart('R41', netz_em_sector_chart1)

    # Create a EMISSIONS sector area chart

    netz_em_sector_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_em_sector_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_sector_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_em_sector_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_em_sector_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(netz_emiss_sector_1_rows):
        netz_em_sector_chart2.add_series({
            'name':       [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 1],
            'categories': [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, 2,\
                (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + 6, netz_emiss_sector_1_cols - 1],
            'values':     [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + i + 7, netz_emiss_sector_1_cols - 1],
            'fill':       {'color': netz_emiss_sector_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    both_worksheet35.insert_chart('B41', netz_em_sector_chart2)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    netz_em_sector_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_em_sector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_em_sector_chart3.set_chartarea({
        'border': {'none': True} 
    })
    
    netz_em_sector_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Proportion of CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    netz_em_sector_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_em_sector_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in Emissions_agg_sectors:
        i = netz_emiss_sector_2[netz_emiss_sector_2['item_code_new'] == component].index[0]
        netz_em_sector_chart3.add_series({
            'name':       [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + i + 10, 1],
            'categories': [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9, 2,\
                (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + 9, netz_emiss_sector_2_cols - 1],
            'values':     [economy + '_Emiss_sector', (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + i + 10, 2,\
                (2 * chart_height) + ref_emiss_sector_1_rows + ref_emiss_sector_2_rows + netz_emiss_sector_1_rows + i + 10, netz_emiss_sector_2_cols - 1],
            'fill':       {'color': netz_emiss_sector_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    both_worksheet35.insert_chart('J41', netz_em_sector_chart3)

    writer.save()

print('Emissions charts are ready for viewing in the results folder')


