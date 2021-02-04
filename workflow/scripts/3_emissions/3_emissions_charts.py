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

EGEDA_emissions = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_emissions_2018.csv')

# Define unique values for economy, fuels, and items columns

Economy_codes = EGEDA_emissions.economy.unique()
Fuels = EGEDA_emissions.fuel_code.unique()
Items = EGEDA_emissions.item_code_new.unique()

# Colours for charting (to be amended later)

colours = pd.read_excel('./data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

# Define month and year to create folder for saving charts/tables

month_year = pd.to_datetime('today').strftime('%B_%Y')

# Subsets for impending df builds

First_level_fuels = ['1_coal', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products',
                     '8_gas', '16_others', '17_electricity', '18_heat', '19_total']

Required_fuels = ['1_coal', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products',
                  '8_gas', '16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources',
                  '17_electricity', '18_heat', '19_total']

Coal_fuels = ['1_coal', '2_coal_products', '3_peat', '4_peat_products']

Oil_fuels = ['6_crude_oil_and_ngl', '7_petroleum_products']

Heat_others_fuels = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources', '18_heat']

# Sectors

Sectors_tfc = ['14_industry_sector', '15_transport_sector', '16_1_commercial_and_public_services', '16_2_residential',
               '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others', '17_nonenergy_use']

Buildings_items = ['16_1_commercial_and_public_services', '16_2_residential']

Ag_items = ['16_3_agriculture', '16_4_fishing']

Subindustry = ['14_industry_sector', '14_1_iron_and_steel', '14_2_chemical_incl_petrochemical', '14_3_non_ferrous_metals',
               '14_4_nonmetallic_mineral_products', '14_5_transportation_equipment', '14_6_machinery', '14_7_mining_and_quarrying',
               '14_8_food_beverages_and_tobacco', '14_9_pulp_paper_and_printing', '14_10_wood_and_wood_products', 
               '14_11_construction', '14_12_textiles_and_leather', '14_13_nonspecified_industry']

Transport_fuels = ['1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products', '7_1_motor_gasoline', '7_2_aviation_gasoline',
                   '7_x_jet_fuel', '7_7_gas_diesel_oil', '7_8_fuel_oil', '7_9_lpg',
                   '7_x_other_petroleum_products', '8_1_natural_gas', '17_electricity']

Transport_fuels_agg = ['Diesel', 'Gasoline', 'LPG', 'Gas', 'Jet fuel', 'Electricity', 'Other']

Other_fuel = ['7_8_fuel_oil', '1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products', '7_x_other_petroleum_products']

Other_industry = ['14_5_transportation_equipment', '14_6_machinery', '14_8_food_beverages_and_tobacco', '14_10_wood_and_wood_products',
                  '14_11_construction', '14_12_textiles_and_leather']

Transport_modal = ['15_1_domestic_air_transport', '15_2_road', '15_3_rail', '15_4_domestic_navigation', '15_5_pipeline_transport',
                   '15_6_nonspecified_transport']

Transport_modal_agg = ['Aviation', 'Road', 'Rail' ,'Marine', 'Pipeline', 'Non-specified']

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written (can change this)

# Define column chart years
col_chart_years = ['2000', '2010', '2017', '2020', '2030', '2040', '2050']

# Define column chart years for transport
col_chart_years_transport = ['2017', '2020', '2030', '2040', '2050']

# FED aggregate fuels

Emissions_agg_fuels = ['Coal', 'Oil', 'Gas', 'Electricity', 'Heat & others']

Emissions_agg_sectors = ['Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-specified']

Industry_eight = ['Iron & steel', 'Chemicals', 'Aluminium', 'Non-metallic minerals', 'Mining', 'Pulp & paper', 'Other', 'Non-specified']

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # First data frame construction: Emissions by fuels
    econ_df1 = EGEDA_emissions[(EGEDA_emissions['economy'] == economy) & 
                          (EGEDA_emissions['item_code_new'].isin(['13_total_final_energy_consumption'])) &
                          (EGEDA_emissions['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    #nrows1 = econ_df1.shape[0]
    #ncols1 = econ_df1.shape[1]

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = econ_df1[econ_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_total_final_energy_consumption')
    
    oil = econ_df1[econ_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_total_final_energy_consumption')
    
    heat_others = econ_df1[econ_df1['fuel_code'].isin(Heat_others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_total_final_energy_consumption')

    # EMISSIONS fuel data frame 1 (data frame 6)

    emissions_fuel_df1 = econ_df1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(econ_df1.loc[:, '2000':])].reset_index(drop = True)

    emissions_fuel_df1.loc[emissions_fuel_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    emissions_fuel_df1.loc[emissions_fuel_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    emissions_fuel_df1 = emissions_fuel_df1[emissions_fuel_df1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()

    nrows6 = emissions_fuel_df1.shape[0]
    ncols6 = emissions_fuel_df1.shape[1]

    emissions_fuel_df2 = emissions_fuel_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows7 = emissions_fuel_df2.shape[0]
    ncols7 = emissions_fuel_df2.shape[1]

    # Second data frame construction: FED by sectors
    econ_df2 = EGEDA_emissions[(EGEDA_emissions['economy'] == economy) &
                               (EGEDA_emissions['item_code_new'].isin(Sectors_tfc)) &
                               (EGEDA_emissions['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    econ_df2 = econ_df2[['fuel_code', 'item_code_new'] + list(econ_df2.loc[:,'2000':])]
    
    nrows2 = econ_df2.shape[0]
    ncols2 = econ_df2.shape[1]

    # Now build aggregate sector variables
    
    buildings = econ_df2[econ_df2['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = econ_df2[econ_df2['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    emissions_sector_df1 = econ_df2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(econ_df2.loc[:, '2000':])].reset_index(drop = True)

    emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    emissions_sector_df1 = emissions_sector_df1[emissions_sector_df1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()
    emissions_sector_df1 = emissions_sector_df1[['fuel_code', 'item_code_new'] + list(emissions_sector_df1.loc[:, '2000':])]

    nrows8 = emissions_sector_df1.shape[0]
    ncols8 = emissions_sector_df1.shape[1]

    emissions_sector_df2 = emissions_sector_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows9 = emissions_sector_df2.shape[0]
    ncols9 = emissions_sector_df2.shape[1]

    # Define directory
    script_dir = './results/' + month_year + '/Emissions/'
    results_dir = os.path.join(script_dir, 'economy_breakdown/', economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_emissions.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Insert the various dataframes into different sheets of the workbook    
    emissions_fuel_df1.to_excel(writer, sheet_name = economy + '_Emissions_fuel', index = False, startrow = chart_height)
    emissions_fuel_df2.to_excel(writer, sheet_name = economy + '_Emissions_fuel', index = False, startrow = chart_height + nrows6 + 3)
    econ_df2.to_excel(writer, sheet_name = economy + '_Emissions_sector', index = False, startrow = chart_height)
    emissions_sector_df1.to_excel(writer, sheet_name = economy + '_Emissions_sector', index = False, startrow = chart_height + nrows2 + 3)
    emissions_sector_df2.to_excel(writer, sheet_name = economy + '_Emissions_sector', index = False, startrow = chart_height + nrows2 + nrows8 + 6)

    # Access the workbook and first sheet with data from df1
    worksheet1 = writer.sheets[economy + '_Emissions_fuel']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    worksheet1.set_column(1, ncols6 + 1, None, comma_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height + nrows6 + 3, None, header_format)
    worksheet1.write(0, 0, economy + ' emissions by fuel', cell_format1)

    ################################################################### CHARTS ###################################################################

    # Create a FED area chart
    em_fuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    em_fuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    em_fuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    em_fuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    em_fuel_chart1.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    em_fuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    em_fuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows6):
        em_fuel_chart1.add_series({
            'name':       [economy + '_Emissions_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_Emissions_fuel', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_Emissions_fuel', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet1.insert_chart('B3', em_fuel_chart1)

    ###################### Create another EMISSIONS chart showing proportional share #################################

    # Create a another chart
    em_fuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    em_fuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    em_fuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    em_fuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    em_fuel_chart2.set_y_axis({
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
        
    em_fuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    em_fuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in Emissions_agg_fuels:
        i = emissions_fuel_df2[emissions_fuel_df2['fuel_code'] == component].index[0]
        em_fuel_chart2.add_series({
            'name':       [economy + '_Emissions_fuel', chart_height + nrows6 + i + 4, 0],
            'categories': [economy + '_Emissions_fuel', chart_height + nrows6 + 3, 2, chart_height + nrows6 + 3, ncols7 - 1],
            'values':     [economy + '_Emissions_fuel', chart_height + nrows6 + i + 4, 2, chart_height + nrows6 + i + 4, ncols7 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet1.insert_chart('J3', em_fuel_chart2)

    # Create a Emissions line chart with higher level aggregation
    em_fuel_chart3 = workbook.add_chart({'type': 'line'})
    em_fuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    em_fuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    em_fuel_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    em_fuel_chart3.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Million Tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    em_fuel_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    em_fuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows6):
        em_fuel_chart3.add_series({
            'name':       [economy + '_Emissions_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_Emissions_fuel', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_Emissions_fuel', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    worksheet1.insert_chart('R3', em_fuel_chart3)


    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    worksheet2 = writer.sheets[economy + '_Emissions_sector']
        
    # Apply comma format and header format to relevant data rows
    worksheet2.set_column(1, ncols2 + 1, None, comma_format)
    worksheet2.set_row(chart_height, None, header_format)
    worksheet2.set_row(chart_height + nrows2 + 3, None, header_format)
    worksheet2.set_row(chart_height + nrows2 + nrows8 + 6, None, header_format)
    worksheet2.write(0, 0, economy + ' emissions by demand sector', cell_format1)
    
    # Create an EMISSIONS sector line chart

    em_sector_chart1 = workbook.add_chart({'type': 'line'})
    em_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    em_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    em_sector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    em_sector_chart1.set_y_axis({
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
        
    em_sector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    em_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows8):
        em_sector_chart1.add_series({
            'name':       [economy + '_Emissions_sector', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_Emissions_sector', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_Emissions_sector', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    worksheet2.insert_chart('R3', em_sector_chart1)

    # Create a EMISSIONS sector area chart

    em_sector_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    em_sector_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    em_sector_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    em_sector_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    em_sector_chart2.set_y_axis({
        'major_tick_mark': 'none', 
        'minor_tick_mark': 'none',
        'name': 'Tonnes CO2',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#bebebe'}
        },
        'line': {'color': '#bebebe'}
    })
        
    em_sector_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    em_sector_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows8):
        em_sector_chart2.add_series({
            'name':       [economy + '_Emissions_sector', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_Emissions_sector', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_Emissions_sector', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet2.insert_chart('B3', em_sector_chart2)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    em_sector_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    em_sector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    em_sector_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    em_sector_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    em_sector_chart3.set_y_axis({
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
        
    em_sector_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    em_sector_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in Emissions_agg_sectors:
        i = emissions_sector_df2[emissions_sector_df2['item_code_new'] == component].index[0]
        em_sector_chart3.add_series({
            'name':       [economy + '_Emissions_sector', chart_height + nrows2 + nrows8 + i + 7, 1],
            'categories': [economy + '_Emissions_sector', chart_height + nrows2 + nrows8 + 6, 2, chart_height + nrows2 + nrows8 + 6, ncols9 - 1],
            'values':     [economy + '_Emissions_sector', chart_height + nrows2 + nrows8 + i + 7, 2, chart_height + nrows2 + nrows8 + i + 7, ncols9 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet2.insert_chart('J3', em_sector_chart3)

    writer.save()

print('Emissions charts are ready for viewing in the results folder')


