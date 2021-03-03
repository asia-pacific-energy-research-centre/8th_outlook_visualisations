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

# Sectors (DEMANDS)

Sectors_tfc = ['9_x_power', '10_losses_and_own_use', 
               '14_industry_sector', '15_transport_sector', '16_1_commercial_and_public_services', '16_2_residential',
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

Emissions_agg_sectors = ['Power', 'Own use', 'Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-specified']

Industry_eight = ['Iron & steel', 'Chemicals', 'Aluminium', 'Non-metallic minerals', 'Mining', 'Pulp & paper', 'Other', 'Non-specified']

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # REFERENCE DATA FRAMES
    # First data frame construction: Emissions by fuels
    ref_econ_df1 = EGEDA_emissions_reference[(EGEDA_emissions_reference['economy'] == economy) & 
                          (EGEDA_emissions_reference['item_code_new'].isin(['13_x_dem_pow_own'])) &
                          (EGEDA_emissions_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    #nrows1 = econ_df1.shape[0]
    #ncols1 = econ_df1.shape[1]

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_x_dem_pow_own')
    
    oil = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_x_dem_pow_own')
    
    heat_others = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Heat_others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_x_dem_pow_own')

    # EMISSIONS fuel data frame 1 (data frame 6)

    ref_emissions_fuel_df1 = ref_econ_df1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(ref_econ_df1.loc[:, '2000':])].reset_index(drop = True)

    ref_emissions_fuel_df1.loc[ref_emissions_fuel_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_emissions_fuel_df1.loc[ref_emissions_fuel_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_emissions_fuel_df1 = ref_emissions_fuel_df1[ref_emissions_fuel_df1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()

    nrows6 = ref_emissions_fuel_df1.shape[0]
    ncols6 = ref_emissions_fuel_df1.shape[1]

    ref_emissions_fuel_df2 = ref_emissions_fuel_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows7 = ref_emissions_fuel_df2.shape[0]
    ncols7 = ref_emissions_fuel_df2.shape[1]

    # Second data frame construction: FED by sectors
    ref_econ_df2 = EGEDA_emissions_reference[(EGEDA_emissions_reference['economy'] == economy) &
                               (EGEDA_emissions_reference['item_code_new'].isin(Sectors_tfc)) &
                               (EGEDA_emissions_reference['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    ref_econ_df2 = ref_econ_df2[['fuel_code', 'item_code_new'] + list(ref_econ_df2.loc[:,'2000':])]
    
    nrows2 = ref_econ_df2.shape[0]
    ncols2 = ref_econ_df2.shape[1]

    # Now build aggregate sector variables
    
    buildings = ref_econ_df2[ref_econ_df2['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = ref_econ_df2[ref_econ_df2['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    ref_emissions_sector_df1 = ref_econ_df2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(ref_econ_df2.loc[:, '2000':])].reset_index(drop = True)

    ref_emissions_sector_df1.loc[ref_emissions_sector_df1['item_code_new'] == '9_x_power', 'item_code_new'] = 'Power'
    ref_emissions_sector_df1.loc[ref_emissions_sector_df1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own use'
    ref_emissions_sector_df1.loc[ref_emissions_sector_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_emissions_sector_df1.loc[ref_emissions_sector_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_emissions_sector_df1.loc[ref_emissions_sector_df1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_emissions_sector_df1 = ref_emissions_sector_df1[ref_emissions_sector_df1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()
    ref_emissions_sector_df1 = ref_emissions_sector_df1[['fuel_code', 'item_code_new'] + list(ref_emissions_sector_df1.loc[:, '2000':])]

    nrows8 = ref_emissions_sector_df1.shape[0]
    ncols8 = ref_emissions_sector_df1.shape[1]

    ref_emissions_sector_df2 = ref_emissions_sector_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows9 = ref_emissions_sector_df2.shape[0]
    ncols9 = ref_emissions_sector_df2.shape[1]

    ##################################################################################################################################
    # NET ZERO DATA FRAMES
    # First data frame construction: Emissions by fuels
    netz_econ_df1 = EGEDA_emissions_netzero[(EGEDA_emissions_netzero['economy'] == economy) & 
                          (EGEDA_emissions_netzero['item_code_new'].isin(['13_x_dem_pow_own'])) &
                          (EGEDA_emissions_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    #nrows1 = econ_df1.shape[0]
    #ncols1 = econ_df1.shape[1]

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_x_dem_pow_own')
    
    oil = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_x_dem_pow_own')
    
    heat_others = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Heat_others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_x_dem_pow_own')

    # EMISSIONS fuel data frame 1 (data frame 6)

    netz_emissions_fuel_df1 = netz_econ_df1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(netz_econ_df1.loc[:, '2000':])].reset_index(drop = True)

    netz_emissions_fuel_df1.loc[netz_emissions_fuel_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_emissions_fuel_df1.loc[netz_emissions_fuel_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_emissions_fuel_df1 = netz_emissions_fuel_df1[netz_emissions_fuel_df1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()

    nrows26 = netz_emissions_fuel_df1.shape[0]
    ncols26 = netz_emissions_fuel_df1.shape[1]

    netz_emissions_fuel_df2 = netz_emissions_fuel_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows27 = netz_emissions_fuel_df2.shape[0]
    ncols27 = netz_emissions_fuel_df2.shape[1]

    # Second data frame construction: FED by sectors
    netz_econ_df2 = EGEDA_emissions_netzero[(EGEDA_emissions_netzero['economy'] == economy) &
                               (EGEDA_emissions_netzero['item_code_new'].isin(Sectors_tfc)) &
                               (EGEDA_emissions_netzero['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    netz_econ_df2 = netz_econ_df2[['fuel_code', 'item_code_new'] + list(netz_econ_df2.loc[:,'2000':])]
    
    nrows22 = netz_econ_df2.shape[0]
    ncols22 = netz_econ_df2.shape[1]

    # Now build aggregate sector variables
    
    buildings = netz_econ_df2[netz_econ_df2['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = netz_econ_df2[netz_econ_df2['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    netz_emissions_sector_df1 = netz_econ_df2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(netz_econ_df2.loc[:, '2000':])].reset_index(drop = True)

    netz_emissions_sector_df1.loc[netz_emissions_sector_df1['item_code_new'] == '9_x_power', 'item_code_new'] = 'Power'
    netz_emissions_sector_df1.loc[netz_emissions_sector_df1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own use'
    netz_emissions_sector_df1.loc[netz_emissions_sector_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_emissions_sector_df1.loc[netz_emissions_sector_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_emissions_sector_df1.loc[netz_emissions_sector_df1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_emissions_sector_df1 = netz_emissions_sector_df1[netz_emissions_sector_df1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()
    netz_emissions_sector_df1 = netz_emissions_sector_df1[['fuel_code', 'item_code_new'] + list(netz_emissions_sector_df1.loc[:, '2000':])]

    nrows28 = netz_emissions_sector_df1.shape[0]
    ncols28 = netz_emissions_sector_df1.shape[1]

    netz_emissions_sector_df2 = netz_emissions_sector_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows29 = netz_emissions_sector_df2.shape[0]
    ncols29 = netz_emissions_sector_df2.shape[1]

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
    ref_emissions_fuel_df1.to_excel(writer, sheet_name = economy + '_Emiss_fuel_ref', index = False, startrow = chart_height)
    netz_emissions_fuel_df1.to_excel(writer, sheet_name = economy + '_Emiss_fuel_netz', index = False, startrow = chart_height)
    ref_emissions_fuel_df2.to_excel(writer, sheet_name = economy + '_Emiss_fuel_ref', index = False, startrow = chart_height + nrows6 + 3)
    netz_emissions_fuel_df2.to_excel(writer, sheet_name = economy + '_Emiss_fuel_netz', index = False, startrow = chart_height + nrows26 + 3)
    ref_econ_df2.to_excel(writer, sheet_name = economy + '_Emiss_sector_ref', index = False, startrow = chart_height)
    netz_econ_df2.to_excel(writer, sheet_name = economy + '_Emiss_sector_netz', index = False, startrow = chart_height)
    ref_emissions_sector_df1.to_excel(writer, sheet_name = economy + '_Emiss_sector_ref', index = False, startrow = chart_height + nrows2 + 3)
    netz_emissions_sector_df1.to_excel(writer, sheet_name = economy + '_Emiss_sector_netz', index = False, startrow = chart_height + nrows22 + 3)
    ref_emissions_sector_df2.to_excel(writer, sheet_name = economy + '_Emiss_sector_ref', index = False, startrow = chart_height + nrows2 + nrows8 + 6)
    netz_emissions_sector_df2.to_excel(writer, sheet_name = economy + '_Emiss_sector_netz', index = False, startrow = chart_height + nrows22 + nrows28 + 6)

    # Access the workbook and first sheet with data from df1
    ref_worksheet1 = writer.sheets[economy + '_Emiss_fuel_ref']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet1.set_column(1, ncols6 + 1, None, comma_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height + nrows6 + 3, None, header_format)
    ref_worksheet1.write(0, 0, economy + ' emissions by fuel reference scenario', cell_format1)

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
    for i in range(nrows6):
        ref_em_fuel_chart1.add_series({
            'name':       [economy + '_Emiss_fuel_ref', chart_height + i + 1, 0],
            'categories': [economy + '_Emiss_fuel_ref', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_Emiss_fuel_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet1.insert_chart('B3', ref_em_fuel_chart1)

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
        i = ref_emissions_fuel_df2[ref_emissions_fuel_df2['fuel_code'] == component].index[0]
        ref_em_fuel_chart2.add_series({
            'name':       [economy + '_Emiss_fuel_ref', chart_height + nrows6 + i + 4, 0],
            'categories': [economy + '_Emiss_fuel_ref', chart_height + nrows6 + 3, 2, chart_height + nrows6 + 3, ncols7 - 1],
            'values':     [economy + '_Emiss_fuel_ref', chart_height + nrows6 + i + 4, 2, chart_height + nrows6 + i + 4, ncols7 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet1.insert_chart('J3', ref_em_fuel_chart2)

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
    for i in range(nrows6):
        ref_em_fuel_chart3.add_series({
            'name':       [economy + '_Emiss_fuel_ref', chart_height + i + 1, 0],
            'categories': [economy + '_Emiss_fuel_ref', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_Emiss_fuel_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    ref_worksheet1.insert_chart('R3', ref_em_fuel_chart3)


    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet2 = writer.sheets[economy + '_Emiss_sector_ref']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet2.set_column(1, ncols2 + 1, None, comma_format)
    ref_worksheet2.set_row(chart_height, None, header_format)
    ref_worksheet2.set_row(chart_height + nrows2 + 3, None, header_format)
    ref_worksheet2.set_row(chart_height + nrows2 + nrows8 + 6, None, header_format)
    ref_worksheet2.write(0, 0, economy + ' emissions by demand sector reference scenario', cell_format1)
    
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
    for i in range(nrows8):
        ref_em_sector_chart1.add_series({
            'name':       [economy + '_Emiss_sector_ref', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_Emiss_sector_ref', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_Emiss_sector_ref', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    ref_worksheet2.insert_chart('R3', ref_em_sector_chart1)

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
    for i in range(nrows8):
        ref_em_sector_chart2.add_series({
            'name':       [economy + '_Emiss_sector_ref', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_Emiss_sector_ref', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_Emiss_sector_ref', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet2.insert_chart('B3', ref_em_sector_chart2)

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
        i = ref_emissions_sector_df2[ref_emissions_sector_df2['item_code_new'] == component].index[0]
        ref_em_sector_chart3.add_series({
            'name':       [economy + '_Emiss_sector_ref', chart_height + nrows2 + nrows8 + i + 7, 1],
            'categories': [economy + '_Emiss_sector_ref', chart_height + nrows2 + nrows8 + 6, 2, chart_height + nrows2 + nrows8 + 6, ncols9 - 1],
            'values':     [economy + '_Emiss_sector_ref', chart_height + nrows2 + nrows8 + i + 7, 2, chart_height + nrows2 + nrows8 + i + 7, ncols9 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet2.insert_chart('J3', ref_em_sector_chart3)

    #############################################################################################################################
    # NET ZERO CHARTS
    # Access the workbook and first sheet with data from df1
    netz_worksheet1 = writer.sheets[economy + '_Emiss_fuel_netz']
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet1.set_column(1, ncols26 + 1, None, comma_format)
    netz_worksheet1.set_row(chart_height, None, header_format)
    netz_worksheet1.set_row(chart_height, None, header_format)
    netz_worksheet1.set_row(chart_height + nrows26 + 3, None, header_format)
    netz_worksheet1.write(0, 0, economy + ' emissions by fuel net zero scenario', cell_format1)

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
    for i in range(nrows26):
        netz_em_fuel_chart1.add_series({
            'name':       [economy + '_Emiss_fuel_netz', chart_height + i + 1, 0],
            'categories': [economy + '_Emiss_fuel_netz', chart_height, 2, chart_height, ncols26 - 1],
            'values':     [economy + '_Emiss_fuel_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols26 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet1.insert_chart('B3', netz_em_fuel_chart1)

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
        i = netz_emissions_fuel_df2[netz_emissions_fuel_df2['fuel_code'] == component].index[0]
        netz_em_fuel_chart2.add_series({
            'name':       [economy + '_Emiss_fuel_netz', chart_height + nrows26 + i + 4, 0],
            'categories': [economy + '_Emiss_fuel_netz', chart_height + nrows26 + 3, 2, chart_height + nrows26 + 3, ncols27 - 1],
            'values':     [economy + '_Emiss_fuel_netz', chart_height + nrows26 + i + 4, 2, chart_height + nrows26 + i + 4, ncols27 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet1.insert_chart('J3', netz_em_fuel_chart2)

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
    for i in range(nrows26):
        netz_em_fuel_chart3.add_series({
            'name':       [economy + '_Emiss_fuel_netz', chart_height + i + 1, 0],
            'categories': [economy + '_Emiss_fuel_netz', chart_height, 2, chart_height, ncols26 - 1],
            'values':     [economy + '_Emiss_fuel_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols26 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    netz_worksheet1.insert_chart('R3', netz_em_fuel_chart3)


    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    netz_worksheet2 = writer.sheets[economy + '_Emiss_sector_netz']
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet2.set_column(1, ncols2 + 1, None, comma_format)
    netz_worksheet2.set_row(chart_height, None, header_format)
    netz_worksheet2.set_row(chart_height + nrows2 + 3, None, header_format)
    netz_worksheet2.set_row(chart_height + nrows2 + nrows8 + 6, None, header_format)
    netz_worksheet2.write(0, 0, economy + ' emissions by demand sector net zero scenario', cell_format1)
    
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
    for i in range(nrows28):
        netz_em_sector_chart1.add_series({
            'name':       [economy + '_Emiss_sector_netz', chart_height + nrows22 + i + 4, 1],
            'categories': [economy + '_Emiss_sector_netz', chart_height + nrows22 + 3, 2, chart_height + nrows22 + 3, ncols28 - 1],
            'values':     [economy + '_Emiss_sector_netz', chart_height + nrows22 + i + 4, 2, chart_height + nrows22 + i + 4, ncols28 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    netz_worksheet2.insert_chart('R3', netz_em_sector_chart1)

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
    for i in range(nrows28):
        netz_em_sector_chart2.add_series({
            'name':       [economy + '_Emiss_sector_netz', chart_height + nrows22 + i + 4, 1],
            'categories': [economy + '_Emiss_sector_netz', chart_height + nrows22 + 3, 2, chart_height + nrows22 + 3, ncols28 - 1],
            'values':     [economy + '_Emiss_sector_netz', chart_height + nrows22 + i + 4, 2, chart_height + nrows22 + i + 4, ncols28 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet2.insert_chart('B3', netz_em_sector_chart2)

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
        i = netz_emissions_sector_df2[netz_emissions_sector_df2['item_code_new'] == component].index[0]
        netz_em_sector_chart3.add_series({
            'name':       [economy + '_Emiss_sector_netz', chart_height + nrows22 + nrows28 + i + 7, 1],
            'categories': [economy + '_Emiss_sector_netz', chart_height + nrows22 + nrows28 + 6, 2, chart_height + nrows22 + nrows28 + 6, ncols29 - 1],
            'values':     [economy + '_Emiss_sector_netz', chart_height + nrows22 + nrows28 + i + 7, 2, chart_height + nrows22 + nrows28 + i + 7, ncols29 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet2.insert_chart('J3', netz_em_sector_chart3)

    writer.save()

print('Emissions charts are ready for viewing in the results folder')


