# EGEDA with OSeMOSYS results to 2050 bolted on 

# FED (tfc; including non-energy) plots for each economy

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
import xlsxwriter
import pandas.io.formats.excel
import glob
from pandas import ExcelWriter

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

Coal_fuels = list(Fuels[[0, 9]])

Oil_fuels = list(Fuels[[17, 24]])

Heat_others_fuels = list(Fuels[[50, 66, 69, 70, 75, 77]])

# Need to amend this to reflect demarcation between modern renewables and traditional biomass renewables 

Renewables_fuels = list(Fuels[[49, 51, 61, 62, 63, 64, 65, 68, 70]])

Modern_renew_primary = list(Fuels[[49, 51, 65, 68, 70]])

Modern_renew_FED = list(Fuels[[49, 51, 65, 68, 70]])

Sectors_tfc = list(Items[[64, 78, 86, 87, 88, 89, 90, 91]])

Buildings_items = list(Items[[86, 87]])

Ag_items = list(Items[[88, 89]])

Subindustry = list(Items[[64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77]])

Transport_fuels = list(Fuels[[25, 26, 27, 33, 46, 71, 72, 35, 76, 29, 34, 2, 6, 8, 9, 38]])

Transport_fuels_agg = ['Diesel', 'Gasoline', 'LPG', 'Gas', 'Jet fuel', 'Electricity', 'Renewables', 'Other']

Renew_fuel = list(Fuels[[71, 72]])

Other_fuel = list(Fuels[[34, 2, 6, 8, 9, 38]])

Other_industry = list(Items[[69, 70, 72, 74, 75, 76]])

Transport_modal = list(Items[[79, 80, 81, 82, 83, 84]])

Transport_modal_agg = ['Aviation', 'Road', 'Rail' ,'Marine', 'Pipeline', 'Non-specified']

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written (can change this)

# Define column chart years
col_chart_years = ['2000', '2010', '2017', '2020', '2030', '2040', '2050']

# Define column chart years for transport
col_chart_years_transport = ['2017', '2020', '2030', '2040', '2050']

# FED aggregate fuels

FED_agg_fuels = ['Coal', 'Oil', 'Gas', 'Renewables', 'Electricity', 'Heat & others']

FED_agg_sectors = ['Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-energy', 'Non-specified']

Industry_eight = ['Iron & steel', 'Chemicals', 'Aluminium', 'Non-metallic minerals', 'Mining', 'Pulp & paper', 'Other', 'Non-specified']

# Final energy demand by fuel and sector for each economy

# This is TFC which includes non-energy

############# Build FED (TFC) dataframes for each economy (TFC) and then build subsequent charts ###########

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # First data frame construction: FED by fuels
    econ_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                          (EGEDA_years['item_code_new'].isin(['11_total_final_consumption'])) &
                          (EGEDA_years['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    #nrows1 = econ_df1.shape[0]
    #ncols1 = econ_df1.shape[1]

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = econ_df1[econ_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                             item_code_new = '11_total_final_consumption')
    
    oil = econ_df1[econ_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                             item_code_new = '11_total_final_consumption')
    
    renewables = econ_df1[econ_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                           item_code_new = '11_total_final_consumption')
    
    heat_others = econ_df1[econ_df1['fuel_code'].isin(Heat_others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                             item_code_new = '11_total_final_consumption')

    # Fed fuel data frame 1 (data frame 6)

    fedfuel_df1 = econ_df1.append([coal, oil, renewables, heat_others])[['fuel_code',
                                                                         'item_code_new'] + list(econ_df1.loc[:, '2000':])].reset_index(drop = True)

    fedfuel_df1.loc[fedfuel_df1['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    fedfuel_df1.loc[fedfuel_df1['fuel_code'] == '10_electricity', 'fuel_code'] = 'Electricity'

    fedfuel_df1 = fedfuel_df1[fedfuel_df1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    nrows6 = fedfuel_df1.shape[0]
    ncols6 = fedfuel_df1.shape[1]

    fedfuel_df2 = fedfuel_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows7 = fedfuel_df2.shape[0]
    ncols7 = fedfuel_df2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    econ_df2 = EGEDA_years[(EGEDA_years['economy'] == economy) &
                        (EGEDA_years['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years['fuel_code'].isin(['12_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    econ_df2 = econ_df2[['fuel_code', 'item_code_new'] + list(econ_df2.loc[:,'2000':])]
    
    nrows2 = econ_df2.shape[0]
    ncols2 = econ_df2.shape[1]

    # Now build aggregate sector variables
    
    buildings = econ_df2[econ_df2['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '12_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = econ_df2[econ_df2['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '12_total',
                                                                                                         item_code_new = 'Agriculture')
    
    # Build aggregate data frame of FED sector

    fedsector_df1 = econ_df2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(econ_df2.loc[:, '2000':])].reset_index(drop = True)

    fedsector_df1.loc[fedsector_df1['item_code_new'] == '13_industry_sector', 'item_code_new'] = 'Industry'
    fedsector_df1.loc[fedsector_df1['item_code_new'] == '14_transport_sector', 'item_code_new'] = 'Transport'
    fedsector_df1.loc[fedsector_df1['item_code_new'] == '16_nonenergy_use', 'item_code_new'] = 'Non-energy'
    fedsector_df1.loc[fedsector_df1['item_code_new'] == '15_4_nonspecified_others', 'item_code_new'] = 'Non-specified'

    fedsector_df1 = fedsector_df1[fedsector_df1['item_code_new'].isin(FED_agg_sectors)].set_index('item_code_new').loc[FED_agg_sectors].reset_index()
    fedsector_df1 = fedsector_df1[['fuel_code', 'item_code_new'] + list(fedsector_df1.loc[:, '2000':])]

    nrows8 = fedsector_df1.shape[0]
    ncols8 = fedsector_df1.shape[1]

    fedsector_df2 = fedsector_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows9 = fedsector_df2.shape[0]
    ncols9 = fedsector_df2.shape[1]
    
    # Third data frame construction: Buildings FED by fuel
    bld_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) &
                         (EGEDA_years['item_code_new'].isin(Buildings_items)) &
                         (EGEDA_years['fuel_code'].isin(Required_fuels))]
    
    for fuel in Required_fuels:
        buildings = bld_df1[bld_df1['fuel_code'] == fuel].groupby(['economy', 'fuel_code']).sum().assign(item_code_new = '15_x_buildings')
        buildings['economy'] = economy
        buildings['fuel_code'] = fuel
        
        bld_df1 = bld_df1.append(buildings).reset_index(drop = True)
        
    bld_df1 = bld_df1[['fuel_code', 'item_code_new'] + col_chart_years]
    
    nrows3 = bld_df1.shape[0]
    ncols3 = bld_df1.shape[1]

    bld_df2 = bld_df1[bld_df1['item_code_new'] == '15_x_buildings']

    coal = bld_df2[bld_df2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal', item_code_new = '15_x_buildings')
    
    oil = bld_df2[bld_df2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil', item_code_new = '15_x_buildings')
    
    renewables = bld_df2[bld_df2['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables', item_code_new = '15_x_buildings')
    
    heat_others = bld_df2[bld_df2['fuel_code'].isin(Heat_others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others', item_code_new = '15_x_buildings')

    bld_df2 = bld_df2.append([coal, oil, renewables, heat_others])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)

    bld_df2.loc[bld_df2['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    bld_df2.loc[bld_df2['fuel_code'] == '10_electricity', 'fuel_code'] = 'Electricity'

    bld_df2 = bld_df2[bld_df2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    nrows12 = bld_df2.shape[0]
    ncols12 = bld_df2.shape[1]

    bld_df3 = bld_df1[(bld_df1['fuel_code'] == '12_total') &
                      (bld_df1['item_code_new'].isin(['15_1_1_commerce_and_public_services', '15_1_2_residential']))].copy()

    bld_df3.loc[bld_df3['item_code_new'] == '15_1_1_commerce_and_public_services', 'item_code_new'] = 'Services' 
    bld_df3.loc[bld_df3['item_code_new'] == '15_1_2_residential', 'item_code_new'] = 'Residential'

    nrows13 = bld_df3.shape[0]
    ncols13 = bld_df3.shape[1]
    
    # Fourth data frame construction: Industry subsector
    ind_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) &
                         (EGEDA_years['item_code_new'].isin(Subindustry)) &
                         (EGEDA_years['fuel_code'] == '12_total')]

    other_industry = ind_df1[ind_df1['item_code_new'].isin(Other_industry)].groupby(['fuel_code']).sum().assign(item_code_new = 'Other',
                                                                                                                fuel_code = '12_total')

    ind_df1 = ind_df1.append([other_industry])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)

    ind_df1.loc[ind_df1['item_code_new'] == '13_1_iron_and_steel', 'item_code_new'] = 'Iron & steel'
    ind_df1.loc[ind_df1['item_code_new'] == '13_2_chemical_incl__petrochemical', 'item_code_new'] = 'Chemicals'
    ind_df1.loc[ind_df1['item_code_new'] == '13_3_nonferrous_metals', 'item_code_new'] = 'Aluminium'
    ind_df1.loc[ind_df1['item_code_new'] == '13_4_nonmetallic_mineral_products', 'item_code_new'] = 'Non-metallic minerals'  
    ind_df1.loc[ind_df1['item_code_new'] == '13_7_mining_and_quarrying', 'item_code_new'] = 'Mining'
    ind_df1.loc[ind_df1['item_code_new'] == '13_9_pulp_paper_and_printing', 'item_code_new'] = 'Pulp & paper'
    ind_df1.loc[ind_df1['item_code_new'] == '13_13_nonspecified_industry', 'item_code_new'] = 'Non-specified'
    
    ind_df1 = ind_df1[ind_df1['item_code_new'].isin(Industry_eight)].set_index('item_code_new').loc[Industry_eight].reset_index()

    ind_df1 = ind_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows4 = ind_df1.shape[0]
    ncols4 = ind_df1.shape[1]
    
    # Fifth data frame construction: Industry by fuel
    ind_df2 = EGEDA_years[(EGEDA_years['economy'] == economy) &
                         (EGEDA_years['item_code_new'].isin(['13_industry_sector'])) &
                         (EGEDA_years['fuel_code'].isin(Required_fuels))]
    
    coal = ind_df2[ind_df2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal', 
                                                                                                             item_code_new = '13_industry_sector')
    
    oil = ind_df2[ind_df2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil', 
                                                                                                item_code_new = '13_industry_sector')
    
    renewables = ind_df2[ind_df2['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables', 
                                                                                                              item_code_new = '13_industry_sector')
    
    heat_others = ind_df2[ind_df2['fuel_code'].isin(Heat_others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others', 
                                                                                                                item_code_new = '13_industry_sector')
    
    ind_df2 = ind_df2.append([coal, oil, renewables, heat_others])[['fuel_code', 
                                                                    'item_code_new'] + col_chart_years].reset_index(drop = True)

    ind_df2.loc[ind_df2['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    ind_df2.loc[ind_df2['fuel_code'] == '10_electricity', 'fuel_code'] = 'Electricity'                                                                    

    ind_df2 = ind_df2[ind_df2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()
    
    nrows5 = ind_df2.shape[0]
    ncols5 = ind_df2.shape[1]

    # Transport data frame construction: FED by fuels
    transport_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                          (EGEDA_years['item_code_new'].isin(['14_transport_sector'])) &
                          (EGEDA_years['fuel_code'].isin(Transport_fuels))]
    
    renewables = transport_df1[transport_df1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '14_transport_sector')
    
    others = transport_df1[transport_df1['fuel_code'].isin(Other_fuel)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '14_transport_sector')
    
    transport_df1 = transport_df1.append([renewables, others])[['fuel_code', 'item_code_new'] + list(transport_df1.loc[:, '2000':])].reset_index(drop = True)

    transport_df1.loc[transport_df1['fuel_code'] == '4_1_gasoline', 'fuel_code'] = 'Gasoline'
    transport_df1.loc[transport_df1['fuel_code'] == '4_5_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    transport_df1.loc[transport_df1['fuel_code'] == '5_1_natural_gas', 'fuel_code'] = 'Gas'
    transport_df1.loc[transport_df1['fuel_code'] == '4_7_lpg', 'fuel_code'] = 'LPG'
    transport_df1.loc[transport_df1['fuel_code'] == '10_electricity', 'fuel_code'] = 'Electricity'
    transport_df1.loc[transport_df1['fuel_code'] == '4_3_jet_fuel', 'fuel_code'] = 'Jet fuel'

    transport_df1 = transport_df1[transport_df1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    nrows10 = transport_df1.shape[0]
    ncols10 = transport_df1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    transport_df2 = EGEDA_years[(EGEDA_years['economy'] == economy) &
                               (EGEDA_years['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years['fuel_code'].isin(['12_total']))].copy()
    
    transport_df2.loc[transport_df2['item_code_new'] == '14_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    transport_df2.loc[transport_df2['item_code_new'] == '14_2_road', 'item_code_new'] = 'Road'
    transport_df2.loc[transport_df2['item_code_new'] == '14_3_rail', 'item_code_new'] = 'Rail'
    transport_df2.loc[transport_df2['item_code_new'] == '14_4_domestic_water_transport', 'item_code_new'] = 'Marine'
    transport_df2.loc[transport_df2['item_code_new'] == '14_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    transport_df2.loc[transport_df2['item_code_new'] == '14_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    transport_df2 = transport_df2[transport_df2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    transport_df2 = transport_df2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    nrows11 = transport_df2.shape[0]
    ncols11 = transport_df2.shape[1]

    # Agriculture data frame

    ag_df1 = EGEDA_years[(EGEDA_years['economy'] == economy) & 
                         (EGEDA_years['item_code_new'].isin(['15_2_agriculture', '15_3_fishing'])) &
                         (EGEDA_years['fuel_code'].isin(Required_fuels))].groupby('fuel_code').sum().assign(item_code_new = 'Agriculture').reset_index()
                     
    coal = ag_df1[ag_df1['fuel_code'].isin(Coal_fuels)].groupby('item_code_new').sum().assign(fuel_code = 'Coal', item_code_new = 'Agriculture')

    oil = ag_df1[ag_df1['fuel_code'].isin(Oil_fuels)].groupby('item_code_new').sum().assign(fuel_code = 'Oil', item_code_new = 'Agriculture')

    renewables = ag_df1[ag_df1['fuel_code'].isin(Renewables_fuels)].groupby('item_code_new').sum().assign(fuel_code = 'Renewables', item_code_new = 'Agriculture')
    
    heat_others = ag_df1[ag_df1['fuel_code'].isin(Heat_others_fuels)].groupby('item_code_new').sum().assign(fuel_code = 'Heat & others', item_code_new = 'Agriculture')
    
    ag_df1 = ag_df1.append([coal, oil, renewables, heat_others])[['fuel_code', 'item_code_new'] + list(ag_df1.loc[:,'2000':'2050'])].reset_index(drop = True)

    ag_df1.loc[ag_df1['fuel_code'] == '5_gas', 'fuel_code'] = 'Gas'
    ag_df1.loc[ag_df1['fuel_code'] == '10_electricity', 'fuel_code'] = 'Electricity'                                                                    

    ag_df1 = ag_df1[ag_df1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()
    
    nrows14 = ag_df1.shape[0]
    ncols14 = ag_df1.shape[1]

    ag_df2 = ag_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows15 = ag_df2.shape[0]
    ncols15 = ag_df2.shape[1]         
    
    # Define directory
    script_dir = './results/' + month_year + '/FED/'
    results_dir = os.path.join(script_dir, 'economy_breakdown/', economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_fed_tfc.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Insert the various dataframes into different sheets of the workbook    
    fedfuel_df1.to_excel(writer, sheet_name = economy + '_FED_fuel', index = False, startrow = chart_height)
    fedfuel_df2.to_excel(writer, sheet_name = economy + '_FED_fuel', index = False, startrow = chart_height + nrows6 + 3)
    econ_df2.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = chart_height)
    fedsector_df1.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = chart_height + nrows2 + 3)
    fedsector_df2.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = chart_height + nrows2 + nrows8 + 6)
    bld_df2.to_excel(writer, sheet_name = economy + '_FED_bld_sector', index = False, startrow = chart_height)
    bld_df3.to_excel(writer, sheet_name = economy + '_FED_bld_sector', index = False, startrow = chart_height + nrows12 + 3)
    ind_df1.to_excel(writer, sheet_name = economy + '_FED_ind', index = False, startrow = chart_height)
    ind_df2.to_excel(writer, sheet_name = economy + '_FED_ind', index = False, startrow = chart_height + nrows4 + 2)
    transport_df1.to_excel(writer, sheet_name = economy + '_FED_transport', index = False, startrow = chart_height)
    transport_df2.to_excel(writer, sheet_name = economy + '_FED_transport', index = False, startrow = chart_height + nrows10 + 3)
    ag_df1.to_excel(writer, sheet_name = economy + '_FED_agriculture', index = False, startrow = chart_height)
    ag_df2.to_excel(writer, sheet_name = economy + '_FED_agriculture', index = False, startrow = chart_height + nrows14 + 3)
    
    # Access the workbook and first sheet with data from df1
    worksheet1 = writer.sheets[economy + '_FED_fuel']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    worksheet1.set_column(1, ncols6 + 1, None, comma_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height + nrows6 + 3, None, header_format)
    worksheet1.write(0, 0, economy + ' FED fuel', cell_format1)

    ################################################################### CHARTS ###################################################################

    # Create a FED area chart
    fedfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    fedfuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    fedfuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    fedfuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    fedfuel_chart1.set_y_axis({
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
        
    fedfuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fedfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows6):
        fedfuel_chart1.add_series({
            'name':       [economy + '_FED_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_FED_fuel', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet1.insert_chart('B3', fedfuel_chart1)

    ###################### Create another FED chart showing proportional share #################################

    # Create a TPES chart
    fedfuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    fedfuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    fedfuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    fedfuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    fedfuel_chart2.set_y_axis({
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
        
    fedfuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fedfuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_fuels:
        i = fedfuel_df2[fedfuel_df2['fuel_code'] == component].index[0]
        fedfuel_chart2.add_series({
            'name':       [economy + '_FED_fuel', chart_height + nrows6 + i + 4, 0],
            'categories': [economy + '_FED_fuel', chart_height + nrows6 + 3, 2, chart_height + nrows6 + 3, ncols7 - 1],
            'values':     [economy + '_FED_fuel', chart_height + nrows6 + i + 4, 2, chart_height + nrows6 + i + 4, ncols7 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet1.insert_chart('J3', fedfuel_chart2)

    # Create a FED line chart with higher level aggregation
    fedfuel_chart3 = workbook.add_chart({'type': 'line'})
    fedfuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    fedfuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    fedfuel_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    fedfuel_chart3.set_y_axis({
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
        
    fedfuel_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fedfuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows6):
        fedfuel_chart3.add_series({
            'name':       [economy + '_FED_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_FED_fuel', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    worksheet1.insert_chart('R3', fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    worksheet2 = writer.sheets[economy + '_FED_sector']
        
    # Apply comma format and header format to relevant data rows
    worksheet2.set_column(1, ncols2 + 1, None, comma_format)
    worksheet2.set_row(chart_height, None, header_format)
    worksheet2.set_row(chart_height + nrows2 + 3, None, header_format)
    worksheet2.set_row(chart_height + nrows2 + nrows8 + 6, None, header_format)
    worksheet2.write(0, 0, economy + ' FED sector', cell_format1)
    
    # Create a FED chart
    fed_sector_chart1 = workbook.add_chart({'type': 'line'})
    fed_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    fed_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    fed_sector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    fed_sector_chart1.set_y_axis({
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
        
    fed_sector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fed_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows2):
        fed_sector_chart1.add_series({
            'name':       [economy + '_FED_sector', chart_height + i + 1, 1],
            'categories': [economy + '_FED_sector', chart_height, 2, chart_height, ncols2 - 1],
            'values':     [economy + '_FED_sector', chart_height + i + 1, 2, chart_height + i + 1, ncols2 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    worksheet2.insert_chart('Z3', fed_sector_chart1)

    # Create a FED sector area chart

    fedsector_chart3 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    fedsector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    fedsector_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    fedsector_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    fedsector_chart3.set_y_axis({
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
        
    fedsector_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fedsector_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows8):
        fedsector_chart3.add_series({
            'name':       [economy + '_FED_sector', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_FED_sector', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_FED_sector', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet2.insert_chart('B3', fedsector_chart3)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    fedsector_chart4 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    fedsector_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    fedsector_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    fedsector_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    fedsector_chart4.set_y_axis({
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
        
    fedsector_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fedsector_chart4.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_sectors:
        i = fedsector_df2[fedsector_df2['item_code_new'] == component].index[0]
        fedsector_chart4.add_series({
            'name':       [economy + '_FED_sector', chart_height + nrows2 + nrows8 + i + 7, 1],
            'categories': [economy + '_FED_sector', chart_height + nrows2 + nrows8 + 6, 2, chart_height + nrows2 + nrows8 + 6, ncols9 - 1],
            'values':     [economy + '_FED_sector', chart_height + nrows2 + nrows8 + i + 7, 2, chart_height + nrows2 + nrows8 + i + 7, ncols9 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet2.insert_chart('J3', fedsector_chart4)

    # Create a FED sector line chart

    fedsector_chart5 = workbook.add_chart({'type': 'line'})
    fedsector_chart5.set_size({
        'width': 500,
        'height': 300
    })
    
    fedsector_chart5.set_chartarea({
        'border': {'none': True}
    })
    
    fedsector_chart5.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    fedsector_chart5.set_y_axis({
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
        
    fedsector_chart5.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fedsector_chart5.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows8):
        fedsector_chart5.add_series({
            'name':       [economy + '_FED_sector', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_FED_sector', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_FED_sector', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    worksheet2.insert_chart('R3', fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    worksheet3 = writer.sheets[economy + '_FED_bld_sector']
    
    # Apply comma format and header format to relevant data rows
    worksheet3.set_column(2, ncols3 + 1, None, comma_format)
    worksheet3.set_row(chart_height, None, header_format)
    worksheet3.set_row(chart_height + nrows12 + 3, None, header_format)
    worksheet3.write(0, 0, economy + ' buildings', cell_format1)
    
    # Create a FED chart
    fed_bld_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    fed_bld_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    fed_bld_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    fed_bld_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    fed_bld_chart1.set_y_axis({
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
        
    fed_bld_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fed_bld_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_fuels:
        i = bld_df2[bld_df2['fuel_code'] == component].index[0]
        fed_bld_chart1.add_series({
            'name':       [economy + '_FED_bld_sector', chart_height + i + 1, 0],
            'categories': [economy + '_FED_bld_sector', chart_height, 2, chart_height, ncols12 - 1],
            'values':     [economy + '_FED_bld_sector', chart_height + i + 1, 2, chart_height + i + 1, ncols12 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })

    worksheet3.insert_chart('B3', fed_bld_chart1)
    
    ################## FED building chart 2 (residential versus services) ###########################################
    
    # Create a second FED building chart
    fed_bld_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    fed_bld_chart2.set_size({
        'width': 500,
        'height': 300
    })

    fed_bld_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    fed_bld_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    fed_bld_chart2.set_y_axis({
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
        
    fed_bld_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fed_bld_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(2):
        fed_bld_chart2.add_series({
            'name':       [economy + '_FED_bld_sector', chart_height + nrows12 + 4 + i, 1],
            'categories': [economy + '_FED_bld_sector', chart_height + nrows12 + 3, 2, chart_height + nrows12 + 3, ncols13 - 1],
            'values':     [economy + '_FED_bld_sector', chart_height + nrows12 + 4 + i, 2, chart_height + nrows12 + 4 + i, ncols13 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    worksheet3.insert_chart('J3', fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # Access the workbook and fourth sheet with data from bld_df1
    worksheet4 = writer.sheets[economy + '_FED_ind']
    
    # Apply comma format and header format to relevant data rows
    worksheet4.set_column(2, ncols4 + 1, None, comma_format)
    worksheet4.set_row(chart_height, None, header_format)
    worksheet4.set_row(chart_height + nrows4 + 2, None, header_format)
    worksheet4.write(0, 0, economy + ' industry', cell_format1)
    
    # Create a industry subsector FED chart
    fed_ind_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    fed_ind_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    fed_ind_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    fed_ind_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    fed_ind_chart1.set_y_axis({
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
        
    fed_ind_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fed_ind_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows4):
        fed_ind_chart1.add_series({
            'name':       [economy + '_FED_ind', chart_height + i + 1, 1],
            'categories': [economy + '_FED_ind', chart_height, 2, chart_height, ncols4 - 1],
            'values':     [economy + '_FED_ind', chart_height + i + 1, 2, chart_height + i + 1, ncols4 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet4.insert_chart('B3', fed_ind_chart1)
    
    ############# FED industry chart 2 (industry by fuel)
    
    # Create a FED industry fuel chart
    fed_ind_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    fed_ind_chart2.set_size({
        'width': 500,
        'height': 300
    })

    fed_ind_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    fed_ind_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    fed_ind_chart2.set_y_axis({
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
        
    fed_ind_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    fed_ind_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for fuel_agg in FED_agg_fuels:
        j = ind_df2[ind_df2['fuel_code'] == fuel_agg].index[0]
        fed_ind_chart2.add_series({
            'name':       [economy + '_FED_ind', chart_height + nrows4 + j + 3, 0],
            'categories': [economy + '_FED_ind', chart_height + nrows4 + 2, 2, chart_height + nrows4 + 2, ncols5 - 1],
            'values':     [economy + '_FED_ind', chart_height + nrows4 + j + 3, 2, chart_height + nrows4 + j + 3, ncols5 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True}
        })
    
    worksheet4.insert_chart('J3', fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    worksheet5 = writer.sheets[economy + '_FED_transport']
        
    # Apply comma format and header format to relevant data rows
    worksheet5.set_column(2, ncols10 + 1, None, comma_format)
    worksheet5.set_row(chart_height, None, header_format)
    worksheet5.set_row(chart_height + nrows10 + 3, None, header_format)
    worksheet5.write(0, 0, economy + ' FED transport', cell_format1)
    
    # Create a transport FED area chart
    transport_chart1 = workbook.add_chart({'type': 'area', 
                                           'subtype': 'stacked'})
    transport_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    transport_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    transport_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    transport_chart1.set_y_axis({
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
        
    transport_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    transport_chart1.set_title({
        'none': True
    })
        
    for fuel_agg in Transport_fuels_agg:
        j = transport_df1[transport_df1['fuel_code'] == fuel_agg].index[0]
        transport_chart1.add_series({
            'name':       [economy + '_FED_transport', chart_height + j + 1, 0],
            'categories': [economy + '_FED_transport', chart_height, 2, chart_height, ncols10 - 1],
            'values':     [economy + '_FED_transport', chart_height + j + 1, 2, chart_height + j + 1, ncols10 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True} 
        })
    
    worksheet5.insert_chart('B3', transport_chart1)
            
    ############# FED transport chart 2 (transport by modality)
    
    # Create a FED transport modality column chart
    transport_chart2 = workbook.add_chart({'type': 'column', 
                                         'subtype': 'stacked'})
    transport_chart2.set_size({
        'width': 500,
        'height': 300
    })

    transport_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    transport_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    transport_chart2.set_y_axis({
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
        
    transport_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    transport_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for modality in Transport_modal_agg:
        j = transport_df2[transport_df2['item_code_new'] == modality].index[0]
        transport_chart2.add_series({
            'name':       [economy + '_FED_transport', chart_height + nrows10 + j + 4, 1],
            'categories': [economy + '_FED_transport', chart_height + nrows10 + 3, 2, chart_height + nrows10 + 3, ncols11 - 1],
            'values':     [economy + '_FED_transport', chart_height + nrows10 + j + 4, 2, chart_height + nrows10 + j + 4, ncols11 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True}
        })
    
    worksheet5.insert_chart('J3', transport_chart2)

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    worksheet6 = writer.sheets[economy + '_FED_agriculture']
        
    # Apply comma format and header format to relevant data rows
    worksheet6.set_column(2, ncols14 + 1, None, comma_format)
    worksheet6.set_row(chart_height, None, header_format)
    worksheet6.set_row(chart_height + nrows14 + 3, None, header_format)
    worksheet6.write(0, 0, economy + ' FED agriculture', cell_format1)

    # Create a Agriculture line chart 
    ag_chart1 = workbook.add_chart({'type': 'line'})
    ag_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ag_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ag_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ag_chart1.set_y_axis({
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
        
    ag_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ag_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows14):
        ag_chart1.add_series({
            'name':       [economy + '_FED_agriculture', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agriculture', chart_height, 2, chart_height, ncols14 - 1],
            'values':     [economy + '_FED_agriculture', chart_height + i + 1, 2, chart_height + i + 1, ncols14 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    worksheet6.insert_chart('B3', ag_chart1)

    # Create a Agriculture area chart
    ag_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ag_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ag_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ag_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ag_chart2.set_y_axis({
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
        
    ag_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ag_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows14):
        ag_chart2.add_series({
            'name':       [economy + '_FED_agriculture', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agriculture', chart_height, 2, chart_height, ncols14 - 1],
            'values':     [economy + '_FED_agriculture', chart_height + i + 1, 2, chart_height + i + 1, ncols14 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet6.insert_chart('J3', ag_chart2)

    # Create a Agriculture stacked column
    ag_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ag_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ag_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ag_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ag_chart3.set_y_axis({
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
        
    ag_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ag_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows15):
        ag_chart3.add_series({
            'name':       [economy + '_FED_agriculture', chart_height + nrows14 + i + 4, 0],
            'categories': [economy + '_FED_agriculture', chart_height + nrows14 + 3, 2, chart_height + nrows14 + 3, ncols15 - 1],
            'values':     [economy + '_FED_agriculture', chart_height + nrows14 + i + 4, 2, chart_height + nrows14 + i + 4, ncols15 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    worksheet6.insert_chart('R3', ag_chart3)
        
    writer.save()




