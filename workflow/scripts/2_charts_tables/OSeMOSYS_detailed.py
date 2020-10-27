# Building more charts in a py script

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
import xlsxwriter
import pandas.io.formats.excel
import glob
import re
import pdb

# Path to look for OSeMOSYS output (for charting and tabling)
# NB: this directory is different to when operating in jupyter notebook
# Directory is directory as defined in git bash rather than where script is saved

path_output = '../../data/3_OSeMOSYS_output'

# They're csv files so use a wild card (*) to grab the filenames
OSeMOSYS_filenames = glob.glob(path_output + "/*.xlsx")

# Create empty dataframe to store multiple UseByTechnology economy data 

use_df1 = pd.DataFrame()

# Read in the OSeMOSYS output files so that that they're all in the currently empty dataframe (use_df1)

for i in range(len(OSeMOSYS_filenames)):
    _df = pd.read_excel(OSeMOSYS_filenames[i], sheet_name = 'UseByTechnology')
    use_df1 = use_df1.append(_df) 

use_df1 = use_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION']).sum().reset_index()

################### THINGS FOR THE BELOW CHART BUILDS ###################

# define month and year for charts
month_year = pd.to_datetime('today').strftime('%B_%Y')

# Chart years for column charts
col_chart_years = [2017, 2020, 2030, 2040, 2050]
    
# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written

# Get maximum year column to build data frames below

year_columns = []

for item in list(use_df1.columns):
    try:
        year_columns.append(int(item))
    except ValueError:
            pass

max_year = max(year_columns)

OSeMOSYS_years = list(range(2017, max_year + 1))

# Colours
# NB again, the directory to the HEX colours depends on the working directory you define in git bash or at the command prompt
# If in Jupyter lab, the address assumes you start in the folder that the jupyter notebook is saved in

colours = pd.read_excel('../../data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

###############################################################################################################

# Charts and tables for the different demand sectors

for economy in use_df1['REGION'].unique():
    # Agriculture
    ag_df1 = use_df1[(use_df1['TECHNOLOGY'].str.startswith('AGR')) &
                     (use_df1['REGION'] == economy)].reset_index(drop = True)\
                         [['TECHNOLOGY', 'FUEL'] + OSeMOSYS_years]

    # Agriculture by fuel
    ag_fuel = ag_df1.groupby(['FUEL'], as_index = False)[OSeMOSYS_years].sum().assign(sector = 'Agriculture')[['sector', 'FUEL'] + OSeMOSYS_years]

    ag_fuel.loc[ag_fuel['FUEL'] == '10_electricity', 'FUEL'] = 'Electricity'
    ag_fuel.loc[ag_fuel['FUEL'] == '4_10_other_petroleum_products', 'FUEL'] = 'Other petroleum products'
    ag_fuel.loc[ag_fuel['FUEL'] == '4_4_other_kerosene', 'FUEL'] = 'Other kerosene'
    ag_fuel.loc[ag_fuel['FUEL'] == '4_5_gas_diesel_oil', 'FUEL'] = 'Gas diesel oil'
    ag_fuel.loc[ag_fuel['FUEL'] == '4_7_lpg', 'FUEL'] = 'LPG'
    ag_fuel.loc[ag_fuel['FUEL'] == '5_1_natural_gas', 'FUEL'] = 'Natural gas'
    ag_fuel.loc[ag_fuel['FUEL'] == '8_1_geothermal_power', 'FUEL'] = 'Geothermal'

    ag_fuel_rows = ag_fuel.shape[0]
    ag_fuel_cols = ag_fuel.shape[1]

    # Agriculture by technology
    ag_nonenergy = ag_df1[ag_df1['TECHNOLOGY'].str.contains('non_energy')].reset_index(drop = True)
    ag_nonenergy['Ag sector'] = 'Non-energy'
    ag_nonenergy = ag_nonenergy.groupby(['Ag sector'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    ag_oil = ag_df1[ag_df1['TECHNOLOGY'].str.contains('crop_oil')].reset_index(drop = True)
    ag_oil['Ag sector'] = 'Oil'
    ag_oil = ag_oil.groupby(['Ag sector'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    ag_starch = ag_df1[ag_df1['TECHNOLOGY'].str.contains('starch')].reset_index(drop = True)
    ag_starch['Ag sector'] = 'Starch'
    ag_starch = ag_starch.groupby(['Ag sector'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    ag_sugar = ag_df1[ag_df1['TECHNOLOGY'].str.contains('sugar')].reset_index(drop = True)
    ag_sugar['Ag sector'] = 'Sugar'
    ag_sugar = ag_sugar.groupby(['Ag sector'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    fishing = ag_df1[ag_df1['TECHNOLOGY'].str.contains('fishing')].reset_index(drop = True)
    fishing['Ag sector'] = 'Fishing'
    fishing = fishing.groupby(['Ag sector'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    ag_tech = pd.DataFrame().append([ag_nonenergy, ag_oil, ag_starch, ag_sugar, fishing])[['Ag sector', 'FUEL'] + OSeMOSYS_years].reset_index(drop = True)

    ag_tech_rows = ag_tech.shape[0]
    ag_tech_cols = ag_tech.shape[1]

    #ind_df1 = use_df1[(use_df1['TECHNOLOGY'].str.startswith('IND')) &
    #                 (use_df1['REGION'] == economy)].reset_index(drop = True)\
    #                     [['TECHNOLOGY', 'FUEL'] + OSeMOSYS_years]

    # Buildings
    bld_df1 = use_df1[(use_df1['TECHNOLOGY'].str.startswith('BLD')) &
                     (use_df1['REGION'] == economy)].reset_index(drop = True)\
                         [['TECHNOLOGY', 'FUEL'] + OSeMOSYS_years]

    # Buildings by fuel
    bld_fuel = bld_df1.groupby(['FUEL'], as_index = False)[OSeMOSYS_years].sum().assign(sector = 'Buildings')[['sector', 'FUEL'] + OSeMOSYS_years]

    bld_fuel.loc[bld_fuel['FUEL'] == '1_1_1_coking_coal', 'FUEL'] = 'Coking coal'
    bld_fuel.loc[bld_fuel['FUEL'] == '1_3_lignite', 'FUEL'] = 'Lignite'
    bld_fuel.loc[bld_fuel['FUEL'] == '4_5_gas_diesel_oil', 'FUEL'] = 'Gas diesel oil'
    bld_fuel.loc[bld_fuel['FUEL'] == '4_7_lpg', 'FUEL'] = 'LPG'
    bld_fuel.loc[bld_fuel['FUEL'] == '5_1_natural_gas', 'FUEL'] = 'Natural gas'
    bld_fuel.loc[bld_fuel['FUEL'] == '8_3_geothermal_heat', 'FUEL'] = 'Geothermal'
    bld_fuel.loc[bld_fuel['FUEL'] == '8_4_solar_heat', 'FUEL'] = 'Solar heat'
    bld_fuel.loc[bld_fuel['FUEL'] == '9_4_other_biomass', 'FUEL'] = 'Biomass'
    bld_fuel.loc[bld_fuel['FUEL'] == '10_electricity_Dx', 'FUEL'] = 'Electricity'
    bld_fuel.loc[bld_fuel['FUEL'] == '11_heat', 'FUEL'] = 'Heat'

    bld_fuel_rows = bld_fuel.shape[0]
    bld_fuel_cols = bld_fuel.shape[1]

    # Buildings by technology
    bld_cook = bld_df1[bld_df1['TECHNOLOGY'].str.contains('Cook')].reset_index(drop = True)
    bld_cook['Bld tech'] = 'Cooking'
    bld_cook = bld_cook.groupby(['Bld tech'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    bld_light = bld_df1[bld_df1['TECHNOLOGY'].str.contains('Light')].reset_index(drop = True)
    bld_light['Bld tech'] = 'Lighting'
    bld_light = bld_light.groupby(['Bld tech'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    bld_sh = bld_df1[bld_df1['TECHNOLOGY'].str.contains('_SH_')].reset_index(drop = True)
    bld_sh['Bld tech'] = 'Space heating'
    bld_sh = bld_sh.groupby(['Bld tech'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    bld_wh = bld_df1[bld_df1['TECHNOLOGY'].str.contains('_WH_')].reset_index(drop = True)
    bld_wh['Bld tech'] = 'Water heating'
    bld_wh = bld_wh.groupby(['Bld tech'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    bld_other = bld_df1[bld_df1['TECHNOLOGY'].str.contains('_Other_')].reset_index(drop = True)
    bld_other['Bld tech'] = 'Other'
    bld_other = bld_other.groupby(['Bld tech'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    bld_sc = bld_df1[bld_df1['TECHNOLOGY'].str.contains('_SC_')].reset_index(drop = True)
    bld_sc['Bld tech'] = 'Space cooling'
    bld_sc = bld_sc.groupby(['Bld tech'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    bld_tech = pd.DataFrame().append([bld_sh, bld_wh, bld_sc, bld_cook, bld_light, bld_other])[['Bld tech', 'FUEL'] + OSeMOSYS_years].reset_index(drop = True)

    bld_tech_rows = bld_tech.shape[0]
    bld_tech_cols = bld_tech.shape[1]

    # Transport
    trn_df1 = use_df1[(use_df1['TECHNOLOGY'].str.startswith('TRN')) &
                     (use_df1['REGION'] == economy)].reset_index(drop = True)\
                         [['TECHNOLOGY', 'FUEL'] + OSeMOSYS_years]

    # Transport disaggregation I
    trn_af = trn_df1[trn_df1['TECHNOLOGY'].str.contains('air_freight')].reset_index(drop = True)
    trn_af['Mode'] = 'Air freight'
    trn_af = trn_af.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_ap = trn_df1[trn_df1['TECHNOLOGY'].str.contains('air_passenger')].reset_index(drop = True)
    trn_ap['Mode'] = 'Air passenger'
    trn_ap = trn_ap.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_bunk = trn_df1[trn_df1['TECHNOLOGY'].str.contains('bunkers')].reset_index(drop = True)
    trn_bunk['Mode'] = 'Bunkers'
    trn_bunk = trn_bunk.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_railf = trn_df1[trn_df1['TECHNOLOGY'].str.contains('rail_freight')].reset_index(drop = True)
    trn_railf['Mode'] = 'Rail freight'
    trn_railf = trn_railf.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_railp = trn_df1[trn_df1['TECHNOLOGY'].str.contains('rail_passenger')].reset_index(drop = True)
    trn_railp['Mode'] = 'Rail passenger'
    trn_railp = trn_railp.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_roadf = trn_df1[trn_df1['TECHNOLOGY'].str.contains('road_freight')].reset_index(drop = True)
    trn_roadf['Mode'] = 'Road freight'
    trn_roadf = trn_roadf.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_roadp = trn_df1[trn_df1['TECHNOLOGY'].str.contains('road_passenger')].reset_index(drop = True)
    trn_roadp['Mode'] = 'Road passenger'
    trn_roadp = trn_roadp.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_ship = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_ship_')].reset_index(drop = True)
    trn_ship['Mode'] = 'Marine'
    trn_ship = trn_ship.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_ag1 = pd.DataFrame().append([trn_af, trn_ap, trn_bunk, trn_railf, trn_railp, trn_roadf, trn_roadp, trn_ship])[['Mode', 'FUEL'] + OSeMOSYS_years].reset_index(drop = True)

    trn_ag1_rows = trn_ag1.shape[0]
    trn_ag1_cols = trn_ag1.shape[1]

    # Reduced number of years for charting

    trn_ag1_years = trn_ag1[['Mode', 'FUEL'] + col_chart_years]

    trn_ag1_years_rows = trn_ag1_years.shape[0]
    trn_ag1_years_cols = trn_ag1_years.shape[1]

    # Transport aggregation II
    trn_air = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_air_')].reset_index(drop = True)
    trn_air['Mode'] = 'Aviation'
    trn_air = trn_air.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    # Bunkers from aggregation above

    trn_rail = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_rail_')].reset_index(drop = True)
    trn_rail['Mode'] = 'Rail'
    trn_rail = trn_rail.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_2w = trn_df1[trn_df1['TECHNOLOGY'].str.contains('2W')].reset_index(drop = True)
    trn_2w['Mode'] = '2-wheeler'
    trn_2w = trn_2w.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_bus = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_BUS_')].reset_index(drop = True)
    trn_bus['Mode'] = 'Bus'
    trn_bus = trn_bus.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_lv = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_LV_')].reset_index(drop = True)
    trn_lv['Mode'] = 'Light vehicles'
    trn_lv = trn_lv.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_lt = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_LT_')].reset_index(drop = True)
    trn_lt['Mode'] = 'Light trucks'
    trn_lt = trn_lt.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    trn_ht = trn_df1[trn_df1['TECHNOLOGY'].str.contains('_HT_')].reset_index(drop = True)
    trn_ht['Mode'] = 'Heavy trucks'
    trn_ht = trn_ht.groupby(['Mode'], as_index = False)[OSeMOSYS_years].sum().assign(FUEL = 'All')

    # Marine from aggregation above

    trn_ag2 = pd.DataFrame().append([trn_air, trn_bunk, trn_rail, trn_2w, trn_bus, trn_lv, trn_lt, trn_ht, trn_ship])[['Mode', 'FUEL'] + OSeMOSYS_years].reset_index(drop = True)

    trn_ag2_rows = trn_ag2.shape[0]
    trn_ag2_cols = trn_ag2.shape[1]

    # Reduced number of years for charting

    trn_ag2_years = trn_ag2[['Mode', 'FUEL'] + col_chart_years]

    trn_ag2_years_rows = trn_ag2_years.shape[0]
    trn_ag2_years_cols = trn_ag2_years.shape[1]

    # Electric vehicles

    # Define directory
    script_dir = '../../results/' + month_year + '/OSeMOSYS_detailed/'
    results_dir = os.path.join(script_dir, 'economy_breakdown/', economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)

    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_demand_detailed.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    ag_fuel.to_excel(writer, sheet_name = economy + '_ag_use', index = False, startrow = chart_height)
    ag_tech.to_excel(writer, sheet_name = economy + '_ag_use', index = False, startrow = chart_height + ag_fuel_rows + 3)
    bld_fuel.to_excel(writer, sheet_name = economy + '_bld_use', index = False, startrow = chart_height)
    bld_tech.to_excel(writer, sheet_name = economy + '_bld_use', index = False, startrow = chart_height + bld_fuel_rows + 3)
    trn_ag1.to_excel(writer, sheet_name = economy + '_trn_use', index = False, startrow = chart_height)
    trn_ag1_years.to_excel(writer, sheet_name = economy + '_trn_use', index = False, startrow = chart_height + trn_ag1_rows + 3)
    trn_ag2.to_excel(writer, sheet_name = economy + '_trn_use', index = False, startrow = chart_height + trn_ag1_rows + trn_ag1_years_rows + 6)
    trn_ag2_years.to_excel(writer, sheet_name = economy + '_trn_use', index = False, startrow = chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + 9)

    # Access the workbook and first sheet with data from df1
    worksheet1 = writer.sheets[economy + '_ag_use']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    worksheet1.set_column(2, ag_fuel_cols + 1, None, comma_format)
    worksheet1.set_row(chart_height, None, header_format)
    worksheet1.set_row(chart_height + ag_fuel_rows + 3, None, header_format)
    worksheet1.write(0, 0, economy + ' Agriculture OSeMOSYS output by fuel and by technology', cell_format1)

    # Create a ag fuel area chart
    agfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    agfuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    agfuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    agfuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    agfuel_chart1.set_y_axis({
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
        
    agfuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    agfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ag_fuel_rows):
        agfuel_chart1.add_series({
            'name':       [economy + '_ag_use', chart_height + i + 1, 1],
            'categories': [economy + '_ag_use', chart_height, 2, chart_height, ag_fuel_cols - 1],
            'values':     [economy + '_ag_use', chart_height + i + 1, 2, chart_height + i + 1, ag_fuel_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet1.insert_chart('B3', agfuel_chart1)

    # Create a ag fuel line chart 
    agfuel_chart2 = workbook.add_chart({'type': 'line'})
    agfuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    agfuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    agfuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    agfuel_chart2.set_y_axis({
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
        
    agfuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    agfuel_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ag_fuel_rows):
        agfuel_chart2.add_series({
            'name':       [economy + '_ag_use', chart_height + i + 1, 1],
            'categories': [economy + '_ag_use', chart_height, 2, chart_height, ag_fuel_cols - 1],
            'values':     [economy + '_ag_use', chart_height + i + 1, 2, chart_height + i + 1, ag_fuel_cols - 1],
            'line':       {'color': colours_hex[i], 'width': 1}
        })    
        
    worksheet1.insert_chart('J3', agfuel_chart2)

    # Create a agriculture tech area chart
    agsector_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    agsector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    agsector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    agsector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    agsector_chart1.set_y_axis({
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
        
    agsector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    agsector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ag_tech_rows):
        agsector_chart1.add_series({
            'name':       [economy + '_ag_use', chart_height + ag_fuel_rows + i + 4, 0],
            'categories': [economy + '_ag_use', chart_height + ag_fuel_rows + 3, 2, chart_height + ag_fuel_rows + 3, ag_tech_cols - 1],
            'values':     [economy + '_ag_use', chart_height + ag_fuel_rows + i + 4, 2, chart_height + ag_fuel_rows + i + 4, ag_tech_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet1.insert_chart('R3', agsector_chart1)

    # Create a ag fuel line chart 
    agsector_chart2 = workbook.add_chart({'type': 'line'})
    agsector_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    agsector_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    agsector_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    agsector_chart2.set_y_axis({
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
        
    agsector_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    agsector_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(ag_tech_rows):
        agsector_chart2.add_series({
            'name':       [economy + '_ag_use', chart_height + ag_fuel_rows + i + 4, 0],
            'categories': [economy + '_ag_use', chart_height + ag_fuel_rows + 3, 2, chart_height + ag_fuel_rows + 3, ag_tech_cols - 1],
            'values':     [economy + '_ag_use', chart_height + ag_fuel_rows + i + 4, 2, chart_height + ag_fuel_rows + i + 4, ag_tech_cols - 1],
            'line':       {'color': colours_hex[i], 'width': 1}
        })    
        
    worksheet1.insert_chart('Z3', agsector_chart2)

    ############################################################################################################################

    # Access the workbook and first sheet with data from df1
    worksheet2 = writer.sheets[economy + '_bld_use']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    worksheet2.set_column(2, bld_fuel_cols + 1, None, comma_format)
    worksheet2.set_row(chart_height, None, header_format)
    worksheet2.set_row(chart_height + bld_fuel_rows + 3, None, header_format)
    worksheet2.write(0, 0, economy + ' Buildings OSeMOSYS output by fuel and by technology', cell_format1)

    # Create a bld fuel area chart
    bldfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    bldfuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    bldfuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    bldfuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    bldfuel_chart1.set_y_axis({
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
        
    bldfuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    bldfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(bld_fuel_rows):
        bldfuel_chart1.add_series({
            'name':       [economy + '_bld_use', chart_height + i + 1, 1],
            'categories': [economy + '_bld_use', chart_height, 2, chart_height, bld_fuel_cols - 1],
            'values':     [economy + '_bld_use', chart_height + i + 1, 2, chart_height + i + 1, bld_fuel_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet2.insert_chart('B3', bldfuel_chart1)

    # Create a bld fuel line chart 
    bldfuel_chart2 = workbook.add_chart({'type': 'line'})
    bldfuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    bldfuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    bldfuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    bldfuel_chart2.set_y_axis({
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
        
    bldfuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    bldfuel_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(bld_fuel_rows):
        bldfuel_chart2.add_series({
            'name':       [economy + '_bld_use', chart_height + i + 1, 1],
            'categories': [economy + '_bld_use', chart_height, 2, chart_height, bld_fuel_cols - 1],
            'values':     [economy + '_bld_use', chart_height + i + 1, 2, chart_height + i + 1, bld_fuel_cols - 1],
            'line':       {'color': colours_hex[i], 'width': 1}
        })    
        
    worksheet2.insert_chart('J3', bldfuel_chart2)

    # Create a Buildings tech area chart
    bldtech_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    bldtech_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    bldtech_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    bldtech_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    bldtech_chart1.set_y_axis({
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
        
    bldtech_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    bldtech_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(bld_tech_rows):
        bldtech_chart1.add_series({
            'name':       [economy + '_bld_use', chart_height + bld_fuel_rows + i + 4, 0],
            'categories': [economy + '_bld_use', chart_height + bld_fuel_rows + 3, 2, chart_height + bld_fuel_rows + 3, bld_tech_cols - 1],
            'values':     [economy + '_bld_use', chart_height + bld_fuel_rows + i + 4, 2, chart_height + bld_fuel_rows + i + 4, bld_tech_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet2.insert_chart('R3', bldtech_chart1)

    # Create a bld fuel line chart 
    bldtech_chart2 = workbook.add_chart({'type': 'line'})
    bldtech_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    bldtech_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    bldtech_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    bldtech_chart2.set_y_axis({
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
        
    bldtech_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    bldtech_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(bld_tech_rows):
        bldtech_chart2.add_series({
            'name':       [economy + '_bld_use', chart_height + bld_fuel_rows + i + 4, 0],
            'categories': [economy + '_bld_use', chart_height + bld_fuel_rows + 3, 2, chart_height + bld_fuel_rows + 3, bld_tech_cols - 1],
            'values':     [economy + '_bld_use', chart_height + bld_fuel_rows + i + 4, 2, chart_height + bld_fuel_rows + i + 4, bld_tech_cols - 1],
            'line':       {'color': colours_hex[i], 'width': 1}
        })    
        
    worksheet2.insert_chart('Z3', bldtech_chart2)

    ####################################################################################################################################

    # Access the workbook and first sheet with data from df1
    worksheet3 = writer.sheets[economy + '_trn_use']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    worksheet3.set_column(2, bld_fuel_cols + 1, None, comma_format)
    worksheet3.set_row(chart_height, None, header_format)
    worksheet3.set_row(chart_height + trn_ag1_rows + 3, None, header_format)
    worksheet3.set_row(chart_height + trn_ag1_rows + trn_ag1_years_rows + 6, None, header_format)
    worksheet3.set_row(chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + 9, None, header_format)
    worksheet3.write(0, 0, economy + ' Transport OSeMOSYS output by modalities', cell_format1)

    # Create a transport fuel line chart 
    trnmodal_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    trnmodal_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    trnmodal_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    trnmodal_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    trnmodal_chart1.set_y_axis({
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
        
    trnmodal_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    trnmodal_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(trn_ag1_rows):
        trnmodal_chart1.add_series({
            'name':       [economy + '_trn_use', chart_height + i + 1, 0],
            'categories': [economy + '_trn_use', chart_height, 2, chart_height, trn_ag1_cols - 1],
            'values':     [economy + '_trn_use', chart_height + i + 1, 2, chart_height + i + 1, trn_ag1_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet3.insert_chart('B3', trnmodal_chart1)

    # Create a transport fuel line chart 
    trnmodal_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    trnmodal_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    trnmodal_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    trnmodal_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    trnmodal_chart2.set_y_axis({
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
        
    trnmodal_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    trnmodal_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(trn_ag1_years_rows):
        trnmodal_chart2.add_series({
            'name':       [economy + '_trn_use', chart_height + trn_ag1_rows + i + 4, 0],
            'categories': [economy + '_trn_use', chart_height + trn_ag1_rows + 3, 2, chart_height + trn_ag1_rows + 3, trn_ag1_years_cols - 1],
            'values':     [economy + '_trn_use', chart_height + trn_ag1_rows + i + 4, 2, chart_height + trn_ag1_rows + i + 4, trn_ag1_years_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet3.insert_chart('J3', trnmodal_chart2)

    # Create a transport fuel line chart 
    trnmodal_chart3 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    trnmodal_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    trnmodal_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    trnmodal_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    trnmodal_chart3.set_y_axis({
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
        
    trnmodal_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    trnmodal_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(trn_ag2_rows):
        trnmodal_chart3.add_series({
            'name':       [economy + '_trn_use', chart_height + trn_ag1_rows + trn_ag1_years_rows + i + 7, 0],
            'categories': [economy + '_trn_use', chart_height + trn_ag1_rows + trn_ag1_years_rows + 6, 2, chart_height + trn_ag1_rows + trn_ag1_years_rows + 6, trn_ag2_cols - 1],
            'values':     [economy + '_trn_use', chart_height + trn_ag1_rows + trn_ag1_years_rows + i + 7, 2, chart_height + trn_ag1_rows + trn_ag1_years_rows + i + 7, trn_ag2_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet3.insert_chart('R3', trnmodal_chart3)

    # Create a transport fuel line chart 
    trnmodal_chart4 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    trnmodal_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    trnmodal_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    trnmodal_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    trnmodal_chart4.set_y_axis({
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
        
    trnmodal_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    trnmodal_chart4.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(trn_ag2_years_rows):
        trnmodal_chart4.add_series({
            'name':       [economy + '_trn_use', chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + i + 10, 0],
            'categories': [economy + '_trn_use', chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + 9, 2, chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + 9, trn_ag2_years_cols - 1],
            'values':     [economy + '_trn_use', chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + i + 10, 2, chart_height + trn_ag1_rows + trn_ag1_years_rows + trn_ag2_rows + i + 10, 
                            trn_ag2_years_cols - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    worksheet3.insert_chart('Z3', trnmodal_chart4)

    writer.save()

    