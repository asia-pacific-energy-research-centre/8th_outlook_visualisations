# EGEDA TPES plots for each economy

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
import xlsxwriter
import pandas.io.formats.excel

# Import the recently created data frame that joins OSeMOSYS results to EGEDA historical 

EGEDA_years_reference = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_2018_reference.csv').loc[:,:'2050']
EGEDA_years_netzero = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_2018_netzero.csv').loc[:,:'2050']

# Define unique values for economy, fuels, and items columns

Economy_codes = EGEDA_years_reference.economy.unique()
Fuels = EGEDA_years_reference.fuel_code.unique()
Items = EGEDA_years_reference.item_code_new.unique()

# Colours for charting (to be amended later)

colours = pd.read_excel('./data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

# Define month and year to create folder for saving charts/tables

month_year = pd.to_datetime('today').strftime('%B_%Y')

# Subsets for impending df builds

First_level_fuels = ['1_coal', '2_coal_products', '5_oil_shale_and_oil_sands', '6_crude_oil_and_ngl', '7_petroleum_products',
                     '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass',
                     '16_others', '17_electricity', '18_heat', '19_total', '20_total_renewables', '21_modern_renewables']
Required_fuels = ['1_coal', '2_coal_products', '5_oil_shale_and_oil_sands', '6_crude_oil_and_ngl', '7_petroleum_products',
                  '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass',
                  '16_1_biogas', '16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable',
                  '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_9_other_sources',
                  '16_x_hydrogen', '17_electricity', '18_heat', '19_total', '20_total_renewables', '21_modern_renewables']

# Aggregate's of different fuels for TPES/prod charting
Coal_fuels = ['1_coal', '2_coal_products', '3_peat', '4_peat_products']

Oil_fuels = ['6_crude_oil_and_ngl', '7_petroleum_products', '5_oil_shale_and_oil_sands']

Other_fuels = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources']

Renewables_fuels = ['10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', '16_1_biogas', 
                    '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                    '16_8_other_liquid_biofuels']

tpes_items = ['1_indigenous_production', '2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers',
              '6_stock_change', '7_total_primary_energy_supply']

Prod_items = tpes_items[:1]

Petroleum_fuels = ['7_petroleum_products', '7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha', '7_4_gasoline_type_jet_fuel',
                   '7_5_kerosene_type_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil', '7_8_fuel_oil', '7_9_lpg',
                   '7_10_refinery_gas_not_liquefied', '7_11_ethane', '7_x_other_petroleum_products', '7_12_white_spirit_sbp',
                   '7_13_lubricants', '7_14_bitumen', '7_15_paraffin_waxes', '7_16_petroleum_coke', '7_17_other_products']

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written

# Define column chart years
col_chart_years = ['2000', '2010', '2018', '2020', '2030', '2040', '2050']

TPES_agg_fuels = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']
TPES_agg_trade = ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']
avi_bunker = ['Aviation gasoline', 'Jet fuel']

# Total Primary Energy Supply fuel breakdown for each economy

########### Build TPES dataframes for each economy providing various breakdowns (by fuel, TPES component, etc)  

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # REFERENCE DATAFRAMES
    # First data frame: TPES by fuels (and also fourth and sixth dataframe with slight tweaks)
    ref_tpes_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]
    
    #nrows1 = tpes_df.shape[0]
    #ncols1 = tpes_df.shape[1]

    coal = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '7_total_primary_energy_supply')
    
    oil = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '7_total_primary_energy_supply')
    
    renewables = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '7_total_primary_energy_supply')
    
    others = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '7_total_primary_energy_supply')
    
    ref_tpes_df1 = ref_tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(ref_tpes_df.loc[:, '2000':])].reset_index(drop = True)

    ref_tpes_df1.loc[ref_tpes_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_tpes_df1.loc[ref_tpes_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_tpes_df1 = ref_tpes_df1[ref_tpes_df1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    nrows4 = ref_tpes_df1.shape[0]
    ncols4 = ref_tpes_df1.shape[1]

    ref_tpes_df2 = ref_tpes_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows6 = ref_tpes_df2.shape[0]
    ncols6 = ref_tpes_df2.shape[1]
    
    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    ref_prod_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]
    
    #nrows2 = prod_df.shape[0]
    #ncols2 = prod_df.shape[1]

    coal = ref_prod_df[ref_prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = ref_prod_df[ref_prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = ref_prod_df[ref_prod_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '1_indigenous_production')
    
    others = ref_prod_df[ref_prod_df['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '1_indigenous_production')
    
    ref_prod_df1 = ref_prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(ref_prod_df.loc[:, '2000':])].reset_index(drop = True)

    ref_prod_df1.loc[ref_prod_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_prod_df1.loc[ref_prod_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_prod_df1 = ref_prod_df1[ref_prod_df1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    nrows5 = ref_prod_df1.shape[0]
    ncols5 = ref_prod_df1.shape[1]

    ref_prod_df2 = ref_prod_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows7 = ref_prod_df2.shape[0]
    ncols7 = ref_prod_df2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    ref_tpes_comp_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                           (EGEDA_years_reference['item_code_new'].isin(tpes_items)) &
                           (EGEDA_years_reference['fuel_code'] == '19_total')]
    
    net_trade = ref_tpes_comp_df1[ref_tpes_comp_df1['item_code_new'].isin(['2_imports', 
                                                                     '3_exports'])].groupby(['economy', 
                                                                                             'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                        item_code_new = 'Net trade')
                           
    bunkers = ref_tpes_comp_df1[ref_tpes_comp_df1['item_code_new'].isin(['4_international_marine_bunkers', 
                                                                 '5_international_aviation_bunkers'])].groupby(['economy', 
                                                                                                                  'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                                             item_code_new = 'Bunkers')
    
    ref_tpes_comp_df1 = ref_tpes_comp_df1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    ref_tpes_comp_df1.loc[ref_tpes_comp_df1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_tpes_comp_df1.loc[ref_tpes_comp_df1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock changes'
    
    ref_tpes_comp_df1 = ref_tpes_comp_df1.loc[ref_tpes_comp_df1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True)
    
    nrows3 = ref_tpes_comp_df1.shape[0]
    ncols3 = ref_tpes_comp_df1.shape[1]

    # Imports/exports data frame

    ref_imports_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '2_imports') & 
                              (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]

    coal = ref_imports_df1[ref_imports_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '2_imports')
    
    renewables = ref_imports_df1[ref_imports_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '2_imports')
    
    others = ref_imports_df1[ref_imports_df1['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '2_imports')
    
    ref_imports_df1 = ref_imports_df1.append([coal, oil, renewables, others]).reset_index(drop = True)

    ref_imports_df1.loc[ref_imports_df1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    ref_imports_df1.loc[ref_imports_df1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    ref_imports_df1.loc[ref_imports_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_imports_df1.loc[ref_imports_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_imports_df1 = ref_imports_df1[ref_imports_df1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_imports_df1.loc[:, '2000':])]

    nrows8 = ref_imports_df1.shape[0]
    ncols8 = ref_imports_df1.shape[1] 

    ref_imports_df2 = ref_imports_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows12 = ref_imports_df2.shape[0]
    ncols12 = ref_imports_df2.shape[1]                             

    ref_exports_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '3_exports') & 
                              (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].copy()

    # Change export values to positive rather than negative

    ref_exports_df1[list(ref_exports_df1.columns[3:])] = ref_exports_df1[list(ref_exports_df1.columns[3:])].apply(lambda x: x * -1)

    coal = ref_exports_df1[ref_exports_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '3_exports')
    
    renewables = ref_exports_df1[ref_exports_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '3_exports')
    
    others = ref_exports_df1[ref_exports_df1['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '3_exports')
    
    ref_exports_df1 = ref_exports_df1.append([coal, oil, renewables, others]).reset_index(drop = True)

    ref_exports_df1.loc[ref_exports_df1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    ref_exports_df1.loc[ref_exports_df1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    ref_exports_df1.loc[ref_exports_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_exports_df1.loc[ref_exports_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_exports_df1 = ref_exports_df1[ref_exports_df1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_exports_df1.loc[:, '2000':])]

    nrows9 = ref_exports_df1.shape[0]
    ncols9 = ref_exports_df1.shape[1]

    ref_exports_df2 = ref_exports_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows13 = ref_exports_df2.shape[0]
    ncols13 = ref_exports_df2.shape[1] 

    # Bunkers dataframes

    ref_bunkers_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '4_international_marine_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(['7_7_gas_diesel_oil', '7_8_fuel_oil']))]

    ref_bunkers_df1 = ref_bunkers_df1[['fuel_code', 'item_code_new'] + list(ref_bunkers_df1.loc[:, '2000':])]

    ref_bunkers_df1.loc[ref_bunkers_df1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    ref_bunkers_df1.loc[ref_bunkers_df1['fuel_code'] == '7_8_fuel_oil', 'fuel_code'] = 'Fuel oil'

    nrows10 = ref_bunkers_df1.shape[0]
    ncols10 = ref_bunkers_df1.shape[1]

    ref_bunkers_df2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '5_international_aviation_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel', '7_2_aviation_gasoline']))]

    jetfuel = ref_bunkers_df2[ref_bunkers_df2['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel'])]\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Jet fuel',
                                                 item_code_new = '5_international_aviation_bunkers')
    
    ref_bunkers_df2 = ref_bunkers_df2.append([jetfuel]).reset_index(drop = True)

    ref_bunkers_df2 = ref_bunkers_df2[['fuel_code', 'item_code_new'] + list(ref_bunkers_df2.loc[:, '2000':])]

    ref_bunkers_df2.loc[ref_bunkers_df2['fuel_code'] == '7_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'

    ref_bunkers_df2 = ref_bunkers_df2[ref_bunkers_df2['fuel_code'].isin(avi_bunker)]\
        .set_index('fuel_code').loc[avi_bunker].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_bunkers_df2.loc[:, '2000':])]

    nrows11 = ref_bunkers_df2.shape[0]
    ncols11 = ref_bunkers_df2.shape[1]

    #########################################################################################################################
    # NETZERO DATAFRAMES

    # First data frame: TPES by fuels (and also fourth and sixth dataframe with slight tweaks)
    netz_tpes_df = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '7_total_primary_energy_supply')
    
    oil = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '7_total_primary_energy_supply')
    
    renewables = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '7_total_primary_energy_supply')
    
    others = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '7_total_primary_energy_supply')
    
    netz_tpes_df1 = netz_tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(netz_tpes_df.loc[:, '2000':])].reset_index(drop = True)

    netz_tpes_df1.loc[netz_tpes_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_tpes_df1.loc[netz_tpes_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_tpes_df1 = netz_tpes_df1[netz_tpes_df1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    nrows24 = netz_tpes_df1.shape[0]
    ncols24 = netz_tpes_df1.shape[1]

    netz_tpes_df2 = netz_tpes_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows26 = netz_tpes_df2.shape[0]
    ncols26 = netz_tpes_df2.shape[1]
    
    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    netz_prod_df = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = netz_prod_df[netz_prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = netz_prod_df[netz_prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = netz_prod_df[netz_prod_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '1_indigenous_production')
    
    others = netz_prod_df[netz_prod_df['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '1_indigenous_production')
    
    netz_prod_df1 = netz_prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(netz_prod_df.loc[:, '2000':])].reset_index(drop = True)

    netz_prod_df1.loc[netz_prod_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_prod_df1.loc[netz_prod_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_prod_df1 = netz_prod_df1[netz_prod_df1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    nrows25 = netz_prod_df1.shape[0]
    ncols25 = netz_prod_df1.shape[1]

    netz_prod_df2 = netz_prod_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows27 = netz_prod_df2.shape[0]
    ncols27 = netz_prod_df2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    netz_tpes_comp_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                           (EGEDA_years_netzero['item_code_new'].isin(tpes_items)) &
                           (EGEDA_years_netzero['fuel_code'] == '19_total')]
    
    net_trade = netz_tpes_comp_df1[netz_tpes_comp_df1['item_code_new'].isin(['2_imports', 
                                                                     '3_exports'])].groupby(['economy', 
                                                                                             'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                        item_code_new = 'Net trade')
                           
    bunkers = netz_tpes_comp_df1[netz_tpes_comp_df1['item_code_new'].isin(['4_international_marine_bunkers', 
                                                                 '5_international_aviation_bunkers'])].groupby(['economy', 
                                                                                                                  'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                                             item_code_new = 'Bunkers')
    
    netz_tpes_comp_df1 = netz_tpes_comp_df1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    netz_tpes_comp_df1.loc[netz_tpes_comp_df1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_tpes_comp_df1.loc[netz_tpes_comp_df1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock changes'
    
    netz_tpes_comp_df1 = netz_tpes_comp_df1.loc[netz_tpes_comp_df1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True)
    
    nrows23 = netz_tpes_comp_df1.shape[0]
    ncols23 = netz_tpes_comp_df1.shape[1]

    # Imports/exports data frame

    netz_imports_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '2_imports') & 
                              (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))]

    coal = netz_imports_df1[netz_imports_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '2_imports')
    
    renewables = netz_imports_df1[netz_imports_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '2_imports')
    
    others = netz_imports_df1[netz_imports_df1['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '2_imports')
    
    netz_imports_df1 = netz_imports_df1.append([coal, oil, renewables, others]).reset_index(drop = True)

    netz_imports_df1.loc[netz_imports_df1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    netz_imports_df1.loc[netz_imports_df1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    netz_imports_df1.loc[netz_imports_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_imports_df1.loc[netz_imports_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_imports_df1 = netz_imports_df1[netz_imports_df1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_imports_df1.loc[:, '2000':])]

    nrows28 = netz_imports_df1.shape[0]
    ncols28 = netz_imports_df1.shape[1] 

    netz_imports_df2 = netz_imports_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows32 = netz_imports_df2.shape[0]
    ncols32 = netz_imports_df2.shape[1]                             

    netz_exports_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '3_exports') & 
                              (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].copy()

    # Change export values to positive rather than negative

    netz_exports_df1[list(netz_exports_df1.columns[3:])] = netz_exports_df1[list(netz_exports_df1.columns[3:])].apply(lambda x: x * -1)

    coal = netz_exports_df1[netz_exports_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '3_exports')
    
    renewables = netz_exports_df1[netz_exports_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '3_exports')
    
    others = netz_exports_df1[netz_exports_df1['fuel_code'].isin(Other_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '3_exports')
    
    netz_exports_df1 = netz_exports_df1.append([coal, oil, renewables, others]).reset_index(drop = True)

    netz_exports_df1.loc[netz_exports_df1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    netz_exports_df1.loc[netz_exports_df1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    netz_exports_df1.loc[netz_exports_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_exports_df1.loc[netz_exports_df1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_exports_df1 = netz_exports_df1[netz_exports_df1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_exports_df1.loc[:, '2000':])]

    nrows29 = netz_exports_df1.shape[0]
    ncols29 = netz_exports_df1.shape[1]

    netz_exports_df2 = netz_exports_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows33 = netz_exports_df2.shape[0]
    ncols33 = netz_exports_df2.shape[1] 

    # Bunkers dataframes

    netz_bunkers_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '4_international_marine_bunkers') & 
                              (EGEDA_years_netzero['fuel_code'].isin(['7_7_gas_diesel_oil', '7_8_fuel_oil']))]

    netz_bunkers_df1 = netz_bunkers_df1[['fuel_code', 'item_code_new'] + list(netz_bunkers_df1.loc[:, '2000':])]

    netz_bunkers_df1.loc[netz_bunkers_df1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    netz_bunkers_df1.loc[netz_bunkers_df1['fuel_code'] == '7_8_fuel_oil', 'fuel_code'] = 'Fuel oil'

    nrows30 = netz_bunkers_df1.shape[0]
    ncols30 = netz_bunkers_df1.shape[1]

    netz_bunkers_df2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '5_international_aviation_bunkers') & 
                              (EGEDA_years_netzero['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel', '7_2_aviation_gasoline']))]

    jetfuel = netz_bunkers_df2[netz_bunkers_df2['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel'])]\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Jet fuel',
                                                 item_code_new = '5_international_aviation_bunkers')
    
    netz_bunkers_df2 = netz_bunkers_df2.append([jetfuel]).reset_index(drop = True)

    netz_bunkers_df2 = netz_bunkers_df2[['fuel_code', 'item_code_new'] + list(netz_bunkers_df2.loc[:, '2000':])]

    netz_bunkers_df2.loc[netz_bunkers_df2['fuel_code'] == '7_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'

    netz_bunkers_df2 = netz_bunkers_df2[netz_bunkers_df2['fuel_code'].isin(avi_bunker)]\
        .set_index('fuel_code').loc[avi_bunker].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_bunkers_df2.loc[:, '2000':])]

    nrows31 = netz_bunkers_df2.shape[0]
    ncols31 = netz_bunkers_df2.shape[1]

    # Define directory
    script_dir = './results/' + month_year + '/TPES/'
    results_dir = os.path.join(script_dir, 'economy_breakdown/', economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
    
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_tpes.xlsx', engine = 'xlsxwriter')
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    
    # REFERENCE and NETZERO
    ref_tpes_df1.to_excel(writer, sheet_name = economy + '_TPES_ref', index = False, startrow = chart_height)
    netz_tpes_df1.to_excel(writer, sheet_name = economy + '_TPES_netz', index = False, startrow = chart_height)
    ref_tpes_df2.to_excel(writer, sheet_name = economy + '_TPES_ref', index = False, startrow = chart_height + nrows4 + 3)
    netz_tpes_df2.to_excel(writer, sheet_name = economy + '_TPES_netz', index = False, startrow = chart_height + nrows24 + 3)
    ref_prod_df1.to_excel(writer, sheet_name = economy + '_prod_ref', index = False, startrow = chart_height)
    netz_prod_df1.to_excel(writer, sheet_name = economy + '_prod_netz', index = False, startrow = chart_height)
    ref_prod_df2.to_excel(writer, sheet_name = economy + '_prod_ref', index = False, startrow = chart_height + nrows5 + 3)
    netz_prod_df2.to_excel(writer, sheet_name = economy + '_prod_netz', index = False, startrow = chart_height + nrows25 + 3)
    ref_tpes_comp_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_I_ref', index = False, startrow = chart_height)
    netz_tpes_comp_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_I_netz', index = False, startrow = chart_height)
    ref_imports_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_I_ref', index = False, startrow = chart_height + nrows3 + 3)
    netz_imports_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_I_netz', index = False, startrow = chart_height + nrows23 + 3)
    ref_imports_df2.to_excel(writer, sheet_name = economy + '_TPES_comp_I_ref', index = False, startrow = chart_height + nrows3 + nrows8 + 6)
    netz_imports_df2.to_excel(writer, sheet_name = economy + '_TPES_comp_I_netz', index = False, startrow = chart_height + nrows23 + nrows28 + 6)
    ref_exports_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_I_ref', index = False, startrow = chart_height + nrows3 + nrows8 + nrows12 + 9)
    netz_exports_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_I_netz', index = False, startrow = chart_height + nrows23 + nrows28 + nrows32 + 9)
    ref_exports_df2.to_excel(writer, sheet_name = economy + '_TPES_comp_I_ref', index = False, startrow = chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12)
    netz_exports_df2.to_excel(writer, sheet_name = economy + '_TPES_comp_I_netz', index = False, startrow = chart_height + nrows23 + nrows28 + nrows32 + nrows29 + 12)
    ref_bunkers_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_II_ref', index = False, startrow = chart_height)
    netz_bunkers_df1.to_excel(writer, sheet_name = economy + '_TPES_comp_II_netz', index = False, startrow = chart_height)
    ref_bunkers_df2.to_excel(writer, sheet_name = economy + '_TPES_comp_II_ref', index = False, startrow = chart_height + nrows10 + 3)
    netz_bunkers_df2.to_excel(writer, sheet_name = economy + '_TPES_comp_II_netz', index = False, startrow = chart_height + nrows30 + 3)
        
    # Access the workbook
    workbook = writer.book
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
    
    ################################################################### CHARTS ###################################################################
    # REFERENCE
    # Access the sheet created using writer above
    ref_worksheet1 = writer.sheets[economy + '_TPES_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet1.set_column(2, ncols4 + 1, None, comma_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height + nrows4 + 3, None, header_format)
    ref_worksheet1.write(0, 0, economy + ' TPES fuel reference', cell_format1)

    ######################################################
    # Create a TPES chart
    ref_tpes_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_tpes_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_tpes_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_tpes_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_chart2.set_y_axis({
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
        
    ref_tpes_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_tpes_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows4):
        ref_tpes_chart2.add_series({
            'name':       [economy + '_TPES_ref', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_ref', chart_height, 2, chart_height, ncols4 - 1],
            'values':     [economy + '_TPES_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols4 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet1.insert_chart('B3', ref_tpes_chart2)

    ######## same chart as above but line

    # TPES line chart
    ref_tpes_chart4 = workbook.add_chart({'type': 'line'})
    ref_tpes_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_tpes_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    ref_tpes_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_chart4.set_y_axis({
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
        
    ref_tpes_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_tpes_chart4.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows4):
        ref_tpes_chart4.add_series({
            'name':       [economy + '_TPES_ref', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_ref', chart_height, 2, chart_height, ncols4 - 1],
            'values':     [economy + '_TPES_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols4 - 1],
            'line':       {'color': colours_hex[i], 
                           'width': 1}
        })    
        
    ref_worksheet1.insert_chart('R3', ref_tpes_chart4)

    ###################### Create another TPES chart showing proportional share #################################

    # Create a TPES chart
    ref_tpes_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_tpes_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_tpes_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_tpes_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_chart3.set_y_axis({
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
        
    ref_tpes_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_tpes_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in TPES_agg_fuels:
        i = ref_tpes_df2[ref_tpes_df2['fuel_code'] == component].index[0]
        ref_tpes_chart3.add_series({
            'name':       [economy + '_TPES_ref', chart_height + nrows4 + i + 4, 0],
            'categories': [economy + '_TPES_ref', chart_height + nrows4 + 3, 2, chart_height + nrows4 + 3, ncols6 - 1],
            'values':     [economy + '_TPES_ref', chart_height + nrows4 + i + 4, 2, chart_height + nrows4 + i + 4, ncols6 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet1.insert_chart('J3', ref_tpes_chart3)

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    ref_worksheet2 = writer.sheets[economy + '_prod_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet2.set_column(2, ncols5 + 1, None, comma_format)
    ref_worksheet2.set_row(chart_height, None, header_format)
    ref_worksheet2.set_row(chart_height + nrows5 + 3, None, header_format)
    ref_worksheet2.write(0, 0, economy + ' prod fuel reference', cell_format1)

    ###################### Create another PRODUCTION chart with only 6 categories #################################

    # Create a PROD chart
    ref_prod_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_prod_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_prod_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_prod_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_prod_chart2.set_y_axis({
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
        
    ref_prod_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_prod_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows5):
        ref_prod_chart2.add_series({
            'name':       [economy + '_prod_ref', chart_height + i + 1, 0],
            'categories': [economy + '_prod_ref', chart_height, 2, chart_height, ncols5 - 1],
            'values':     [economy + '_prod_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols5 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet2.insert_chart('B3', ref_prod_chart2)

    ############ Same as above but with a line ###########

    # Create a PROD chart
    ref_prod_chart2 = workbook.add_chart({'type': 'line'})
    ref_prod_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_prod_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_prod_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_prod_chart2.set_y_axis({
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
        
    ref_prod_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_prod_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows5):
        ref_prod_chart2.add_series({
            'name':       [economy + '_prod_ref', chart_height + i + 1, 0],
            'categories': [economy + '_prod_ref', chart_height, 2, chart_height, ncols5 - 1],
            'values':     [economy + '_prod_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols5 - 1],
            'line':       {'color': colours_hex[i],
                           'width': 1} 
        })    
        
    ref_worksheet2.insert_chart('R3', ref_prod_chart2)

    ###################### Create another PRODUCTION chart showing proportional share #################################

    # Create a production chart
    ref_prod_chart3 = workbook.add_chart({'type': 'column', 
                                      'subtype': 'percent_stacked'})
    ref_prod_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_prod_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_prod_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_prod_chart3.set_y_axis({
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
        
    ref_prod_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_prod_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in TPES_agg_fuels:
        i = ref_prod_df2[ref_prod_df2['fuel_code'] == component].index[0]
        ref_prod_chart3.add_series({
            'name':       [economy + '_prod_ref', chart_height + nrows5 + i + 4, 0],
            'categories': [economy + '_prod_ref', chart_height + nrows5 + 3, 2, chart_height + nrows5 + 3, ncols7 - 1],
            'values':     [economy + '_prod_ref', chart_height + nrows5 + i + 4, 2, chart_height + nrows5 + i + 4, ncols7 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet2.insert_chart('J3', ref_prod_chart3)
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    ref_worksheet3 = writer.sheets[economy + '_TPES_comp_I_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet3.set_column(2, ncols8 + 1, None, comma_format)
    ref_worksheet3.set_row(chart_height, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows3 + 3, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows3 + nrows8 + 6, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows3 + nrows8 + nrows12 + 9, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12, None, header_format)
    ref_worksheet3.write(0, 0, economy + ' TPES components I reference', cell_format1)
    
    # Create a TPES components chart
    ref_tpes_comp_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    ref_tpes_comp_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_tpes_comp_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_tpes_comp_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    ref_tpes_comp_chart1.set_y_axis({
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
        
    ref_tpes_comp_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_tpes_comp_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Net trade', 'Bunkers', 'Stock changes']:
        i = ref_tpes_comp_df1[ref_tpes_comp_df1['item_code_new'] == component].index[0]
        ref_tpes_comp_chart1.add_series({
            'name':       [economy + '_TPES_comp_I_ref', chart_height + i + 1, 1],
            'categories': [economy + '_TPES_comp_I_ref', chart_height, 2, chart_height, ncols3 - 1],
            'values':     [economy + '_TPES_comp_I_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols3 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('B3', ref_tpes_comp_chart1)

    # IMPORTS: Create a line chart subset by fuel
    
    ref_imports_line = workbook.add_chart({'type': 'line'})
    ref_imports_line.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_imports_line.set_chartarea({
        'border': {'none': True}
    })
    
    ref_imports_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_imports_line.set_y_axis({
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
        
    ref_imports_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_imports_line.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for fuel in ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']:
        i = ref_imports_df1[ref_imports_df1['fuel_code'] == fuel].index[0]
        ref_imports_line.add_series({
            'name':       [economy + '_TPES_comp_I_ref', chart_height + nrows3 + i + 4, 0],
            'categories': [economy + '_TPES_comp_I_ref', chart_height + nrows3 + 3, 2, chart_height + nrows3 + 3, ncols8 - 1],
            'values':     [economy + '_TPES_comp_I_ref', chart_height + nrows3 + i + 4, 2, chart_height + nrows3 + i + 4, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    ref_worksheet3.insert_chart('J3', ref_imports_line)

    # Create a imports by fuel column
    ref_imports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    ref_imports_column.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_imports_column.set_chartarea({
        'border': {'none': True}
    })
    
    ref_imports_column.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    ref_imports_column.set_y_axis({
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
        
    ref_imports_column.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_imports_column.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows12):
        ref_imports_column.add_series({
            'name':       [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + i + 7, 0],
            'categories': [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + 6, 2, chart_height + nrows3 + nrows8 + 6, ncols12 - 1],
            'values':     [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + i + 7, 2, chart_height + nrows3 + nrows8 + i + 7, ncols12 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('R3', ref_imports_column)

    # EXPORTS: Create a line chart subset by fuel
    
    ref_exports_line = workbook.add_chart({'type': 'line'})
    ref_exports_line.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_exports_line.set_chartarea({
        'border': {'none': True}
    })
    
    ref_exports_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_exports_line.set_y_axis({
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
        
    ref_exports_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_exports_line.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for fuel in ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']:
        i = ref_exports_df1[ref_exports_df1['fuel_code'] == fuel].index[0]
        ref_exports_line.add_series({
            'name':       [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + nrows12 + i + 10, 0],
            'categories': [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + nrows12 + 9, 2, chart_height + nrows3 + nrows8 + nrows12 + 9, ncols8 - 1],
            'values':     [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + nrows12 + i + 10, 2, chart_height + nrows3 + nrows8 + nrows12 + i + 10, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    ref_worksheet3.insert_chart('Z3', ref_exports_line)

    # Create a imports by fuel column
    ref_exports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    ref_exports_column.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_exports_column.set_chartarea({
        'border': {'none': True}
    })
    
    ref_exports_column.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    ref_exports_column.set_y_axis({
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
        
    ref_exports_column.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_exports_column.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows13):
        ref_exports_column.add_series({
            'name':       [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + nrows12 + nrows9 + i + 13, 0],
            'categories': [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12, 2, chart_height + nrows3 + nrows8 + nrows12 + nrows9 + 12, ncols13 - 1],
            'values':     [economy + '_TPES_comp_I_ref', chart_height + nrows3 + nrows8 + nrows12 + nrows9 + i + 13, 2, chart_height + nrows3 + nrows8 + nrows12 + nrows9 + i + 13, ncols13 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('AH3', ref_exports_column)

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    ref_worksheet4 = writer.sheets[economy + '_TPES_comp_II_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet4.set_column(2, ncols10 + 1, None, comma_format)
    ref_worksheet4.set_row(chart_height, None, header_format)
    ref_worksheet4.set_row(chart_height + nrows10 + 3, None, header_format)
    ref_worksheet4.write(0, 0, economy + ' TPES components II reference', cell_format1)
    
    # MARINE BUNKER: Create a line chart subset by fuel
    
    ref_marine_line = workbook.add_chart({'type': 'line'})
    ref_marine_line.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_marine_line.set_chartarea({
        'border': {'none': True}
    })
    
    ref_marine_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_marine_line.set_y_axis({
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
        
    ref_marine_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_marine_line.set_title({
        'none': True
    }) 

    # Configure the series of the chart from the dataframe data.
    for i in range(nrows10):
        ref_marine_line.add_series({
            'name':       [economy + '_TPES_comp_II_ref', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_comp_II_ref', chart_height, 2, chart_height, ncols10 - 1],
            'values':     [economy + '_TPES_comp_II_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols10 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    ref_worksheet4.insert_chart('B3', ref_marine_line)

    # AVIATION BUNKER: Create a line chart subset by fuel
    
    ref_aviation_line = workbook.add_chart({'type': 'line'})
    ref_aviation_line.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_aviation_line.set_chartarea({
        'border': {'none': True}
    })
    
    ref_aviation_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_aviation_line.set_y_axis({
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
        
    ref_aviation_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_aviation_line.set_title({
        'none': True
    }) 

    # Configure the series of the chart from the dataframe data.
    for i in range(nrows11):
        ref_aviation_line.add_series({
            'name':       [economy + '_TPES_comp_II_ref', chart_height + nrows10 + i + 4, 0],
            'categories': [economy + '_TPES_comp_II_ref', chart_height + nrows10 + 3, 2, chart_height + nrows10 + 3, ncols11 - 1],
            'values':     [economy + '_TPES_comp_II_ref', chart_height + nrows10 + i + 4, 2, chart_height + nrows10 + i + 4, ncols11 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    ref_worksheet4.insert_chart('J3', ref_aviation_line)

    #####################################################################################################################################
    # NET ZERO CHARTS

    # Access the sheet created using writer above
    netz_worksheet1 = writer.sheets[economy + '_TPES_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet1.set_column(2, ncols24 + 1, None, comma_format)
    netz_worksheet1.set_row(chart_height, None, header_format)
    netz_worksheet1.set_row(chart_height + nrows24 + 3, None, header_format)
    netz_worksheet1.write(0, 0, economy + ' TPES fuel net-zero', cell_format1)

    ######################################################
    # Create a TPES chart
    netz_tpes_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_tpes_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_chart2.set_y_axis({
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
        
    netz_tpes_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows24):
        netz_tpes_chart2.add_series({
            'name':       [economy + '_TPES_netz', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_netz', chart_height, 2, chart_height, ncols24 - 1],
            'values':     [economy + '_TPES_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols24 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet1.insert_chart('B3', netz_tpes_chart2)

    ######## same chart as above but line

    # TPES line chart
    netz_tpes_chart4 = workbook.add_chart({'type': 'line'})
    netz_tpes_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_chart4.set_y_axis({
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
        
    netz_tpes_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_chart4.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows24):
        netz_tpes_chart4.add_series({
            'name':       [economy + '_TPES_netz', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_netz', chart_height, 2, chart_height, ncols24 - 1],
            'values':     [economy + '_TPES_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols24 - 1],
            'line':       {'color': colours_hex[i], 
                           'width': 1}
        })    
        
    netz_worksheet1.insert_chart('R3', netz_tpes_chart4)

    ###################### Create another TPES chart showing proportional share #################################

    # Create a TPES chart
    netz_tpes_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_tpes_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_chart3.set_y_axis({
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
        
    netz_tpes_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in TPES_agg_fuels:
        i = netz_tpes_df2[netz_tpes_df2['fuel_code'] == component].index[0]
        netz_tpes_chart3.add_series({
            'name':       [economy + '_TPES_netz', chart_height + nrows24 + i + 4, 0],
            'categories': [economy + '_TPES_netz', chart_height + nrows24 + 3, 2, chart_height + nrows24 + 3, ncols26 - 1],
            'values':     [economy + '_TPES_netz', chart_height + nrows24 + i + 4, 2, chart_height + nrows24 + i + 4, ncols26 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet1.insert_chart('J3', netz_tpes_chart3)

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    netz_worksheet2 = writer.sheets[economy + '_prod_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet2.set_column(2, ncols25 + 1, None, comma_format)
    netz_worksheet2.set_row(chart_height, None, header_format)
    netz_worksheet2.set_row(chart_height + nrows25 + 3, None, header_format)
    netz_worksheet2.write(0, 0, economy + ' prod fuel net-zero', cell_format1)

    ###################### Create another PRODUCTION chart with only 6 categories #################################

    # Create a PROD chart
    netz_prod_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_prod_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_prod_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_prod_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_prod_chart2.set_y_axis({
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
        
    netz_prod_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_prod_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows25):
        netz_prod_chart2.add_series({
            'name':       [economy + '_prod_netz', chart_height + i + 1, 0],
            'categories': [economy + '_prod_netz', chart_height, 2, chart_height, ncols25 - 1],
            'values':     [economy + '_prod_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols25 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet2.insert_chart('B3', netz_prod_chart2)

    ############ Same as above but with a line ###########

    # Create a PROD chart
    netz_prod_chart2 = workbook.add_chart({'type': 'line'})
    netz_prod_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_prod_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_prod_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_prod_chart2.set_y_axis({
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
        
    netz_prod_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_prod_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows25):
        netz_prod_chart2.add_series({
            'name':       [economy + '_prod_netz', chart_height + i + 1, 0],
            'categories': [economy + '_prod_netz', chart_height, 2, chart_height, ncols25 - 1],
            'values':     [economy + '_prod_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols25 - 1],
            'line':       {'color': colours_hex[i],
                           'width': 1} 
        })    
        
    netz_worksheet2.insert_chart('R3', netz_prod_chart2)

    ###################### Create another PRODUCTION chart showing proportional share #################################

    # Create a production chart
    netz_prod_chart3 = workbook.add_chart({'type': 'column', 
                                      'subtype': 'percent_stacked'})
    netz_prod_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_prod_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_prod_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_prod_chart3.set_y_axis({
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
        
    netz_prod_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_prod_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in TPES_agg_fuels:
        i = netz_prod_df2[netz_prod_df2['fuel_code'] == component].index[0]
        netz_prod_chart3.add_series({
            'name':       [economy + '_prod_netz', chart_height + nrows25 + i + 4, 0],
            'categories': [economy + '_prod_netz', chart_height + nrows25 + 3, 2, chart_height + nrows25 + 3, ncols27 - 1],
            'values':     [economy + '_prod_netz', chart_height + nrows25 + i + 4, 2, chart_height + nrows25 + i + 4, ncols27 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet2.insert_chart('J3', netz_prod_chart3)
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    netz_worksheet3 = writer.sheets[economy + '_TPES_comp_I_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet3.set_column(2, ncols28 + 1, None, comma_format)
    netz_worksheet3.set_row(chart_height, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows23 + 3, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows23 + nrows28 + 6, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows23 + nrows28 + nrows32 + 9, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows23 + nrows28 + nrows32 + nrows29 + 12, None, header_format)
    netz_worksheet3.write(0, 0, economy + ' TPES components I net-zero', cell_format1)
    
    # Create a TPES components chart
    netz_tpes_comp_chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_tpes_comp_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_tpes_comp_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_tpes_comp_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_tpes_comp_chart1.set_y_axis({
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
        
    netz_tpes_comp_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_tpes_comp_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in ['Production', 'Net trade', 'Bunkers', 'Stock changes']:
        i = netz_tpes_comp_df1[netz_tpes_comp_df1['item_code_new'] == component].index[0]
        netz_tpes_comp_chart1.add_series({
            'name':       [economy + '_TPES_comp_I_netz', chart_height + i + 1, 1],
            'categories': [economy + '_TPES_comp_I_netz', chart_height, 2, chart_height, ncols23 - 1],
            'values':     [economy + '_TPES_comp_I_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols23 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    netz_worksheet3.insert_chart('B3', netz_tpes_comp_chart1)

    # IMPORTS: Create a line chart subset by fuel
    
    netz_imports_line = workbook.add_chart({'type': 'line'})
    netz_imports_line.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_imports_line.set_chartarea({
        'border': {'none': True}
    })
    
    netz_imports_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_imports_line.set_y_axis({
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
        
    netz_imports_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_imports_line.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for fuel in ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']:
        i = netz_imports_df1[netz_imports_df1['fuel_code'] == fuel].index[0]
        netz_imports_line.add_series({
            'name':       [economy + '_TPES_comp_I_netz', chart_height + nrows23 + i + 4, 0],
            'categories': [economy + '_TPES_comp_I_netz', chart_height + nrows23 + 3, 2, chart_height + nrows23 + 3, ncols28 - 1],
            'values':     [economy + '_TPES_comp_I_netz', chart_height + nrows23 + i + 4, 2, chart_height + nrows23 + i + 4, ncols28 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    netz_worksheet3.insert_chart('J3', netz_imports_line)

    # Create a imports by fuel column
    netz_imports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_imports_column.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_imports_column.set_chartarea({
        'border': {'none': True}
    })
    
    netz_imports_column.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_imports_column.set_y_axis({
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
        
    netz_imports_column.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_imports_column.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows32):
        netz_imports_column.add_series({
            'name':       [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + i + 7, 0],
            'categories': [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + 6, 2, chart_height + nrows23 + nrows28 + 6, ncols32 - 1],
            'values':     [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + i + 7, 2, chart_height + nrows23 + nrows28 + i + 7, ncols32 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    netz_worksheet3.insert_chart('R3', netz_imports_column)

    # EXPORTS: Create a line chart subset by fuel
    
    netz_exports_line = workbook.add_chart({'type': 'line'})
    netz_exports_line.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_exports_line.set_chartarea({
        'border': {'none': True}
    })
    
    netz_exports_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_exports_line.set_y_axis({
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
        
    netz_exports_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_exports_line.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for fuel in ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']:
        i = netz_exports_df1[netz_exports_df1['fuel_code'] == fuel].index[0]
        netz_exports_line.add_series({
            'name':       [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + nrows32 + i + 10, 0],
            'categories': [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + nrows32 + 9, 2, chart_height + nrows23 + nrows28 + nrows32 + 9, ncols28 - 1],
            'values':     [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + nrows32 + i + 10, 2, chart_height + nrows23 + nrows28 + nrows32 + i + 10, ncols28 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    netz_worksheet3.insert_chart('Z3', netz_exports_line)

    # Create a imports by fuel column
    netz_exports_column = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    netz_exports_column.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_exports_column.set_chartarea({
        'border': {'none': True}
    })
    
    netz_exports_column.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'line': {'color': '#bebebe'}
    })
        
    netz_exports_column.set_y_axis({
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
        
    netz_exports_column.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_exports_column.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows33):
        netz_exports_column.add_series({
            'name':       [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + nrows32 + nrows29 + i + 13, 0],
            'categories': [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + nrows32 + nrows29 + 12, 2, chart_height + nrows23 + nrows28 + nrows32 + nrows29 + 12, ncols33 - 1],
            'values':     [economy + '_TPES_comp_I_netz', chart_height + nrows23 + nrows28 + nrows32 + nrows29 + i + 13, 2, chart_height + nrows23 + nrows28 + nrows32 + nrows29 + i + 13, ncols33 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    netz_worksheet3.insert_chart('AH3', netz_exports_column)

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    netz_worksheet4 = writer.sheets[economy + '_TPES_comp_II_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet4.set_column(2, ncols30 + 1, None, comma_format)
    netz_worksheet4.set_row(chart_height, None, header_format)
    netz_worksheet4.set_row(chart_height + nrows30 + 3, None, header_format)
    netz_worksheet4.write(0, 0, economy + ' TPES components II net-zero', cell_format1)
    
    # MARINE BUNKER: Create a line chart subset by fuel
    
    netz_marine_line = workbook.add_chart({'type': 'line'})
    netz_marine_line.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_marine_line.set_chartarea({
        'border': {'none': True}
    })
    
    netz_marine_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_marine_line.set_y_axis({
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
        
    netz_marine_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_marine_line.set_title({
        'none': True
    }) 

    # Configure the series of the chart from the dataframe data.
    for i in range(nrows30):
        netz_marine_line.add_series({
            'name':       [economy + '_TPES_comp_II_netz', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_comp_II_netz', chart_height, 2, chart_height, ncols30 - 1],
            'values':     [economy + '_TPES_comp_II_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols30 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    netz_worksheet4.insert_chart('B3', netz_marine_line)

    # AVIATION BUNKER: Create a line chart subset by fuel
    
    netz_aviation_line = workbook.add_chart({'type': 'line'})
    netz_aviation_line.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_aviation_line.set_chartarea({
        'border': {'none': True}
    })
    
    netz_aviation_line.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_aviation_line.set_y_axis({
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
        
    netz_aviation_line.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_aviation_line.set_title({
        'none': True
    }) 

    # Configure the series of the chart from the dataframe data.
    for i in range(nrows31):
        netz_aviation_line.add_series({
            'name':       [economy + '_TPES_comp_II_netz', chart_height + nrows30 + i + 4, 0],
            'categories': [economy + '_TPES_comp_II_netz', chart_height + nrows30 + 3, 2, chart_height + nrows30 + 3, ncols31 - 1],
            'values':     [economy + '_TPES_comp_II_netz', chart_height + nrows30 + i + 4, 2, chart_height + nrows30 + i + 4, ncols31 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25},
        })    
        
    netz_worksheet4.insert_chart('J3', netz_aviation_line)
    
    writer.save()

print('Bling blang blaow, you have some TPES charts now')


