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
# 2 Dataframes: REFERENCE and NET ZERO 

EGEDA_years_reference = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_2018_reference.csv').loc[:,:'2050']
EGEDA_years_netzero = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_2018_netzero.csv').loc[:,:'2050']

# Define unique values for economy, fuels, and items columns
# only looking at one dataframe which should be sufficient as both have same structure

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

Coal_fuels = ['1_coal', '2_coal_products', '3_peat', '4_peat_products']

Oil_fuels = ['6_crude_oil_and_ngl', '7_petroleum_products', '5_oil_shale_and_oil_sands']

Others_fuels = ['9_nuclear', '16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable']

# Need to amend this to reflect demarcation between modern renewables and traditional biomass renewables 

Renewables_fuels = ['10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', '16_1_biogas', 
                    '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                    '16_8_other_liquid_biofuels']

Renewables_fuels_nobiomass = ['10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '16_1_biogas', 
                          '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                          '16_8_other_liquid_biofuels']

trad_bio_sectors = ['16_1_commercial_and_public_services', '16_2_residential',
                  '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others']

no_trad_bio_sectors = ['14_industry_sector', '15_transport_sector', '17_nonenergy_use']

# Modern_renew_primary = to be completed

# Modern_renew_FED = to be completed

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
                   '7_x_other_petroleum_products', '8_1_natural_gas', '16_5_biogasoline', '16_6_biodiesel',
                   '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_9_other_sources', '17_electricity'] 

Transport_fuels_agg = ['Diesel', 'Gasoline', 'LPG', 'Gas', 'Jet fuel', 'Electricity', 'Renewables', 'Hydrogen', 'Other']

Renew_fuel = ['16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels']

Other_fuel = ['7_8_fuel_oil', '1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products', '7_x_other_petroleum_products']

Other_industry = ['14_5_transportation_equipment', '14_6_machinery', '14_8_food_beverages_and_tobacco', '14_10_wood_and_wood_products',
                  '14_11_construction', '14_12_textiles_and_leather']

Transport_modal = ['15_1_domestic_air_transport', '15_2_road', '15_3_rail', '15_4_domestic_navigation', '15_5_pipeline_transport',
                   '15_6_nonspecified_transport']

Transport_modal_agg = ['Aviation', 'Road', 'Rail' ,'Marine', 'Pipeline', 'Non-specified']

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written (can change this)

# Define column chart years
col_chart_years = ['2000', '2010', '2018', '2020', '2030', '2040', '2050']

# Define column chart years for transport
col_chart_years_transport = ['2018', '2020', '2030', '2040', '2050']

# FED aggregate fuels

FED_agg_fuels = ['Coal', 'Oil', 'Gas', 'Modern renewables', 'Traditional biomass', 'Hydrogen', 'Electricity', 'Heat', 'Others']
FED_agg_fuels_ind = ['Coal', 'Oil', 'Gas', 'Renewables', 'Hydrogen', 'Electricity', 'Heat', 'Others']

FED_agg_sectors = ['Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-energy', 'Non-specified']

Industry_eight = ['Iron & steel', 'Chemicals', 'Aluminium', 'Non-metallic minerals', 'Mining', 'Pulp & paper', 'Other', 'Non-specified']

# Final energy demand by fuel and sector for each economy

# This is TFC which includes non-energy

############# Build FED (TFC) dataframes for each economy (TFC) and then build subsequent charts ###########

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # REFERENCE DATA FRAMES
    # First data frame construction: FED by fuels
    ref_econ_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(no_trad_bio_sectors)) &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    ref_econ_df1 = ref_econ_df1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Industry, transport, NE').reset_index()

    #nrows1 = econ_df1.shape[0]
    #ncols1 = econ_df1.shape[1]

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Industry, transport, NE')
    
    oil = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Industry, transport, NE')
    
    renewables = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Modern renewables', item_code_new = 'Industry, transport, NE')
    
    others = ref_econ_df1[ref_econ_df1['fuel_code'].isin(Others_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Industry, transport, NE')

    # Fed fuel data frame 1 (data frame 6)

    ref_fedfuel_df1 = ref_econ_df1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_econ_df1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_fedfuel_df1.loc[ref_fedfuel_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_fedfuel_df1.loc[ref_fedfuel_df1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_fedfuel_df1.loc[ref_fedfuel_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_fedfuel_df1.loc[ref_fedfuel_df1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    # Insert 0 traditional biomass row
    new_row = ['Traditional biomass', 'Industry, transport, NE'] + [0] * 51
    new_series = pd.Series(new_row, index = ref_fedfuel_df1.columns)

    ref_fedfuel_df1 = ref_fedfuel_df1.append(new_series, ignore_index = True)

    ref_fedfuel_df1 = ref_fedfuel_df1[ref_fedfuel_df1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ##### No biomass fix for dataframe

    ref_tradbio_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(trad_bio_sectors)) &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)

    ref_tradbio_df1 = ref_tradbio_df1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Trad bio sectors').reset_index()

    # build aggregate with altered vector to account for no biomass in renewables
    coal_tradbio = ref_tradbio_df1[ref_tradbio_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Coal', item_code_new = 'Trad bio sectors')

    oil_tradbio = ref_tradbio_df1[ref_tradbio_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Oil', item_code_new = 'Trad bio sectors')

    renew_tradbio = ref_tradbio_df1[ref_tradbio_df1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Modern renewables', item_code_new = 'Trad bio sectors')

    others_tradbio = ref_tradbio_df1[ref_tradbio_df1['fuel_code'].isin(Others_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Others', item_code_new = 'Trad bio sectors')

    # Fed fuel no biomass in other sector renewables
    ref_tradbio_df2 = ref_tradbio_df1.append([coal_tradbio, oil_tradbio, renew_tradbio, others_tradbio])\
        [['fuel_code', 'item_code_new'] + list(ref_tradbio_df1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_tradbio_df2.loc[ref_tradbio_df2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_tradbio_df2.loc[ref_tradbio_df2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Traditional biomass'
    ref_tradbio_df2.loc[ref_tradbio_df2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_tradbio_df2.loc[ref_tradbio_df2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_tradbio_df2.loc[ref_tradbio_df2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_tradbio_df2 = ref_tradbio_df2[ref_tradbio_df2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ref_fedfuel_df1 = ref_fedfuel_df1.append(ref_tradbio_df2)

    # Combine the two dataframes that account for Modern renewables
    ref_fedfuel_df1 = ref_fedfuel_df1.copy().groupby(['fuel_code']).sum().assign(item_code_new = '12_total_final_consumption')\
        .reset_index()[['fuel_code', 'item_code_new'] + list(ref_fedfuel_df1.loc[:,'2000':'2050'])]\
            .set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    nrows6 = ref_fedfuel_df1.shape[0]
    ncols6 = ref_fedfuel_df1.shape[1]

    ref_fedfuel_df2 = ref_fedfuel_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows7 = ref_fedfuel_df2.shape[0]
    ncols7 = ref_fedfuel_df2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    ref_econ_df2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                        (EGEDA_years_reference['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years_reference['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    ref_econ_df2 = ref_econ_df2[['fuel_code', 'item_code_new'] + list(ref_econ_df2.loc[:,'2000':])]
    
    nrows2 = ref_econ_df2.shape[0]
    ncols2 = ref_econ_df2.shape[1]

    # Now build aggregate sector variables
    
    buildings = ref_econ_df2[ref_econ_df2['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = ref_econ_df2[ref_econ_df2['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')
    
    # Build aggregate data frame of FED sector

    ref_fedsector_df1 = ref_econ_df2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(ref_econ_df2.loc[:, '2000':])].reset_index(drop = True)

    ref_fedsector_df1.loc[ref_fedsector_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_fedsector_df1.loc[ref_fedsector_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_fedsector_df1.loc[ref_fedsector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_fedsector_df1.loc[ref_fedsector_df1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_fedsector_df1 = ref_fedsector_df1[ref_fedsector_df1['item_code_new'].isin(FED_agg_sectors)].set_index('item_code_new').loc[FED_agg_sectors].reset_index()
    ref_fedsector_df1 = ref_fedsector_df1[['fuel_code', 'item_code_new'] + list(ref_fedsector_df1.loc[:, '2000':])]

    nrows8 = ref_fedsector_df1.shape[0]
    ncols8 = ref_fedsector_df1.shape[1]

    ref_fedsector_df2 = ref_fedsector_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows9 = ref_fedsector_df2.shape[0]
    ncols9 = ref_fedsector_df2.shape[1]
    
    # Third data frame construction: Buildings FED by fuel
    ref_bld_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                         (EGEDA_years_reference['item_code_new'].isin(Buildings_items)) &
                         (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]
    
    for fuel in Required_fuels:
        buildings = ref_bld_df1[ref_bld_df1['fuel_code'] == fuel].groupby(['economy', 'fuel_code']).sum().assign(item_code_new = '16_x_buildings')
        buildings['economy'] = economy
        buildings['fuel_code'] = fuel
        
        ref_bld_df1 = ref_bld_df1.append(buildings).reset_index(drop = True)
        
    ref_bld_df1 = ref_bld_df1[['fuel_code', 'item_code_new'] + list(ref_bld_df1.loc[:, '2000':])]
    
    nrows3 = ref_bld_df1.shape[0]
    ncols3 = ref_bld_df1.shape[1]

    ref_bld_df2 = ref_bld_df1[ref_bld_df1['item_code_new'] == '16_x_buildings']

    coal = ref_bld_df2[ref_bld_df2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = '16_x_buildings')
    
    oil = ref_bld_df2[ref_bld_df2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = '16_x_buildings')
    
    renewables = ref_bld_df2[ref_bld_df2['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Modern renewables', item_code_new = '16_x_buildings')
    
    others = ref_bld_df2[ref_bld_df2['fuel_code'].isin(Others_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = '16_x_buildings')

    ref_bld_df2 = ref_bld_df2.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_bld_df2.loc[:, '2000':])].reset_index(drop = True)

    ref_bld_df2.loc[ref_bld_df2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_bld_df2.loc[ref_bld_df2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Traditional biomass'
    ref_bld_df2.loc[ref_bld_df2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_bld_df2.loc[ref_bld_df2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_bld_df2.loc[ref_bld_df2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_bld_df2 = ref_bld_df2[ref_bld_df2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code')\
        .loc[FED_agg_fuels].reset_index()

    nrows12 = ref_bld_df2.shape[0]
    ncols12 = ref_bld_df2.shape[1]

    ref_bld_df3 = ref_bld_df1[(ref_bld_df1['fuel_code'] == '19_total') &
                      (ref_bld_df1['item_code_new'].isin(Buildings_items))].copy()

    ref_bld_df3.loc[ref_bld_df3['item_code_new'] == '16_1_commercial_and_public_services', 'item_code_new'] = 'Services' 
    ref_bld_df3.loc[ref_bld_df3['item_code_new'] == '16_2_residential', 'item_code_new'] = 'Residential'

    nrows13 = ref_bld_df3.shape[0]
    ncols13 = ref_bld_df3.shape[1]
    
    # Fourth data frame construction: Industry subsector
    ref_ind_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                         (EGEDA_years_reference['item_code_new'].isin(Subindustry)) &
                         (EGEDA_years_reference['fuel_code'] == '19_total')]

    other_industry = ref_ind_df1[ref_ind_df1['item_code_new'].isin(Other_industry)].groupby(['fuel_code']).sum().assign(item_code_new = 'Other',
                                                                                                                fuel_code = '19_total')

    ref_ind_df1 = ref_ind_df1.append([other_industry])[['fuel_code', 'item_code_new'] + list(ref_ind_df1.loc[:, '2000':])].reset_index(drop = True)

    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_1_iron_and_steel', 'item_code_new'] = 'Iron & steel'
    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_2_chemical_incl_petrochemical', 'item_code_new'] = 'Chemicals'
    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_3_non_ferrous_metals', 'item_code_new'] = 'Aluminium'
    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_4_nonmetallic_mineral_products', 'item_code_new'] = 'Non-metallic minerals'  
    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_7_mining_and_quarrying', 'item_code_new'] = 'Mining'
    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_9_pulp_paper_and_printing', 'item_code_new'] = 'Pulp & paper'
    ref_ind_df1.loc[ref_ind_df1['item_code_new'] == '14_13_nonspecified_industry', 'item_code_new'] = 'Non-specified'
    
    ref_ind_df1 = ref_ind_df1[ref_ind_df1['item_code_new'].isin(Industry_eight)].set_index('item_code_new').loc[Industry_eight].reset_index()

    ref_ind_df1 = ref_ind_df1[['fuel_code', 'item_code_new'] + list(ref_ind_df1.loc[:, '2000':])]

    nrows4 = ref_ind_df1.shape[0]
    ncols4 = ref_ind_df1.shape[1]
    
    # Fifth data frame construction: Industry by fuel
    ref_ind_df2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                         (EGEDA_years_reference['item_code_new'].isin(['14_industry_sector'])) &
                         (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]
    
    coal = ref_ind_df2[ref_ind_df2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal', 
                                                                                                  item_code_new = '14_industry_sector')
    
    oil = ref_ind_df2[ref_ind_df2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil', 
                                                                                                item_code_new = '14_industry_sector')
    
    renewables = ref_ind_df2[ref_ind_df2['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables', 
                                                                                                              item_code_new = '14_industry_sector')
    
    others = ref_ind_df2[ref_ind_df2['fuel_code'].isin(Others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Others', 
                                                                                                                item_code_new = '14_industry_sector')
    
    ref_ind_df2 = ref_ind_df2.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_ind_df2.loc[:, '2000':])].reset_index(drop = True)

    ref_ind_df2.loc[ref_ind_df2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_ind_df2.loc[ref_ind_df2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_ind_df2.loc[ref_ind_df2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_ind_df2.loc[ref_ind_df2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_ind_df2 = ref_ind_df2[ref_ind_df2['fuel_code'].isin(FED_agg_fuels_ind)].set_index('fuel_code').loc[FED_agg_fuels_ind].reset_index()
    
    nrows5 = ref_ind_df2.shape[0]
    ncols5 = ref_ind_df2.shape[1]

    # Transport data frame construction: FED by fuels
    ref_transport_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                          (EGEDA_years_reference['fuel_code'].isin(Transport_fuels))]
    
    renewables = ref_transport_df1[ref_transport_df1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '15_transport_sector')
    
    others = ref_transport_df1[ref_transport_df1['fuel_code'].isin(Other_fuel)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '15_transport_sector')

    trans_gasoline = ref_transport_df1[ref_transport_df1['fuel_code'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Gasoline', 
                                                            item_code_new = '15_transport_sector')

    trans_jetfuel = ref_transport_df1[ref_transport_df1['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Jet fuel', 
                                                            item_code_new = '15_transport_sector')
    
    ref_transport_df1 = ref_transport_df1.append([renewables, trans_gasoline, trans_jetfuel, others])[['fuel_code', 'item_code_new'] + list(ref_transport_df1.loc[:, '2000':])].reset_index(drop = True) 

    ref_transport_df1.loc[ref_transport_df1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    ref_transport_df1.loc[ref_transport_df1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    ref_transport_df1.loc[ref_transport_df1['fuel_code'] == '7_9_lpg', 'fuel_code'] = 'LPG'
    ref_transport_df1.loc[ref_transport_df1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_transport_df1.loc[ref_transport_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_transport_df1 = ref_transport_df1[ref_transport_df1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    nrows10 = ref_transport_df1.shape[0]
    ncols10 = ref_transport_df1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    ref_transport_df2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                               (EGEDA_years_reference['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years_reference['fuel_code'].isin(['19_total']))].copy()
    
    ref_transport_df2.loc[ref_transport_df2['item_code_new'] == '15_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    ref_transport_df2.loc[ref_transport_df2['item_code_new'] == '15_2_road', 'item_code_new'] = 'Road'
    ref_transport_df2.loc[ref_transport_df2['item_code_new'] == '15_3_rail', 'item_code_new'] = 'Rail'
    ref_transport_df2.loc[ref_transport_df2['item_code_new'] == '15_4_domestic_navigation', 'item_code_new'] = 'Marine'
    ref_transport_df2.loc[ref_transport_df2['item_code_new'] == '15_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    ref_transport_df2.loc[ref_transport_df2['item_code_new'] == '15_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    ref_transport_df2 = ref_transport_df2[ref_transport_df2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    ref_transport_df2 = ref_transport_df2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    nrows11 = ref_transport_df2.shape[0]
    ncols11 = ref_transport_df2.shape[1]

    # Agriculture data frame

    ref_ag_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                         (EGEDA_years_reference['item_code_new'].isin(Ag_items)) &
                         (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].groupby('fuel_code').sum().assign(item_code_new = 'Agriculture').reset_index()
                     
    coal = ref_ag_df1[ref_ag_df1['fuel_code'].isin(Coal_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Agriculture')

    oil = ref_ag_df1[ref_ag_df1['fuel_code'].isin(Oil_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Agriculture')

    renewables = ref_ag_df1[ref_ag_df1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Modern renewables', item_code_new = 'Agriculture')
    
    others = ref_ag_df1[ref_ag_df1['fuel_code'].isin(Others_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Agriculture')
    
    ref_ag_df1 = ref_ag_df1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_ag_df1.loc[:,'2000':'2050'])].reset_index(drop = True)

    ref_ag_df1.loc[ref_ag_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_ag_df1.loc[ref_ag_df1['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Traditional biomass'
    ref_ag_df1.loc[ref_ag_df1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_ag_df1.loc[ref_ag_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_ag_df1.loc[ref_ag_df1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_ag_df1 = ref_ag_df1[ref_ag_df1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()
    
    nrows14 = ref_ag_df1.shape[0]
    ncols14 = ref_ag_df1.shape[1]

    ref_ag_df2 = ref_ag_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows15 = ref_ag_df2.shape[0]
    ncols15 = ref_ag_df2.shape[1]

    # Hydrogen data frame reference

    ref_hyd_df1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                                        (EGEDA_years_reference['item_code_new'].isin(Sectors_tfc)) &
                                        (EGEDA_years_reference['fuel_code'] == '16_9_other_sources')].groupby('item_code_new').sum().assign(fuel_code = 'Hydrogen').reset_index()

    buildings_hy = ref_hyd_df1[ref_hyd_df1['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Buildings', fuel_code = 'Hydrogen')

    ag_hy = ref_hyd_df1[ref_hyd_df1['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Agriculture', fuel_code = 'Hydrogen')

    ref_hyd_df1 = ref_hyd_df1.append([buildings_hy, ag_hy])\
        [['fuel_code', 'item_code_new'] + list(ref_hyd_df1.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_hyd_df1.loc[ref_hyd_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_hyd_df1.loc[ref_hyd_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'

    ref_hyd_df1 = ref_hyd_df1[ref_hyd_df1['item_code_new'].isin(['Agriculture', 'Buildings', 'Industry', 'Transport'])]

    nrows16 = ref_hyd_df1.shape[0]
    ncols16 = ref_hyd_df1.shape[1]

    ###############################################################################################################

    # NET ZERO DATA FRAMES
    # First data frame construction: FED by fuels
    netz_econ_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'].isin(no_trad_bio_sectors)) &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    netz_econ_df1 = netz_econ_df1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Industry, transport, NE').reset_index()

    #nrows1 = econ_df1.shape[0]
    #ncols1 = econ_df1.shape[1]

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Industry, transport, NE')
    
    oil = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Industry, transport, NE')
    
    renewables = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Modern renewables', item_code_new = 'Industry, transport, NE')
    
    others = netz_econ_df1[netz_econ_df1['fuel_code'].isin(Others_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Industry, transport, NE')

    # Fed fuel data frame 1 (data frame 6)

    netz_fedfuel_df1 = netz_econ_df1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_econ_df1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_fedfuel_df1.loc[netz_fedfuel_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_fedfuel_df1.loc[netz_fedfuel_df1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_fedfuel_df1.loc[netz_fedfuel_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_fedfuel_df1.loc[netz_fedfuel_df1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    # Insert 0 traditional biomass row
    new_row = ['Traditional biomass', 'Industry, transport, NE'] + [0] * 51
    new_series = pd.Series(new_row, index = netz_fedfuel_df1.columns)

    netz_fedfuel_df1 = netz_fedfuel_df1.append(new_series, ignore_index = True)

    netz_fedfuel_df1 = netz_fedfuel_df1[netz_fedfuel_df1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ##### No biomass fix for dataframe

    netz_tradbio_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(trad_bio_sectors)) &
                                           (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)

    netz_tradbio_df1 = netz_tradbio_df1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Trad bio sectors').reset_index()

    # build aggregate with altered vector to account for no biomass in renewables
    coal_tradbio = netz_tradbio_df1[netz_tradbio_df1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Coal', item_code_new = 'Trad bio sectors')

    oil_tradbio = netz_tradbio_df1[netz_tradbio_df1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Oil', item_code_new = 'Trad bio sectors')

    renew_tradbio = netz_tradbio_df1[netz_tradbio_df1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Modern renewables', item_code_new = 'Trad bio sectors')

    others_tradbio = netz_tradbio_df1[netz_tradbio_df1['fuel_code'].isin(Others_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Others', item_code_new = 'Trad bio sectors')

    # Fed fuel no biomass in other sector renewables
    netz_tradbio_df2 = netz_tradbio_df1.append([coal_tradbio, oil_tradbio, renew_tradbio, others_tradbio])\
        [['fuel_code', 'item_code_new'] + list(netz_tradbio_df1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_tradbio_df2.loc[netz_tradbio_df2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_tradbio_df2.loc[netz_tradbio_df2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Traditional biomass'
    netz_tradbio_df2.loc[netz_tradbio_df2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_tradbio_df2.loc[netz_tradbio_df2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_tradbio_df2.loc[netz_tradbio_df2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_tradbio_df2 = netz_tradbio_df2[netz_tradbio_df2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    netz_fedfuel_df1 = netz_fedfuel_df1.append(netz_tradbio_df2)

    # Combine the two dataframes that account for Modern renewables
    netz_fedfuel_df1 = netz_fedfuel_df1.copy().groupby(['fuel_code']).sum().assign(item_code_new = '12_total_final_consumption')\
        .reset_index()[['fuel_code', 'item_code_new'] + list(netz_fedfuel_df1.loc[:,'2000':'2050'])]\
            .set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    nrows26 = netz_fedfuel_df1.shape[0]
    ncols26 = netz_fedfuel_df1.shape[1]

    netz_fedfuel_df2 = netz_fedfuel_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows27 = netz_fedfuel_df2.shape[0]
    ncols27 = netz_fedfuel_df2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    netz_econ_df2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                        (EGEDA_years_netzero['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years_netzero['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    netz_econ_df2 = netz_econ_df2[['fuel_code', 'item_code_new'] + list(netz_econ_df2.loc[:,'2000':])]
    
    nrows22 = netz_econ_df2.shape[0]
    ncols22 = netz_econ_df2.shape[1]

    # Now build aggregate sector variables
    
    buildings = netz_econ_df2[netz_econ_df2['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = netz_econ_df2[netz_econ_df2['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')
    
    # Build aggregate data frame of FED sector

    netz_fedsector_df1 = netz_econ_df2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(netz_econ_df2.loc[:, '2000':])].reset_index(drop = True)

    netz_fedsector_df1.loc[netz_fedsector_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_fedsector_df1.loc[netz_fedsector_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_fedsector_df1.loc[netz_fedsector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_fedsector_df1.loc[netz_fedsector_df1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_fedsector_df1 = netz_fedsector_df1[netz_fedsector_df1['item_code_new'].isin(FED_agg_sectors)].set_index('item_code_new').loc[FED_agg_sectors].reset_index()
    netz_fedsector_df1 = netz_fedsector_df1[['fuel_code', 'item_code_new'] + list(netz_fedsector_df1.loc[:, '2000':])]

    nrows28 = netz_fedsector_df1.shape[0]
    ncols28 = netz_fedsector_df1.shape[1]

    netz_fedsector_df2 = netz_fedsector_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows29 = netz_fedsector_df2.shape[0]
    ncols29 = netz_fedsector_df2.shape[1]
    
    # Third data frame construction: Buildings FED by fuel
    netz_bld_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                         (EGEDA_years_netzero['item_code_new'].isin(Buildings_items)) &
                         (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))]
    
    for fuel in Required_fuels:
        buildings = netz_bld_df1[netz_bld_df1['fuel_code'] == fuel].groupby(['economy', 'fuel_code']).sum().assign(item_code_new = '16_x_buildings')
        buildings['economy'] = economy
        buildings['fuel_code'] = fuel
        
        netz_bld_df1 = netz_bld_df1.append(buildings).reset_index(drop = True)
        
    netz_bld_df1 = netz_bld_df1[['fuel_code', 'item_code_new'] + list(netz_bld_df1.loc[:, '2000':])]
    
    nrows23 = netz_bld_df1.shape[0]
    ncols23 = netz_bld_df1.shape[1]

    netz_bld_df2 = netz_bld_df1[netz_bld_df1['item_code_new'] == '16_x_buildings']

    coal = netz_bld_df2[netz_bld_df2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = '16_x_buildings')
    
    oil = netz_bld_df2[netz_bld_df2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = '16_x_buildings')
    
    renewables = netz_bld_df2[netz_bld_df2['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Modern renewables', item_code_new = '16_x_buildings')
    
    others = netz_bld_df2[netz_bld_df2['fuel_code'].isin(Others_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = '16_x_buildings')

    netz_bld_df2 = netz_bld_df2.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_bld_df2.loc[:, '2000':])].reset_index(drop = True)

    netz_bld_df2.loc[netz_bld_df2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_bld_df2.loc[netz_bld_df2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Traditional biomass'
    netz_bld_df2.loc[netz_bld_df2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_bld_df2.loc[netz_bld_df2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_bld_df2.loc[netz_bld_df2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_bld_df2 = netz_bld_df2[netz_bld_df2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code')\
        .loc[FED_agg_fuels].reset_index()
    nrows32 = netz_bld_df2.shape[0]
    ncols32 = netz_bld_df2.shape[1]

    netz_bld_df3 = netz_bld_df1[(netz_bld_df1['fuel_code'] == '19_total') &
                      (netz_bld_df1['item_code_new'].isin(Buildings_items))].copy()

    netz_bld_df3.loc[netz_bld_df3['item_code_new'] == '16_1_commercial_and_public_services', 'item_code_new'] = 'Services' 
    netz_bld_df3.loc[netz_bld_df3['item_code_new'] == '16_2_residential', 'item_code_new'] = 'Residential'

    nrows33 = netz_bld_df3.shape[0]
    ncols33 = netz_bld_df3.shape[1]
    
    # Fourth data frame construction: Industry subsector
    netz_ind_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                         (EGEDA_years_netzero['item_code_new'].isin(Subindustry)) &
                         (EGEDA_years_netzero['fuel_code'] == '19_total')]

    other_industry = netz_ind_df1[netz_ind_df1['item_code_new'].isin(Other_industry)].groupby(['fuel_code']).sum().assign(item_code_new = 'Other',
                                                                                                                fuel_code = '19_total')

    netz_ind_df1 = netz_ind_df1.append([other_industry])[['fuel_code', 'item_code_new'] + list(netz_ind_df1.loc[:, '2000':])].reset_index(drop = True)

    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_1_iron_and_steel', 'item_code_new'] = 'Iron & steel'
    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_2_chemical_incl_petrochemical', 'item_code_new'] = 'Chemicals'
    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_3_non_ferrous_metals', 'item_code_new'] = 'Aluminium'
    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_4_nonmetallic_mineral_products', 'item_code_new'] = 'Non-metallic minerals'  
    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_7_mining_and_quarrying', 'item_code_new'] = 'Mining'
    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_9_pulp_paper_and_printing', 'item_code_new'] = 'Pulp & paper'
    netz_ind_df1.loc[netz_ind_df1['item_code_new'] == '14_13_nonspecified_industry', 'item_code_new'] = 'Non-specified'
    
    netz_ind_df1 = netz_ind_df1[netz_ind_df1['item_code_new'].isin(Industry_eight)].set_index('item_code_new').loc[Industry_eight].reset_index()

    netz_ind_df1 = netz_ind_df1[['fuel_code', 'item_code_new'] + list(netz_ind_df1.loc[:, '2000':])]

    nrows24 = netz_ind_df1.shape[0]
    ncols24 = netz_ind_df1.shape[1]
    
    # Fifth data frame construction: Industry by fuel
    netz_ind_df2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                         (EGEDA_years_netzero['item_code_new'].isin(['14_industry_sector'])) &
                         (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))]
    
    coal = netz_ind_df2[netz_ind_df2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal', 
                                                                                                  item_code_new = '14_industry_sector')
    
    oil = netz_ind_df2[netz_ind_df2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil', 
                                                                                                item_code_new = '14_industry_sector')
    
    renewables = netz_ind_df2[netz_ind_df2['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables', 
                                                                                                              item_code_new = '14_industry_sector')
    
    others = netz_ind_df2[netz_ind_df2['fuel_code'].isin(Others_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Others', 
                                                                                                                item_code_new = '14_industry_sector')
    
    netz_ind_df2 = netz_ind_df2.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_ind_df2.loc[:, '2000':])].reset_index(drop = True)

    netz_ind_df2.loc[netz_ind_df2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_ind_df2.loc[netz_ind_df2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_ind_df2.loc[netz_ind_df2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_ind_df2.loc[netz_ind_df2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_ind_df2 = netz_ind_df2[netz_ind_df2['fuel_code'].isin(FED_agg_fuels_ind)].set_index('fuel_code').loc[FED_agg_fuels_ind].reset_index()
    
    nrows25 = netz_ind_df2.shape[0]
    ncols25 = netz_ind_df2.shape[1]

    # Transport data frame construction: FED by fuels
    netz_transport_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                          (EGEDA_years_netzero['fuel_code'].isin(Transport_fuels))]
    
    renewables = netz_transport_df1[netz_transport_df1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '15_transport_sector')
    
    others = netz_transport_df1[netz_transport_df1['fuel_code'].isin(Other_fuel)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '15_transport_sector')

    trans_gasoline = netz_transport_df1[netz_transport_df1['fuel_code'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Gasoline', 
                                                            item_code_new = '15_transport_sector')

    trans_jetfuel = netz_transport_df1[netz_transport_df1['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Jet fuel', 
                                                            item_code_new = '15_transport_sector')
    
    netz_transport_df1 = netz_transport_df1.append([renewables, trans_gasoline, trans_jetfuel, others])[['fuel_code', 'item_code_new'] + list(netz_transport_df1.loc[:, '2000':])].reset_index(drop = True) 

    netz_transport_df1.loc[netz_transport_df1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    netz_transport_df1.loc[netz_transport_df1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    netz_transport_df1.loc[netz_transport_df1['fuel_code'] == '7_9_lpg', 'fuel_code'] = 'LPG'
    netz_transport_df1.loc[netz_transport_df1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_transport_df1.loc[netz_transport_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_transport_df1 = netz_transport_df1[netz_transport_df1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    nrows30 = netz_transport_df1.shape[0]
    ncols30 = netz_transport_df1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    netz_transport_df2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                               (EGEDA_years_netzero['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years_netzero['fuel_code'].isin(['19_total']))].copy()
    
    netz_transport_df2.loc[netz_transport_df2['item_code_new'] == '15_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    netz_transport_df2.loc[netz_transport_df2['item_code_new'] == '15_2_road', 'item_code_new'] = 'Road'
    netz_transport_df2.loc[netz_transport_df2['item_code_new'] == '15_3_rail', 'item_code_new'] = 'Rail'
    netz_transport_df2.loc[netz_transport_df2['item_code_new'] == '15_4_domestic_navigation', 'item_code_new'] = 'Marine'
    netz_transport_df2.loc[netz_transport_df2['item_code_new'] == '15_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    netz_transport_df2.loc[netz_transport_df2['item_code_new'] == '15_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    netz_transport_df2 = netz_transport_df2[netz_transport_df2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    netz_transport_df2 = netz_transport_df2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    nrows31 = netz_transport_df2.shape[0]
    ncols31 = netz_transport_df2.shape[1]

    # Agriculture data frame

    netz_ag_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                         (EGEDA_years_netzero['item_code_new'].isin(Ag_items)) &
                         (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].groupby('fuel_code').sum().assign(item_code_new = 'Agriculture').reset_index()
                     
    coal = netz_ag_df1[netz_ag_df1['fuel_code'].isin(Coal_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Agriculture')

    oil = netz_ag_df1[netz_ag_df1['fuel_code'].isin(Oil_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Agriculture')

    renewables = netz_ag_df1[netz_ag_df1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Modern renewables', item_code_new = 'Agriculture')
    
    others = netz_ag_df1[netz_ag_df1['fuel_code'].isin(Others_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Agriculture')
    
    netz_ag_df1 = netz_ag_df1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_ag_df1.loc[:,'2000':'2050'])].reset_index(drop = True)

    netz_ag_df1.loc[netz_ag_df1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_ag_df1.loc[netz_ag_df1['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Traditional biomass'
    netz_ag_df1.loc[netz_ag_df1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_ag_df1.loc[netz_ag_df1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_ag_df1.loc[netz_ag_df1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_ag_df1 = netz_ag_df1[netz_ag_df1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()
    
    nrows34 = netz_ag_df1.shape[0]
    ncols34 = netz_ag_df1.shape[1]

    netz_ag_df2 = netz_ag_df1[['fuel_code', 'item_code_new'] + col_chart_years]

    nrows35 = netz_ag_df2.shape[0]
    ncols35 = netz_ag_df2.shape[1]

    # Hydrogen data frame net zero

    netz_hyd_df1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                                        (EGEDA_years_netzero['item_code_new'].isin(Sectors_tfc)) &
                                        (EGEDA_years_netzero['fuel_code'] == '16_9_other_sources')].groupby('item_code_new').sum().assign(fuel_code = 'Hydrogen').reset_index()

    buildings_hy = netz_hyd_df1[netz_hyd_df1['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Buildings', fuel_code = 'Hydrogen')

    ag_hy = netz_hyd_df1[netz_hyd_df1['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Agriculture', fuel_code = 'Hydrogen')

    netz_hyd_df1 = netz_hyd_df1.append([buildings_hy, ag_hy])\
        [['fuel_code', 'item_code_new'] + list(netz_hyd_df1.loc[:, '2017':'2050'])].reset_index(drop = True)

    netz_hyd_df1.loc[netz_hyd_df1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_hyd_df1.loc[netz_hyd_df1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'

    netz_hyd_df1 = netz_hyd_df1[netz_hyd_df1['item_code_new'].isin(['Agriculture', 'Buildings', 'Industry', 'Transport'])]

    nrows36 = netz_hyd_df1.shape[0]
    ncols36 = netz_hyd_df1.shape[1]

    ############################################################################################################################
    
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
    # REFERENCE and NETZERO
    ref_fedfuel_df1.to_excel(writer, sheet_name = economy + '_FED_fuel_ref', index = False, startrow = chart_height)
    netz_fedfuel_df1.to_excel(writer, sheet_name = economy + '_FED_fuel_netz', index = False, startrow = chart_height)
    ref_fedfuel_df2.to_excel(writer, sheet_name = economy + '_FED_fuel_ref', index = False, startrow = chart_height + nrows6 + 3)
    netz_fedfuel_df2.to_excel(writer, sheet_name = economy + '_FED_fuel_netz', index = False, startrow = chart_height + nrows26 + 3)
    ref_econ_df2.to_excel(writer, sheet_name = economy + '_FED_sector_ref', index = False, startrow = chart_height)
    netz_econ_df2.to_excel(writer, sheet_name = economy + '_FED_sector_netz', index = False, startrow = chart_height)
    ref_fedsector_df1.to_excel(writer, sheet_name = economy + '_FED_sector_ref', index = False, startrow = chart_height + nrows2 + 3)
    netz_fedsector_df1.to_excel(writer, sheet_name = economy + '_FED_sector_netz', index = False, startrow = chart_height + nrows22 + 3)
    ref_fedsector_df2.to_excel(writer, sheet_name = economy + '_FED_sector_ref', index = False, startrow = chart_height + nrows2 + nrows8 + 6)
    netz_fedsector_df2.to_excel(writer, sheet_name = economy + '_FED_sector_netz', index = False, startrow = chart_height + nrows22 + nrows28 + 6)
    ref_bld_df2.to_excel(writer, sheet_name = economy + '_FED_bld_ref', index = False, startrow = chart_height)
    netz_bld_df2.to_excel(writer, sheet_name = economy + '_FED_bld_netz', index = False, startrow = chart_height)
    ref_bld_df3.to_excel(writer, sheet_name = economy + '_FED_bld_ref', index = False, startrow = chart_height + nrows12 + 3)
    netz_bld_df3.to_excel(writer, sheet_name = economy + '_FED_bld_netz', index = False, startrow = chart_height + nrows32 + 3)
    ref_ind_df1.to_excel(writer, sheet_name = economy + '_FED_ind_ref', index = False, startrow = chart_height)
    netz_ind_df1.to_excel(writer, sheet_name = economy + '_FED_ind_netz', index = False, startrow = chart_height)
    ref_ind_df2.to_excel(writer, sheet_name = economy + '_FED_ind_ref', index = False, startrow = chart_height + nrows4 + 2)
    netz_ind_df2.to_excel(writer, sheet_name = economy + '_FED_ind_netz', index = False, startrow = chart_height + nrows24 + 2)
    ref_transport_df1.to_excel(writer, sheet_name = economy + '_FED_trn_ref', index = False, startrow = chart_height)
    netz_transport_df1.to_excel(writer, sheet_name = economy + '_FED_trn_netz', index = False, startrow = chart_height)
    ref_transport_df2.to_excel(writer, sheet_name = economy + '_FED_trn_ref', index = False, startrow = chart_height + nrows10 + 3)
    netz_transport_df2.to_excel(writer, sheet_name = economy + '_FED_trn_netz', index = False, startrow = chart_height + nrows30 + 3)
    ref_ag_df1.to_excel(writer, sheet_name = economy + '_FED_agr_ref', index = False, startrow = chart_height)
    netz_ag_df1.to_excel(writer, sheet_name = economy + '_FED_agr_netz', index = False, startrow = chart_height)
    ref_ag_df2.to_excel(writer, sheet_name = economy + '_FED_agr_ref', index = False, startrow = chart_height + nrows14 + 3)
    netz_ag_df2.to_excel(writer, sheet_name = economy + '_FED_agr_netz', index = False, startrow = chart_height + nrows34 + 3)
    ref_hyd_df1.to_excel(writer, sheet_name = economy + '_FED_hyd', index = False, startrow = chart_height)
    netz_hyd_df1.to_excel(writer, sheet_name = economy + '_FED_hyd', index = False, startrow = chart_height + nrows16 + 3)
    
    ################################################################################################################################

    # CHARTS
    # REFERENCE

    # Access the workbook and first sheet with data from df1
    ref_worksheet1 = writer.sheets[economy + '_FED_fuel_ref']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet1.set_column(1, ncols6 + 1, None, comma_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height + nrows6 + 3, None, header_format)
    ref_worksheet1.write(0, 0, economy + ' FED fuel', cell_format1)

    # FED Fuel REFERENCE charts

    # Create a FED area chart
    ref_fedfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_fedfuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fedfuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fedfuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart1.set_y_axis({
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
        
    ref_fedfuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fedfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows6):
        ref_fedfuel_chart1.add_series({
            'name':       [economy + '_FED_fuel_ref', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel_ref', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_FED_fuel_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet1.insert_chart('B3', ref_fedfuel_chart1)

    ###################### Create another FED chart showing proportional share #################################

    # Create a another chart
    ref_fedfuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_fedfuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fedfuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fedfuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart2.set_y_axis({
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
        
    ref_fedfuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fedfuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_fuels:
        i = ref_fedfuel_df2[ref_fedfuel_df2['fuel_code'] == component].index[0]
        ref_fedfuel_chart2.add_series({
            'name':       [economy + '_FED_fuel_ref', chart_height + nrows6 + i + 4, 0],
            'categories': [economy + '_FED_fuel_ref', chart_height + nrows6 + 3, 2, chart_height + nrows6 + 3, ncols7 - 1],
            'values':     [economy + '_FED_fuel_ref', chart_height + nrows6 + i + 4, 2, chart_height + nrows6 + i + 4, ncols7 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet1.insert_chart('J3', ref_fedfuel_chart2)

    # Create a FED line chart with higher level aggregation
    ref_fedfuel_chart3 = workbook.add_chart({'type': 'line'})
    ref_fedfuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fedfuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fedfuel_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedfuel_chart3.set_y_axis({
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
        
    ref_fedfuel_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fedfuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows6):
        ref_fedfuel_chart3.add_series({
            'name':       [economy + '_FED_fuel_ref', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel_ref', chart_height, 2, chart_height, ncols6 - 1],
            'values':     [economy + '_FED_fuel_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols6 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    ref_worksheet1.insert_chart('R3', ref_fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet2 = writer.sheets[economy + '_FED_sector_ref']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet2.set_column(1, ncols2 + 1, None, comma_format)
    ref_worksheet2.set_row(chart_height, None, header_format)
    ref_worksheet2.set_row(chart_height + nrows2 + 3, None, header_format)
    ref_worksheet2.set_row(chart_height + nrows2 + nrows8 + 6, None, header_format)
    ref_worksheet2.write(0, 0, economy + ' FED sector', cell_format1)
    
    # Create a FED chart
    ref_fed_sector_chart1 = workbook.add_chart({'type': 'line'})
    ref_fed_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fed_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fed_sector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_sector_chart1.set_y_axis({
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
        
    ref_fed_sector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fed_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows2):
        ref_fed_sector_chart1.add_series({
            'name':       [economy + '_FED_sector_ref', chart_height + i + 1, 1],
            'categories': [economy + '_FED_sector_ref', chart_height, 2, chart_height, ncols2 - 1],
            'values':     [economy + '_FED_sector_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols2 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    ref_worksheet2.insert_chart('Z3', ref_fed_sector_chart1)

    # Create a FED sector area chart

    ref_fedsector_chart3 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_fedsector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fedsector_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fedsector_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart3.set_y_axis({
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
        
    ref_fedsector_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fedsector_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows8):
        ref_fedsector_chart3.add_series({
            'name':       [economy + '_FED_sector_ref', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_FED_sector_ref', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_FED_sector_ref', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet2.insert_chart('B3', ref_fedsector_chart3)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    ref_fedsector_chart4 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_fedsector_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fedsector_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fedsector_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart4.set_y_axis({
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
        
    ref_fedsector_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fedsector_chart4.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_sectors:
        i = ref_fedsector_df2[ref_fedsector_df2['item_code_new'] == component].index[0]
        ref_fedsector_chart4.add_series({
            'name':       [economy + '_FED_sector_ref', chart_height + nrows2 + nrows8 + i + 7, 1],
            'categories': [economy + '_FED_sector_ref', chart_height + nrows2 + nrows8 + 6, 2, chart_height + nrows2 + nrows8 + 6, ncols9 - 1],
            'values':     [economy + '_FED_sector_ref', chart_height + nrows2 + nrows8 + i + 7, 2, chart_height + nrows2 + nrows8 + i + 7, ncols9 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet2.insert_chart('J3', ref_fedsector_chart4)

    # Create a FED sector line chart

    ref_fedsector_chart5 = workbook.add_chart({'type': 'line'})
    ref_fedsector_chart5.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fedsector_chart5.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fedsector_chart5.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_fedsector_chart5.set_y_axis({
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
        
    ref_fedsector_chart5.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fedsector_chart5.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows8):
        ref_fedsector_chart5.add_series({
            'name':       [economy + '_FED_sector_ref', chart_height + nrows2 + i + 4, 1],
            'categories': [economy + '_FED_sector_ref', chart_height + nrows2 + 3, 2, chart_height + nrows2 + 3, ncols8 - 1],
            'values':     [economy + '_FED_sector_ref', chart_height + nrows2 + i + 4, 2, chart_height + nrows2 + i + 4, ncols8 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    ref_worksheet2.insert_chart('R3', ref_fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    ref_worksheet3 = writer.sheets[economy + '_FED_bld_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet3.set_column(2, ncols3 + 1, None, comma_format)
    ref_worksheet3.set_row(chart_height, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows12 + 3, None, header_format)
    ref_worksheet3.write(0, 0, economy + ' buildings', cell_format1)
    
    # Create a FED chart
    ref_fed_bld_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_fed_bld_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fed_bld_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fed_bld_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_bld_chart1.set_y_axis({
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
        
    ref_fed_bld_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fed_bld_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_fuels:
        i = ref_bld_df2[ref_bld_df2['fuel_code'] == component].index[0]
        ref_fed_bld_chart1.add_series({
            'name':       [economy + '_FED_bld_ref', chart_height + i + 1, 0],
            'categories': [economy + '_FED_bld_ref', chart_height, 2, chart_height, ncols12 - 1],
            'values':     [economy + '_FED_bld_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols12 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })

    ref_worksheet3.insert_chart('B3', ref_fed_bld_chart1)
    
    ################## FED building chart 2 (residential versus services) ###########################################
    
    # Create a second FED building chart
    ref_fed_bld_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_fed_bld_chart2.set_size({
        'width': 500,
        'height': 300
    })

    ref_fed_bld_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fed_bld_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_bld_chart2.set_y_axis({
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
        
    ref_fed_bld_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fed_bld_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(2):
        ref_fed_bld_chart2.add_series({
            'name':       [economy + '_FED_bld_ref', chart_height + nrows12 + 4 + i, 1],
            'categories': [economy + '_FED_bld_ref', chart_height + nrows12 + 3, 2, chart_height + nrows12 + 3, ncols13 - 1],
            'values':     [economy + '_FED_bld_ref', chart_height + nrows12 + 4 + i, 2, chart_height + nrows12 + 4 + i, ncols13 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('J3', ref_fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # Access the workbook and fourth sheet with data from bld_df1
    ref_worksheet4 = writer.sheets[economy + '_FED_ind_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet4.set_column(2, ncols4 + 1, None, comma_format)
    ref_worksheet4.set_row(chart_height, None, header_format)
    ref_worksheet4.set_row(chart_height + nrows4 + 2, None, header_format)
    ref_worksheet4.write(0, 0, economy + ' industry', cell_format1)
    
    # Create a industry subsector FED chart
    ref_fed_ind_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_fed_ind_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_fed_ind_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fed_ind_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_ind_chart1.set_y_axis({
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
        
    ref_fed_ind_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fed_ind_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows4):
        ref_fed_ind_chart1.add_series({
            'name':       [economy + '_FED_ind_ref', chart_height + i + 1, 1],
            'categories': [economy + '_FED_ind_ref', chart_height, 2, chart_height, ncols4 - 1],
            'values':     [economy + '_FED_ind_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols4 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet4.insert_chart('B3', ref_fed_ind_chart1)
    
    ############# FED industry chart 2 (industry by fuel)
    
    # Create a FED industry fuel chart
    ref_fed_ind_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_fed_ind_chart2.set_size({
        'width': 500,
        'height': 300
    })

    ref_fed_ind_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_fed_ind_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_fed_ind_chart2.set_y_axis({
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
        
    ref_fed_ind_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_fed_ind_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for fuel_agg in FED_agg_fuels_ind:
        j = ref_ind_df2[ref_ind_df2['fuel_code'] == fuel_agg].index[0]
        ref_fed_ind_chart2.add_series({
            'name':       [economy + '_FED_ind_ref', chart_height + nrows4 + j + 3, 0],
            'categories': [economy + '_FED_ind_ref', chart_height + nrows4 + 2, 2, chart_height + nrows4 + 2, ncols5 - 1],
            'values':     [economy + '_FED_ind_ref', chart_height + nrows4 + j + 3, 2, chart_height + nrows4 + j + 3, ncols5 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True}
        })
    
    ref_worksheet4.insert_chart('J3', ref_fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    ref_worksheet5 = writer.sheets[economy + '_FED_trn_ref']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet5.set_column(2, ncols10 + 1, None, comma_format)
    ref_worksheet5.set_row(chart_height, None, header_format)
    ref_worksheet5.set_row(chart_height + nrows10 + 3, None, header_format)
    ref_worksheet5.write(0, 0, economy + ' FED transport', cell_format1)
    
    # Create a transport FED area chart
    ref_transport_chart1 = workbook.add_chart({'type': 'area', 
                                           'subtype': 'stacked'})
    ref_transport_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_transport_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_transport_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_transport_chart1.set_y_axis({
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
        
    ref_transport_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_transport_chart1.set_title({
        'none': True
    })
        
    for fuel_agg in Transport_fuels_agg:
        j = ref_transport_df1[ref_transport_df1['fuel_code'] == fuel_agg].index[0]
        ref_transport_chart1.add_series({
            'name':       [economy + '_FED_trn_ref', chart_height + j + 1, 0],
            'categories': [economy + '_FED_trn_ref', chart_height, 2, chart_height, ncols10 - 1],
            'values':     [economy + '_FED_trn_ref', chart_height + j + 1, 2, chart_height + j + 1, ncols10 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True} 
        })
    
    ref_worksheet5.insert_chart('B3', ref_transport_chart1)
            
    ############# FED transport chart 2 (transport by modality)
    
    # Create a FED transport modality column chart
    ref_transport_chart2 = workbook.add_chart({'type': 'column', 
                                         'subtype': 'stacked'})
    ref_transport_chart2.set_size({
        'width': 500,
        'height': 300
    })

    ref_transport_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_transport_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    ref_transport_chart2.set_y_axis({
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
        
    ref_transport_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_transport_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for modality in Transport_modal_agg:
        j = ref_transport_df2[ref_transport_df2['item_code_new'] == modality].index[0]
        ref_transport_chart2.add_series({
            'name':       [economy + '_FED_trn_ref', chart_height + nrows10 + j + 4, 1],
            'categories': [economy + '_FED_trn_ref', chart_height + nrows10 + 3, 2, chart_height + nrows10 + 3, ncols11 - 1],
            'values':     [economy + '_FED_trn_ref', chart_height + nrows10 + j + 4, 2, chart_height + nrows10 + j + 4, ncols11 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True}
        })
    
    ref_worksheet5.insert_chart('J3', ref_transport_chart2)

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    ref_worksheet6 = writer.sheets[economy + '_FED_agr_ref']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet6.set_column(2, ncols14 + 1, None, comma_format)
    ref_worksheet6.set_row(chart_height, None, header_format)
    ref_worksheet6.set_row(chart_height + nrows14 + 3, None, header_format)
    ref_worksheet6.write(0, 0, economy + ' FED agriculture', cell_format1)

    # Create a Agriculture line chart 
    ref_ag_chart1 = workbook.add_chart({'type': 'line'})
    ref_ag_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_ag_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_ag_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_ag_chart1.set_y_axis({
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
        
    ref_ag_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_ag_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows14):
        ref_ag_chart1.add_series({
            'name':       [economy + '_FED_agr_ref', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agr_ref', chart_height, 2, chart_height, ncols14 - 1],
            'values':     [economy + '_FED_agr_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols14 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    ref_worksheet6.insert_chart('B3', ref_ag_chart1)

    # Create a Agriculture area chart
    ref_ag_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_ag_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_ag_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    ref_ag_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_ag_chart2.set_y_axis({
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
        
    ref_ag_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_ag_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows14):
        ref_ag_chart2.add_series({
            'name':       [economy + '_FED_agr_ref', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agr_ref', chart_height, 2, chart_height, ncols14 - 1],
            'values':     [economy + '_FED_agr_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols14 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet6.insert_chart('J3', ref_ag_chart2)

    # Create a Agriculture stacked column
    ref_ag_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    ref_ag_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_ag_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    ref_ag_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    ref_ag_chart3.set_y_axis({
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
        
    ref_ag_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_ag_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows15):
        ref_ag_chart3.add_series({
            'name':       [economy + '_FED_agr_ref', chart_height + nrows14 + i + 4, 0],
            'categories': [economy + '_FED_agr_ref', chart_height + nrows14 + 3, 2, chart_height + nrows14 + 3, ncols15 - 1],
            'values':     [economy + '_FED_agr_ref', chart_height + nrows14 + i + 4, 2, chart_height + nrows14 + i + 4, ncols15 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet6.insert_chart('R3', ref_ag_chart3)

    # HYDROGEN CHARTS

    # Access the workbook and first sheet with data from df1
    hyd_worksheet1 = writer.sheets[economy + '_FED_hyd']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    hyd_worksheet1.set_column(1, ncols16 + 1, None, comma_format)
    hyd_worksheet1.set_row(chart_height, None, header_format)
    hyd_worksheet1.set_row(chart_height, None, header_format)
    hyd_worksheet1.set_row(chart_height + nrows16 + 3, None, header_format)
    hyd_worksheet1.write(0, 0, economy + ' FED hydrogen', cell_format1)

    # Create a HYDROGEN area chart
    ref_hyd_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    ref_hyd_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    ref_hyd_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    ref_hyd_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    ref_hyd_chart1.set_y_axis({
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
        
    ref_hyd_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    ref_hyd_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows16):
        ref_hyd_chart1.add_series({
            'name':       [economy + '_FED_hyd', chart_height + i + 1, 1],
            'categories': [economy + '_FED_hyd', chart_height, 2, chart_height, ncols16 - 1],
            'values':     [economy + '_FED_hyd', chart_height + i + 1, 2, chart_height + i + 1, ncols16 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    hyd_worksheet1.insert_chart('B3', ref_hyd_chart1)

    # Create a HYDROGEN area chart
    netz_hyd_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_hyd_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_hyd_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_hyd_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_hyd_chart1.set_y_axis({
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
        
    netz_hyd_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_hyd_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows36):
        netz_hyd_chart1.add_series({
            'name':       [economy + '_FED_hyd', chart_height + nrows16 + i + 4, 1],
            'categories': [economy + '_FED_hyd', chart_height + nrows16 + 3, 2, chart_height + nrows16 + 3, ncols36 - 1],
            'values':     [economy + '_FED_hyd', chart_height + nrows16 + i + 4, 2, chart_height + nrows16 + i + 4, ncols36 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    hyd_worksheet1.insert_chart('J3', netz_hyd_chart1)

    ##############################################################################################################################

    ##############################################################################################################################

    ##############################################################################################################################

    # CHARTS
    # NET ZERO

    # Access the workbook and first sheet with data from df1
    netz_worksheet1 = writer.sheets[economy + '_FED_fuel_netz']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet1.set_column(1, ncols26 + 1, None, comma_format)
    netz_worksheet1.set_row(chart_height, None, header_format)
    netz_worksheet1.set_row(chart_height, None, header_format)
    netz_worksheet1.set_row(chart_height + nrows26 + 3, None, header_format)
    netz_worksheet1.write(0, 0, economy + ' FED fuel', cell_format1)

    # FED Fuel REFERENCE charts

    # Create a FED area chart
    netz_fedfuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_fedfuel_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fedfuel_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fedfuel_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart1.set_y_axis({
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
        
    netz_fedfuel_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fedfuel_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows26):
        netz_fedfuel_chart1.add_series({
            'name':       [economy + '_FED_fuel_netz', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel_netz', chart_height, 2, chart_height, ncols26 - 1],
            'values':     [economy + '_FED_fuel_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols26 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet1.insert_chart('B3', netz_fedfuel_chart1)

    ###################### Create another FED chart showing proportional share #################################

    # Create a another chart
    netz_fedfuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_fedfuel_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fedfuel_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fedfuel_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart2.set_y_axis({
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
        
    netz_fedfuel_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fedfuel_chart2.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_fuels:
        i = netz_fedfuel_df2[netz_fedfuel_df2['fuel_code'] == component].index[0]
        netz_fedfuel_chart2.add_series({
            'name':       [economy + '_FED_fuel_netz', chart_height + nrows26 + i + 4, 0],
            'categories': [economy + '_FED_fuel_netz', chart_height + nrows26 + 3, 2, chart_height + nrows26 + 3, ncols27 - 1],
            'values':     [economy + '_FED_fuel_netz', chart_height + nrows26 + i + 4, 2, chart_height + nrows26 + i + 4, ncols27 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet1.insert_chart('J3', netz_fedfuel_chart2)

    # Create a FED line chart with higher level aggregation
    netz_fedfuel_chart3 = workbook.add_chart({'type': 'line'})
    netz_fedfuel_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fedfuel_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fedfuel_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedfuel_chart3.set_y_axis({
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
        
    netz_fedfuel_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fedfuel_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows26):
        netz_fedfuel_chart3.add_series({
            'name':       [economy + '_FED_fuel_netz', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel_netz', chart_height, 2, chart_height, ncols26 - 1],
            'values':     [economy + '_FED_fuel_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols26 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    netz_worksheet1.insert_chart('R3', netz_fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    netz_worksheet2 = writer.sheets[economy + '_FED_sector_netz']
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet2.set_column(1, ncols22 + 1, None, comma_format)
    netz_worksheet2.set_row(chart_height, None, header_format)
    netz_worksheet2.set_row(chart_height + nrows22 + 3, None, header_format)
    netz_worksheet2.set_row(chart_height + nrows22 + nrows28 + 6, None, header_format)
    netz_worksheet2.write(0, 0, economy + ' FED sector', cell_format1)
    
    # Create a FED chart
    netz_fed_sector_chart1 = workbook.add_chart({'type': 'line'})
    netz_fed_sector_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fed_sector_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fed_sector_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_sector_chart1.set_y_axis({
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
        
    netz_fed_sector_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fed_sector_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows22):
        netz_fed_sector_chart1.add_series({
            'name':       [economy + '_FED_sector_netz', chart_height + i + 1, 1],
            'categories': [economy + '_FED_sector_netz', chart_height, 2, chart_height, ncols22 - 1],
            'values':     [economy + '_FED_sector_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols22 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    netz_worksheet2.insert_chart('Z3', netz_fed_sector_chart1)

    # Create a FED sector area chart

    netz_fedsector_chart3 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_fedsector_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fedsector_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fedsector_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart3.set_y_axis({
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
        
    netz_fedsector_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fedsector_chart3.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows28):
        netz_fedsector_chart3.add_series({
            'name':       [economy + '_FED_sector_netz', chart_height + nrows22 + i + 4, 1],
            'categories': [economy + '_FED_sector_netz', chart_height + nrows22 + 3, 2, chart_height + nrows22 + 3, ncols28 - 1],
            'values':     [economy + '_FED_sector_netz', chart_height + nrows22 + i + 4, 2, chart_height + nrows22 + i + 4, ncols28 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet2.insert_chart('B3', netz_fedsector_chart3)

    ###################### Create another FED chart showing proportional share #################################

    # Create a FED chart
    netz_fedsector_chart4 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_fedsector_chart4.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fedsector_chart4.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fedsector_chart4.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart4.set_y_axis({
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
        
    netz_fedsector_chart4.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fedsector_chart4.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_sectors:
        i = netz_fedsector_df2[netz_fedsector_df2['item_code_new'] == component].index[0]
        netz_fedsector_chart4.add_series({
            'name':       [economy + '_FED_sector_netz', chart_height + nrows22 + nrows28 + i + 7, 1],
            'categories': [economy + '_FED_sector_netz', chart_height + nrows22 + nrows28 + 6, 2, chart_height + nrows22 + nrows28 + 6, ncols29 - 1],
            'values':     [economy + '_FED_sector_netz', chart_height + nrows22 + nrows28 + i + 7, 2, chart_height + nrows22 + nrows28 + i + 7, ncols29 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet2.insert_chart('J3', netz_fedsector_chart4)

    # Create a FED sector line chart

    netz_fedsector_chart5 = workbook.add_chart({'type': 'line'})
    netz_fedsector_chart5.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fedsector_chart5.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fedsector_chart5.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_fedsector_chart5.set_y_axis({
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
        
    netz_fedsector_chart5.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fedsector_chart5.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows28):
        netz_fedsector_chart5.add_series({
            'name':       [economy + '_FED_sector_netz', chart_height + nrows22 + i + 4, 1],
            'categories': [economy + '_FED_sector_netz', chart_height + nrows22 + 3, 2, chart_height + nrows22 + 3, ncols28 - 1],
            'values':     [economy + '_FED_sector_netz', chart_height + nrows22 + i + 4, 2, chart_height + nrows22 + i + 4, ncols28 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    netz_worksheet2.insert_chart('R3', netz_fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    netz_worksheet3 = writer.sheets[economy + '_FED_bld_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet3.set_column(2, ncols23 + 1, None, comma_format)
    netz_worksheet3.set_row(chart_height, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows32 + 3, None, header_format)
    netz_worksheet3.write(0, 0, economy + ' buildings', cell_format1)
    
    # Create a FED chart
    netz_fed_bld_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_fed_bld_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fed_bld_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fed_bld_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_bld_chart1.set_y_axis({
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
        
    netz_fed_bld_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fed_bld_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for component in FED_agg_fuels:
        i = netz_bld_df2[netz_bld_df2['fuel_code'] == component].index[0]
        netz_fed_bld_chart1.add_series({
            'name':       [economy + '_FED_bld_netz', chart_height + i + 1, 0],
            'categories': [economy + '_FED_bld_netz', chart_height, 2, chart_height, ncols32 - 1],
            'values':     [economy + '_FED_bld_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols32 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })

    netz_worksheet3.insert_chart('B3', netz_fed_bld_chart1)
    
    ################## FED building chart 2 (residential versus services) ###########################################
    
    # Create a second FED building chart
    netz_fed_bld_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_fed_bld_chart2.set_size({
        'width': 500,
        'height': 300
    })

    netz_fed_bld_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fed_bld_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_bld_chart2.set_y_axis({
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
        
    netz_fed_bld_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fed_bld_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(2):
        netz_fed_bld_chart2.add_series({
            'name':       [economy + '_FED_bld_netz', chart_height + nrows32 + 4 + i, 1],
            'categories': [economy + '_FED_bld_netz', chart_height + nrows32 + 3, 2, chart_height + nrows32 + 3, ncols33 - 1],
            'values':     [economy + '_FED_bld_netz', chart_height + nrows32 + 4 + i, 2, chart_height + nrows32 + 4 + i, ncols33 - 1],
            'fill':       {'color': colours_hex[i + 5]},
            'border':     {'none': True}
        })
    
    netz_worksheet3.insert_chart('J3', netz_fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # Access the workbook and fourth sheet with data from bld_df1
    netz_worksheet4 = writer.sheets[economy + '_FED_ind_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet4.set_column(2, ncols24 + 1, None, comma_format)
    netz_worksheet4.set_row(chart_height, None, header_format)
    netz_worksheet4.set_row(chart_height + nrows24 + 2, None, header_format)
    netz_worksheet4.write(0, 0, economy + ' industry', cell_format1)
    
    # Create a industry subsector FED chart
    netz_fed_ind_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_fed_ind_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_fed_ind_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fed_ind_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_ind_chart1.set_y_axis({
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
        
    netz_fed_ind_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fed_ind_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows24):
        netz_fed_ind_chart1.add_series({
            'name':       [economy + '_FED_ind_netz', chart_height + i + 1, 1],
            'categories': [economy + '_FED_ind_netz', chart_height, 2, chart_height, ncols24 - 1],
            'values':     [economy + '_FED_ind_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols24 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet4.insert_chart('B3', netz_fed_ind_chart1)
    
    ############# FED industry chart 2 (industry by fuel)
    
    # Create a FED industry fuel chart
    netz_fed_ind_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_fed_ind_chart2.set_size({
        'width': 500,
        'height': 300
    })

    netz_fed_ind_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_fed_ind_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_fed_ind_chart2.set_y_axis({
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
        
    netz_fed_ind_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_fed_ind_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for fuel_agg in FED_agg_fuels_ind:
        j = netz_ind_df2[netz_ind_df2['fuel_code'] == fuel_agg].index[0]
        netz_fed_ind_chart2.add_series({
            'name':       [economy + '_FED_ind_netz', chart_height + nrows24 + j + 3, 0],
            'categories': [economy + '_FED_ind_netz', chart_height + nrows24 + 2, 2, chart_height + nrows24 + 2, ncols25 - 1],
            'values':     [economy + '_FED_ind_netz', chart_height + nrows24 + j + 3, 2, chart_height + nrows24 + j + 3, ncols25 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True}
        })
    
    netz_worksheet4.insert_chart('J3', netz_fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    netz_worksheet5 = writer.sheets[economy + '_FED_trn_netz']
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet5.set_column(2, ncols30 + 1, None, comma_format)
    netz_worksheet5.set_row(chart_height, None, header_format)
    netz_worksheet5.set_row(chart_height + nrows30 + 3, None, header_format)
    netz_worksheet5.write(0, 0, economy + ' FED transport', cell_format1)
    
    # Create a transport FED area chart
    netz_transport_chart1 = workbook.add_chart({'type': 'area', 
                                           'subtype': 'stacked'})
    netz_transport_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_transport_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_transport_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_transport_chart1.set_y_axis({
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
        
    netz_transport_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_transport_chart1.set_title({
        'none': True
    })
        
    for fuel_agg in Transport_fuels_agg:
        j = netz_transport_df1[netz_transport_df1['fuel_code'] == fuel_agg].index[0]
        netz_transport_chart1.add_series({
            'name':       [economy + '_FED_trn_netz', chart_height + j + 1, 0],
            'categories': [economy + '_FED_trn_netz', chart_height, 2, chart_height, ncols30 - 1],
            'values':     [economy + '_FED_trn_netz', chart_height + j + 1, 2, chart_height + j + 1, ncols30 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True} 
        })
    
    netz_worksheet5.insert_chart('B3', netz_transport_chart1)
            
    ############# FED transport chart 2 (transport by modality)
    
    # Create a FED transport modality column chart
    netz_transport_chart2 = workbook.add_chart({'type': 'column', 
                                         'subtype': 'stacked'})
    netz_transport_chart2.set_size({
        'width': 500,
        'height': 300
    })

    netz_transport_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_transport_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
        'line': {'color': '#bebebe'}
    })
        
    netz_transport_chart2.set_y_axis({
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
        
    netz_transport_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_transport_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.    
    for modality in Transport_modal_agg:
        j = netz_transport_df2[netz_transport_df2['item_code_new'] == modality].index[0]
        netz_transport_chart2.add_series({
            'name':       [economy + '_FED_trn_netz', chart_height + nrows30 + j + 4, 1],
            'categories': [economy + '_FED_trn_netz', chart_height + nrows30 + 3, 2, chart_height + nrows30 + 3, ncols31 - 1],
            'values':     [economy + '_FED_trn_netz', chart_height + nrows30 + j + 4, 2, chart_height + nrows30 + j + 4, ncols31 - 1],
            'fill':       {'color': colours_hex[j]},
            'border':     {'none': True}
        })
    
    netz_worksheet5.insert_chart('J3', netz_transport_chart2)

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    netz_worksheet6 = writer.sheets[economy + '_FED_agr_netz']
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet6.set_column(2, ncols34 + 1, None, comma_format)
    netz_worksheet6.set_row(chart_height, None, header_format)
    netz_worksheet6.set_row(chart_height + nrows34 + 3, None, header_format)
    netz_worksheet6.write(0, 0, economy + ' FED agriculture', cell_format1)

    # Create a Agriculture line chart 
    netz_ag_chart1 = workbook.add_chart({'type': 'line'})
    netz_ag_chart1.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_ag_chart1.set_chartarea({
        'border': {'none': True}
    })
    
    netz_ag_chart1.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_ag_chart1.set_y_axis({
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
        
    netz_ag_chart1.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_ag_chart1.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows34):
        netz_ag_chart1.add_series({
            'name':       [economy + '_FED_agr_netz', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agr_netz', chart_height, 2, chart_height, ncols34 - 1],
            'values':     [economy + '_FED_agr_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols34 - 1],
            'line':       {'color': colours_hex[i], 'width': 1.25}
        })    
        
    netz_worksheet6.insert_chart('B3', netz_ag_chart1)

    # Create a Agriculture area chart
    netz_ag_chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
    netz_ag_chart2.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_ag_chart2.set_chartarea({
        'border': {'none': True}
    })
    
    netz_ag_chart2.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'position_axis': 'on_tick',
        'interval_unit': 4,
        'line': {'color': '#bebebe'}
    })
        
    netz_ag_chart2.set_y_axis({
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
        
    netz_ag_chart2.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_ag_chart2.set_title({
        'none': True
    })
    
    # Configure the series of the chart from the dataframe data.
    for i in range(nrows34):
        netz_ag_chart2.add_series({
            'name':       [economy + '_FED_agr_netz', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agr_netz', chart_height, 2, chart_height, ncols34 - 1],
            'values':     [economy + '_FED_agr_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols34 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })    
        
    netz_worksheet6.insert_chart('J3', netz_ag_chart2)

    # Create a Agriculture stacked column
    netz_ag_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
    netz_ag_chart3.set_size({
        'width': 500,
        'height': 300
    })
    
    netz_ag_chart3.set_chartarea({
        'border': {'none': True}
    })
    
    netz_ag_chart3.set_x_axis({
        'name': 'Year',
        'label_position': 'low',
        'major_tick_mark': 'none',
        'minor_tick_mark': 'none',
        'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
        'interval_unit': 1,
        'line': {'color': '#bebebe'}
    })
        
    netz_ag_chart3.set_y_axis({
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
        
    netz_ag_chart3.set_legend({
        'font': {'font': 'Segoe UI', 'size': 10}
        #'none': True
    })
        
    netz_ag_chart3.set_title({
        'none': True
    })

    # Configure the series of the chart from the dataframe data.    
    for i in range(nrows35):
        netz_ag_chart3.add_series({
            'name':       [economy + '_FED_agr_netz', chart_height + nrows34 + i + 4, 0],
            'categories': [economy + '_FED_agr_netz', chart_height + nrows34 + 3, 2, chart_height + nrows34 + 3, ncols35 - 1],
            'values':     [economy + '_FED_agr_netz', chart_height + nrows34 + i + 4, 2, chart_height + nrows34 + i + 4, ncols35 - 1],
            'fill':       {'color': colours_hex[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet6.insert_chart('R3', netz_ag_chart3)
        
    writer.save()

print('Bling blang blaow, you have some FED charts now')
