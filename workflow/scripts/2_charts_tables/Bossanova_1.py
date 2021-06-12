# An attempt to move all portion of FED, Supply and Transformation into one script

# import dependencies

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

ref_power_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_power_reference.csv').loc[:,:'2050']
ref_refownsup_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_refownsup_reference.csv').loc[:,:'2050']
ref_pow_capacity_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_powcapacity_reference.csv').loc[:,:'2050']
ref_trans_df1 = = pd.read_csv('./data/4_Joined/OSeMOSYS_transformation_reference.csv').loc[:,:'2050']

netz_power_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_power_netzero.csv').loc[:,:'2050']
netz_refownsup_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_refownsup_netzero.csv').loc[:,:'2050']
netz_pow_capacity_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_powcapacity_netzero.csv').loc[:,:'2050']
netz_trans_df1 = = pd.read_csv('./data/4_Joined/OSeMOSYS_transformation_netzero.csv').loc[:,:'2050']

# Define unique values for economy, fuels, and items columns
# only looking at one dataframe which should be sufficient as both have same structure

Economy_codes = EGEDA_years_reference.economy.unique()
Fuels = EGEDA_years_reference.fuel_code.unique()
Items = EGEDA_years_reference.item_code_new.unique()

# Define colour palette

colours_dict = pd.read_csv('./data/2_Mapping_and_other/colours_dict.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

# FED: Subsets for impending df builds

# Fuels

First_level_fuels = ['1_coal', '2_coal_products', '5_oil_shale_and_oil_sands', '6_crude_oil_and_ngl', '7_petroleum_products',
                     '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass',
                     '16_others', '17_electricity', '18_heat', '19_total', '20_total_renewables', '21_modern_renewables']

Required_fuels = ['1_coal', '2_coal_products', '5_oil_shale_and_oil_sands', '6_crude_oil_and_ngl', '7_petroleum_products',
                  '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass',
                  '16_1_biogas', '16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable',
                  '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_9_other_sources',
                  '16_x_hydrogen', '17_electricity', '18_heat', '19_total', '20_total_renewables', '21_modern_renewables']

required_fuels_elec = ['1_coal', '1_5_lignite', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products', 
                       '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', 
                       '15_solid_biomass', '16_others', '18_heat']                  

Coal_fuels = ['1_coal', '2_coal_products', '3_peat', '4_peat_products']

Oil_fuels = ['6_crude_oil_and_ngl', '7_petroleum_products', '5_oil_shale_and_oil_sands']

Other_fuels_FED = ['9_nuclear', '16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable']

Other_fuels_TPES = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources']

Other_fuels_industry = ['9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '16_1_biogas',
                         '16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable', 
                         '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels']

Renewables_fuels = ['10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', '16_1_biogas', 
                    '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                    '16_8_other_liquid_biofuels']

Renewables_fuels_nobiomass = ['10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '16_1_biogas', 
                          '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                          '16_8_other_liquid_biofuels']

Petroleum_fuels = ['7_petroleum_products', '7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha', '7_4_gasoline_type_jet_fuel',
                   '7_5_kerosene_type_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil', '7_8_fuel_oil', '7_9_lpg',
                   '7_10_refinery_gas_not_liquefied', '7_11_ethane', '7_x_other_petroleum_products', '7_12_white_spirit_sbp',
                   '7_13_lubricants', '7_14_bitumen', '7_15_paraffin_waxes', '7_16_petroleum_coke', '7_17_other_products']

### Transport fuel vectors

Transport_fuels = ['1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products', '7_1_motor_gasoline', '7_2_aviation_gasoline',
                   '7_x_jet_fuel', '7_7_gas_diesel_oil', '7_8_fuel_oil', '7_9_lpg',
                   '7_x_other_petroleum_products', '8_1_natural_gas', '16_5_biogasoline', '16_6_biodiesel',
                   '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_9_other_sources', '17_electricity'] 

Renew_fuel = ['16_5_biogasoline', '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels']

Other_fuel_trans = ['7_8_fuel_oil', '1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products', '7_x_other_petroleum_products']

# Sectors

trad_bio_sectors = ['16_1_commercial_and_public_services', '16_2_residential',
                  '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others']

no_trad_bio_sectors = ['14_industry_sector', '15_transport_sector', '17_nonenergy_use']

Sectors_tfc = ['14_industry_sector', '15_transport_sector', '16_1_commercial_and_public_services', '16_2_residential',
               '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others', '17_nonenergy_use']

Buildings_items = ['16_1_commercial_and_public_services', '16_2_residential']

Ag_items = ['16_3_agriculture', '16_4_fishing']

Subindustry = ['14_industry_sector', '14_1_iron_and_steel', '14_2_chemical_incl_petrochemical', '14_3_non_ferrous_metals',
               '14_4_nonmetallic_mineral_products', '14_5_transportation_equipment', '14_6_machinery', '14_7_mining_and_quarrying',
               '14_8_food_beverages_and_tobacco', '14_9_pulp_paper_and_printing', '14_10_wood_and_wood_products', 
               '14_11_construction', '14_12_textiles_and_leather', '14_13_nonspecified_industry']

Other_industry = ['14_5_transportation_equipment', '14_6_machinery', '14_8_food_beverages_and_tobacco', '14_10_wood_and_wood_products',
                  '14_11_construction', '14_12_textiles_and_leather']

Transport_modal = ['15_1_domestic_air_transport', '15_2_road', '15_3_rail', '15_4_domestic_navigation', '15_5_pipeline_transport',
                   '15_6_nonspecified_transport']

tpes_items = ['1_indigenous_production', '2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers',
              '6_stock_change', '7_total_primary_energy_supply']

Prod_items = tpes_items[:1]

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written (can change this)

# Define column chart years
col_chart_years = ['2000', '2010', '2018', '2020', '2030', '2040', '2050']

# Define column chart years for transport
col_chart_years_transport = ['2018', '2020', '2030', '2040', '2050']

# FED aggregate fuels

FED_agg_fuels = ['Coal', 'Oil', 'Gas', 'Other renewables', 'Biomass', 'Hydrogen', 'Electricity', 'Heat', 'Others']
FED_agg_fuels_ind = ['Coal', 'Oil', 'Gas', 'Biomass', 'Hydrogen', 'Electricity', 'Heat', 'Others']
Transport_fuels_agg = ['Diesel', 'Gasoline', 'LPG', 'Gas', 'Jet fuel', 'Electricity', 'Renewables', 'Hydrogen', 'Other']

# FED aggregate sectors

FED_agg_sectors = ['Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-energy', 'Non-specified']
Industry_eight = ['Iron & steel', 'Chemicals', 'Aluminium', 'Non-metallic minerals', 'Mining', 'Pulp & paper', 'Other', 'Non-specified']
Transport_modal_agg = ['Aviation', 'Road', 'Rail' ,'Marine', 'Pipeline', 'Non-specified']

# TPES

TPES_agg_fuels = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']
TPES_agg_trade = ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']
avi_bunker = ['Aviation gasoline', 'Jet fuel']

########################### Create historical electricity generation dataframe for use later ###########################

EGEDA_hist_gen = pd.read_csv('./data/1_EGEDA/EGEDA_2018_years.csv', 
                             names = ['economy', 'fuel_code', 'item_code_new'] + list(range(1980, 2019)),
                             header = 0)

EGEDA_hist_gen = EGEDA_hist_gen[(EGEDA_hist_gen['item_code_new'] == '18_electricity_output_in_pj') & 
                                (EGEDA_hist_gen['fuel_code'].isin(required_fuels_elec))].reset_index(drop = True)

# China only having data for 1_coal requires workaround to keep lignite data
lignite_alt = EGEDA_hist_gen[EGEDA_hist_gen['fuel_code'] == '1_5_lignite'].copy()\
    .set_index(['economy', 'fuel_code', 'item_code_new']) * -1

lignite_alt = lignite_alt.reset_index()

new_coal = EGEDA_hist_gen[EGEDA_hist_gen['fuel_code'] == '1_coal'].copy().reset_index(drop = True)

lig_coal = new_coal.append(lignite_alt).reset_index(drop = True).groupby(['economy', 'item_code_new']).sum().reset_index()
lig_coal['fuel_code'] = '1_coal'

no_coal = EGEDA_hist_gen[EGEDA_hist_gen['fuel_code'] != '1_coal'].copy().reset_index(drop = True)

EGEDA_hist_gen = no_coal.append(lig_coal).reset_index(drop = True)

EGEDA_hist_gen['TECHNOLOGY'] = EGEDA_hist_gen['fuel_code'].map({'1_coal': 'Coal', 
                                                                '1_5_lignite': 'Lignite', 
                                                                '2_coal_products': 'Coal',
                                                                '6_crude_oil_and_ngl': 'Oil',
                                                                '7_petroleum_products': 'Oil',
                                                                '8_gas': 'Gas', 
                                                                '9_nuclear': 'Nuclear', 
                                                                '10_hydro': 'Hydro', 
                                                                '11_geothermal': 'Geothermal', 
                                                                '12_solar': 'Solar', 
                                                                '13_tide_wave_ocean': 'Other', 
                                                                '14_wind': 'Wind', 
                                                                '15_solid_biomass': 'Biomass', 
                                                                '16_others': 'Other', 
                                                                '18_heat': 'Other'})

EGEDA_hist_gen['Generation'] = 'Electricity'

EGEDA_hist_gen = EGEDA_hist_gen[['economy', 'TECHNOLOGY', 'Generation'] + list(range(2000, 2019))].\
    groupby(['economy', 'TECHNOLOGY', 'Generation']).sum().reset_index()

#########################################################################################################################################

# Now build the subset dataframes for charts and tables

for economy in Economy_codes:
    ################################################################### DATAFRAMES ###################################################################
    # FED REFERENCE DATA FRAMES
    # First data frame construction: FED by fuels
    ref_notrad_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(no_trad_bio_sectors)) &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    ref_notrad_1 = ref_notrad_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Industry, transport, NE').reset_index()

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = ref_notrad_1[ref_notrad_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Industry, transport, NE')
    
    oil = ref_notrad_1[ref_notrad_1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Industry, transport, NE')
    
    renewables = ref_notrad_1[ref_notrad_1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = 'Industry, transport, NE')
    
    others = ref_notrad_1[ref_notrad_1['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Industry, transport, NE')

    # Fed fuel data frame 1

    ref_fedfuel_1 = ref_notrad_1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_notrad_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    # Insert 0 traditional biomass row
    new_row = ['Biomass', 'Industry, transport, NE'] + [0] * 51
    new_series = pd.Series(new_row, index = ref_fedfuel_1.columns)

    ref_fedfuel_1 = ref_fedfuel_1.append(new_series, ignore_index = True)

    ref_fedfuel_1 = ref_fedfuel_1[ref_fedfuel_1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ##### No biomass fix for dataframe

    ref_tradbio_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(trad_bio_sectors)) &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)

    ref_tradbio_1 = ref_tradbio_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Trad bio sectors').reset_index()

    # build aggregate with altered vector to account for no biomass in renewables
    coal_tradbio = ref_tradbio_1[ref_tradbio_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Coal', item_code_new = 'Trad bio sectors')

    oil_tradbio = ref_tradbio_1[ref_tradbio_1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Oil', item_code_new = 'Trad bio sectors')

    renew_tradbio = ref_tradbio_1[ref_tradbio_1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Other renewables', item_code_new = 'Trad bio sectors')

    others_tradbio = ref_tradbio_1[ref_tradbio_1['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Others', item_code_new = 'Trad bio sectors')

    # Fed fuel no biomass in other sector renewables
    ref_tradbio_2 = ref_tradbio_1.append([coal_tradbio, oil_tradbio, renew_tradbio, others_tradbio])\
        [['fuel_code', 'item_code_new'] + list(ref_tradbio_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_tradbio_2 = ref_tradbio_2[ref_tradbio_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ref_fedfuel_1 = ref_fedfuel_1.append(ref_tradbio_2)

    # Combine the two dataframes that account for Modern renewables
    ref_fedfuel_1 = ref_fedfuel_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = '12_total_final_consumption')\
        .reset_index()[['fuel_code', 'item_code_new'] + list(ref_fedfuel_1.loc[:,'2000':'2050'])]\
            .set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ref_fedfuel_1_rows = ref_fedfuel_1.shape[0]
    ref_fedfuel_1_cols = ref_fedfuel_1.shape[1]

    ref_fedfuel_2 = ref_fedfuel_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_fedfuel_2_rows = ref_fedfuel_2.shape[0]
    ref_fedfuel_2_cols = ref_fedfuel_2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    ref_fedsector_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                        (EGEDA_years_reference['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years_reference['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    ref_fedsector_1 = ref_fedsector_1[['fuel_code', 'item_code_new'] + list(ref_fedsector_1.loc[:,'2000':])]
    
    ref_fedsector_1_rows = ref_fedsector_1.shape[0]
    ref_fedsector_1_cols = ref_fedsector_1.shape[1]

    # Now build aggregate sector variables
    
    buildings = ref_fedsector_1[ref_fedsector_1['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = ref_fedsector_1[ref_fedsector_1['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')
    
    # Build aggregate data frame of FED sector

    ref_fedsector_2 = ref_fedsector_1.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(ref_fedsector_1.loc[:, '2000':])].reset_index(drop = True)

    ref_fedsector_2.loc[ref_fedsector_2['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_fedsector_2.loc[ref_fedsector_2['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_fedsector_2.loc[ref_fedsector_2['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_fedsector_2.loc[ref_fedsector_2['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_fedsector_2 = ref_fedsector_2[ref_fedsector_2['item_code_new'].isin(FED_agg_sectors)].set_index('item_code_new').loc[FED_agg_sectors].reset_index()
    ref_fedsector_2 = ref_fedsector_2[['fuel_code', 'item_code_new'] + list(ref_fedsector_2.loc[:, '2000':])]

    ref_fedsector_2_rows = ref_fedsector_2.shape[0]
    ref_fedsector_2_cols = ref_fedsector_2.shape[1]

    ref_fedsector_3 = ref_fedsector_2[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_fedsector_3_rows = ref_fedsector_3.shape[0]
    ref_fedsector_3_cols = ref_fedsector_3.shape[1]

    # New FED by sector (not including non-energy)

    ref_tfec_1 = ref_fedsector_2[ref_fedsector_2['item_code_new'] != 'Non-energy'].copy().groupby(['fuel_code'])\
        .sum().assign(item_code_new = 'TFEC', fuel_code = 'Total').reset_index(drop = True)

    ref_tfec_1 = ref_tfec_1[['fuel_code', 'item_code_new'] + list(ref_tfec_1.loc[:, '2000':'2050'])]

    ref_tfec_1_rows = ref_tfec_1.shape[0]
    ref_tfec_1_cols = ref_tfec_1.shape[1] 
    
    # Third data frame construction: Buildings FED by fuel
    ref_bld_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                         (EGEDA_years_reference['item_code_new'].isin(Buildings_items)) &
                         (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]
    
    for fuel in Required_fuels:
        buildings = ref_bld_1[ref_bld_1['fuel_code'] == fuel].groupby(['economy', 'fuel_code']).sum().assign(item_code_new = '16_x_buildings')
        buildings['economy'] = economy
        buildings['fuel_code'] = fuel
        
        ref_bld_1 = ref_bld_1.append(buildings).reset_index(drop = True)
        
    ref_bld_1 = ref_bld_1[['fuel_code', 'item_code_new'] + list(ref_bld_1.loc[:, '2000':])]

    # Create data fram with commercial and residential aggregated together 

    ref_bld_2 = ref_bld_1[ref_bld_1['item_code_new'] == '16_x_buildings']

    coal = ref_bld_2[ref_bld_2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = '16_x_buildings')
    
    oil = ref_bld_2[ref_bld_2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = '16_x_buildings')
    
    renewables = ref_bld_2[ref_bld_2['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = '16_x_buildings')
    
    others = ref_bld_2[ref_bld_2['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = '16_x_buildings')

    ref_bld_2 = ref_bld_2.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_bld_2.loc[:, '2000':])].reset_index(drop = True)

    ref_bld_2.loc[ref_bld_2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_bld_2 = ref_bld_2[ref_bld_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code')\
        .loc[FED_agg_fuels].reset_index()

    ref_bld_2_rows = ref_bld_2.shape[0]
    ref_bld_2_cols = ref_bld_2.shape[1]

    ref_bld_3 = ref_bld_1[(ref_bld_1['fuel_code'] == '19_total') &
                      (ref_bld_1['item_code_new'].isin(Buildings_items))].copy().reset_index(drop = True)

    ref_bld_3.loc[ref_bld_3['item_code_new'] == '16_1_commercial_and_public_services', 'item_code_new'] = 'Services' 
    ref_bld_3.loc[ref_bld_3['item_code_new'] == '16_2_residential', 'item_code_new'] = 'Residential'

    ref_bld_3_rows = ref_bld_3.shape[0]
    ref_bld_3_cols = ref_bld_3.shape[1]
    
    # Fourth data frame: Industry subsector
    ref_ind_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                         (EGEDA_years_reference['item_code_new'].isin(Subindustry)) &
                         (EGEDA_years_reference['fuel_code'] == '19_total')]

    other_industry = ref_ind_1[ref_ind_1['item_code_new'].isin(Other_industry)].groupby(['fuel_code']).sum().assign(item_code_new = 'Other',
                                                                                                                fuel_code = '19_total')

    ref_ind_1 = ref_ind_1.append([other_industry])[['fuel_code', 'item_code_new'] + list(ref_ind_1.loc[:, '2000':])].reset_index(drop = True)

    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_1_iron_and_steel', 'item_code_new'] = 'Iron & steel'
    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_2_chemical_incl_petrochemical', 'item_code_new'] = 'Chemicals'
    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_3_non_ferrous_metals', 'item_code_new'] = 'Aluminium'
    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_4_nonmetallic_mineral_products', 'item_code_new'] = 'Non-metallic minerals'  
    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_7_mining_and_quarrying', 'item_code_new'] = 'Mining'
    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_9_pulp_paper_and_printing', 'item_code_new'] = 'Pulp & paper'
    ref_ind_1.loc[ref_ind_1['item_code_new'] == '14_13_nonspecified_industry', 'item_code_new'] = 'Non-specified'
    
    ref_ind_1 = ref_ind_1[ref_ind_1['item_code_new'].isin(Industry_eight)].set_index('item_code_new').loc[Industry_eight].reset_index()

    ref_ind_1 = ref_ind_1[['fuel_code', 'item_code_new'] + list(ref_ind_1.loc[:, '2000':])]

    ref_ind_1_rows = ref_ind_1.shape[0]
    ref_ind_1_cols = ref_ind_1.shape[1]
    
    # Fifth data frame construction: Industry by fuel
    ref_ind_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                         (EGEDA_years_reference['item_code_new'].isin(['14_industry_sector'])) &
                         (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]
    
    coal = ref_ind_2[ref_ind_2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal', 
                                                                                                  item_code_new = '14_industry_sector')
    
    oil = ref_ind_2[ref_ind_2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil', 
                                                                                                item_code_new = '14_industry_sector')
    
    biomass = ref_ind_2[ref_ind_2['fuel_code'].isin(['15_solid_biomass'])].groupby(['item_code_new']).sum().assign(fuel_code = 'Biomass', 
                                                                                                              item_code_new = '14_industry_sector')
    
    others = ref_ind_2[ref_ind_2['fuel_code'].isin(Other_fuels_industry)].groupby(['item_code_new']).sum().assign(fuel_code = 'Others', 
                                                                                                                item_code_new = '14_industry_sector')
    
    ref_ind_2 = ref_ind_2.append([coal, oil, biomass, others])\
        [['fuel_code', 'item_code_new'] + list(ref_ind_2.loc[:, '2000':])].reset_index(drop = True)

    ref_ind_2.loc[ref_ind_2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_ind_2.loc[ref_ind_2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_ind_2.loc[ref_ind_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_ind_2.loc[ref_ind_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_ind_2 = ref_ind_2[ref_ind_2['fuel_code'].isin(FED_agg_fuels_ind)].set_index('fuel_code').loc[FED_agg_fuels_ind].reset_index()
    
    ref_ind_2_rows = ref_ind_2.shape[0]
    ref_ind_2_cols = ref_ind_2.shape[1]

    # Transport data frame construction: FED by fuels
    ref_trans_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                          (EGEDA_years_reference['fuel_code'].isin(Transport_fuels))]
    
    renewables = ref_trans_1[ref_trans_1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '15_transport_sector')
    
    others = ref_trans_1[ref_trans_1['fuel_code'].isin(Other_fuel_trans)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '15_transport_sector')

    trans_gasoline = ref_trans_1[ref_trans_1['fuel_code'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Gasoline', 
                                                            item_code_new = '15_transport_sector')

    trans_jetfuel = ref_trans_1[ref_trans_1['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Jet fuel', 
                                                            item_code_new = '15_transport_sector')
    
    ref_trans_1 = ref_trans_1.append([renewables, trans_gasoline, trans_jetfuel, others])[['fuel_code', 'item_code_new'] + list(ref_trans_1.loc[:, '2000':])].reset_index(drop = True) 

    ref_trans_1.loc[ref_trans_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    ref_trans_1.loc[ref_trans_1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    ref_trans_1.loc[ref_trans_1['fuel_code'] == '7_9_lpg', 'fuel_code'] = 'LPG'
    ref_trans_1.loc[ref_trans_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_trans_1.loc[ref_trans_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_trans_1 = ref_trans_1[ref_trans_1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    ref_trans_1_rows = ref_trans_1.shape[0]
    ref_trans_1_cols = ref_trans_1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    ref_trans_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                               (EGEDA_years_reference['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years_reference['fuel_code'].isin(['19_total']))].copy()
    
    ref_trans_2.loc[ref_trans_2['item_code_new'] == '15_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    ref_trans_2.loc[ref_trans_2['item_code_new'] == '15_2_road', 'item_code_new'] = 'Road'
    ref_trans_2.loc[ref_trans_2['item_code_new'] == '15_3_rail', 'item_code_new'] = 'Rail'
    ref_trans_2.loc[ref_trans_2['item_code_new'] == '15_4_domestic_navigation', 'item_code_new'] = 'Marine'
    ref_trans_2.loc[ref_trans_2['item_code_new'] == '15_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    ref_trans_2.loc[ref_trans_2['item_code_new'] == '15_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    ref_trans_2 = ref_trans_2[ref_trans_2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    ref_trans_2 = ref_trans_2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    ref_trans_2_rows = ref_trans_2.shape[0]
    ref_trans_2_cols = ref_trans_2.shape[1]

    # Agriculture data frame

    ref_ag_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                         (EGEDA_years_reference['item_code_new'].isin(Ag_items)) &
                         (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].groupby('fuel_code').sum().assign(item_code_new = 'Agriculture').reset_index()
                     
    coal = ref_ag_1[ref_ag_1['fuel_code'].isin(Coal_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Agriculture')

    oil = ref_ag_1[ref_ag_1['fuel_code'].isin(Oil_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Agriculture')

    renewables = ref_ag_1[ref_ag_1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = 'Agriculture')
    
    others = ref_ag_1[ref_ag_1['fuel_code'].isin(Other_fuels_FED)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Agriculture')
    
    ref_ag_1 = ref_ag_1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_ag_1.loc[:,'2000':'2050'])].reset_index(drop = True)

    ref_ag_1.loc[ref_ag_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_ag_1 = ref_ag_1[ref_ag_1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()
    
    ref_ag_1_rows = ref_ag_1.shape[0]
    ref_ag_1_cols = ref_ag_1.shape[1]

    ref_ag_2 = ref_ag_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_ag_2_rows = ref_ag_2.shape[0]
    ref_ag_2_cols = ref_ag_2.shape[1]

    # Hydrogen data frame reference

    ref_hyd_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                                        (EGEDA_years_reference['item_code_new'].isin(Sectors_tfc)) &
                                        (EGEDA_years_reference['fuel_code'] == '16_9_other_sources')].groupby('item_code_new').sum().assign(fuel_code = 'Hydrogen').reset_index()

    buildings_hy = ref_hyd_1[ref_hyd_1['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Buildings', fuel_code = 'Hydrogen')

    ag_hy = ref_hyd_1[ref_hyd_1['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Agriculture', fuel_code = 'Hydrogen')

    ref_hyd_1 = ref_hyd_1.append([buildings_hy, ag_hy])\
        [['fuel_code', 'item_code_new'] + list(ref_hyd_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_hyd_1.loc[ref_hyd_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_hyd_1.loc[ref_hyd_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'

    ref_hyd_1 = ref_hyd_1[ref_hyd_1['item_code_new'].isin(['Agriculture', 'Buildings', 'Industry', 'Transport'])]\
        .copy().reset_index(drop = True)

    ref_hyd_1_rows = ref_hyd_1.shape[0]
    ref_hyd_1_cols = ref_hyd_1.shape[1]

    ###############################################################################################################

    # NET ZERO DATA FRAMES
    # First data frame construction: FED by fuels
    netz_notrad_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'].isin(no_trad_bio_sectors)) &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)
    
    netz_notrad_1 = netz_notrad_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Industry, transport, NE').reset_index()

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = netz_notrad_1[netz_notrad_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Industry, transport, NE')
    
    oil = netz_notrad_1[netz_notrad_1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Industry, transport, NE')
    
    renewables = netz_notrad_1[netz_notrad_1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = 'Industry, transport, NE')
    
    others = netz_notrad_1[netz_notrad_1['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Industry, transport, NE')

    # Fed fuel data frame 1 (data frame 6)

    netz_fedfuel_1 = netz_notrad_1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_notrad_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    # Insert 0 traditional biomass row
    new_row = ['Biomass', 'Industry, transport, NE'] + [0] * 51
    new_series = pd.Series(new_row, index = netz_fedfuel_1.columns)

    netz_fedfuel_1 = netz_fedfuel_1.append(new_series, ignore_index = True)

    netz_fedfuel_1 = netz_fedfuel_1[netz_fedfuel_1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ##### No biomass fix for dataframe

    netz_tradbio_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                           (EGEDA_years_netzero['item_code_new'].isin(trad_bio_sectors)) &
                                           (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':].reset_index(drop = True)

    netz_tradbio_1 = netz_tradbio_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = 'Trad bio sectors').reset_index()

    # build aggregate with altered vector to account for no biomass in renewables
    coal_tradbio = netz_tradbio_1[netz_tradbio_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Coal', item_code_new = 'Trad bio sectors')

    oil_tradbio = netz_tradbio_1[netz_tradbio_1['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Oil', item_code_new = 'Trad bio sectors')

    renew_tradbio = netz_tradbio_1[netz_tradbio_1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Other renewables', item_code_new = 'Trad bio sectors')

    others_tradbio = netz_tradbio_1[netz_tradbio_1['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new']).\
        sum().assign(fuel_code = 'Others', item_code_new = 'Trad bio sectors')

    # Fed fuel no biomass in other sector renewables
    netz_tradbio_2 = netz_tradbio_1.append([coal_tradbio, oil_tradbio, renew_tradbio, others_tradbio])\
        [['fuel_code', 'item_code_new'] + list(netz_tradbio_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_tradbio_2 = netz_tradbio_2[netz_tradbio_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    netz_fedfuel_1 = netz_fedfuel_1.append(netz_tradbio_2)

    # Combine the two dataframes that account for Modern renewables
    netz_fedfuel_1 = netz_fedfuel_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = '12_total_final_consumption')\
        .reset_index()[['fuel_code', 'item_code_new'] + list(netz_fedfuel_1.loc[:,'2000':'2050'])]\
            .set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    netz_fedfuel_1_rows = netz_fedfuel_1.shape[0]
    netz_fedfuel_1_cols = netz_fedfuel_1.shape[1]

    netz_fedfuel_2 = netz_fedfuel_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_fedfuel_2_rows = netz_fedfuel_2.shape[0]
    netz_fedfuel_2_cols = netz_fedfuel_2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    netz_fedsector_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                        (EGEDA_years_netzero['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years_netzero['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    netz_fedsector_1 = netz_fedsector_1[['fuel_code', 'item_code_new'] + list(netz_fedsector_1.loc[:,'2000':])]
    
    netz_fedsector_1_rows = netz_fedsector_1.shape[0]
    netz_fedsector_1_cols = netz_fedsector_1.shape[1]

    # Now build aggregate sector variables
    
    buildings = netz_fedsector_1[netz_fedsector_1['item_code_new'].isin(Buildings_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = netz_fedsector_1[netz_fedsector_1['item_code_new'].isin(Ag_items)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')
    
    # Build aggregate data frame of FED sector

    netz_fedsector_2 = netz_fedsector_1.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(netz_fedsector_1.loc[:, '2000':])].reset_index(drop = True)

    netz_fedsector_2.loc[netz_fedsector_2['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_fedsector_2.loc[netz_fedsector_2['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_fedsector_2.loc[netz_fedsector_2['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_fedsector_2.loc[netz_fedsector_2['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_fedsector_2 = netz_fedsector_2[netz_fedsector_2['item_code_new'].isin(FED_agg_sectors)].set_index('item_code_new').loc[FED_agg_sectors].reset_index()
    netz_fedsector_2 = netz_fedsector_2[['fuel_code', 'item_code_new'] + list(netz_fedsector_2.loc[:, '2000':])]

    netz_fedsector_2_rows = netz_fedsector_2.shape[0]
    netz_fedsector_2_cols = netz_fedsector_2.shape[1]

    netz_fedsector_3 = netz_fedsector_2[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_fedsector_3_rows = netz_fedsector_3.shape[0]
    netz_fedsector_3_cols = netz_fedsector_3.shape[1]

    # New FED by sector (not including non-energy)

    netz_tfec_1 = netz_fedsector_2[netz_fedsector_2['item_code_new'] != 'Non-energy'].copy().groupby(['fuel_code'])\
        .sum().assign(item_code_new = 'TFEC', fuel_code = 'Total').reset_index(drop = True)

    netz_tfec_1 = netz_tfec_1[['fuel_code', 'item_code_new'] + list(netz_tfec_1.loc[:, '2000':'2050'])]

    netz_tfec_1_rows = netz_tfec_1.shape[0]
    netz_tfec_1_cols = netz_tfec_1.shape[1] 
    
    # Third data frame construction: Buildings FED by fuel
    netz_bld_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                         (EGEDA_years_netzero['item_code_new'].isin(Buildings_items)) &
                         (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))]
    
    for fuel in Required_fuels:
        buildings = netz_bld_1[netz_bld_1['fuel_code'] == fuel].groupby(['economy', 'fuel_code']).sum().assign(item_code_new = '16_x_buildings')
        buildings['economy'] = economy
        buildings['fuel_code'] = fuel
        
        netz_bld_1 = netz_bld_1.append(buildings).reset_index(drop = True)
        
    netz_bld_1 = netz_bld_1[['fuel_code', 'item_code_new'] + list(netz_bld_1.loc[:, '2000':])]

    netz_bld_2 = netz_bld_1[netz_bld_1['item_code_new'] == '16_x_buildings']

    coal = netz_bld_2[netz_bld_2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Coal', item_code_new = '16_x_buildings')
    
    oil = netz_bld_2[netz_bld_2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Oil', item_code_new = '16_x_buildings')
    
    renewables = netz_bld_2[netz_bld_2['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = '16_x_buildings')
    
    others = netz_bld_2[netz_bld_2['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = '16_x_buildings')

    netz_bld_2 = netz_bld_2.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_bld_2.loc[:, '2000':])].reset_index(drop = True)

    netz_bld_2.loc[netz_bld_2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_bld_2 = netz_bld_2[netz_bld_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code')\
        .loc[FED_agg_fuels].reset_index()
    netz_bld_2_rows = netz_bld_2.shape[0]
    netz_bld_2_cols = netz_bld_2.shape[1]

    netz_bld_3 = netz_bld_1[(netz_bld_1['fuel_code'] == '19_total') &
                      (netz_bld_1['item_code_new'].isin(Buildings_items))].copy().reset_index(drop = True)

    netz_bld_3.loc[netz_bld_3['item_code_new'] == '16_1_commercial_and_public_services', 'item_code_new'] = 'Services' 
    netz_bld_3.loc[netz_bld_3['item_code_new'] == '16_2_residential', 'item_code_new'] = 'Residential'

    netz_bld_3_rows = netz_bld_3.shape[0]
    netz_bld_3_cols = netz_bld_3.shape[1]
    
    # Fourth data frame construction: Industry subsector
    netz_ind_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                         (EGEDA_years_netzero['item_code_new'].isin(Subindustry)) &
                         (EGEDA_years_netzero['fuel_code'] == '19_total')]

    other_industry = netz_ind_1[netz_ind_1['item_code_new'].isin(Other_industry)].groupby(['fuel_code']).sum().assign(item_code_new = 'Other',
                                                                                                                fuel_code = '19_total')

    netz_ind_1 = netz_ind_1.append([other_industry])[['fuel_code', 'item_code_new'] + list(netz_ind_1.loc[:, '2000':])].reset_index(drop = True)

    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_1_iron_and_steel', 'item_code_new'] = 'Iron & steel'
    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_2_chemical_incl_petrochemical', 'item_code_new'] = 'Chemicals'
    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_3_non_ferrous_metals', 'item_code_new'] = 'Aluminium'
    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_4_nonmetallic_mineral_products', 'item_code_new'] = 'Non-metallic minerals'  
    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_7_mining_and_quarrying', 'item_code_new'] = 'Mining'
    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_9_pulp_paper_and_printing', 'item_code_new'] = 'Pulp & paper'
    netz_ind_1.loc[netz_ind_1['item_code_new'] == '14_13_nonspecified_industry', 'item_code_new'] = 'Non-specified'
    
    netz_ind_1 = netz_ind_1[netz_ind_1['item_code_new'].isin(Industry_eight)].set_index('item_code_new').loc[Industry_eight].reset_index()

    netz_ind_1 = netz_ind_1[['fuel_code', 'item_code_new'] + list(netz_ind_1.loc[:, '2000':])]

    netz_ind_1_rows = netz_ind_1.shape[0]
    netz_ind_1_cols = netz_ind_1.shape[1]
    
    # Fifth data frame construction: Industry by fuel
    netz_ind_2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                         (EGEDA_years_netzero['item_code_new'].isin(['14_industry_sector'])) &
                         (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))]
    
    coal = netz_ind_2[netz_ind_2['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal', 
                                                                                                  item_code_new = '14_industry_sector')
    
    oil = netz_ind_2[netz_ind_2['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil', 
                                                                                                item_code_new = '14_industry_sector')
    
    biomass = netz_ind_2[netz_ind_2['fuel_code'].isin(['15_solid_biomass'])].groupby(['item_code_new']).sum().assign(fuel_code = 'Biomass', 
                                                                                                              item_code_new = '14_industry_sector')
    
    others = netz_ind_2[netz_ind_2['fuel_code'].isin(Other_fuels_industry)].groupby(['item_code_new']).sum().assign(fuel_code = 'Others', 
                                                                                                                item_code_new = '14_industry_sector')
    
    netz_ind_2 = netz_ind_2.append([coal, oil, biomass, others])\
        [['fuel_code', 'item_code_new'] + list(netz_ind_2.loc[:, '2000':])].reset_index(drop = True)

    netz_ind_2.loc[netz_ind_2['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_ind_2.loc[netz_ind_2['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_ind_2.loc[netz_ind_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_ind_2.loc[netz_ind_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_ind_2 = netz_ind_2[netz_ind_2['fuel_code'].isin(FED_agg_fuels_ind)].set_index('fuel_code').loc[FED_agg_fuels_ind].reset_index()
    
    netz_ind_2_rows = netz_ind_2.shape[0]
    netz_ind_2_cols = netz_ind_2.shape[1]

    # Transport data frame construction: FED by fuels
    netz_trans_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                          (EGEDA_years_netzero['fuel_code'].isin(Transport_fuels))]
    
    renewables = netz_trans_1[netz_trans_1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '15_transport_sector')
    
    others = netz_trans_1[netz_trans_1['fuel_code'].isin(Other_fuel_trans)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '15_transport_sector')

    trans_gasoline = netz_trans_1[netz_trans_1['fuel_code'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Gasoline', 
                                                            item_code_new = '15_transport_sector')

    trans_jetfuel = netz_trans_1[netz_trans_1['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Jet fuel', 
                                                            item_code_new = '15_transport_sector')
    
    netz_trans_1 = netz_trans_1.append([renewables, trans_gasoline, trans_jetfuel, others])[['fuel_code', 'item_code_new'] + list(netz_trans_1.loc[:, '2000':])].reset_index(drop = True) 

    netz_trans_1.loc[netz_trans_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    netz_trans_1.loc[netz_trans_1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    netz_trans_1.loc[netz_trans_1['fuel_code'] == '7_9_lpg', 'fuel_code'] = 'LPG'
    netz_trans_1.loc[netz_trans_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_trans_1.loc[netz_trans_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_trans_1 = netz_trans_1[netz_trans_1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    netz_trans_1_rows = netz_trans_1.shape[0]
    netz_trans_1_cols = netz_trans_1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    netz_trans_2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                               (EGEDA_years_netzero['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years_netzero['fuel_code'].isin(['19_total']))].copy()
    
    netz_trans_2.loc[netz_trans_2['item_code_new'] == '15_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    netz_trans_2.loc[netz_trans_2['item_code_new'] == '15_2_road', 'item_code_new'] = 'Road'
    netz_trans_2.loc[netz_trans_2['item_code_new'] == '15_3_rail', 'item_code_new'] = 'Rail'
    netz_trans_2.loc[netz_trans_2['item_code_new'] == '15_4_domestic_navigation', 'item_code_new'] = 'Marine'
    netz_trans_2.loc[netz_trans_2['item_code_new'] == '15_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    netz_trans_2.loc[netz_trans_2['item_code_new'] == '15_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    netz_trans_2 = netz_trans_2[netz_trans_2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    netz_trans_2 = netz_trans_2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    netz_trans_2_rows = netz_trans_2.shape[0]
    netz_trans_2_cols = netz_trans_2.shape[1]

    # Agriculture data frame

    netz_ag_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                         (EGEDA_years_netzero['item_code_new'].isin(Ag_items)) &
                         (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].groupby('fuel_code').sum().assign(item_code_new = 'Agriculture').reset_index()
                     
    coal = netz_ag_1[netz_ag_1['fuel_code'].isin(Coal_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Coal', item_code_new = 'Agriculture')

    oil = netz_ag_1[netz_ag_1['fuel_code'].isin(Oil_fuels)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Oil', item_code_new = 'Agriculture')

    renewables = netz_ag_1[netz_ag_1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = 'Agriculture')
    
    others = netz_ag_1[netz_ag_1['fuel_code'].isin(Other_fuels_FED)].groupby('item_code_new')\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Agriculture')
    
    netz_ag_1 = netz_ag_1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_ag_1.loc[:,'2000':'2050'])].reset_index(drop = True)

    netz_ag_1.loc[netz_ag_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_ag_1 = netz_ag_1[netz_ag_1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()
    
    netz_ag_1_rows = netz_ag_1.shape[0]
    netz_ag_1_cols = netz_ag_1.shape[1]

    netz_ag_2 = netz_ag_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_ag_2_rows = netz_ag_2.shape[0]
    netz_ag_2_cols = netz_ag_2.shape[1]

    # Hydrogen data frame net zero

    netz_hyd_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                                        (EGEDA_years_netzero['item_code_new'].isin(Sectors_tfc)) &
                                        (EGEDA_years_netzero['fuel_code'] == '16_9_other_sources')].groupby('item_code_new').sum().assign(fuel_code = 'Hydrogen').reset_index()

    buildings_hy = netz_hyd_1[netz_hyd_1['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Buildings', fuel_code = 'Hydrogen')

    ag_hy = netz_hyd_1[netz_hyd_1['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Agriculture', fuel_code = 'Hydrogen')

    netz_hyd_1 = netz_hyd_1.append([buildings_hy, ag_hy])\
        [['fuel_code', 'item_code_new'] + list(netz_hyd_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    netz_hyd_1.loc[netz_hyd_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_hyd_1.loc[netz_hyd_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'

    netz_hyd_1 = netz_hyd_1[netz_hyd_1['item_code_new'].isin(['Agriculture', 'Buildings', 'Industry', 'Transport'])]\
        .copy().reset_index(drop = True)

    netz_hyd_1_rows = netz_hyd_1.shape[0]
    netz_hyd_1_cols = netz_hyd_1.shape[1]

    #############################################################################################################################################

    # TPES REFERENCE DATA FRAMES
    # First data frame: TPES by fuels (and also fourth and sixth dataframe with slight tweaks)
    ref_tpes_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '7_total_primary_energy_supply')
    
    oil = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '7_total_primary_energy_supply')
    
    renewables = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '7_total_primary_energy_supply')
    
    others = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '7_total_primary_energy_supply')
    
    ref_tpes_1 = ref_tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(ref_tpes_df.loc[:, '2000':])].reset_index(drop = True)

    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_tpes_1 = ref_tpes_1[ref_tpes_1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    ref_tpes_1_rows = ref_tpes_1.shape[0]
    ref_tpes_1_cols = ref_tpes_1.shape[1]

    ref_tpes_2 = ref_tpes_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_tpes_2_rows = ref_tpes_2.shape[0]
    ref_tpes_2_cols = ref_tpes_2.shape[1]
    
    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    ref_prod_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = ref_prod_df[ref_prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = ref_prod_df[ref_prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = ref_prod_df[ref_prod_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '1_indigenous_production')
    
    others = ref_prod_df[ref_prod_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '1_indigenous_production')
    
    ref_prod_1 = ref_prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(ref_prod_df.loc[:, '2000':])].reset_index(drop = True)

    ref_prod_1.loc[ref_prod_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_prod_1.loc[ref_prod_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_prod_1 = ref_prod_1[ref_prod_1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    ref_prod_1_rows = ref_prod_1.shape[0]
    ref_prod_1_cols = ref_prod_1.shape[1]

    ref_prod_2 = ref_prod_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_prod_2_rows = ref_prod_2.shape[0]
    ref_prod_2_cols = ref_prod_2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    ref_tpes_comp_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                           (EGEDA_years_reference['item_code_new'].isin(tpes_items)) &
                           (EGEDA_years_reference['fuel_code'] == '19_total')]
    
    net_trade = ref_tpes_comp_1[ref_tpes_comp_1['item_code_new'].isin(['2_imports', 
                                                                     '3_exports'])].groupby(['economy', 
                                                                                             'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                        item_code_new = 'Net trade')
                           
    bunkers = ref_tpes_comp_1[ref_tpes_comp_1['item_code_new'].isin(['4_international_marine_bunkers', 
                                                                 '5_international_aviation_bunkers'])].groupby(['economy', 
                                                                                                                  'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                                             item_code_new = 'Bunkers')
    
    ref_tpes_comp_1 = ref_tpes_comp_1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    ref_tpes_comp_1.loc[ref_tpes_comp_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_tpes_comp_1.loc[ref_tpes_comp_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock changes'
    
    ref_tpes_comp_1 = ref_tpes_comp_1.loc[ref_tpes_comp_1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True)
    
    ref_tpes_comp_1_rows = ref_tpes_comp_1.shape[0]
    ref_tpes_comp_1_cols = ref_tpes_comp_1.shape[1]

    # Imports/exports data frame

    ref_imports_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '2_imports') & 
                              (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]

    coal = ref_imports_1[ref_imports_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '2_imports')
    
    renewables = ref_imports_1[ref_imports_1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '2_imports')
    
    others = ref_imports_1[ref_imports_1['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '2_imports')
    
    ref_imports_1 = ref_imports_1.append([coal, oil, renewables, others]).reset_index(drop = True)

    ref_imports_1.loc[ref_imports_1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    ref_imports_1.loc[ref_imports_1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    ref_imports_1.loc[ref_imports_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_imports_1.loc[ref_imports_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_imports_1 = ref_imports_1[ref_imports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_imports_1.loc[:, '2000':])]

    ref_imports_1_rows = ref_imports_1.shape[0]
    ref_imports_1_cols = ref_imports_1.shape[1] 

    ref_imports_2 = ref_imports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_imports_2_rows = ref_imports_2.shape[0]
    ref_imports_2_cols = ref_imports_2.shape[1]                             

    ref_exports_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '3_exports') & 
                              (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].copy()

    # Change export values to positive rather than negative

    ref_exports_1[list(ref_exports_1.columns[3:])] = ref_exports_1[list(ref_exports_1.columns[3:])].apply(lambda x: x * -1)

    coal = ref_exports_1[ref_exports_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '3_exports')
    
    renewables = ref_exports_1[ref_exports_1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '3_exports')
    
    others = ref_exports_1[ref_exports_1['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '3_exports')
    
    ref_exports_1 = ref_exports_1.append([coal, oil, renewables, others]).reset_index(drop = True)

    ref_exports_1.loc[ref_exports_1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    ref_exports_1.loc[ref_exports_1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    ref_exports_1.loc[ref_exports_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_exports_1.loc[ref_exports_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    ref_exports_1 = ref_exports_1[ref_exports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_exports_1.loc[:, '2000':])]

    ref_exports_1_rows = ref_exports_1.shape[0]
    ref_exports_1_cols = ref_exports_1.shape[1]

    ref_exports_2 = ref_exports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_exports_2_rows = ref_exports_2.shape[0]
    ref_exports_2_cols = ref_exports_2.shape[1] 

    # Bunkers dataframes

    ref_bunkers_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '4_international_marine_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(['7_7_gas_diesel_oil', '7_8_fuel_oil']))]

    ref_bunkers_1 = ref_bunkers_1[['fuel_code', 'item_code_new'] + list(ref_bunkers_1.loc[:, '2000':])]

    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '7_8_fuel_oil', 'fuel_code'] = 'Fuel oil'

    ref_bunkers_1_rows = ref_bunkers_1.shape[0]
    ref_bunkers_1_cols = ref_bunkers_1.shape[1]

    ref_bunkers_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '5_international_aviation_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel', '7_2_aviation_gasoline']))]

    jetfuel = ref_bunkers_2[ref_bunkers_2['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel'])]\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Jet fuel',
                                                 item_code_new = '5_international_aviation_bunkers')
    
    ref_bunkers_2 = ref_bunkers_2.append([jetfuel]).reset_index(drop = True)

    ref_bunkers_2 = ref_bunkers_2[['fuel_code', 'item_code_new'] + list(ref_bunkers_2.loc[:, '2000':])]

    ref_bunkers_2.loc[ref_bunkers_2['fuel_code'] == '7_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'

    ref_bunkers_2 = ref_bunkers_2[ref_bunkers_2['fuel_code'].isin(avi_bunker)]\
        .set_index('fuel_code').loc[avi_bunker].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_bunkers_2.loc[:, '2000':])]

    ref_bunkers_2_rows = ref_bunkers_2.shape[0]
    ref_bunkers_2_cols = ref_bunkers_2.shape[1]

    ######################################################################################################################

    # TPES NET-ZERO DATA FRAMES
    # First data frame: TPES by fuels (and also fourth and sixth dataframe with slight tweaks)
    netz_tpes_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '7_total_primary_energy_supply')
    
    oil = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '7_total_primary_energy_supply')
    
    renewables = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '7_total_primary_energy_supply')
    
    others = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '7_total_primary_energy_supply')
    
    netz_tpes_1 = netz_tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(netz_tpes_df.loc[:, '2000':])].reset_index(drop = True)

    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_tpes_1 = netz_tpes_1[netz_tpes_1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    netz_tpes_1_rows = netz_tpes_1.shape[0]
    netz_tpes_1_cols = netz_tpes_1.shape[1]

    netz_tpes_2 = netz_tpes_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_tpes_2_rows = netz_tpes_2.shape[0]
    netz_tpes_2_cols = netz_tpes_2.shape[1]
    
    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    netz_prod_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = netz_prod_df[netz_prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = netz_prod_df[netz_prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = netz_prod_df[netz_prod_df['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                              item_code_new = '1_indigenous_production')
    
    others = netz_prod_df[netz_prod_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                     item_code_new = '1_indigenous_production')
    
    netz_prod_1 = netz_prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(netz_prod_df.loc[:, '2000':])].reset_index(drop = True)

    netz_prod_1.loc[netz_prod_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_prod_1.loc[netz_prod_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_prod_1 = netz_prod_1[netz_prod_1['fuel_code'].isin(TPES_agg_fuels)].set_index('fuel_code').loc[TPES_agg_fuels].reset_index()

    netz_prod_1_rows = netz_prod_1.shape[0]
    netz_prod_1_cols = netz_prod_1.shape[1]

    netz_prod_2 = netz_prod_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_prod_2_rows = netz_prod_2.shape[0]
    netz_prod_2_cols = netz_prod_2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    netz_tpes_comp_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                           (EGEDA_years_reference['item_code_new'].isin(tpes_items)) &
                           (EGEDA_years_reference['fuel_code'] == '19_total')]
    
    net_trade = netz_tpes_comp_1[netz_tpes_comp_1['item_code_new'].isin(['2_imports', 
                                                                     '3_exports'])].groupby(['economy', 
                                                                                             'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                        item_code_new = 'Net trade')
                           
    bunkers = netz_tpes_comp_1[netz_tpes_comp_1['item_code_new'].isin(['4_international_marine_bunkers', 
                                                                 '5_international_aviation_bunkers'])].groupby(['economy', 
                                                                                                                  'fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                                                             item_code_new = 'Bunkers')
    
    netz_tpes_comp_1 = netz_tpes_comp_1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    netz_tpes_comp_1.loc[netz_tpes_comp_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_tpes_comp_1.loc[netz_tpes_comp_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock changes'
    
    netz_tpes_comp_1 = netz_tpes_comp_1.loc[netz_tpes_comp_1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True)
    
    netz_tpes_comp_1_rows = netz_tpes_comp_1.shape[0]
    netz_tpes_comp_1_cols = netz_tpes_comp_1.shape[1]

    # Imports/exports data frame

    netz_imports_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '2_imports') & 
                              (EGEDA_years_reference['fuel_code'].isin(Required_fuels))]

    coal = netz_imports_1[netz_imports_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '2_imports')
    
    renewables = netz_imports_1[netz_imports_1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '2_imports')
    
    others = netz_imports_1[netz_imports_1['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '2_imports')
    
    netz_imports_1 = netz_imports_1.append([coal, oil, renewables, others]).reset_index(drop = True)

    netz_imports_1.loc[netz_imports_1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    netz_imports_1.loc[netz_imports_1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    netz_imports_1.loc[netz_imports_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_imports_1.loc[netz_imports_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_imports_1 = netz_imports_1[netz_imports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_imports_1.loc[:, '2000':])]

    netz_imports_1_rows = netz_imports_1.shape[0]
    netz_imports_1_cols = netz_imports_1.shape[1] 

    netz_imports_2 = netz_imports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_imports_2_rows = netz_imports_2.shape[0]
    netz_imports_2_cols = netz_imports_2.shape[1]                             

    netz_exports_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '3_exports') & 
                              (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].copy()

    # Change export values to positive rather than negative

    netz_exports_1[list(netz_exports_1.columns[3:])] = netz_exports_1[list(netz_exports_1.columns[3:])].apply(lambda x: x * -1)

    coal = netz_exports_1[netz_exports_1['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                          item_code_new = '3_exports')
    
    renewables = netz_exports_1[netz_exports_1['fuel_code'].isin(Renewables_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                      item_code_new = '3_exports')
    
    others = netz_exports_1[netz_exports_1['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other fuels',
                                                                                                             item_code_new = '3_exports')
    
    netz_exports_1 = netz_exports_1.append([coal, oil, renewables, others]).reset_index(drop = True)

    netz_exports_1.loc[netz_exports_1['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'
    netz_exports_1.loc[netz_exports_1['fuel_code'] == '7_petroleum_products', 'fuel_code'] = 'Petroleum products'
    netz_exports_1.loc[netz_exports_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_exports_1.loc[netz_exports_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'

    netz_exports_1 = netz_exports_1[netz_exports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_exports_1.loc[:, '2000':])]

    netz_exports_1_rows = netz_exports_1.shape[0]
    netz_exports_1_cols = netz_exports_1.shape[1]

    netz_exports_2 = netz_exports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_exports_2_rows = netz_exports_2.shape[0]
    netz_exports_2_cols = netz_exports_2.shape[1] 

    # Bunkers dataframes

    netz_bunkers_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '4_international_marine_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(['7_7_gas_diesel_oil', '7_8_fuel_oil']))]

    netz_bunkers_1 = netz_bunkers_1[['fuel_code', 'item_code_new'] + list(netz_bunkers_1.loc[:, '2000':])]

    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '7_8_fuel_oil', 'fuel_code'] = 'Fuel oil'

    netz_bunkers_1_rows = netz_bunkers_1.shape[0]
    netz_bunkers_1_cols = netz_bunkers_1.shape[1]

    netz_bunkers_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '5_international_aviation_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel', '7_2_aviation_gasoline']))]

    jetfuel = netz_bunkers_2[netz_bunkers_2['fuel_code'].isin(['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel'])]\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Jet fuel',
                                                 item_code_new = '5_international_aviation_bunkers')
    
    netz_bunkers_2 = netz_bunkers_2.append([jetfuel]).reset_index(drop = True)

    netz_bunkers_2 = netz_bunkers_2[['fuel_code', 'item_code_new'] + list(netz_bunkers_2.loc[:, '2000':])]

    netz_bunkers_2.loc[netz_bunkers_2['fuel_code'] == '7_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'

    netz_bunkers_2 = netz_bunkers_2[netz_bunkers_2['fuel_code'].isin(avi_bunker)]\
        .set_index('fuel_code').loc[avi_bunker].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_bunkers_2.loc[:, '2000':])]

    netz_bunkers_2_rows = netz_bunkers_2.shape[0]
    netz_bunkers_2_cols = netz_bunkers_2.shape[1]