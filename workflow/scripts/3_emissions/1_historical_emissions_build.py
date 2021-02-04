# Build a cleaned up historical emissions dataframe based on EGEDA

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

# Path for OSeMOSYS output
path_output = './data/3_OSeMOSYS_output'

# Path for OSeMOSYS to EGEDA mapping
path_mapping = './data/2_Mapping_and_other'

# Where to save finalised dataframe
path_final = './data/4_Joined'

# OSeMOSYS results files
OSeMOSYS_filenames = glob.glob(path_output + "/*.xlsx")

# Load historical emissions

EGEDA_emissions = pd.read_csv('./data/1_EGEDA/EGEDA_FC_CO2_Emissions_years_2018.csv')

# Remove all aggregate variables as they're zero

agg_fuel = ['1_coal', '1_x_coal_thermal', '2_coal_products', '6_crude_oil_and_ngl', '6_x_ngls',
            '7_petroleum_products', '7_x_jet_fuel', '7_x_other_petroleum_products', '8_gas', '16_others', '19_total']

EGEDA_emissions = EGEDA_emissions[~EGEDA_emissions['fuel_code'].isin(agg_fuel)].reset_index(drop = True)

########################## fuel_code aggregations ##########################

# lowest level

thermal_coal = ['1_2_other_bituminous_coal', '1_3_subbituminous_coal', '1_4_anthracite', '3_peat', '4_peat_products']
ngl = ['6_2_natural_gas_liquids', '6_3_refinery_feedstocks', '6_4_additives_oxygenates', '6_5_other_hydrocarbons']
other_petrol = ['7_12_white_spirit_sbp', '7_13_lubricants', '7_14_bitumen', '7_15_paraffin_waxes', '7_16_petroleum_coke', '7_17_other_products']
jetfuel = ['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel']

# First level
coal_fuels = ['1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal']
coal_prod_fuels = ['2_1_coke_oven_coke', '2_2_coke_oven_gas', '2_3_blast_furnace_gas', '2_4_other_recovered_gases', '2_5_patent_fuel', '2_6_coal_tar', '2_7_bkb_pb']
oil_fuels = ['6_1_crude_oil', '6_x_ngls']
petrol_fuels = ['7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha', '7_x_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil',
                '7_8_fuel_oil', '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_11_ethane', '7_x_other_petroleum_products']
gas_fuels = ['8_1_natural_gas', '8_2_lng', '8_3_gas_works_gas']
other_fuels = ['16_1_biogas', '16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable', '16_5_biogasoline', '16_6_biodiesel',
               '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_9_other_sources', '16_x_hydrogen']

# Total
total_fuels = ['1_coal', '2_coal_products', '5_oil_shale_and_oil_sands', '6_crude_oil_and_ngl', '7_petroleum_products', '8_gas', '9_nuclear', '10_hydro', '11_geothermal',
               '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', '16_others', '17_electricity', '18_heat']

# item_code_new aggregations
tfc_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector', '17_nonenergy_use']

tfec_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector']

# Aggregations

EGEDA_aggregate = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in EGEDA_emissions['economy'].unique():
    interim_df1 = EGEDA_emissions[EGEDA_emissions['economy'] == region]
    
    thermal_agg = interim_df1[interim_df1['fuel_code'].isin(thermal_coal)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '1_x_coal_thermal').reset_index()

    ngl_agg = interim_df1[interim_df1['fuel_code'].isin(ngl)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '6_x_ngls').reset_index()

    oth_pet_agg = interim_df1[interim_df1['fuel_code'].isin(other_petrol)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '7_x_other_petroleum_products').reset_index()

    jetfuel_agg = interim_df1[interim_df1['fuel_code'].isin(jetfuel)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '7_x_jet_fuel').reset_index()

    interim_df2 = interim_df1.append([thermal_agg, ngl_agg, oth_pet_agg, jetfuel_agg]).reset_index(drop = True)
        
    coal = interim_df2[interim_df2['fuel_code'].isin(coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '1_coal').reset_index()

    coal_prod = interim_df2[interim_df2['fuel_code'].isin(coal_prod_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '2_coal_products').reset_index()
        
    oil = interim_df2[interim_df2['fuel_code'].isin(oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '6_crude_oil_and_ngl').reset_index()
        
    petrol = interim_df2[interim_df2['fuel_code'].isin(petrol_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '7_petroleum_products').reset_index()
        
    gas = interim_df2[interim_df2['fuel_code'].isin(gas_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '8_gas').reset_index()

    others = interim_df2[interim_df2['fuel_code'].isin(other_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '16_others').reset_index()

    interim_df3 = interim_df2.append([coal, coal_prod, oil, petrol, gas, others]).reset_index(drop = True)

    # Now add in the totals
    total = interim_df3[interim_df3['fuel_code'].isin(total_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '19_total').reset_index()

    interim_df4 = interim_df3.append([total]).reset_index(drop = True)

    # # Totals by sector aggregation
    # tfc = interim_df5[interim_df5['item_code_new'].isin(tfc_agg)].groupby(['fuel_code'])\
    #     .sum().assign(item_code_new = '12_total_final_consumption').reset_index()

    # tfec = interim_df5[interim_df5['item_code_new'].isin(tfec_agg)].groupby(['fuel_code'])\
    #     .sum().assign(item_code_new = '13_total_final_energy_consumption').reset_index()

    # interim_df6 = interim_df5.append([tpes, tfc, tfec]).reset_index(drop = True)                                      

    interim_df4['economy'] = region

    EGEDA_aggregate = EGEDA_aggregate.append(interim_df4).reset_index(drop = True)

# Load correct order of fuel code and item code. Update this csv based on new entries or desired order

ordered = pd.read_csv('./data/2_Mapping_and_other/order_2018.csv')

# This grabs the unique values of fuel_code and item_code_new in the order they appear in the original dataframe. It removes 'na' by calling '[:-1]' 

order1 = list(ordered['fuel_code'].unique())[:-1]
order2 = list(ordered['item_code_new'])

# Take order defined above and define each of the variables as categorical in that already established order (for the benefit of viewing data later)

EGEDA_aggregate['fuel_code'] = pd.Categorical(EGEDA_aggregate['fuel_code'], 
                                                categories = order1, 
                                                ordered = True)

EGEDA_aggregate['item_code_new'] = pd.Categorical(EGEDA_aggregate['item_code_new'],
                                                    categories = order2,
                                                    ordered = True)

EGEDA_aggregate_sorted = EGEDA_aggregate.sort_values(['economy', 'fuel_code', 'item_code_new']).reset_index(drop = True)

# Write file
EGEDA_aggregate_sorted.to_csv('./data/1_EGEDA/EGEDA_2018_emissions.csv', index = False)

########################################################################################################################################



