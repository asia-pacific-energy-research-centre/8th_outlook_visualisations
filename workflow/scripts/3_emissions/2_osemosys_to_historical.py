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

# Path for OSeMOSYS output
path_output = './data/3_OSeMOSYS_output'

# Path for OSeMOSYS to EGEDA mapping
path_mapping = './data/2_Mapping_and_other'

# Where to save finalised dataframe
path_final = './data/4_Joined'

# OSeMOSYS results files
OSeMOSYS_filenames = glob.glob(path_output + "/*.xlsx")

########################################################################################################################################

Mapping_sheets = list(pd.read_excel(path_mapping + '/OSeMOSYS_mapping_2021.xlsx', sheet_name = None).keys())[1:]

Mapping_file = pd.DataFrame()

for sheet in Mapping_sheets:
    interim_map = pd.read_excel(path_mapping + '/OSeMOSYS_mapping_2021.xlsx', sheet_name = sheet, skiprows = 1)
    Mapping_file = Mapping_file.append(interim_map).reset_index(drop = True)

# Now moving everything from OSeMOSYS to EGEDA (Only demand sectors and own use for now)

Mapping_file = Mapping_file[Mapping_file['Sector'].isin(['AGR', 'BLD', 'IND', 'TRN', 'PIP', 'NON', 'OWN', 'POW'])].copy() 
Mapping_file = Mapping_file[Mapping_file['item_code_new'].notna()].copy().reset_index(drop = True)

# Define unique workbook and sheet combinations
Unique_combo = Mapping_file.groupby(['Workbook', 'Sheet_emissions']).size().reset_index().loc[:, ['Workbook', 'Sheet_emissions']]

# Determine list of files to read based on the workbooks identified in the mapping file
file_df = pd.DataFrame()

for i in range(len(Unique_combo['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in OSeMOSYS_filenames if Unique_combo['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_combo['Workbook'].unique()[i]})
    file_df = file_df.append(_file)

file_df = file_df.merge(Unique_combo, how = 'outer', on = 'Workbook')

# Create empty dataframe to store aggregated results 
aggregate_df1 = pd.DataFrame()

# Now read in the OSeMOSYS output files so that that they're all in one data frame (aggregate_df1)
for i in range(file_df.shape[0]):
    _df = pd.read_excel(file_df.iloc[i, 0], sheet_name = file_df.iloc[i, 2])
    _df['Workbook'] = file_df.iloc[i, 1]
    _df['Sheet_emissions'] = file_df.iloc[i, 2]
    aggregate_df1 = aggregate_df1.append(_df) 

# Now aggregate all the results for APEC

APEC = aggregate_df1.groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()
APEC['REGION'] = 'APEC'

aggregate_df1 = aggregate_df1.append(APEC).reset_index(drop = True)

# Get maximum year column to build data frame below
year_columns = []

for item in list(aggregate_df1.columns):
    try:
        year_columns.append(int(item))
    except ValueError:
            pass

max_year = max(year_columns)

OSeMOSYS_years = list(range(2017, max_year + 1))

########################## fuel_code aggregations ##########################

# First level
coal_fuels = ['1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal']
oil_fuels = ['6_1_crude_oil', '6_x_ngls']
petrol_fuels = ['7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha', '7_x_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil',
                '7_8_fuel_oil', '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_11_ethane', '7_x_other_petroleum_products']
gas_fuels = ['8_1_natural_gas', '8_2_lng', '8_3_gas_works_gas']
biomass_fuels = ['15_1_fuelwood_and_woodwaste', '15_2_bagasse', '15_3_charcoal', '15_4_black_liquor', '15_5_other_biomass']
other_fuels = ['16_1_biogas', '16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable', '16_5_biogasoline', '16_6_biodiesel',
               '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_9_other_sources', '16_x_hydrogen']

# Total
total_fuels = ['1_coal', '2_coal_products', '5_oil_shale_and_oil_sands', '6_crude_oil_and_ngl', '7_petroleum_products', '8_gas', '9_nuclear', '10_hydro', '11_geothermal',
               '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', '16_others', '17_electricity', '18_heat']

# total_renewables to be completed

##############################################################################

# item_code_new aggregations

# Lowest level
industry_agg = ['14_1_iron_and_steel', '14_2_chemical_incl_petrochemical', '14_3_non_ferrous_metals', '14_4_nonmetallic_mineral_products', '14_5_transportation_equipment',
                '14_6_machinery', '14_7_mining_and_quarrying', '14_8_food_beverages_and_tobacco', '14_9_pulp_paper_and_printing', '14_10_wood_and_wood_products',
                '14_11_construction', '14_12_textiles_and_leather', '14_13_nonspecified_industry']
transport_agg = ['15_1_domestic_air_transport', '15_2_road', '15_3_rail', '15_4_domestic_navigation', '15_5_pipeline_transport', '15_6_nonspecified_transport']
others_agg = ['16_1_commercial_and_public_services', '16_2_residential', '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others']
dem_pow_own_agg = ['9_x_power', '10_losses_and_own_use', '13_total_final_energy_consumption']

# Then first level
tpes_agg = ['1_indigenous_production', '2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers']

tfc_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector', '17_nonenergy_use']

tfec_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector']

###################################################################################################

# Now aggregate data based on the mapping
# That is group by REGION, TECHNOLOGY and EMISSION
# First create empty dataframe
aggregate_df2 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in aggregate_df1['REGION'].unique():
    interim_df1 = aggregate_df1[aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Mapping_file, how = 'left', on = ['TECHNOLOGY', 'EMISSION'])
    interim_df1 = interim_df1.groupby(['item_code_new', 'fuel_code']).sum().reset_index()

    ########################### Aggregate fuel_code for new variables ###################################

    # First level fuels

    coal = interim_df1[interim_df1['fuel_code'].isin(coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '1_coal').reset_index()

    oil = interim_df1[interim_df1['fuel_code'].isin(oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '6_crude_oil_and_ngl').reset_index()

    petrol = interim_df1[interim_df1['fuel_code'].isin(petrol_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '7_petroleum_products').reset_index()

    gas = interim_df1[interim_df1['fuel_code'].isin(gas_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '8_gas').reset_index()

    # biomass = interim_df1[interim_df1['fuel_code'].isin(biomass_fuels)].groupby(['item_code_new'])\
    #     .sum().assign(fuel_code = '15_solid_biomass').reset_index()

    others = interim_df1[interim_df1['fuel_code'].isin(other_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '16_others').reset_index()

    interim_df2 = interim_df1.append([coal, oil, petrol, gas, others]).reset_index(drop = True)

    # And total fuels

    total_f = interim_df2[interim_df2['fuel_code'].isin(total_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '19_total').reset_index()

    interim_df3 = interim_df2.append(total_f).reset_index(drop = True)

    ################################ And now item_code_new ######################################

    # Start with lowest level

    industry = interim_df3[interim_df3['item_code_new'].isin(industry_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '14_industry_sector').reset_index()

    transport = interim_df3[interim_df3['item_code_new'].isin(transport_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '15_transport_sector').reset_index()

    bld_ag_other = interim_df3[interim_df3['item_code_new'].isin(others_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '16_other_sector').reset_index()

    interim_df4 = interim_df3.append([industry, transport, bld_ag_other]).reset_index(drop = True)

    # Now higher level agg

    tpes = interim_df4[interim_df4['item_code_new'].isin(tpes_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '7_total_primary_energy_supply').reset_index()

    tfc = interim_df4[interim_df4['item_code_new'].isin(tfc_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '12_total_final_consumption').reset_index()

    tfec = interim_df4[interim_df4['item_code_new'].isin(tfec_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '13_total_final_energy_consumption').reset_index()

    interim_df5 = interim_df4.append([tpes, tfc, tfec]).reset_index(drop = True)

    dem_pow_own = interim_df5[interim_df5['item_code_new'].isin(dem_pow_own_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '13_x_dem_pow_own').reset_index()

    interim_df6 = interim_df5.append(dem_pow_own).reset_index(drop = True)

    # Now add in economy reference
    interim_df6['economy'] = region

    # Now append economy dataframe to communal data frame 
    aggregate_df2 = aggregate_df2.append(interim_df6)

key_variables = ['economy', 'fuel_code', 'item_code_new']

# aggregate_df2 = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years]
aggregate_df2 = aggregate_df2.loc[:, key_variables + OSeMOSYS_years]

# Now load the EGEDA_2018_emissions data frame
EGEDA_emissions = pd.read_csv('./data/1_EGEDA/EGEDA_2018_emissions.csv')

# Join EGEDA historical to OSeMOSYS results (line below removes 2017 and 2018 from historical)
Joined_df = EGEDA_emissions.iloc[:, :-2].merge(aggregate_df2, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')
Joined_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_emissions_2018.csv', index = False)

print('OSeMOSYS_to_EGEDA.csv file successfully created')



