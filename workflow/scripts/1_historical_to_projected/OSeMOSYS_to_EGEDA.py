# Notebook to transform OSeMOSYS output to same format as EGEDA

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

# Read in mapping file
Mapping_file = pd.read_excel(path_mapping + '/OSeMOSYS mapping.xlsx', sheet_name = 'Mapping',  skiprows = 1)
Mapping_file = Mapping_file[Mapping_file['Balance'].isin(['TFC', 'TPES'])]

# Define unique workbook and sheet combinations
Unique_combo = Mapping_file.groupby(['Workbook', 'Sheet']).size().reset_index().loc[:, ['Workbook', 'Sheet']]

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
    _df['Sheet'] = file_df.iloc[i, 2]
    aggregate_df1 = aggregate_df1.append(_df) 

interim_df1 = aggregate_df1[aggregate_df1['TIMESLICE'] != 'ONE']
interim_df2 = aggregate_df1[aggregate_df1['TIMESLICE'] == 'ONE']

interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION', 'Workbook', 'Sheet']).sum().reset_index()

aggregate_df1 = interim_df2.append(interim_df1).reset_index(drop = True)

# Now aggregate all the results for APEC

APEC = aggregate_df1.groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
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

# Start at lowest level first
gasoline_fuels = ['4_1_1_motor_gasoline', '4_1_2_aviation_gasoline']               
liquid_biofuels = ['9_8_1_biogasoline', '9_8_2_biodiesel', '9_8_4_other_liquid_biofuels']
other_renew = ['8_2_1_photovoltaic', '8_2_2_tide_wave_ocean', '8_2_3_wind', '8_2_4_solar']

# Then first level
coal_fuels = ['1_1_1_coking_coal', '1_x_coal_thermal', '1_3_lignite']
oil_fuels = ['3_1_crude_oil', '3_x_ngls']
petrol_fuels = ['4_1_gasoline', '4_2_naphtha', '4_3_jet_fuel', '4_4_other_kerosene', '4_5_gas_diesel_oil', 
                  '4_6_fuel_oil', '4_7_lpg', '4_8_refinery_gas_not_liq', '4_9_ethane', '4_10_other_petroleum_products']
gas_fuels = ['5_1_natural_gas', '5_2_lng']
renew_fuels = ['8_1_geothermal_power', '8_2_other_power', '8_3_geothermal_heat', '8_4_solar_heat']
other_fuels = ['9_1_fuel_wood_and_woodwaste', '9_2_bagasse', '9_3_charcoal', '9_4_other_biomass', '9_5_biogas', '9_6_industrial_waste', '9_7_municipal_solid_waste',
               '9_8_liquid_biofuels', '9_9_other_sources']

# Total
total_fuels = ['1_coal', '2_coal_products', '3_crude_oil_and_ngl', '4_petroleum_products', '5_gas', '6_hydro', '7_nuclear', '8_geothermal_solar_etc', '9_others',
               '10_electricity', '11_heat']

# total_renewables to be completed

##############################################################################

# item_code_new aggregations

# Lowest level
industry_agg = ['13_1_iron_and_steel', '13_2_chemical_incl__petrochemical', '13_3_nonferrous_metals', '13_4_nonmetallic_mineral_products', 
                '13_5_transportation_equipment', '13_6_machinery', '13_7_mining_and_quarrying', '13_8_food_beverages_and_tobacco',
                '13_9_pulp_paper_and_printing', '13_10_wood_and_wood_products', '13_11_construction', '13_12_textiles_and_leather',
                '13_13_nonspecified_industry']

transport_agg = ['14_1_domestic_air_transport', '14_2_road', '14_3_rail', '14_4_domestic_water_transport', '14_5_pipeline_transport', '14_6_nonspecified_transport']

others_agg = ['15_1_1_commerce_and_public_services', '15_1_2_residential', '15_2_agriculture', '15_3_fishing', '15_4_nonspecified_others']

# Then first level
tpes_agg = ['1_indigenous_production', '2_imports', '3_exports', '4_1_international_marine_bunkers', '4_2_international_aviation_bunkers']

tfc_agg = ['13_industry_sector', '14_transport_sector', '15_other_sector', '16_nonenergy_use']

tfec_agg = ['13_industry_sector', '14_transport_sector', '15_other_sector']

##############################################################################

# Now aggregate data based on the mapping
# That is group by REGION, TECHNOLOGY and FUEL
# First create empty dataframe
aggregate_df2 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in aggregate_df1['REGION'].unique():
    interim_df1 = aggregate_df1[aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Mapping_file, how = 'left', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['item_code_new', 'fuel_code']).sum().reset_index()

    # Change export data to negative values
    exports_bunkers = interim_df1[interim_df1['item_code_new'].isin(['3_exports', '4_1_international_marine_bunkers', '4_2_international_aviation_bunkers'])]\
        .set_index(['item_code_new', 'fuel_code'])
    everything_else = interim_df1[~interim_df1['item_code_new'].isin(['3_exports', '4_1_international_marine_bunkers', '4_2_international_aviation_bunkers'])]

    exports_bunkers = exports_bunkers * -1
    exports_bunkers = exports_bunkers.reset_index()
    interim_df1 = everything_else.append(exports_bunkers)

    ########################### Aggregate fuel_code for new variables ###################################

    # Start with lowest level of aggregation (so that it can then be grabbed at higher levels of aggregation later)
    
    gasoline = interim_df1[interim_df1['fuel_code'].isin(gasoline_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '4_1_gasoline').reset_index()
    
    liquids = interim_df1[interim_df1['fuel_code'].isin(liquid_biofuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '9_8_liquid_biofuels').reset_index()

    renewables_others = interim_df1[interim_df1['fuel_code'].isin(other_renew)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '8_2_other_power').reset_index()

    interim_df2 = interim_df1.append([gasoline, liquids, renewables_others]).reset_index(drop = True)

    # Now first level fuels

    coal = interim_df2[interim_df2['fuel_code'].isin(coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '1_coal').reset_index()

    oil = interim_df2[interim_df2['fuel_code'].isin(oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '3_crude_oil_and_ngl').reset_index()

    petrol = interim_df2[interim_df2['fuel_code'].isin(petrol_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '4_petroleum_products').reset_index()

    gas = interim_df2[interim_df2['fuel_code'].isin(gas_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '5_gas').reset_index()

    renew = interim_df2[interim_df2['fuel_code'].isin(renew_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '8_geothermal_solar_etc').reset_index()

    others = interim_df2[interim_df2['fuel_code'].isin(other_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '9_others').reset_index()

    interim_df3 = interim_df2.append([coal, oil, petrol, gas, renew, others]).reset_index(drop = True)

    # And total fuels

    total_f = interim_df3[interim_df3['fuel_code'].isin(total_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '12_total').reset_index()

    interim_df4 = interim_df3.append(total_f).reset_index(drop = True)

    ################################ And now item_code_new ######################################

    # Start with lowest level

    industry = interim_df4[interim_df4['item_code_new'].isin(industry_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '13_industry_sector').reset_index()

    transport = interim_df4[interim_df4['item_code_new'].isin(transport_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '14_transport_sector').reset_index()

    bld_ag_other = interim_df4[interim_df4['item_code_new'].isin(others_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '15_other_sector').reset_index()

    interim_df5 = interim_df4.append([industry, transport, bld_ag_other]).reset_index(drop = True)

    # Now higher level agg

    #Might need to check this depending on whether exports is negative
    tpes = interim_df5[interim_df5['item_code_new'].isin(tpes_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '6_total_primary_energy_supply').reset_index()

    tfc = interim_df5[interim_df5['item_code_new'].isin(tfc_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '11_total_final_consumption').reset_index()

    tfec = interim_df5[interim_df5['item_code_new'].isin(tfec_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '12_total_final_energy_consumption').reset_index()

    interim_df6 = interim_df5.append([tpes, tfc, tfec]).reset_index(drop = True)

    # Now add in economy reference
    interim_df6['economy'] = region

    # Now append economy dataframe to communal data frame 
    aggregate_df2 = aggregate_df2.append(interim_df6)
    
key_variables = ['economy', 'fuel_code', 'item_code_new']

# aggregate_df2 = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years]
aggregate_df2 = aggregate_df2.loc[:, key_variables + OSeMOSYS_years]

# Now load the EGEDA_years data frame
EGEDA_years = pd.read_csv('./data/1_EGEDA/EGEDA_2020_June_22_wide_years_PJ.csv')

########################### Special amendment to EGEDA historical ############################

# HONG KONG imported electricity shifts to both Nuclear and Hydro in a 75:25 split

# Hydro

EGEDA_hkc_hydro = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '6_hydro') &
                                (EGEDA_years['item_code_new'] == '1_indigenous_production')].copy()

EGEDA_hkc_elec_imports = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '10_electricity') &
                                (EGEDA_years['item_code_new'] == '2_imports')].copy()

EGEDA_hkc_elec_imports[EGEDA_hkc_elec_imports.select_dtypes(include = ['number']).columns] *= 0.25

append1 = EGEDA_hkc_hydro.append(EGEDA_hkc_elec_imports).reset_index(drop = True)
d = append1.dtypes

append1.loc[44002] = append1.sum(numeric_only = True)
append1.astype(d)

append1.loc[44002, 'economy'] = append1.loc[0, 'economy']
append1.loc[44002, 'fuel_code'] = append1.loc[0, 'fuel_code']
append1.loc[44002, 'item_code_new'] = append1.loc[0, 'item_code_new']

append1 = append1.drop([0, 1])

EGEDA_years = append1.combine_first(EGEDA_years)

# Nuclear

EGEDA_hkc_nuclear = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '7_nuclear') &
                                (EGEDA_years['item_code_new'] == '1_indigenous_production')].copy()

EGEDA_hkc_elec_imports = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '10_electricity') &
                                (EGEDA_years['item_code_new'] == '2_imports')].copy()

EGEDA_hkc_elec_imports[EGEDA_hkc_elec_imports.select_dtypes(include = ['number']).columns] *= 0.75

append2 = EGEDA_hkc_nuclear.append(EGEDA_hkc_elec_imports).reset_index(drop = True)
d = append2.dtypes

append2.loc[44100] = append2.sum(numeric_only = True)
append2.astype(d)

append2.loc[44100, 'economy'] = append2.loc[0, 'economy']
append2.loc[44100, 'fuel_code'] = append2.loc[0, 'fuel_code']
append2.loc[44100, 'item_code_new'] = append2.loc[0, 'item_code_new']

append2 = append2.drop([0, 1])

EGEDA_years = append2.combine_first(EGEDA_years)

# Now change electricity imports to zero

EGEDA_hkc_elec_imports = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '10_electricity') &
                                (EGEDA_years['item_code_new'] == '2_imports')].copy()

EGEDA_hkc_elec_imports[EGEDA_hkc_elec_imports.select_dtypes(include = ['number']).columns] *= 0

EGEDA_years = EGEDA_hkc_elec_imports.combine_first(EGEDA_years)

# Now amend TPES category 6 to reflect changes to hydro, nuclear (TPES should now equal indigenous prod) and electricity (TPES should only equal exports)

# Hydro TPES
EGEDA_hkc_hydro_tpes = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '6_hydro') &
                                (EGEDA_years['item_code_new'].isin(['1_indigenous_production', '2_imports', '3_exports']))].copy()

EGEDA_hkc_hydro_tpes = EGEDA_hkc_hydro_tpes.append(EGEDA_hkc_hydro_tpes.sum(numeric_only = True), ignore_index = True)

EGEDA_hkc_hydro_tpes.loc[3, 'economy'] = '06_HKC'
EGEDA_hkc_hydro_tpes.loc[3, 'fuel_code'] = '6_hydro'
EGEDA_hkc_hydro_tpes.loc[3, 'item_code_new'] = '6_total_primary_energy_supply'

EGEDA_hkc_hydro_tpes = EGEDA_hkc_hydro_tpes.rename(index = {3: 44010})

EGEDA_hkc_hydro_tpes = EGEDA_hkc_hydro_tpes.drop([0, 1, 2])

EGEDA_years = EGEDA_hkc_hydro_tpes.combine_first(EGEDA_years)

# Nuclear TPES
EGEDA_hkc_nuclear_tpes = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '7_nuclear') &
                                (EGEDA_years['item_code_new'].isin(['1_indigenous_production', '2_imports', '3_exports']))].copy()

EGEDA_hkc_nuclear_tpes = EGEDA_hkc_nuclear_tpes.append(EGEDA_hkc_nuclear_tpes.sum(numeric_only = True), ignore_index = True)

EGEDA_hkc_nuclear_tpes.loc[3, 'economy'] = '06_HKC'
EGEDA_hkc_nuclear_tpes.loc[3, 'fuel_code'] = '7_nuclear'
EGEDA_hkc_nuclear_tpes.loc[3, 'item_code_new'] = '6_total_primary_energy_supply'

EGEDA_hkc_nuclear_tpes = EGEDA_hkc_nuclear_tpes.rename(index = {3: 44108})

EGEDA_hkc_nuclear_tpes = EGEDA_hkc_nuclear_tpes.drop([0, 1, 2])

EGEDA_years = EGEDA_hkc_nuclear_tpes.combine_first(EGEDA_years)

# Electricity TPES
EGEDA_hkc_elec_tpes = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
                                (EGEDA_years['fuel_code'] == '10_electricity') &
                                (EGEDA_years['item_code_new'].isin(['1_indigenous_production', '2_imports', '3_exports']))].copy()

EGEDA_hkc_elec_tpes = EGEDA_hkc_elec_tpes.append(EGEDA_hkc_elec_tpes.sum(numeric_only = True), ignore_index = True)

EGEDA_hkc_elec_tpes.loc[3, 'economy'] = '06_HKC'
EGEDA_hkc_elec_tpes.loc[3, 'fuel_code'] = '10_electricity'
EGEDA_hkc_elec_tpes.loc[3, 'item_code_new'] = '6_total_primary_energy_supply'

EGEDA_hkc_elec_tpes = EGEDA_hkc_elec_tpes.rename(index = {3: 46656})

EGEDA_hkc_elec_tpes = EGEDA_hkc_elec_tpes.drop([0, 1, 2])

EGEDA_years = EGEDA_hkc_elec_tpes.combine_first(EGEDA_years)

# Remove 2017 which is already in the EGEDA historical
# aggregate_df2_tojoin = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years[1:]]
# aggregate_df2_tojoin = aggregate_df2.loc[:, key_variables + OSeMOSYS_years[1:]] # New line below keeps 2017 in OSeMOSYS
aggregate_df2_tojoin = aggregate_df2.loc[:, key_variables + OSeMOSYS_years]

# Join EGEDA historical to OSeMOSYS results (line below removes 2017 from historical)
Joined_df = EGEDA_years.iloc[:, :-1].merge(aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')
Joined_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA.csv', index = False)

print('OSeMOSYS_to_EGEDA.csv file successfully created')


