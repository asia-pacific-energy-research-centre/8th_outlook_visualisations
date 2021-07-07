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

# Reference filenames and net zero filenames

reference_filenames = list(filter(lambda k: 'reference' in k, OSeMOSYS_filenames))
netzero_filenames = list(filter(lambda y: 'net-zero' in y, OSeMOSYS_filenames))

# New 2018 data variable names 

Mapping_sheets = list(pd.read_excel(path_mapping + '/OSeMOSYS_mapping_2021.xlsx', sheet_name = None).keys())[1:]

Mapping_file = pd.DataFrame()

for sheet in Mapping_sheets:
    interim_map = pd.read_excel(path_mapping + '/OSeMOSYS_mapping_2021.xlsx', sheet_name = sheet, skiprows = 1)
    Mapping_file = Mapping_file.append(interim_map).reset_index(drop = True)

# Moving everything from OSeMOSYS to EGEDA for TFC and TPES

Mapping_TFC_TPES = Mapping_file[Mapping_file['Balance'].isin(['TFC', 'TPES'])]

# And for transformation

Map_trans = Mapping_file[Mapping_file['Balance'] == 'TRANS'].reset_index(drop = True)

# A mapping just for i) power and ii) ref, own, sup

Map_power = Map_trans[Map_trans['Sector'] == 'POW'].reset_index(drop = True)
Map_refownsup = Map_trans[Map_trans['Sector'].isin(['REF', 'SUP', 'OWN'])].reset_index(drop = True)

# Define unique workbook and sheet combinations for TFC and TPES
Unique_TFC_TPES = Mapping_TFC_TPES.groupby(['Workbook', 'Sheet_energy']).size().reset_index().loc[:, ['Workbook', 'Sheet_energy']]

# Define unique workbook and sheet combinations for Transformation
Unique_trans = Map_trans.groupby(['Workbook', 'Sheet_energy']).size().reset_index().loc[:, ['Workbook', 'Sheet_energy']]

################################### TFC and TPES #############################################################
# Determine list of files to read based on the workbooks identified in the mapping file for REFERENCE scenario
ref_file_df = pd.DataFrame()

for i in range(len(Unique_TFC_TPES['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in reference_filenames if Unique_TFC_TPES['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_TFC_TPES['Workbook'].unique()[i]})
    ref_file_df = ref_file_df.append(_file)

ref_file_df = ref_file_df.merge(Unique_TFC_TPES, how = 'outer', on = 'Workbook')

# Determine list of files to read based on the workbooks identified in the mapping file for NET-ZERO scenario
netz_file_df = pd.DataFrame()

for i in range(len(Unique_TFC_TPES['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in netzero_filenames if Unique_TFC_TPES['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_TFC_TPES['Workbook'].unique()[i]})
    netz_file_df = netz_file_df.append(_file)

netz_file_df = netz_file_df.merge(Unique_TFC_TPES, how = 'outer', on = 'Workbook')

# Create empty dataframe to store REFERENCE aggregated results 
ref_aggregate_df1 = pd.DataFrame(columns = ['TECHNOLOGY', 'FUEL', 'REGION', 2050])

# Now read in the OSeMOSYS output files so that that they're all in one data frame (ref_aggregate_df1)
if ref_file_df['File'].isna().any() == False:
    for i in range(ref_file_df.shape[0]):
        _df = pd.read_excel(ref_file_df.iloc[i, 0], sheet_name = ref_file_df.iloc[i, 2])
        _df['Workbook'] = ref_file_df.iloc[i, 1]
        _df['Sheet_energy'] = ref_file_df.iloc[i, 2]
        ref_aggregate_df1 = ref_aggregate_df1.append(_df) 
        
    interim_df1 = ref_aggregate_df1[ref_aggregate_df1['TIMESLICE'] != 'ONE']
    interim_df2 = ref_aggregate_df1[ref_aggregate_df1['TIMESLICE'] == 'ONE']
    
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION', 'Workbook', 'Sheet_energy']).sum().reset_index()
    
    ref_aggregate_df1 = interim_df2.append(interim_df1).reset_index(drop = True)

# Create empty dataframe to store NET ZERO aggregated results 
netz_aggregate_df1 = pd.DataFrame(columns = ['TECHNOLOGY', 'FUEL', 'REGION', 2050])

# Now read in the OSeMOSYS output files so that that they're all in one data frame (netz_aggregate_df1)
if netz_file_df['File'].isna().any() == False:
    for i in range(netz_file_df.shape[0]):
        _df = pd.read_excel(netz_file_df.iloc[i, 0], sheet_name = netz_file_df.iloc[i, 2])
        _df['Workbook'] = netz_file_df.iloc[i, 1]
        _df['Sheet_energy'] = netz_file_df.iloc[i, 2]
        netz_aggregate_df1 = netz_aggregate_df1.append(_df) 

    interim_df1 = netz_aggregate_df1[netz_aggregate_df1['TIMESLICE'] != 'ONE']
    interim_df2 = netz_aggregate_df1[netz_aggregate_df1['TIMESLICE'] == 'ONE']

    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION', 'Workbook', 'Sheet_energy']).sum().reset_index()

    netz_aggregate_df1 = interim_df2.append(interim_df1).reset_index(drop = True)

# Now aggregate all the results for APEC

# REFERENCE
APEC_ref = ref_aggregate_df1.groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
APEC_ref['REGION'] = 'APEC'

ref_aggregate_df1 = ref_aggregate_df1.append(APEC_ref).reset_index(drop = True)

# NET ZERO
APEC_netz = netz_aggregate_df1.groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
APEC_netz['REGION'] = 'APEC'

netz_aggregate_df1 = netz_aggregate_df1.append(APEC_netz).reset_index(drop = True)

# Now aggregate results for 22_SEA
# Southeast Asia: 02, 07, 10, 15, 17, 19, 21

# REFERENCE
SEA_ref = ref_aggregate_df1[ref_aggregate_df1['REGION']\
    .isin(['02_BD', '07_INA', '10_MAS', '15_RP', '17_SIN', '19_THA', '21_VN'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
SEA_ref['REGION'] = '22_SEA'

ref_aggregate_df1 = ref_aggregate_df1.append(SEA_ref).reset_index(drop = True)

# NET ZERO
SEA_netz = netz_aggregate_df1[netz_aggregate_df1['REGION']\
    .isin(['02_BD', '07_INA', '10_MAS', '15_RP', '17_SIN', '19_THA', '21_VN'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
SEA_netz['REGION'] = '22_SEA'

netz_aggregate_df1 = netz_aggregate_df1.append(SEA_netz).reset_index(drop = True)

# Aggregate results for 23_NEA
# Northeast Asia: 06, 08, 09, 18

# REFERENCE
NEA_ref = ref_aggregate_df1[ref_aggregate_df1['REGION']\
    .isin(['06_HKC', '08_JPN', '09_ROK', '18_CT'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
NEA_ref['REGION'] = '23_NEA'

ref_aggregate_df1 = ref_aggregate_df1.append(NEA_ref).reset_index(drop = True)

# NET ZERO
NEA_netz = netz_aggregate_df1[netz_aggregate_df1['REGION']\
    .isin(['06_HKC', '08_JPN', '09_ROK', '18_CT'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
NEA_netz['REGION'] = '23_NEA'

netz_aggregate_df1 = netz_aggregate_df1.append(NEA_netz).reset_index(drop = True)


# Aggregate results for 23b_ONEA
# ONEA: 06, 09, 18

# REFERENCE
ONEA_ref = ref_aggregate_df1[ref_aggregate_df1['REGION']\
    .isin(['06_HKC', '09_ROK', '18_CT'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
ONEA_ref['REGION'] = '23b_ONEA'

ref_aggregate_df1 = ref_aggregate_df1.append(ONEA_ref).reset_index(drop = True)

# NET ZERO
ONEA_netz = netz_aggregate_df1[netz_aggregate_df1['REGION']\
    .isin(['06_HKC', '09_ROK', '18_CT'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
ONEA_netz['REGION'] = '23b_ONEA'

netz_aggregate_df1 = netz_aggregate_df1.append(ONEA_netz).reset_index(drop = True)

# Aggregate results for 24_OAM
# OAM: 03, 04, 11, 14

# REFERENCE
OAM_ref = ref_aggregate_df1[ref_aggregate_df1['REGION']\
    .isin(['03_CDA', '04_CHL', '11_MEX', '14_PE'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
OAM_ref['REGION'] = '24_OAM'

ref_aggregate_df1 = ref_aggregate_df1.append(OAM_ref).reset_index(drop = True)

# NET ZERO
OAM_netz = netz_aggregate_df1[netz_aggregate_df1['REGION']\
    .isin(['03_CDA', '04_CHL', '11_MEX', '14_PE'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
OAM_netz['REGION'] = '24_OAM'

netz_aggregate_df1 = netz_aggregate_df1.append(OAM_netz).reset_index(drop = True)

# Aggregate results for 24b_OOAM
# OOAM: 04, 11, 14

# REFERENCE
OOAM_ref = ref_aggregate_df1[ref_aggregate_df1['REGION']\
    .isin(['04_CHL', '11_MEX', '14_PE'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
OOAM_ref['REGION'] = '24b_OOAM'

ref_aggregate_df1 = ref_aggregate_df1.append(OOAM_ref).reset_index(drop = True)

# NET ZERO
OOAM_netz = netz_aggregate_df1[netz_aggregate_df1['REGION']\
    .isin(['04_CHL', '11_MEX', '14_PE'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
OOAM_netz['REGION'] = '24b_OOAM'

netz_aggregate_df1 = netz_aggregate_df1.append(OOAM_netz).reset_index(drop = True)

# Aggregate results for 25_OCE
# Oceania: 01, 12, 13

# REFERENCE
OCE_ref = ref_aggregate_df1[ref_aggregate_df1['REGION']\
    .isin(['01_AUS', '12_NZ', '13_PNG'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
OCE_ref['REGION'] = '25_OCE'

ref_aggregate_df1 = ref_aggregate_df1.append(OCE_ref).reset_index(drop = True)

# NET ZERO
OCE_netz = netz_aggregate_df1[netz_aggregate_df1['REGION']\
    .isin(['01_AUS', '12_NZ', '13_PNG'])]\
        .groupby(['TECHNOLOGY', 'FUEL']).sum().reset_index()
OCE_netz['REGION'] = '25_OCE'

netz_aggregate_df1 = netz_aggregate_df1.append(OCE_netz).reset_index(drop = True)


# Get maximum REFERENCE year column to build data frame below
ref_year_columns = []

for item in list(ref_aggregate_df1.columns):
    try:
        ref_year_columns.append(int(item))
    except ValueError:
            pass

max_year_ref = max(ref_year_columns)

OSeMOSYS_years_ref = list(range(2017, max_year_ref + 1))

# Get maximum NET ZERO year column to build data frame below
netz_year_columns = []

for item in list(netz_aggregate_df1.columns):
    try:
        netz_year_columns.append(int(item))
    except ValueError:
            pass

max_year_netz = max(netz_year_columns)

OSeMOSYS_years_netz = list(range(2017, max_year_netz + 1))

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

# Then first level
tpes_agg = ['1_indigenous_production', '2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers']

tfc_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector', '17_nonenergy_use']

tfec_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector']

# For dataframe finalising
key_variables = ['economy', 'fuel_code', 'item_code_new']

#######################################################################################################################
# REFERENCE

# Now aggregate data based on the mapping
# That is group by REGION, TECHNOLOGY and FUEL
# First create empty dataframe
ref_aggregate_df2 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in ref_aggregate_df1['REGION'].unique():
    interim_df1 = ref_aggregate_df1[ref_aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Mapping_TFC_TPES, how = 'left', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['item_code_new', 'fuel_code']).sum().reset_index()

    # Change export data to negative values
    exports_bunkers = interim_df1[interim_df1['item_code_new'].isin(['3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers'])]\
        .set_index(['item_code_new', 'fuel_code'])
    everything_else = interim_df1[~interim_df1['item_code_new'].isin(['3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers'])]

    exports_bunkers = exports_bunkers * -1
    exports_bunkers = exports_bunkers.reset_index()
    interim_df2 = everything_else.append(exports_bunkers)

    ########################### Aggregate fuel_code for new variables ###################################

    # First level fuels

    coal = interim_df2[interim_df2['fuel_code'].isin(coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '1_coal').reset_index()

    oil = interim_df2[interim_df2['fuel_code'].isin(oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '6_crude_oil_and_ngl').reset_index()

    petrol = interim_df2[interim_df2['fuel_code'].isin(petrol_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '7_petroleum_products').reset_index()

    gas = interim_df2[interim_df2['fuel_code'].isin(gas_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '8_gas').reset_index()

    biomass = interim_df2[interim_df2['fuel_code'].isin(biomass_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '15_solid_biomass').reset_index()

    others = interim_df2[interim_df2['fuel_code'].isin(other_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '16_others').reset_index()

    interim_df3 = interim_df2.append([coal, oil, petrol, gas, biomass, others]).reset_index(drop = True)

    # And total fuels

    total_f = interim_df3[interim_df3['fuel_code'].isin(total_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '19_total').reset_index()

    interim_df4 = interim_df3.append(total_f).reset_index(drop = True)

    ################################ And now item_code_new ######################################

    # Start with lowest level

    industry = interim_df4[interim_df4['item_code_new'].isin(industry_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '14_industry_sector').reset_index()

    transport = interim_df4[interim_df4['item_code_new'].isin(transport_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '15_transport_sector').reset_index()

    bld_ag_other = interim_df4[interim_df4['item_code_new'].isin(others_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '16_other_sector').reset_index()

    interim_df5 = interim_df4.append([industry, transport, bld_ag_other]).reset_index(drop = True)

    # Now higher level agg

    #Might need to check this depending on whether exports is negative
    tpes = interim_df5[interim_df5['item_code_new'].isin(tpes_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '7_total_primary_energy_supply').reset_index()

    tfc = interim_df5[interim_df5['item_code_new'].isin(tfc_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '12_total_final_consumption').reset_index()

    tfec = interim_df5[interim_df5['item_code_new'].isin(tfec_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '13_total_final_energy_consumption').reset_index()

    interim_df6 = interim_df5.append([tpes, tfc, tfec]).reset_index(drop = True)

    # Now add in economy reference
    interim_df6['economy'] = region

    # Now append economy dataframe to communal data frame 
    ref_aggregate_df2 = ref_aggregate_df2.append(interim_df6)

# aggregate_df2 = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years]

if ref_aggregate_df2.empty:
    ref_aggregate_df2
else:
    ref_aggregate_df2 = ref_aggregate_df2.loc[:, key_variables + OSeMOSYS_years_ref]

#######################################################################################################################
# NET ZERO

# Now aggregate data based on the mapping
# That is group by REGION, TECHNOLOGY and FUEL
# First create empty dataframe
netz_aggregate_df2 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in netz_aggregate_df1['REGION'].unique():
    interim_df1 = netz_aggregate_df1[netz_aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Mapping_TFC_TPES, how = 'left', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['item_code_new', 'fuel_code']).sum().reset_index()

    # Change export data to negative values
    exports_bunkers = interim_df1[interim_df1['item_code_new'].isin(['3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers'])]\
        .set_index(['item_code_new', 'fuel_code'])
    everything_else = interim_df1[~interim_df1['item_code_new'].isin(['3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers'])]

    exports_bunkers = exports_bunkers * -1
    exports_bunkers = exports_bunkers.reset_index()
    interim_df2 = everything_else.append(exports_bunkers)

    ########################### Aggregate fuel_code for new variables ###################################

    # First level fuels

    coal = interim_df2[interim_df2['fuel_code'].isin(coal_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '1_coal').reset_index()

    oil = interim_df2[interim_df2['fuel_code'].isin(oil_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '6_crude_oil_and_ngl').reset_index()

    petrol = interim_df2[interim_df2['fuel_code'].isin(petrol_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '7_petroleum_products').reset_index()

    gas = interim_df2[interim_df2['fuel_code'].isin(gas_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '8_gas').reset_index()

    biomass = interim_df2[interim_df2['fuel_code'].isin(biomass_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '15_solid_biomass').reset_index()

    others = interim_df2[interim_df2['fuel_code'].isin(other_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '16_others').reset_index()

    interim_df3 = interim_df2.append([coal, oil, petrol, gas, biomass, others]).reset_index(drop = True)

    # And total fuels

    total_f = interim_df3[interim_df3['fuel_code'].isin(total_fuels)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = '19_total').reset_index()

    interim_df4 = interim_df3.append(total_f).reset_index(drop = True)

    ################################ And now item_code_new ######################################

    # Start with lowest level

    industry = interim_df4[interim_df4['item_code_new'].isin(industry_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '14_industry_sector').reset_index()

    transport = interim_df4[interim_df4['item_code_new'].isin(transport_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '15_transport_sector').reset_index()

    bld_ag_other = interim_df4[interim_df4['item_code_new'].isin(others_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '16_other_sector').reset_index()

    interim_df5 = interim_df4.append([industry, transport, bld_ag_other]).reset_index(drop = True)

    # Now higher level agg

    #Might need to check this depending on whether exports is negative
    tpes = interim_df5[interim_df5['item_code_new'].isin(tpes_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '7_total_primary_energy_supply').reset_index()

    tfc = interim_df5[interim_df5['item_code_new'].isin(tfc_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '12_total_final_consumption').reset_index()

    tfec = interim_df5[interim_df5['item_code_new'].isin(tfec_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '13_total_final_energy_consumption').reset_index()

    interim_df6 = interim_df5.append([tpes, tfc, tfec]).reset_index(drop = True)

    # Now add in economy reference
    interim_df6['economy'] = region

    # Now append economy dataframe to communal data frame 
    netz_aggregate_df2 = netz_aggregate_df2.append(interim_df6)

# aggregate_df2 = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years]
if netz_aggregate_df2.empty == True:
    netz_aggregate_df2
else:
    netz_aggregate_df2 = netz_aggregate_df2.loc[:, key_variables + OSeMOSYS_years_netz]

# Now load the EGEDA_years data frame
EGEDA_years = pd.read_csv('./data/1_EGEDA/EGEDA_2018_years.csv')

# REFERENCE
if ref_aggregate_df2.empty == True:
    ref_aggregate_df2_tojoin = ref_aggregate_df2.copy()
else:
    ref_aggregate_df2_tojoin = ref_aggregate_df2.copy().loc[:, key_variables + OSeMOSYS_years_ref]

# NET ZERO
if netz_aggregate_df2.empty == True:
    netz_aggregate_df2_tojoin = netz_aggregate_df2.copy()
else:
    netz_aggregate_df2_tojoin = netz_aggregate_df2.copy().loc[:, key_variables + OSeMOSYS_years_netz]

# Join EGEDA historical to OSeMOSYS results (line below removes 2017 and 2018 from historical)
# REFERENCE
if ref_aggregate_df2_tojoin.empty == True:
    Joined_ref_df = EGEDA_years.copy().reindex(columns = EGEDA_years.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_ref_df = EGEDA_years.copy().iloc[:, :-2].merge(ref_aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_ref_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_2018_reference.csv', index = False)

# NET ZERO
if netz_aggregate_df2_tojoin.empty == True:
    Joined_netz_df = EGEDA_years.copy().reindex(columns = EGEDA_years.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_netz_df = EGEDA_years.copy().iloc[:, :-2].merge(netz_aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_netz_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_2018_netzero.csv', index = False)

###############################################################################################################################

# Moving beyond TFC and TPES and Transformation

# Determine list of files to read based on the workbooks identified in the mapping file
# REFERENCE
ref_file_trans = pd.DataFrame()

for i in range(len(Unique_trans['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in reference_filenames if Unique_trans['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_trans['Workbook'].unique()[i]})
    ref_file_trans = ref_file_trans.append(_file)

ref_file_trans = ref_file_trans.merge(Unique_trans, how = 'outer', on = 'Workbook')

# NET ZERO
netz_file_trans = pd.DataFrame()

for i in range(len(Unique_trans['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in netzero_filenames if Unique_trans['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_trans['Workbook'].unique()[i]})
    netz_file_trans = netz_file_trans.append(_file)

netz_file_trans = netz_file_trans.merge(Unique_trans, how = 'outer', on = 'Workbook')

# Create empty dataframe to store aggregated results 
# REFERENCE

ref_aggregate_df1 = pd.DataFrame()

# Now read in the OSeMOSYS output files so that that they're all in one data frame (aggregate_df1)

for i in range(ref_file_trans.shape[0]):
    _df = pd.read_excel(ref_file_trans.iloc[i, 0], sheet_name = ref_file_trans.iloc[i, 2])
    _df['Workbook'] = ref_file_trans.iloc[i, 1]
    _df['Sheet_energy'] = ref_file_trans.iloc[i, 2]
    ref_aggregate_df1 = ref_aggregate_df1.append(_df) 

ref_aggregate_df1 = ref_aggregate_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION']).sum().reset_index()

# NET ZERO

netz_aggregate_df1 = pd.DataFrame()

# Now read in the OSeMOSYS output files so that that they're all in one data frame (aggregate_df1)

for i in range(netz_file_trans.shape[0]):
    _df = pd.read_excel(netz_file_trans.iloc[i, 0], sheet_name = netz_file_trans.iloc[i, 2])
    _df['Workbook'] = netz_file_trans.iloc[i, 1]
    _df['Sheet_energy'] = netz_file_trans.iloc[i, 2]
    netz_aggregate_df1 = netz_aggregate_df1.append(_df) 

netz_aggregate_df1 = netz_aggregate_df1.groupby(['TECHNOLOGY', 'FUEL', 'REGION']).sum().reset_index()

# Read in capacity data
# REFERENCE
ref_capacity_df1 = pd.DataFrame()

# Populate the above blank dataframe with capacity data from the results workbook

for i in range(len(reference_filenames)):
    _df = pd.read_excel(reference_filenames[i], sheet_name = 'TotalCapacityAnnual')
    ref_capacity_df1 = ref_capacity_df1.append(_df)

# Now just extract the power capacity

ref_pow_capacity_df1 = ref_capacity_df1[ref_capacity_df1['TECHNOLOGY'].str.startswith('POW')].reset_index(drop = True)

# NET ZERO
netz_capacity_df1 = pd.DataFrame()

# Populate the above blank dataframe with capacity data from the results workbook

for i in range(len(netzero_filenames)):
    _df = pd.read_excel(netzero_filenames[i], sheet_name = 'TotalCapacityAnnual')
    netz_capacity_df1 = netz_capacity_df1.append(_df)

# Now just extract the power capacity

netz_pow_capacity_df1 = netz_capacity_df1[netz_capacity_df1['TECHNOLOGY'].str.startswith('POW')].reset_index(drop = True)

# Get maximum year column to build data frame below
# REFERENCE
ref_year_columns = []

for item in list(ref_aggregate_df1.columns):
    try:
        ref_year_columns.append(int(item))
    except ValueError:
            pass

max_year_ref = min(2050, max(ref_year_columns))

OSeMOSYS_years_ref = list(range(2017, max_year_ref + 1))

# NET ZERO
netz_year_columns = []

for item in list(netz_aggregate_df1.columns):
    try:
        netz_year_columns.append(int(item))
    except ValueError:
            pass

max_year_netz = min(2050, max(netz_year_columns))

OSeMOSYS_years_netz = list(range(2017, max_year_netz + 1))

#################################################################################################

# Now create the dataframes to save and use in the later bossanova script

################################ POWER SECTOR ############################### 

# Aggregate data based on the Map_power mapping

# That is group by REGION, TECHNOLOGY and FUEL

# First create empty dataframe
# REFERENCE
ref_power_df1 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in ref_aggregate_df1['REGION'].unique():
    interim_df1 = ref_aggregate_df1[ref_aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Map_power, how = 'right', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector']).sum().reset_index()

    # Now add in economy reference
    interim_df1['economy'] = region
    
    # Now append economy dataframe to communal data frame 
    ref_power_df1 = ref_power_df1.append(interim_df1)
    
ref_power_df1 = ref_power_df1[['economy', 'TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector'] + OSeMOSYS_years_ref]

# NET ZERO
netz_power_df1 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in netz_aggregate_df1['REGION'].unique():
    interim_df1 = netz_aggregate_df1[netz_aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Map_power, how = 'right', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector']).sum().reset_index()

    # Now add in economy reference
    interim_df1['economy'] = region
    
    # Now append economy dataframe to communal data frame 
    netz_power_df1 = netz_power_df1.append(interim_df1)
    
netz_power_df1 = netz_power_df1[['economy', 'TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector'] + OSeMOSYS_years_netz]

################################ REFINERY, OWN USE and SUPPLY TRANSFORMATION SECTOR ############################### 

# Aggregate data based on REGION, TECHNOLOGY and FUEL

# First create empty dataframe
# REFERENCE
ref_refownsup_df1 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in ref_aggregate_df1['REGION'].unique():
    interim_df1 = ref_aggregate_df1[ref_aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Map_refownsup, how = 'right', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector']).sum().reset_index()

    # Now add in economy reference
    interim_df1['economy'] = region
    
    # Now append economy dataframe to communal data frame 
    ref_refownsup_df1 = ref_refownsup_df1.append(interim_df1)
    
ref_refownsup_df1 = ref_refownsup_df1[['economy', 'TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector'] + OSeMOSYS_years_ref]

# REFERENCE
netz_refownsup_df1 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in netz_aggregate_df1['REGION'].unique():
    interim_df1 = netz_aggregate_df1[netz_aggregate_df1['REGION'] == region]
    interim_df1 = interim_df1.merge(Map_refownsup, how = 'right', on = ['TECHNOLOGY', 'FUEL'])
    interim_df1 = interim_df1.groupby(['TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector']).sum().reset_index()

    # Now add in economy reference
    interim_df1['economy'] = region
    
    # Now append economy dataframe to communal data frame 
    netz_refownsup_df1 = netz_refownsup_df1.append(interim_df1)
    
netz_refownsup_df1 = netz_refownsup_df1[['economy', 'TECHNOLOGY', 'FUEL', 'Sheet_energy', 'Sector'] + OSeMOSYS_years_netz]

# Refinery, own-use, supply and power

ref_trans_df1 = ref_power_df1.append(ref_refownsup_df1)
netz_trans_df1 = netz_power_df1.append(netz_refownsup_df1)

# Save the required dataframes for transformation charts in bossanova script

# Reference
ref_power_df1.to_csv(path_final + '/OSeMOSYS_power_reference.csv', index = False)
ref_refownsup_df1.to_csv(path_final + '/OSeMOSYS_refownsup_reference.csv', index = False)
ref_pow_capacity_df1.to_csv(path_final + '/OSeMOSYS_powcapacity_reference.csv', index = False)
ref_trans_df1.to_csv(path_final + '/OSeMOSYS_transformation_reference.csv', index = False)

# Net-zero
netz_power_df1.to_csv(path_final + '/OSeMOSYS_power_netzero.csv', index = False)
netz_refownsup_df1.to_csv(path_final + '/OSeMOSYS_refownsup_netzero.csv', index = False)
netz_pow_capacity_df1.to_csv(path_final + '/OSeMOSYS_powcapacity_netzero.csv', index = False)
netz_trans_df1.to_csv(path_final + '/OSeMOSYS_transformation_netzero.csv', index = False)

# Dataframes for demand sectors

# Save OSeMOSYS results dataframes 

ref_aggregate_df1.to_csv(path_final + '/OSeMOSYS_only_reference.csv', index = False)
netz_aggregate_df1.to_csv(path_final + '/OSeMOSYS_only_netzero.csv', index = False)

# # Macro dataframes (opens in Bossanova)

# macro_GDP = pd.read_excel(path_mapping + '/Key Inputs.xlsx', sheet_name = 'GDP')
# macro_GDP.columns = macro_GDP.columns.astype(str) 
# macro_GDP['Series'] = 'GDP 2018 USD PPP'
# macro_GDP = macro_GDP[['Economy', 'Series'] + list(macro_GDP.loc[:, '2000':'2050'])]
# macro_GDP = macro_GDP[macro_GDP['Economy'].isin(list(macro_GDP['Economy'].unique()))]
# macro_GDP.to_csv(path_final + '/macro_GDP.csv', index = False)

# macro_GDP_growth = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'GDP_growth')
# macro_GDP_growth.columns = macro_GDP_growth.columns.astype(str) 
# macro_GDP_growth['Series'] = 'GDP growth'
# macro_GDP_growth = macro_GDP_growth[['Economy', 'Series'] + list(macro_GDP_growth.loc[:, '2000':'2050'])]

# macro_pop = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'Population')
# macro_pop.columns = macro_pop.columns.astype(str) 
# macro_pop['Series'] = 'Population'
# macro_pop = macro_pop[['Economy', 'Series'] + list(macro_pop.loc[:, '2000':'2050'])]

# macro_GDPpc = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'GDP per capita')
# macro_GDPpc.columns = macro_GDPpc.columns.astype(str)
# macro_GDPpc['Series'] = 'GDP per capita' 
# macro_GDPpc = macro_GDPpc[['Economy', 'Series'] + list(macro_GDPpc.loc[:, '2000':'2050'])]

################################################################################################

# EMISSIONS

EGEDA_emissions = pd.read_csv('./data/1_EGEDA/EGEDA_FC_CO2_Emissions_years_2018.csv')

agg_fuel = ['1_coal', '1_x_coal_thermal', '2_coal_products', '6_crude_oil_and_ngl', '6_x_ngls',
            '7_petroleum_products', '7_x_jet_fuel', '7_x_other_petroleum_products', '8_gas', '16_others', '19_total']

EGEDA_emissions = EGEDA_emissions[~EGEDA_emissions['fuel_code'].isin(agg_fuel)].reset_index(drop = True)

########################## fuel_code aggregations ##########################

# lowest level

thermal_coal = ['1_2_other_bituminous_coal', '1_3_subbituminous_coal', '1_4_anthracite', '3_peat', '4_peat_products']
ngl = ['6_2_natural_gas_liquids', '6_3_refinery_feedstocks', '6_4_additives_oxygenates', '6_5_other_hydrocarbons']
other_petrol = ['7_12_white_spirit_sbp', '7_13_lubricants', '7_14_bitumen', '7_15_paraffin_waxes', '7_16_petroleum_coke', '7_17_other_products']
jetfuel = ['7_4_gasoline_type_jet_fuel', '7_5_kerosene_type_jet_fuel']

# First level and Total vetor(s) defined at beginning of script

coal_prod_fuels = ['2_1_coke_oven_coke', '2_2_coke_oven_gas', '2_3_blast_furnace_gas', '2_4_other_recovered_gases', '2_5_patent_fuel', '2_6_coal_tar', '2_7_bkb_pb']
power_agg = ['9_1_main_activity_producer', '9_2_autoproducers']

# Change from negative to positive

neg_to_pos = ['9_x_power',
              '9_1_main_activity_producer', '9_1_1_electricity_plants', '9_1_2_chp_plants', '9_1_3_heat_plants', '9_2_autoproducers',
              '9_2_1_electricity_plants', '9_2_2_chp_plants', '9_2_3_heat_plants', '9_3_gas_processing_plants', '9_3_1_gas_works_plants',
              '9_3_2_liquefaction_plants', '9_3_3_regasification_plants', '9_3_4_natural_gas_blending_plants', '9_3_5_gastoliquids_plants',
              '9_4_oil_refineries', '9_5_coal_transformation', '9_5_1_coke_ovens', '9_5_2_blast_furnaces', '9_5_3_patent_fuel_plants',
              '9_5_4_bkb_pb_plants', '9_5_5_liquefaction_coal_to_oil', '9_6_petrochemical_industry', '9_7_biofuels_processing', 
              '9_8_charcoal_processing', '9_9_nonspecified_transformation', '10_losses_and_own_use']

# Aggregations for Emissions dataframe

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

    # Totals by sector aggregation
    
    power_total = interim_df4[interim_df4['item_code_new'].isin(power_agg)].groupby(['fuel_code'])\
        .sum().assign(item_code_new = '9_x_power').reset_index()

    # tfc = interim_df5[interim_df5['item_code_new'].isin(tfc_agg)].groupby(['fuel_code'])\
    #     .sum().assign(item_code_new = '12_total_final_consumption').reset_index()

    # tfec = interim_df5[interim_df5['item_code_new'].isin(tfec_agg)].groupby(['fuel_code'])\
    #     .sum().assign(item_code_new = '13_total_final_energy_consumption').reset_index()

    interim_df5 = interim_df4.append([power_total]).reset_index(drop = True)                                      

    interim_df5['economy'] = region

    EGEDA_aggregate = EGEDA_aggregate.append(interim_df5).reset_index(drop = True)

# Now change main activity producer and own use to positive

change_to_negative = EGEDA_aggregate[EGEDA_aggregate['item_code_new'].\
    isin(neg_to_pos)].copy().reset_index(drop = True)
everything_else = EGEDA_aggregate[~EGEDA_aggregate['item_code_new'].\
    isin(neg_to_pos)].copy().reset_index(drop = True)

s = change_to_negative.select_dtypes(include=[np.number]) * -1 
change_to_negative[s.columns] = s

EGEDA_aggregate = everything_else.append(change_to_negative).reset_index(drop = True)

# Aggregate for demand sectors, power and own use and losses (EMISSIONS)

EGEDA_aggregate2 = pd.DataFrame()

for region in EGEDA_aggregate['economy'].unique():
    interim_df1 = EGEDA_aggregate[EGEDA_aggregate['economy'] == region]

    dem_pow_own = interim_df1[interim_df1['item_code_new']\
        .isin(['9_x_power', '10_losses_and_own_use', '13_total_final_energy_consumption'])]\
            .groupby(['fuel_code'])\
                .sum().assign(item_code_new = '13_x_dem_pow_own').reset_index()

    interim_df2 = interim_df1.append([dem_pow_own]).reset_index(drop = True)

    interim_df2['economy'] = region

    EGEDA_aggregate2 = EGEDA_aggregate2.append(interim_df2).reset_index(drop = True)    

# Load correct order of fuel code and item code. Update this csv based on new entries or desired order

ordered = pd.read_csv('./data/2_Mapping_and_other/emissions_order_2018.csv')

# This grabs the unique values of fuel_code and item_code_new in the order they appear in the original dataframe. It removes 'na' by calling '[:-1]' 

order1 = list(ordered['fuel_code'].unique())[:-1]
order2 = list(ordered['item_code_new'])

# Take order defined above and define each of the variables as categorical in that already established order (for the benefit of viewing data later)

EGEDA_aggregate2['fuel_code'] = pd.Categorical(EGEDA_aggregate2['fuel_code'], 
                                                categories = order1, 
                                                ordered = True)

EGEDA_aggregate2['item_code_new'] = pd.Categorical(EGEDA_aggregate2['item_code_new'],
                                                    categories = order2,
                                                    ordered = True)

EGEDA_aggregate_sorted = EGEDA_aggregate2.sort_values(['economy', 'fuel_code', 'item_code_new']).reset_index(drop = True)

# Write emissions file
EGEDA_aggregate_sorted.to_csv('./data/1_EGEDA/EGEDA_2018_emissions.csv', index = False)

###############################################################################################

# Now grab the OSeMOSYS output emissions

# Now moving everything from OSeMOSYS to EGEDA (Only demand sectors and own use for now)

Mapping_file_emiss = Mapping_file[Mapping_file['Sector'].isin(['AGR', 'BLD', 'IND', 'TRN', 'PIP', 'NON', 'OWN', 'POW'])].copy() 
Mapping_file_emiss = Mapping_file_emiss[Mapping_file_emiss['item_code_new'].notna()].copy().reset_index(drop = True)

# Define unique workbook and sheet combinations
Unique_combo = Mapping_file_emiss.groupby(['Workbook', 'Sheet_emissions']).size().reset_index().loc[:, ['Workbook', 'Sheet_emissions']]

# Determine list of files to read based on the workbooks identified in the mapping file
# REFERENCE
ref_file_emiss_df = pd.DataFrame()

for i in range(len(Unique_combo['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in reference_filenames if Unique_combo['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_combo['Workbook'].unique()[i]})
    ref_file_emiss_df = ref_file_emiss_df.append(_file)

ref_file_emiss_df = ref_file_emiss_df.merge(Unique_combo, how = 'outer', on = 'Workbook')

# NET ZERO
netz_file_emiss_df = pd.DataFrame()

for i in range(len(Unique_combo['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in netzero_filenames if Unique_combo['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_combo['Workbook'].unique()[i]})
    netz_file_emiss_df = netz_file_emiss_df.append(_file)

netz_file_emiss_df = netz_file_emiss_df.merge(Unique_combo, how = 'outer', on = 'Workbook')

# Create empty dataframe to store aggregated results 
# REFERENCE
ref_aggemiss_df1 = pd.DataFrame()

# Now read in the OSeMOSYS output files so that that they're all in one data frame (aggregate_df1)
for i in range(ref_file_emiss_df.shape[0]):
    _df = pd.read_excel(ref_file_emiss_df.iloc[i, 0], sheet_name = ref_file_emiss_df.iloc[i, 2])
    _df['Workbook'] = ref_file_emiss_df.iloc[i, 1]
    _df['Sheet_emissions'] = ref_file_emiss_df.iloc[i, 2]
    ref_aggemiss_df1 = ref_aggemiss_df1.append(_df) 

# Now aggregate all the results for APEC

APEC_ref_emiss = ref_aggemiss_df1.groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()
APEC_ref_emiss['REGION'] = 'APEC'

ref_aggemiss_df1 = ref_aggemiss_df1.append(APEC_ref_emiss).reset_index(drop = True)

# Create empty dataframe to store aggregated results 
# NET ZERO
netz_aggemiss_df1 = pd.DataFrame()

# Now read in the OSeMOSYS output files so that that they're all in one data frame (aggregate_df1)
for i in range(netz_file_emiss_df.shape[0]):
    _df = pd.read_excel(netz_file_emiss_df.iloc[i, 0], sheet_name = netz_file_emiss_df.iloc[i, 2])
    _df['Workbook'] = netz_file_emiss_df.iloc[i, 1]
    _df['Sheet_emissions'] = netz_file_emiss_df.iloc[i, 2]
    netz_aggemiss_df1 = netz_aggemiss_df1.append(_df) 

# Now aggregate all the results for APEC

APEC_netz_emiss = netz_aggemiss_df1.groupby(['TECHNOLOGY', 'EMISSION']).sum().reset_index()
APEC_netz_emiss['REGION'] = 'APEC'

netz_aggemiss_df1 = netz_aggemiss_df1.append(APEC_netz_emiss).reset_index(drop = True)

# Get maximum year column to build data frame below
# REFERENCE
ref_year_columns = []

for item in list(ref_aggemiss_df1.columns):
    try:
        ref_year_columns.append(int(item))
    except ValueError:
            pass

ref_max_year = max(ref_year_columns)

OSeMOSYS_years_ref = list(range(2017, ref_max_year + 1))

# NET ZERO
netz_year_columns = []

for item in list(netz_aggemiss_df1.columns):
    try:
        netz_year_columns.append(int(item))
    except ValueError:
            pass

netz_max_year = max(netz_year_columns)

OSeMOSYS_years_netz = list(range(2017, netz_max_year + 1))

# Vector not defined above

dem_pow_own_agg = ['9_x_power', '10_losses_and_own_use', '13_total_final_energy_consumption']

###################################################################################################

# Now aggregate data based on the mapping
# That is group by REGION, TECHNOLOGY and EMISSION
# First create empty dataframe
# REFERENCE
ref_aggemiss_df2 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in ref_aggemiss_df1['REGION'].unique():
    interim_df1 = ref_aggemiss_df1[ref_aggemiss_df1['REGION'] == region]
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
    ref_aggemiss_df2 = ref_aggemiss_df2.append(interim_df6)

#############################################################################################################

# NET ZERO
netz_aggemiss_df2 = pd.DataFrame()

# Then loop through based on different regions/economies and stitch back together

for region in netz_aggemiss_df1['REGION'].unique():
    interim_df1 = netz_aggemiss_df1[netz_aggemiss_df1['REGION'] == region]
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
    netz_aggemiss_df2 = netz_aggemiss_df2.append(interim_df6)

# Finalise dataframes
# Already defined above: key_variables = ['economy', 'fuel_code', 'item_code_new']

# aggregate_df2 = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years]

if ref_aggemiss_df2.empty:
    ref_aggemiss_df2
else:
    ref_aggemiss_df2 = ref_aggemiss_df2.loc[:, key_variables + OSeMOSYS_years_ref]

if netz_aggemiss_df2.empty:
    netz_aggemiss_df2
else:
    netz_aggemiss_df2 = netz_aggemiss_df2.loc[:, key_variables + OSeMOSYS_years_netz]

# Now load the EGEDA_2018_emissions data frame
EGEDA_emissions = pd.read_csv('./data/1_EGEDA/EGEDA_2018_emissions.csv')

# Join EGEDA historical to OSeMOSYS results (line below removes 2017 and 2018 from historical)
# REFERENCE
if ref_aggemiss_df2.empty:
    Joined_emissdf_ref = EGEDA_emissions.reindex(columns = EGEDA_emissions.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_emissdf_ref = EGEDA_emissions.iloc[:, :-2].merge(ref_aggemiss_df2, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_emissdf_ref.to_csv(path_final + '/OSeMOSYS_to_EGEDA_emissions_2018_reference.csv', index = False)

# NET ZERO
if netz_aggemiss_df2.empty:
    Joined_emissdf_netz = EGEDA_emissions.reindex(columns = EGEDA_emissions.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_emissdf_netz = EGEDA_emissions.iloc[:, :-2].merge(netz_aggemiss_df2, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_emissdf_netz.to_csv(path_final + '/OSeMOSYS_to_EGEDA_emissions_2018_netzero.csv', index = False)


print('Requisite dataframes created and saved ready for Bossanova script')