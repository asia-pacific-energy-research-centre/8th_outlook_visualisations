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

# Now moving everything from OSeMOSYS to EGEDA only requires TFC and TPES for now

Mapping_file = Mapping_file[Mapping_file['Balance'].isin(['TFC', 'TPES'])]

# Define unique workbook and sheet combinations
Unique_combo = Mapping_file.groupby(['Workbook', 'Sheet_energy']).size().reset_index().loc[:, ['Workbook', 'Sheet_energy']]

# Determine list of files to read based on the workbooks identified in the mapping file for REFERENCE scenario
ref_file_df = pd.DataFrame()

for i in range(len(Unique_combo['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in reference_filenames if Unique_combo['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_combo['Workbook'].unique()[i]})
    ref_file_df = ref_file_df.append(_file)

ref_file_df = ref_file_df.merge(Unique_combo, how = 'outer', on = 'Workbook')

# Determine list of files to read based on the workbooks identified in the mapping file for NET-ZERO scenario
netz_file_df = pd.DataFrame()

for i in range(len(Unique_combo['Workbook'].unique())):
    _file = pd.DataFrame({'File': [entry for entry in netzero_filenames if Unique_combo['Workbook'].unique()[i] in entry],
                         'Workbook': Unique_combo['Workbook'].unique()[i]})
    netz_file_df = netz_file_df.append(_file)

netz_file_df = netz_file_df.merge(Unique_combo, how = 'outer', on = 'Workbook')

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

# Now read in the OSeMOSYS output files so that that they're all in one data frame (ref_aggregate_df1)
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
OSeMOSYS_check_ref = list(range(2019, max_year_ref + 1))

# Get maximum NET ZERO year column to build data frame below
netz_year_columns = []

for item in list(netz_aggregate_df1.columns):
    try:
        netz_year_columns.append(int(item))
    except ValueError:
            pass

max_year_netz = max(netz_year_columns)

OSeMOSYS_years_netz = list(range(2017, max_year_netz + 1))
OSeMOSYS_years_netz = list(range(2019, max_year_netz + 1))

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
    interim_df1 = interim_df1.merge(Mapping_file, how = 'left', on = ['TECHNOLOGY', 'FUEL'])
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
    interim_df1 = interim_df1.merge(Mapping_file, how = 'left', on = ['TECHNOLOGY', 'FUEL'])
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

########################### Special amendment to EGEDA historical ############################

# # HONG KONG imported electricity shifts to both Nuclear and Hydro in a 75:25 split

# # Hydro

# EGEDA_hkc_hydro = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '10_hydro') &
#                                 (EGEDA_years['item_code_new'] == '2_imports')].copy()

# EGEDA_hkc_elec_imports = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '17_electricity') &
#                                 (EGEDA_years['item_code_new'] == '2_imports')].copy()

# EGEDA_hkc_elec_imports[EGEDA_hkc_elec_imports.select_dtypes(include = ['number']).columns] *= 0.25

# append1 = EGEDA_hkc_hydro.append(EGEDA_hkc_elec_imports).reset_index(drop = True)
# d = append1.dtypes

# # locate index of row to overwrite

# append1.loc[EGEDA_hkc_hydro.index[0]] = append1.sum(numeric_only = True)
# append1.astype(d)

# append1.loc[EGEDA_hkc_hydro.index[0], 'economy'] = append1.loc[0, 'economy']
# append1.loc[EGEDA_hkc_hydro.index[0], 'fuel_code'] = append1.loc[0, 'fuel_code']
# append1.loc[EGEDA_hkc_hydro.index[0], 'item_code_new'] = append1.loc[0, 'item_code_new']

# append1 = append1.drop([0, 1])

# EGEDA_years = append1.combine_first(EGEDA_years)

# # Nuclear

# EGEDA_hkc_nuclear = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '9_nuclear') &
#                                 (EGEDA_years['item_code_new'] == '2_imports')].copy()

# EGEDA_hkc_elec_imports = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '17_electricity') &
#                                 (EGEDA_years['item_code_new'] == '2_imports')].copy()

# EGEDA_hkc_elec_imports[EGEDA_hkc_elec_imports.select_dtypes(include = ['number']).columns] *= 0.75

# append2 = EGEDA_hkc_nuclear.append(EGEDA_hkc_elec_imports).reset_index(drop = True)
# d = append2.dtypes

# append2.loc[EGEDA_hkc_nuclear.index[0]] = append2.sum(numeric_only = True)
# append2.astype(d)

# append2.loc[EGEDA_hkc_nuclear.index[0], 'economy'] = append2.loc[0, 'economy']
# append2.loc[EGEDA_hkc_nuclear.index[0], 'fuel_code'] = append2.loc[0, 'fuel_code']
# append2.loc[EGEDA_hkc_nuclear.index[0], 'item_code_new'] = append2.loc[0, 'item_code_new']

# append2 = append2.drop([0, 1])

# EGEDA_years = append2.combine_first(EGEDA_years)

# # Now change electricity imports to zero

# EGEDA_hkc_elec_imports = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '17_electricity') &
#                                 (EGEDA_years['item_code_new'] == '2_imports')].copy()

# EGEDA_hkc_elec_imports[EGEDA_hkc_elec_imports.select_dtypes(include = ['number']).columns] *= 0

# EGEDA_years = EGEDA_hkc_elec_imports.combine_first(EGEDA_years)

# # Now amend TPES category 6 to reflect changes to hydro, nuclear (TPES should now equal indigenous prod) and electricity (TPES should only equal exports)

# # Hydro TPES
# EGEDA_hkc_hydro_tpes = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '10_hydro') &
#                                 (EGEDA_years['item_code_new'].isin(['1_indigenous_production', '2_imports', '3_exports']))].copy()

# EGEDA_hkc_hydro_tpes = EGEDA_hkc_hydro_tpes.append(EGEDA_hkc_hydro_tpes.sum(numeric_only = True), ignore_index = True)

# EGEDA_hkc_hydro_tpes.loc[3, 'economy'] = '06_HKC'
# EGEDA_hkc_hydro_tpes.loc[3, 'fuel_code'] = '10_hydro'
# EGEDA_hkc_hydro_tpes.loc[3, 'item_code_new'] = '7_total_primary_energy_supply'

# # Find index for hydro TPES hkc

# index1 = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '10_hydro') &
#                                 (EGEDA_years['item_code_new'] == '7_total_primary_energy_supply')].copy().index[0]

# EGEDA_hkc_hydro_tpes = EGEDA_hkc_hydro_tpes.rename(index = {3: index1})

# EGEDA_hkc_hydro_tpes = EGEDA_hkc_hydro_tpes.drop([0, 1, 2])

# EGEDA_years = EGEDA_hkc_hydro_tpes.combine_first(EGEDA_years)

# # Nuclear TPES
# EGEDA_hkc_nuclear_tpes = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '9_nuclear') &
#                                 (EGEDA_years['item_code_new'].isin(['1_indigenous_production', '2_imports', '3_exports']))].copy()

# EGEDA_hkc_nuclear_tpes = EGEDA_hkc_nuclear_tpes.append(EGEDA_hkc_nuclear_tpes.sum(numeric_only = True), ignore_index = True)

# EGEDA_hkc_nuclear_tpes.loc[3, 'economy'] = '06_HKC'
# EGEDA_hkc_nuclear_tpes.loc[3, 'fuel_code'] = '9_nuclear'
# EGEDA_hkc_nuclear_tpes.loc[3, 'item_code_new'] = '7_total_primary_energy_supply'

# # Find index for nuclear TPES hkc

# index2 = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '9_nuclear') &
#                                 (EGEDA_years['item_code_new'] == '7_total_primary_energy_supply')].copy().index[0]

# EGEDA_hkc_nuclear_tpes = EGEDA_hkc_nuclear_tpes.rename(index = {3: index2})

# EGEDA_hkc_nuclear_tpes = EGEDA_hkc_nuclear_tpes.drop([0, 1, 2])

# EGEDA_years = EGEDA_hkc_nuclear_tpes.combine_first(EGEDA_years)

# # Electricity TPES
# EGEDA_hkc_elec_tpes = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '17_electricity') &
#                                 (EGEDA_years['item_code_new'].isin(['1_indigenous_production', '2_imports', '3_exports']))].copy()

# EGEDA_hkc_elec_tpes = EGEDA_hkc_elec_tpes.append(EGEDA_hkc_elec_tpes.sum(numeric_only = True), ignore_index = True)

# EGEDA_hkc_elec_tpes.loc[3, 'economy'] = '06_HKC'
# EGEDA_hkc_elec_tpes.loc[3, 'fuel_code'] = '17_electricity'
# EGEDA_hkc_elec_tpes.loc[3, 'item_code_new'] = '7_total_primary_energy_supply'

# # Find index for elec TPES hkc

# index3 = EGEDA_years[(EGEDA_years['economy'] == '06_HKC') & 
#                                 (EGEDA_years['fuel_code'] == '17_electricity') &
#                                 (EGEDA_years['item_code_new'] == '7_total_primary_energy_supply')].copy().index[0]

# EGEDA_hkc_elec_tpes = EGEDA_hkc_elec_tpes.rename(index = {3: index3})

# EGEDA_hkc_elec_tpes = EGEDA_hkc_elec_tpes.drop([0, 1, 2])

# EGEDA_years = EGEDA_hkc_elec_tpes.combine_first(EGEDA_years)

# Remove 2017 which is already in the EGEDA historical
# aggregate_df2_tojoin = aggregate_df2[['economy', 'fuel_code', 'item_code_new'] + OSeMOSYS_years[1:]]
# aggregate_df2_tojoin = aggregate_df2.loc[:, key_variables + OSeMOSYS_years[1:]] # New line below keeps 2017 in OSeMOSYS

# REFERENCE
if ref_aggregate_df2.empty == True:
    ref_aggregate_df2_tojoin = ref_aggregate_df2
else:
    ref_aggregate_df2_tojoin = ref_aggregate_df2.loc[:, key_variables + OSeMOSYS_years_ref]

# NET ZERO
if netz_aggregate_df2.empty == True:
    netz_aggregate_df2_tojoin = netz_aggregate_df2
else:
    netz_aggregate_df2_tojoin = netz_aggregate_df2.loc[:, key_variables + OSeMOSYS_years_netz]

# Join EGEDA historical to OSeMOSYS results (line below removes 2017 and 2018 from historical)
# REFERENCE
if ref_aggregate_df2_tojoin.empty == True:
    Joined_ref_df = EGEDA_years.reindex(columns = EGEDA_years.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_ref_df = EGEDA_years.iloc[:, :-2].merge(ref_aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_ref_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_2018_reference.csv', index = False)

# NET ZERO
if netz_aggregate_df2_tojoin.empty == True:
    Joined_netz_df = EGEDA_years.reindex(columns = EGEDA_years.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_netz_df = EGEDA_years.iloc[:, :-2].merge(netz_aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_netz_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_2018_netzero.csv', index = False)

###################### FOR CHECKING #########################################

# Join EGEDA historical to OSeMOSYS results (line below removes 2017 and 2018 from historical)
# REFERENCE
if ref_aggregate_df2_tojoin.empty == True:
    Joined_ref_df = EGEDA_years.reindex(columns = EGEDA_years.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_ref_df = EGEDA_years.iloc[:, :].merge(ref_aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_ref_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_2018_reference_CHECK.csv', index = False)

# NET ZERO
if netz_aggregate_df2_tojoin.empty == True:
    Joined_netz_df = EGEDA_years.reindex(columns = EGEDA_years.columns.tolist() + list(range(2019, 2051)))
else:
    Joined_netz_df = EGEDA_years.iloc[:, :].merge(netz_aggregate_df2_tojoin, on = ['economy', 'fuel_code', 'item_code_new'], how = 'left')

Joined_netz_df.to_csv(path_final + '/OSeMOSYS_to_EGEDA_2018_netzero_CHECK.csv', index = False)

print('OSeMOSYS_to_EGEDA_2018_reference.csv and OSeMOSYS_to_EGEDA_2018_netzero.csv file successfully created')





