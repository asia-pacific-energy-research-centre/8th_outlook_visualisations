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
path_final = './data/5_CCS_grab'

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

# # Moving everything from OSeMOSYS to EGEDA for TFC and TPES

# Mapping_TFC_TPES = Mapping_file[Mapping_file['Balance'].isin(['TFC', 'TPES'])]

# # And for transformation

# Map_trans = Mapping_file[Mapping_file['Balance'] == 'TRANS'].reset_index(drop = True)

# # A mapping just for i) power, ii) ref, own, sup and iii) hydrogen

# Map_power = Map_trans[Map_trans['Sector'] == 'POW'].reset_index(drop = True)
# Map_refownsup = Map_trans[Map_trans['Sector'].isin(['REF', 'SUP', 'OWN', 'HYD'])].reset_index(drop = True)
# Map_hydrogen = Map_trans[Map_trans['Sector'] == 'HYD'].reset_index(drop = True)

# # Define unique workbook and sheet combinations for TFC and TPES
# Unique_TFC_TPES = Mapping_TFC_TPES.groupby(['Workbook', 'Sheet_energy']).size().reset_index().loc[:, ['Workbook', 'Sheet_energy']]

# # Define unique workbook and sheet combinations for Transformation
# Unique_trans = Map_trans.groupby(['Workbook', 'Sheet_energy']).size().reset_index().loc[:, ['Workbook', 'Sheet_energy']]



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
tpes_agg = ['1_indigenous_production', '2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers', '6_stock_change']

tfc_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector', '17_nonenergy_use']

tfec_agg = ['14_industry_sector', '15_transport_sector', '16_other_sector']

# For dataframe finalising
key_variables = ['economy', 'fuel_code', 'item_code_new']


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



###############################################################################################

# Now grab the OSeMOSYS output emissions

# Now moving everything from OSeMOSYS to EGEDA (Only demand sectors and own use for now)

Mapping_file_emiss = Mapping_file[Mapping_file['Sector'].isin(['AGR', 'BLD', 'IND', 'TRN', 'PIP', 'NON', 'OWN', 'POW', 'HYD'])].copy() 
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

dem_pow_own_hyd_agg = ['9_x_power', '9_x_hydrogen', '10_losses_and_own_use', '13_total_final_energy_consumption']

ref_ccs_own_ind_1 = ref_aggemiss_df1[(ref_aggemiss_df1['EMISSION'].str.contains('_captured'))].copy().reset_index(drop = True)
cn_ccs_own_ind_1 = netz_aggemiss_df1[(netz_aggemiss_df1['EMISSION'].str.contains('_captured'))].copy().reset_index(drop = True)

ref_ccs_own_ind_1.to_csv(path_final + '/ccs_own_industry_reference.csv', index = False)
cn_ccs_own_ind_1.to_csv(path_final + '/ccs_own_industry_carbonneutrality.csv', index = False)

print('Requisite data saved for captured CCS emissions for industry and own-use')
