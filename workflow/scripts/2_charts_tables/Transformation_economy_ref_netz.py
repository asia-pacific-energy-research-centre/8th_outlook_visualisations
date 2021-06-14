# Charting OSeMOSYS transformation data
# These charts won't necessarily need to be mapped back to EGEDA historical.
# Will effectively be base year and out
# But will be good to incorporate some historical generation before the base year eventually

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

# They're csv files so use a wild card (*) to grab the filenames

OSeMOSYS_filenames = glob.glob(path_output + "/*.xlsx")

# Reference filenames and net zero filenames

reference_filenames = list(filter(lambda k: 'reference' in k, OSeMOSYS_filenames))
netzero_filenames = list(filter(lambda y: 'net-zero' in y, OSeMOSYS_filenames))

# Read in mapping file

# New 2018 data variable names 

Mapping_sheets = list(pd.read_excel(path_mapping + '/OSeMOSYS_mapping_2021.xlsx', sheet_name = None).keys())[1:]

Mapping_file = pd.DataFrame()

for sheet in Mapping_sheets:
    interim_map = pd.read_excel(path_mapping + '/OSeMOSYS_mapping_2021.xlsx', sheet_name = sheet, skiprows = 1)
    Mapping_file = Mapping_file.append(interim_map).reset_index(drop = True)

# Subset the mapping file so that it's just transformation

Map_trans = Mapping_file[Mapping_file['Balance'] == 'TRANS'].reset_index(drop = True)

# Define unique workbook and sheet combinations

Unique_trans = Map_trans.groupby(['Workbook', 'Sheet_energy']).size().reset_index().loc[:, ['Workbook', 'Sheet_energy']]

########################################################################################################################
########################### Create historical electricity generation dataframe for use later ###########################

required_fuels_elec = ['1_coal', '1_5_lignite', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products', 
                  '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', 
                  '15_solid_biomass', '16_others', '18_heat']

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

########################################################################################################################

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

# Colours for charting (to be amended later)

colours = pd.read_excel('./data/2_Mapping_and_other/colour_template_7th.xlsx')
colours_hex = colours['hex']

# Colour dictionary
colours_dict = {
    'Coal': '#323232',
    'Oil': '#be280a',
    'Gas': '#f59300',
    'Modern renewables': '#3c7896',
    'Traditional biomass': '#828282',
    'Hydrogen': '#28825a',
    'Electricity': '#a5cdf0',
    'Heat': '#cd6477',
    'Others': '#bebebe',
    'Industry': '#ffc305',
    'Transport': '#bebebe',
    'Buildings': '#3c7896',
    'Agriculture': '#323232',
    'Non-energy': '#cd6477',
    'Non-specified': '#872355',
    'Services': '#a5cdf0',
    'Residential': '#28825a',
    'Iron & steel': '#8c0000',
    'Chemicals': '#a5cdf0',
    'Aluminium': '#bebebe',
    'Non-metallic minerals': '#1e465a',
    'Mining': '#f59300',
    'Pulp & paper': '#28825a',
    'Other': '#cd6477',
    'Biomass': '#828282',
    'Jet fuel': '#323232',
    'LPG': '#ffdc96',
    'Gasoline': '#be280a',
    'Diesel': '#3c7896',
    'Renewables': '#1e465a',
    'Aviation': '#ffc305',
    'Road': '#1e465a',
    'Rail': '#be280a',
    'Marine': '#28825a',
    'Pipeline': '#bebebe',
    # Transformation unique
    'Geothermal': '#3c7896',
    'Hydro': '#a5cdf0',
    'Lignite': '#833C0C',
    'Nuclear': '#872355',
    'Other renewables': '#1e465a',
    'Solar': '#ffc305',
    'Wind': '#28825a',
    'Storage': '#ffdc96',
    'Imports': '#641964',
    'Crude oil': '#be280a',
    'NGLs': '#3c7896',
    'Motor gasoline': '#1e465a',
    'Aviation gasoline': '#3c7896',
    'Naphtha': '#a5cdf0',
    'Other kerosene': '#8c0000',
    'Gas diesel oil': '#be280a',
    'Fuel oil': '#f59300',
    'Refinery gas': '#ffc305',
    'Ethane': '#872355',
    'Power': '#1e465a',
    'Refining': '#3c7896'
    }

Map_power = Map_trans[Map_trans['Sector'] == 'POW'].reset_index(drop = True)

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

Map_refownsup = Map_trans[Map_trans['Sector'].isin(['REF', 'SUP', 'OWN'])].reset_index(drop = True)

# Aggregate data based on the Map_power mapping

# That is group by REGION, TECHNOLOGY and FUEL

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

# FUEL aggregations for UseByTechnology

# First aggregation (13 fuels)
coal_fuel_1 = ['1_x_coal_thermal', '2_coal_products']
lignite_fuel_1 = ['1_5_lignite']
oil_fuel_1 = ['7_7_gas_diesel_oil','7_3_naphtha', '7_8_fuel_oil', '6_1_crude_oil', '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_x_other_petroleum_products']
gas_fuel_1 = ['8_1_natural_gas']
nuclear_fuel_1 = ['9_nuclear']
hydro_fuel_1 = ['10_hydro']
solar_fuel_1 = ['12_1_of_which_photovoltaics']
wind_fuel_1 = ['14_wind']
biomass_fuel_1 = ['15_1_fuelwood_and_woodwaste', '15_2_bagasse', '15_4_black_liquor', '15_5_other_biomass']
geothermal_fuel_1 = ['11_geothermal']
other_renew_fuel_1 = ['13_tide_wave_ocean', '16_3_municipal_solid_waste_renewable', '16_1_biogas']
other_fuel_1 = ['16_4_municipal_solid_waste_nonrenewable', '17_electricity', '18_heat', '16_x_hydrogen', '16_2_industrial_waste']
imports_fuel_1 = ['17_electricity_import']

# Second aggreagtion: Oil, Gas, Nuclear, Imports, Other from above and below two new aggregations (7 fuels)
coal_fuel_2 = ['1_x_coal_thermal', '1_5_lignite', '2_coal_products']
renewables_fuel_2 = ['10_hydro', '11_geothermal', '12_1_of_which_photovoltaics', '13_tide_wave_ocean', '14_wind', '15_1_fuelwood_and_woodwaste', 
                     '15_2_bagasse', '15_4_black_liquor', '15_5_other_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable']

# Own use fuels
coal_ou = ['1_x_coal_thermal', '1_5_lignite', '2_coal_products', '1_1_coking_coal']
oil_ou = ['6_1_crude_oil', '6_x_ngls', '7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha', '7_6_kerosene',
          '7_7_gas_diesel_oil', '7_8_fuel_oil', '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_11_ethane',
          '7_x_jet_fuel', '7_x_other_petroleum_products']
gas_ou = ['8_1_natural_gas']
renew_ou = ['15_1_fuelwood_and_woodwaste', '15_2_bagasse', '15_3_charcoal', '15_4_black_liquor', '15_5_other_biomass', 
            '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', '16_6_biodiesel',
            '16_8_other_liquid_biofuels']
elec_ou = ['17_electricity']
heat_ou = ['18_heat']
other_ou = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable']

own_use_fuels = ['Coal', 'Oil', 'Gas', 'Renewables', 'Electricity', 'Heat', 'Other']

# Note, 12_1_of_which_photovoltaics is a subset of 12_solar so including will lead to double counting

use_agg_fuels_1 = ['Coal', 'Lignite', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Solar', 'Wind', 
                   'Biomass', 'Geothermal', 'Other renewables', 'Other', 'Imports']
use_agg_fuels_2 = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Other', 'Imports']

# TECHNOLOGY aggregations for ProductionByTechnology

coal_tech = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']
oil_tech = ['POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP', 'POW_PetCoke_PP']
gas_tech = ['POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP', 'POW_CCGT_CCS_PP']
nuclear_tech = ['POW_Nuclear_PP', 'POW_IMP_Nuclear_PP']
hydro_tech = ['POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP']
solar_tech = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP']
wind_tech = ['POW_WindOff_PP', 'POW_Wind_PP']
bio_tech = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP']
geo_tech = ['POW_Geothermal_PP']
storage_tech = ['POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage']
other_tech = ['POW_IPP_PP', 'POW_TIDAL_PP', 'POW_WasteToEnergy_PP', 'POW_CHP_PP']
# chp_tech = ['POW_CHP_PP']
im_tech = ['POW_IMPORTS_PP', 'POW_IMPORT_ELEC_PP']

lignite_tech = ['POW_Sub_Brown_PP']
thermal_coal_tech = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']
solar_roof_tech = ['POW_SolarRoofPV_PP']
solar_nr_tech = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP']


# POW_EXPORT_ELEC_PP need to work this in

prod_agg_tech = ['Coal', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Wind', 'Solar', 'Biomass', 'Geothermal', 'Storage', 'Other', 'Imports']
prod_agg_tech2 = ['Coal', 'Lignite', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Wind', 'Solar', 
                 'Biomass', 'Geothermal', 'Storage', 'Other', 'Imports']

# Refinery vectors

refinery_input = ['6_1_crude_oil', '6_x_ngls']
refinery_output = ['7_1_motor_gasoline', '7_2_aviation_gasoline', '7_3_naphtha', '7_x_jet_fuel', '7_6_kerosene', '7_7_gas_diesel_oil', '7_8_fuel_oil',
              '7_9_lpg', '7_10_refinery_gas_not_liquefied', '7_11_ethane', '7_x_other_petroleum_products']

refinery_new_output = ['7_1_from_ref', '7_2_from_ref', '7_3_from_ref', '7_jet_from_ref', '7_6_from_ref', '7_7_from_ref',
                       '7_8_from_ref', '7_9_from_ref', '7_10_from_ref', '7_11_from_ref', '7_other_from_ref']

# Capacity vectors
    
coal_cap = ['POW_Black_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_CHP_COAL_PP', 'POW_Other_Coal_PP', 'POW_Ultra_BituCoal_PP', 'POW_Ultra_CHP_PP']
gas_cap = ['POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP', 'POW_CCGT_CCS_PP']
oil_cap = ['POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP', 'POW_PetCoke_PP']
nuclear_cap = ['POW_Nuclear_PP', 'POW_IMP_Nuclear_PP']
hydro_cap = ['POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP']
bio_cap = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP']
wind_cap = ['POW_Wind_PP', 'POW_WindOff_PP']
solar_cap = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP']
geo_cap = ['POW_Geothermal_PP']
storage_cap = ['POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage']
other_cap = ['POW_WasteToEnergy_PP', 'POW_IPP_PP', 'POW_TIDAL_PP', 'POW_CHP_PP']
# chp_cap = ['POW_CHP_PP']
# 'POW_HEAT_HP' not in electricity capacity
transmission_cap = ['POW_Transmission']

lignite_cap = ['POW_Sub_Brown_PP']
thermal_coal_cap = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']


pow_capacity_agg = ['Coal', 'Gas', 'Oil', 'Nuclear', 'Hydro', 'Biomass', 'Wind', 'Solar', 'Geothermal', 'Storage', 'Other']
pow_capacity_agg2 = ['Coal', 'Lignite', 'Gas', 'Oil', 'Nuclear', 'Hydro', 'Biomass', 'Wind', 
                     'Solar', 'Geothermal', 'Storage', 'Other']

# Chart years for column charts

col_chart_years = [2018, 2020, 2030, 2040, 2050]
gen_col_chart_years = [2000, 2010, 2018, 2020, 2030, 2040, 2050]

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written

# TRANSFORMATION SECTOR: Build use, capacity and production dataframes with appropriate aggregations to chart

for economy in ref_power_df1['economy'].unique():
    ref_use_df1 = ref_power_df1[(ref_power_df1['economy'] == economy) &
                        (ref_power_df1['Sheet_energy'] == 'UseByTechnology') &
                        (ref_power_df1['TECHNOLOGY'] != 'POW_Transmission')].reset_index(drop = True)

    # Now build aggregate variables of the FUELS

    # First level aggregations
    coal = ref_use_df1[ref_use_df1['FUEL'].isin(coal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    lignite = ref_use_df1[ref_use_df1['FUEL'].isin(lignite_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Lignite',
                                                                                              TECHNOLOGY = 'Lignite power')                                                                                      

    oil = ref_use_df1[ref_use_df1['FUEL'].isin(oil_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Oil power')

    gas = ref_use_df1[ref_use_df1['FUEL'].isin(gas_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Gas',
                                                                                      TECHNOLOGY = 'Gas power')

    nuclear = ref_use_df1[ref_use_df1['FUEL'].isin(nuclear_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Nuclear',
                                                                                    TECHNOLOGY = 'Nuclear power')

    hydro = ref_use_df1[ref_use_df1['FUEL'].isin(hydro_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Hydro',
                                                                                    TECHNOLOGY = 'Hydro power')

    solar = ref_use_df1[ref_use_df1['FUEL'].isin(solar_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Solar',
                                                                                        TECHNOLOGY = 'Solar power')

    wind = ref_use_df1[ref_use_df1['FUEL'].isin(wind_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Wind',
                                                                                    TECHNOLOGY = 'Wind power')

    geothermal = ref_use_df1[ref_use_df1['FUEL'].isin(geothermal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Geothermal',
                                                                                    TECHNOLOGY = 'Geothermal power')

    biomass = ref_use_df1[ref_use_df1['FUEL'].isin(biomass_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Biomass',
                                                                                    TECHNOLOGY = 'Biomass power')

    other_renew = ref_use_df1[ref_use_df1['FUEL'].isin(other_renew_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other renewables',
                                                                                    TECHNOLOGY = 'Other renewable power')

    other = ref_use_df1[ref_use_df1['FUEL'].isin(other_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other',
                                                                                        TECHNOLOGY = 'Other power')

    imports = ref_use_df1[ref_use_df1['FUEL'].isin(imports_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Imports',
                                                                                        TECHNOLOGY = 'Electricity imports')                                                                                         

    # Second level aggregations

    coal2 = ref_use_df1[ref_use_df1['FUEL'].isin(coal_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    renew2 = ref_use_df1[ref_use_df1['FUEL'].isin(renewables_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Renewables',
                                                                                      TECHNOLOGY = 'Renewables power')

    # Use by fuel data frame number 1

    ref_usefuel_df1 = ref_use_df1.append([coal, lignite, oil, gas, nuclear, hydro, solar, wind, geothermal, biomass, other_renew, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + OSeMOSYS_years_ref].reset_index(drop = True)

    ref_usefuel_df1 = ref_usefuel_df1[ref_usefuel_df1['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    ref_usefuel_df1 = ref_usefuel_df1.groupby('FUEL').sum().reset_index()
    ref_usefuel_df1['Transformation'] = 'Input fuel'
    ref_usefuel_df1['FUEL'] = pd.Categorical(ref_usefuel_df1['FUEL'], use_agg_fuels_1)

    ref_usefuel_df1 = ref_usefuel_df1.sort_values('FUEL').reset_index(drop = True)

    ref_usefuel_df1 = ref_usefuel_df1[['FUEL', 'Transformation'] + OSeMOSYS_years_ref]

    nrows1 = ref_usefuel_df1.shape[0]
    ncols1 = ref_usefuel_df1.shape[1]

    ref_usefuel_df2 = ref_usefuel_df1[['FUEL', 'Transformation'] + col_chart_years]

    nrows2 = ref_usefuel_df2.shape[0]
    ncols2 = ref_usefuel_df2.shape[1]

    # Use by fuel data frame number 1

    ref_usefuel_df3 = ref_use_df1.append([coal2, oil, gas, nuclear, renew2, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + OSeMOSYS_years_ref].reset_index(drop = True)

    ref_usefuel_df3 = ref_usefuel_df3[ref_usefuel_df3['FUEL'].isin(use_agg_fuels_2)].copy().set_index('FUEL').reset_index() 

    ref_usefuel_df3 = ref_usefuel_df3.groupby('FUEL').sum().reset_index()
    ref_usefuel_df3['Transformation'] = 'Input fuel'
    ref_usefuel_df3 = ref_usefuel_df3[['FUEL', 'Transformation'] + OSeMOSYS_years_ref]

    nrows10 = ref_usefuel_df3.shape[0]
    ncols10 = ref_usefuel_df3.shape[1]

    ref_usefuel_df4 = ref_usefuel_df3[['FUEL', 'Transformation'] + col_chart_years]

    nrows11 = ref_usefuel_df4.shape[0]
    ncols11 = ref_usefuel_df4.shape[1]

    # Now build production dataframe
    ref_prodelec_df1 = ref_power_df1[(ref_power_df1['economy'] == economy) &
                             (ref_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                             (ref_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx']))].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    oil_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    storage_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    # chp_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Cogeneration')
    nuclear_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    other_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    geo_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(geo_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Geothermal')
    misc = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')

    coal_pp2 = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(thermal_coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_pp2 = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(lignite_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    roof_pp2 = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(solar_roof_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar roof')
    nonroof_pp = ref_prodelec_df1[ref_prodelec_df1['TECHNOLOGY'].isin(solar_nr_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_prodelec_bytech_df1 = ref_prodelec_df1.append([coal_pp2, lignite_pp2, oil_pp, gas_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + OSeMOSYS_years_ref].reset_index(drop = True)                                                                                                    

    ref_prodelec_bytech_df1['Generation'] = 'Electricity'
    ref_prodelec_bytech_df1 = ref_prodelec_bytech_df1[['TECHNOLOGY', 'Generation'] + OSeMOSYS_years_ref] 

    ref_prodelec_bytech_df1 = ref_prodelec_bytech_df1[ref_prodelec_bytech_df1['TECHNOLOGY'].isin(prod_agg_tech2)].\
        set_index('TECHNOLOGY')

    ref_prodelec_bytech_df1 = ref_prodelec_bytech_df1.loc[ref_prodelec_bytech_df1.index.intersection(prod_agg_tech2)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_gen[EGEDA_hist_gen['economy'] == economy].copy().\
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(range(2000, 2017))]

    ref_prodelec_bytech_df1 = historical_gen.merge(ref_prodelec_bytech_df1, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    ref_prodelec_bytech_df1['TECHNOLOGY'] = pd.Categorical(ref_prodelec_bytech_df1['TECHNOLOGY'], prod_agg_tech2)

    ref_prodelec_bytech_df1 = ref_prodelec_bytech_df1.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = ref_prodelec_bytech_df1.select_dtypes(include=[np.number]) / 3.6 
    ref_prodelec_bytech_df1[s.columns] = s

    nrows3 = ref_prodelec_bytech_df1.shape[0]
    ncols3 = ref_prodelec_bytech_df1.shape[1]

    ref_prodelec_bytech_df2 = ref_prodelec_bytech_df1[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    nrows4 = ref_prodelec_bytech_df2.shape[0]
    ncols4 = ref_prodelec_bytech_df2.shape[1]

    ##################################################################################################################################################################

    # Now create some refinery dataframes

    ref_refinery_df1 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    ref_refinery_df1['Transformation'] = 'Input to refinery'
    ref_refinery_df1 = ref_refinery_df1[['FUEL', 'Transformation'] + OSeMOSYS_years_ref].reset_index(drop = True)

    ref_refinery_df1.loc[ref_refinery_df1['FUEL'] == '6_1_crude_oil', 'FUEL'] = 'Crude oil'
    ref_refinery_df1.loc[ref_refinery_df1['FUEL'] == '6_x_ngls', 'FUEL'] = 'NGLs'

    nrows5 = ref_refinery_df1.shape[0]
    ncols5 = ref_refinery_df1.shape[1]

    ref_refinery_df2 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_new_output))].copy()

    ref_refinery_df2['Transformation'] = 'Output from refinery'
    ref_refinery_df2 = ref_refinery_df2[['FUEL', 'Transformation'] + OSeMOSYS_years_ref].reset_index(drop = True)

    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_1_from_ref', 'FUEL'] = 'Motor gasoline'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_2_from_ref', 'FUEL'] = 'Aviation gasoline'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_3_from_ref', 'FUEL'] = 'Naphtha'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_jet_from_ref', 'FUEL'] = 'Jet fuel'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_6_from_ref', 'FUEL'] = 'Other kerosene'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_7_from_ref', 'FUEL'] = 'Gas diesel oil'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_8_from_ref', 'FUEL'] = 'Fuel oil'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_9_from_ref', 'FUEL'] = 'LPG'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_10_from_ref', 'FUEL'] = 'Refinery gas'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_11_from_ref', 'FUEL'] = 'Ethane'
    ref_refinery_df2.loc[ref_refinery_df2['FUEL'] == '7_other_from_ref', 'FUEL'] = 'Other'

    ref_refinery_df2['FUEL'] = pd.Categorical(
        ref_refinery_df2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other'], 
        ordered = True)

    ref_refinery_df2 = ref_refinery_df2.sort_values('FUEL')

    nrows6 = ref_refinery_df2.shape[0]
    ncols6 = ref_refinery_df2.shape[1]

    ref_refinery_df3 = ref_refinery_df2[['FUEL', 'Transformation'] + col_chart_years]

    nrows7 = ref_refinery_df3.shape[0]
    ncols7 = ref_refinery_df3.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    ref_powcap_df1 = ref_pow_capacity_df1[ref_pow_capacity_df1['REGION'] == economy]

    coal_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    oil_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    hydro_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Biomass')
    geo_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(geo_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Geothermal')
    #chp_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(chp_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Cogeneration')
    other_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')

    lignite_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(lignite_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Lignite')
    thermal_capacity = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(thermal_coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')

    # Capacity by tech dataframe (with the above aggregations added)

    ref_powcap_df1 = ref_powcap_df1.append([coal_capacity, gas_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, other_capacity])\
        [['TECHNOLOGY'] + OSeMOSYS_years_ref].reset_index(drop = True) 

    ref_powcap_df1 = ref_powcap_df1[ref_powcap_df1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    ref_powcap_df1['TECHNOLOGY'] = pd.Categorical(ref_powcap_df1['TECHNOLOGY'], prod_agg_tech[:-1])

    ref_powcap_df1 = ref_powcap_df1.sort_values('TECHNOLOGY').reset_index(drop = True)

    nrows8 = ref_powcap_df1.shape[0]
    ncols8 = ref_powcap_df1.shape[1]

    ref_powcap_df2 = ref_powcap_df1[['TECHNOLOGY'] + col_chart_years]

    nrows9 = ref_powcap_df2.shape[0]
    ncols9 = ref_powcap_df2.shape[1]

    #########################################################################################################################################
    ############ NEW DATAFRAMES #############################################################################################################

    # Refining, supply and own-use, and power
    # SHould this include POW_Transmission
    ref_transformation_df1 = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                           (ref_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                           (ref_trans_df1['TECHNOLOGY'] != 'POW_Transmission')]

    ref_transmission1 = ref_trans_df1[(ref_trans_df1['economy'] == economy) &
                                     (ref_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                     (ref_trans_df1['TECHNOLOGY'] == 'POW_Transmission')]

    ref_transmission1 = ref_transmission1.groupby('Sector').sum().copy().reset_index()
    ref_transmission1.loc[ref_transmission1['Sector'] == 'POW', 'Sector'] = 'Transmission'

    ref_transformation_sector = ref_transformation_df1.groupby('Sector').sum().copy().reset_index().append(ref_transmission1)

    ref_transformation_sector.loc[ref_transformation_sector['Sector'] == 'OWN', 'Sector'] = 'Own-use'
    ref_transformation_sector.loc[ref_transformation_sector['Sector'] == 'POW', 'Sector'] = 'Power'
    ref_transformation_sector.loc[ref_transformation_sector['Sector'] == 'REF', 'Sector'] = 'Refining'

    ref_transformation_sector1 = ref_transformation_sector[ref_transformation_sector['Sector'].isin(['Power', 'Refining'])]\
        .reset_index(drop = True)

    nrows12 = ref_transformation_sector1.shape[0]
    ncols12 = ref_transformation_sector1.shape[1]

    ref_transformation_sector2 = ref_transformation_sector1[['Sector'] + col_chart_years]

    nrows13 = ref_transformation_sector2.shape[0]
    ncols13 = ref_transformation_sector2.shape[1]

    # Own-use
    ref_ownuse_df1 = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                   (ref_trans_df1['Sector'] == 'OWN')]

    coal_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(coal_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Coal', Sector = 'Own-use and losses')
    oil_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(oil_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Oil', Sector = 'Own-use and losses')
    gas_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(gas_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Gas', Sector = 'Own-use and losses')
    renewables_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(renew_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Renewables', Sector = 'Own-use and losses')
    elec_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(elec_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Electricity', Sector = 'Own-use and losses')
    heat_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(heat_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Heat', Sector = 'Own-use and losses')
    other_own = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(other_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Other', Sector = 'Own-use and losses')

    ref_ownuse_df1 = ref_ownuse_df1.append([coal_own, oil_own, gas_own, renewables_own, elec_own, heat_own, other_own])\
        [['FUEL', 'Sector'] + OSeMOSYS_years_ref].reset_index(drop = True)

    ref_ownuse_df1 = ref_ownuse_df1[ref_ownuse_df1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    nrows14 = ref_ownuse_df1.shape[0]
    ncols14 = ref_ownuse_df1.shape[1]

    ref_ownuse_df2 = ref_ownuse_df1[['FUEL', 'Sector'] + col_chart_years]

    nrows15 = ref_ownuse_df2.shape[0]
    ncols15 = ref_ownuse_df2.shape[1]

    ######################################################################################################################
    # Net zero dataframes
    netz_use_df1 = netz_power_df1[(netz_power_df1['economy'] == economy) &
                        (netz_power_df1['Sheet_energy'] == 'UseByTechnology') &
                        (netz_power_df1['TECHNOLOGY'] != 'POW_Transmission')].reset_index(drop = True)

    # Now build aggregate variables of the FUELS

    # First level aggregations
    coal = netz_use_df1[netz_use_df1['FUEL'].isin(coal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    lignite = netz_use_df1[netz_use_df1['FUEL'].isin(lignite_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Lignite',
                                                                                              TECHNOLOGY = 'Lignite power')                                                                                      

    oil = netz_use_df1[netz_use_df1['FUEL'].isin(oil_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Oil power')

    gas = netz_use_df1[netz_use_df1['FUEL'].isin(gas_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Gas',
                                                                                      TECHNOLOGY = 'Gas power')

    nuclear = netz_use_df1[netz_use_df1['FUEL'].isin(nuclear_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Nuclear',
                                                                                    TECHNOLOGY = 'Nuclear power')

    hydro = netz_use_df1[netz_use_df1['FUEL'].isin(hydro_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Hydro',
                                                                                    TECHNOLOGY = 'Hydro power')

    solar = netz_use_df1[netz_use_df1['FUEL'].isin(solar_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Solar',
                                                                                        TECHNOLOGY = 'Solar power')

    wind = netz_use_df1[netz_use_df1['FUEL'].isin(wind_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Wind',
                                                                                    TECHNOLOGY = 'Wind power')

    geothermal = netz_use_df1[netz_use_df1['FUEL'].isin(geothermal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Geothermal',
                                                                                    TECHNOLOGY = 'Geothermal power')

    biomass = netz_use_df1[netz_use_df1['FUEL'].isin(biomass_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Biomass',
                                                                                    TECHNOLOGY = 'Biomass power')

    other_renew = netz_use_df1[netz_use_df1['FUEL'].isin(other_renew_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other renewables',
                                                                                    TECHNOLOGY = 'Other renewable power')

    other = netz_use_df1[netz_use_df1['FUEL'].isin(other_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other',
                                                                                        TECHNOLOGY = 'Other power')

    imports = netz_use_df1[netz_use_df1['FUEL'].isin(imports_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Imports',
                                                                                        TECHNOLOGY = 'Electricity imports')                                                                                         

    # Second level aggregations

    coal2 = netz_use_df1[netz_use_df1['FUEL'].isin(coal_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    renew2 = netz_use_df1[netz_use_df1['FUEL'].isin(renewables_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Renewables',
                                                                                      TECHNOLOGY = 'Renewables power')

    # Use by fuel data frame number 1

    netz_usefuel_df1 = netz_use_df1.append([coal, lignite, oil, gas, nuclear, hydro, solar, wind, geothermal, biomass, other_renew, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + OSeMOSYS_years_netz].reset_index(drop = True)

    netz_usefuel_df1 = netz_usefuel_df1[netz_usefuel_df1['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index() 

    netz_usefuel_df1 = netz_usefuel_df1.groupby('FUEL').sum().reset_index()
    netz_usefuel_df1['Transformation'] = 'Input fuel'
    netz_usefuel_df1['FUEL'] = pd.Categorical(netz_usefuel_df1['FUEL'], use_agg_fuels_1)

    netz_usefuel_df1 = netz_usefuel_df1.sort_values('FUEL').reset_index(drop = True)
    
    netz_usefuel_df1 = netz_usefuel_df1[['FUEL', 'Transformation'] + OSeMOSYS_years_netz]

    nrows21 = netz_usefuel_df1.shape[0]
    ncols21 = netz_usefuel_df1.shape[1]

    netz_usefuel_df2 = netz_usefuel_df1[['FUEL', 'Transformation'] + col_chart_years]

    nrows22 = netz_usefuel_df2.shape[0]
    ncols22 = netz_usefuel_df2.shape[1]

    # Use by fuel data frame number 1

    netz_usefuel_df3 = netz_use_df1.append([coal2, oil, gas, nuclear, renew2, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + OSeMOSYS_years_netz].reset_index(drop = True)

    netz_usefuel_df3 = netz_usefuel_df3[netz_usefuel_df3['FUEL'].isin(use_agg_fuels_2)].copy().set_index('FUEL').reset_index() 

    netz_usefuel_df3 = netz_usefuel_df3.groupby('FUEL').sum().reset_index()
    netz_usefuel_df3['Transformation'] = 'Input fuel'
    netz_usefuel_df3 = netz_usefuel_df3[['FUEL', 'Transformation'] + OSeMOSYS_years_netz]

    nrows30 = netz_usefuel_df3.shape[0]
    ncols30 = netz_usefuel_df3.shape[1]

    netz_usefuel_df4 = netz_usefuel_df3[['FUEL', 'Transformation'] + col_chart_years]

    nrows31 = netz_usefuel_df4.shape[0]
    ncols31 = netz_usefuel_df4.shape[1]

    # Now build production dataframe
    netz_prodelec_df1 = netz_power_df1[(netz_power_df1['economy'] == economy) &
                             (netz_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                             (netz_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx']))].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    oil_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    storage_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    # chp_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Cogeneration')
    nuclear_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    other_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    geo_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(geo_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Geothermal')
    misc = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')

    coal_pp2 = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(thermal_coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_pp2 = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(lignite_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    roof_pp2 = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(solar_roof_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar roof')
    nonroof_pp = netz_prodelec_df1[netz_prodelec_df1['TECHNOLOGY'].isin(solar_nr_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_prodelec_bytech_df1 = netz_prodelec_df1.append([coal_pp2, lignite_pp2, oil_pp, gas_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + OSeMOSYS_years_netz].reset_index(drop = True)                                                                                                    

    netz_prodelec_bytech_df1['Generation'] = 'Electricity'
    netz_prodelec_bytech_df1 = netz_prodelec_bytech_df1[['TECHNOLOGY', 'Generation'] + OSeMOSYS_years_netz] 

    netz_prodelec_bytech_df1 = netz_prodelec_bytech_df1[netz_prodelec_bytech_df1['TECHNOLOGY'].isin(prod_agg_tech2)].\
        set_index('TECHNOLOGY')

    netz_prodelec_bytech_df1 = netz_prodelec_bytech_df1.loc[netz_prodelec_bytech_df1.index.intersection(prod_agg_tech2)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_gen[EGEDA_hist_gen['economy'] == economy].copy().\
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(range(2000, 2017))]

    netz_prodelec_bytech_df1 = historical_gen.merge(netz_prodelec_bytech_df1, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    netz_prodelec_bytech_df1['TECHNOLOGY'] = pd.Categorical(netz_prodelec_bytech_df1['TECHNOLOGY'], prod_agg_tech2)

    netz_prodelec_bytech_df1 = netz_prodelec_bytech_df1.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = netz_prodelec_bytech_df1.select_dtypes(include=[np.number]) / 3.6 
    netz_prodelec_bytech_df1[s.columns] = s

    nrows23 = netz_prodelec_bytech_df1.shape[0]
    ncols23 = netz_prodelec_bytech_df1.shape[1]

    netz_prodelec_bytech_df2 = netz_prodelec_bytech_df1[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    nrows24 = netz_prodelec_bytech_df2.shape[0]
    ncols24 = netz_prodelec_bytech_df2.shape[1]

    ##################################################################################################################################################################

    # Now create some refinery dataframes

    netz_refinery_df1 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                 (netz_refownsup_df1['Sector'] == 'REF') & 
                                 (netz_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    netz_refinery_df1['Transformation'] = 'Input to refinery'
    netz_refinery_df1 = netz_refinery_df1[['FUEL', 'Transformation'] + OSeMOSYS_years_netz].reset_index(drop = True)

    netz_refinery_df1.loc[netz_refinery_df1['FUEL'] == '6_1_crude_oil', 'FUEL'] = 'Crude oil'
    netz_refinery_df1.loc[netz_refinery_df1['FUEL'] == '6_x_ngls', 'FUEL'] = 'NGLs'

    nrows25 = netz_refinery_df1.shape[0]
    ncols25 = netz_refinery_df1.shape[1]

    netz_refinery_df2 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                 (netz_refownsup_df1['Sector'] == 'REF') & 
                                 (netz_refownsup_df1['FUEL'].isin(refinery_new_output))].copy()

    netz_refinery_df2['Transformation'] = 'Output from refinery'
    netz_refinery_df2 = netz_refinery_df2[['FUEL', 'Transformation'] + OSeMOSYS_years_netz].reset_index(drop = True)

    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_1_from_ref', 'FUEL'] = 'Motor gasoline'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_2_from_ref', 'FUEL'] = 'Aviation gasoline'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_3_from_ref', 'FUEL'] = 'Naphtha'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_jet_from_ref', 'FUEL'] = 'Jet fuel'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_6_from_ref', 'FUEL'] = 'Other kerosene'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_7_from_ref', 'FUEL'] = 'Gas diesel oil'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_8_from_ref', 'FUEL'] = 'Fuel oil'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_9_from_ref', 'FUEL'] = 'LPG'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_10_from_ref', 'FUEL'] = 'Refinery gas'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_11_from_ref', 'FUEL'] = 'Ethane'
    netz_refinery_df2.loc[netz_refinery_df2['FUEL'] == '7_other_from_ref', 'FUEL'] = 'Other'

    netz_refinery_df2['FUEL'] = pd.Categorical(
        netz_refinery_df2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other'], 
        ordered = True)

    netz_refinery_df2 = netz_refinery_df2.sort_values('FUEL')

    nrows26 = netz_refinery_df2.shape[0]
    ncols26 = netz_refinery_df2.shape[1]

    netz_refinery_df3 = netz_refinery_df2[['FUEL', 'Transformation'] + col_chart_years]

    nrows27 = netz_refinery_df3.shape[0]
    ncols27 = netz_refinery_df3.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    netz_powcap_df1 = netz_pow_capacity_df1[netz_pow_capacity_df1['REGION'] == economy]

    coal_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    oil_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    hydro_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Biomass')
    geo_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(geo_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Geothermal')
    #chp_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(chp_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Cogeneration')
    other_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')

    lignite_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(lignite_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Lignite')
    thermal_capacity = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(thermal_coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')

    # Capacity by tech dataframe (with the above aggregations added)

    netz_powcap_df1 = netz_powcap_df1.append([coal_capacity, gas_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, other_capacity])\
        [['TECHNOLOGY'] + OSeMOSYS_years_netz].reset_index(drop = True) 

    netz_powcap_df1 = netz_powcap_df1[netz_powcap_df1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    netz_powcap_df1['TECHNOLOGY'] = pd.Categorical(netz_powcap_df1['TECHNOLOGY'], prod_agg_tech[:-1])

    netz_powcap_df1 = netz_powcap_df1.sort_values('TECHNOLOGY').reset_index(drop = True)

    nrows28 = netz_powcap_df1.shape[0]
    ncols28 = netz_powcap_df1.shape[1]

    netz_powcap_df2 = netz_powcap_df1[['TECHNOLOGY'] + col_chart_years]

    nrows29 = netz_powcap_df2.shape[0]
    ncols29 = netz_powcap_df2.shape[1]

    #########################################################################################################################################
    ############ NEW DATAFRAMES #############################################################################################################

    # Refining, supply and own-use, and power
    # SHould this include POW_Transmission
    netz_transformation_df1 = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                           (netz_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                           (netz_trans_df1['TECHNOLOGY'] != 'POW_Transmission')]

    netz_transmission1 = netz_trans_df1[(netz_trans_df1['economy'] == economy) &
                                     (netz_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                     (netz_trans_df1['TECHNOLOGY'] == 'POW_Transmission')]

    netz_transmission1 = netz_transmission1.groupby('Sector').sum().copy().reset_index()
    netz_transmission1.loc[netz_transmission1['Sector'] == 'POW', 'Sector'] = 'Transmission'

    netz_transformation_sector = netz_transformation_df1.groupby('Sector').sum().copy().reset_index().append(netz_transmission1)

    netz_transformation_sector.loc[netz_transformation_sector['Sector'] == 'OWN', 'Sector'] = 'Own-use'
    netz_transformation_sector.loc[netz_transformation_sector['Sector'] == 'POW', 'Sector'] = 'Power'
    netz_transformation_sector.loc[netz_transformation_sector['Sector'] == 'REF', 'Sector'] = 'Refining'

    netz_transformation_sector1 = netz_transformation_sector[netz_transformation_sector['Sector'].isin(['Power', 'Refining'])]\
        .reset_index(drop = True)

    nrows32 = netz_transformation_sector1.shape[0]
    ncols32 = netz_transformation_sector1.shape[1]

    netz_transformation_sector2 = netz_transformation_sector1[['Sector'] + col_chart_years]

    nrows33 = netz_transformation_sector2.shape[0]
    ncols33 = netz_transformation_sector2.shape[1]

    # Own-use
    netz_ownuse_df1 = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                   (netz_trans_df1['Sector'] == 'OWN')]

    coal_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(coal_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Coal', Sector = 'Own-use and losses')
    oil_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(oil_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Oil', Sector = 'Own-use and losses')
    gas_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(gas_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Gas', Sector = 'Own-use and losses')
    renewables_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(renew_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Renewables', Sector = 'Own-use and losses')
    elec_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(elec_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Electricity', Sector = 'Own-use and losses')
    heat_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(heat_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Heat', Sector = 'Own-use and losses')
    other_own = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(other_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Other', Sector = 'Own-use and losses')

    netz_ownuse_df1 = netz_ownuse_df1.append([coal_own, oil_own, gas_own, renewables_own, elec_own, heat_own, other_own])\
        [['FUEL', 'Sector'] + OSeMOSYS_years_netz].reset_index(drop = True)

    netz_ownuse_df1 = netz_ownuse_df1[netz_ownuse_df1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    nrows34 = netz_ownuse_df1.shape[0]
    ncols34 = netz_ownuse_df1.shape[1]

    netz_ownuse_df2 = netz_ownuse_df1[['FUEL', 'Sector'] + col_chart_years]

    nrows35 = netz_ownuse_df2.shape[0]
    ncols35 = netz_ownuse_df2.shape[1]

    # Define directory
    script_dir = './results/'
    results_dir = os.path.join(script_dir, economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)

    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_transform.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None
    ref_usefuel_df1.to_excel(writer, sheet_name = economy + '_use_fuel_ref', index = False, startrow = chart_height)
    netz_usefuel_df1.to_excel(writer, sheet_name = economy + '_use_fuel_netz', index = False, startrow = chart_height)
    ref_usefuel_df2.to_excel(writer, sheet_name = economy + '_use_fuel_ref', index = False, startrow = chart_height + nrows1 + 3)
    netz_usefuel_df2.to_excel(writer, sheet_name = economy + '_use_fuel_netz', index = False, startrow = chart_height + nrows21 + 3)
    ref_prodelec_bytech_df1.to_excel(writer, sheet_name = economy + '_elec_gen_ref', index = False, startrow = chart_height)
    netz_prodelec_bytech_df1.to_excel(writer, sheet_name = economy + '_elec_gen_netz', index = False, startrow = chart_height)
    ref_prodelec_bytech_df2.to_excel(writer, sheet_name = economy + '_elec_gen_ref', index = False, startrow = chart_height + nrows3 + 3)
    netz_prodelec_bytech_df2.to_excel(writer, sheet_name = economy + '_elec_gen_netz', index = False, startrow = chart_height + nrows23 + 3)
    ref_refinery_df1.to_excel(writer, sheet_name = economy + '_refining_ref', index = False, startrow = chart_height)
    netz_refinery_df1.to_excel(writer, sheet_name = economy + '_refining_netz', index = False, startrow = chart_height)
    ref_refinery_df2.to_excel(writer, sheet_name = economy + '_refining_ref', index = False, startrow = chart_height + nrows5 + 3)
    netz_refinery_df2.to_excel(writer, sheet_name = economy + '_refining_netz', index = False, startrow = chart_height + nrows25 + 3)
    ref_refinery_df3.to_excel(writer, sheet_name = economy + '_refining_ref', index = False, startrow = chart_height + nrows5 + nrows6 + 6)
    netz_refinery_df3.to_excel(writer, sheet_name = economy + '_refining_netz', index = False, startrow = chart_height + nrows25 + nrows26 + 6)
    ref_powcap_df1.to_excel(writer, sheet_name = economy + '_pow_cap_ref', index = False, startrow = chart_height)
    netz_powcap_df1.to_excel(writer, sheet_name = economy + '_pow_cap_netz', index = False, startrow = chart_height)
    ref_powcap_df2.to_excel(writer, sheet_name = economy + '_pow_cap_ref', index = False, startrow = chart_height + nrows8 + 3)
    netz_powcap_df2.to_excel(writer, sheet_name = economy + '_pow_cap_netz', index = False, startrow = chart_height + nrows28 + 3)

    ref_transformation_sector1.to_excel(writer, sheet_name = economy + '_trnsfrm_ref', index = False, startrow = chart_height)
    netz_transformation_sector1.to_excel(writer, sheet_name = economy + '_trnsfrm_netz', index = False, startrow = chart_height)
    ref_transformation_sector2.to_excel(writer, sheet_name = economy + '_trnsfrm_ref', index = False, startrow = chart_height + nrows12 + 3)
    netz_transformation_sector2.to_excel(writer, sheet_name = economy + '_trnsfrm_netz', index = False, startrow = chart_height + nrows32 + 3)
    ref_ownuse_df1.to_excel(writer, sheet_name = economy + '_own_ref', index = False, startrow = chart_height)
    netz_ownuse_df1.to_excel(writer, sheet_name = economy + '_own_netz', index = False, startrow = chart_height)
    ref_ownuse_df2.to_excel(writer, sheet_name = economy + '_own_ref', index = False, startrow = chart_height + nrows14 + 3)
    netz_ownuse_df2.to_excel(writer, sheet_name = economy + '_own_netz', index = False, startrow = chart_height + nrows34 + 3)

    ############################################################################################################################
    
    # Access the workbook and first sheet with data from df1
    ref_worksheet1 = writer.sheets[economy + '_use_fuel_ref']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet1.set_column(2, ncols1 + 1, None, comma_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height + nrows1 + 3, None, header_format)
    ref_worksheet1.write(0, 0, economy + ' transformation use fuel', cell_format1)

    # Create a use by fuel area chart
    if nrows1 > 0:
        usefuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        usefuel_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        usefuel_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        usefuel_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        usefuel_chart1.set_y_axis({
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
            
        usefuel_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        usefuel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows1):
            usefuel_chart1.add_series({
                'name':       [economy + '_use_fuel_ref', chart_height + i + 1, 0],
                'categories': [economy + '_use_fuel_ref', chart_height, 2, chart_height, ncols1 - 1],
                'values':     [economy + '_use_fuel_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols1 - 1],
                'fill':       {'color': ref_usefuel_df1['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet1.insert_chart('B3', usefuel_chart1)

    else:
        pass

    # Create a use column chart
    if nrows2 > 0:
        usefuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        usefuel_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        usefuel_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        usefuel_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        usefuel_chart2.set_y_axis({
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
            
        usefuel_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        usefuel_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(nrows2):
            usefuel_chart2.add_series({
                'name':       [economy + '_use_fuel_ref', chart_height + nrows1 + i + 4, 0],
                'categories': [economy + '_use_fuel_ref', chart_height + nrows1 + 3, 2, chart_height + nrows1 + 3, ncols2 - 1],
                'values':     [economy + '_use_fuel_ref', chart_height + nrows1 + i + 4, 2, chart_height + nrows1 + i + 4, ncols2 - 1],
                'fill':       {'color': ref_usefuel_df2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        ref_worksheet1.insert_chart('J3', usefuel_chart2)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    ref_worksheet2 = writer.sheets[economy + '_elec_gen_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet2.set_column(2, ncols3 + 1, None, comma_format)
    ref_worksheet2.set_row(chart_height, None, header_format)
    ref_worksheet2.set_row(chart_height + nrows3 + 3, None, header_format)
    ref_worksheet2.write(0, 0, economy + ' electricity generation by technology', cell_format1)
    
    # Create a electricity production area chart
    if nrows3 > 0:
        prodelec_bytech_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        prodelec_bytech_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        prodelec_bytech_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        prodelec_bytech_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'TWh',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        prodelec_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows3):
            prodelec_bytech_chart1.add_series({
                'name':       [economy + '_elec_gen_ref', chart_height + i + 1, 0],
                'categories': [economy + '_elec_gen_ref', chart_height, 2, chart_height, ncols3 - 1],
                'values':     [economy + '_elec_gen_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols3 - 1],
                'fill':       {'color': ref_prodelec_bytech_df1['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet2.insert_chart('B3', prodelec_bytech_chart1)

    else: 
        pass

    # Create a industry subsector FED chart
    if nrows4 > 0:
        prodelec_bytech_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        prodelec_bytech_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        prodelec_bytech_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        prodelec_bytech_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'TWh',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        prodelec_bytech_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        prodelec_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows4):
            prodelec_bytech_chart2.add_series({
                'name':       [economy + '_elec_gen_ref', chart_height + nrows3 + i + 4, 0],
                'categories': [economy + '_elec_gen_ref', chart_height + nrows3 + 3, 2, chart_height + nrows3 + 3, ncols4 - 1],
                'values':     [economy + '_elec_gen_ref', chart_height + nrows3 + i + 4, 2, chart_height + nrows3 + i + 4, ncols4 - 1],
                'fill':       {'color': ref_prodelec_bytech_df2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet2.insert_chart('J3', prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    ref_worksheet3 = writer.sheets[economy + '_refining_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet3.set_column(2, ncols5 + 1, None, comma_format)
    ref_worksheet3.set_row(chart_height, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows5 + 3, None, header_format)
    ref_worksheet3.set_row(chart_height + nrows5 + nrows6 + 6, None, header_format)
    ref_worksheet3.write(0, 0, economy + ' refining', cell_format1)

    # Create ainput refining line chart
    if nrows5 > 0:
        refinery_chart1 = workbook.add_chart({'type': 'line'})
        refinery_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        refinery_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        refinery_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart1.set_y_axis({
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
            
        refinery_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        refinery_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows5):
            refinery_chart1.add_series({
                'name':       [economy + '_refining_ref', chart_height + i + 1, 0],
                'categories': [economy + '_refining_ref', chart_height, 2, chart_height, ncols5 - 1],
                'values':     [economy + '_refining_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols5 - 1],
                'line':       {'color': ref_refinery_df1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet3.insert_chart('B3', refinery_chart1)

    else:
        pass

    # Create an output refining line chart
    if nrows6 > 0:
        refinery_chart2 = workbook.add_chart({'type': 'line'})
        refinery_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        refinery_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        refinery_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart2.set_y_axis({
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
            
        refinery_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        refinery_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows6):
            refinery_chart2.add_series({
                'name':       [economy + '_refining_ref', chart_height + nrows5 + i + 4, 0],
                'categories': [economy + '_refining_ref', chart_height + nrows5 + 3, 2, chart_height + nrows5 + 3, ncols6 - 1],
                'values':     [economy + '_refining_ref', chart_height + nrows5 + i + 4, 2, chart_height + nrows5 + i + 4, ncols6 - 1],
                'line':       {'color': ref_refinery_df2['FUEL'].map(colours_dict).loc[i],
                               'width': 1}
            })    
            
        ref_worksheet3.insert_chart('J3', refinery_chart2)

    else: 
        pass

    # Create refinery output column stacked
    if nrows7 > 0:
        refinery_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        refinery_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        refinery_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        refinery_chart3.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        refinery_chart3.set_y_axis({
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
            
        refinery_chart3.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        refinery_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows7):
            refinery_chart3.add_series({
                'name':       [economy + '_refining_ref', chart_height + nrows5 + nrows6 + i + 7, 0],
                'categories': [economy + '_refining_ref', chart_height + nrows5 + nrows6 + 6, 2, chart_height + nrows5 + nrows6 + 6, ncols7 - 1],
                'values':     [economy + '_refining_ref', chart_height + nrows5 + nrows6 + i + 7, 2, chart_height + nrows5 + nrows6 + i + 7, ncols7 - 1],
                'fill':       {'color': ref_refinery_df3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet3.insert_chart('R3', refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    ref_worksheet4 = writer.sheets[economy + '_pow_cap_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet4.set_column(1, ncols8 + 1, None, comma_format)
    ref_worksheet4.set_row(chart_height, None, header_format)
    ref_worksheet4.set_row(chart_height + nrows8 + 3, None, header_format)
    ref_worksheet4.write(0, 0, economy + ' electricity capacity by technology', cell_format1)
    
    # Create a electricity production area chart
    if nrows8 > 0:
        pow_cap_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        pow_cap_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        pow_cap_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        pow_cap_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'GW',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        pow_cap_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows8):
            pow_cap_chart1.add_series({
                'name':       [economy + '_pow_cap_ref', chart_height + i + 1, 0],
                'categories': [economy + '_pow_cap_ref', chart_height, 1, chart_height, ncols8 - 1],
                'values':     [economy + '_pow_cap_ref', chart_height + i + 1, 1, chart_height + i + 1, ncols8 - 1],
                'fill':       {'color': ref_powcap_df1['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet4.insert_chart('B3', pow_cap_chart1)

    else:
        pass

    # Create a industry subsector FED chart
    if nrows9 > 0:
        pow_cap_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        pow_cap_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        pow_cap_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        pow_cap_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'GW',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        pow_cap_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        pow_cap_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows9):
            pow_cap_chart2.add_series({
                'name':       [economy + '_pow_cap_ref', chart_height + nrows8 + i + 4, 0],
                'categories': [economy + '_pow_cap_ref', chart_height + nrows8 + 3, 1, chart_height + nrows8 + 3, ncols9 - 1],
                'values':     [economy + '_pow_cap_ref', chart_height + nrows8 + i + 4, 1, chart_height + nrows8 + i + 4, ncols9 - 1],
                'fill':       {'color': ref_powcap_df2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet4.insert_chart('J3', pow_cap_chart2)

    else:
        pass

    ############################# Next sheet: Transformation sector ##################################
    
    # Access the workbook and second sheet
    ref_worksheet5 = writer.sheets[economy + '_trnsfrm_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet5.set_column(1, ncols12 + 1, None, comma_format)
    ref_worksheet5.set_row(chart_height, None, header_format)
    ref_worksheet5.set_row(chart_height + nrows12 + 3, None, header_format)
    ref_worksheet5.write(0, 0, economy + ' transformation', cell_format1)

    # Create a transformation area chart
    if nrows12 > 0:
        ref_trnsfrm_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_trnsfrm_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_trnsfrm_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_trnsfrm_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart1.set_y_axis({
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
            
        ref_trnsfrm_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        ref_trnsfrm_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows12):
            ref_trnsfrm_chart1.add_series({
                'name':       [economy + '_trnsfrm_ref', chart_height + i + 1, 0],
                'categories': [economy + '_trnsfrm_ref', chart_height, 1, chart_height, ncols12 - 1],
                'values':     [economy + '_trnsfrm_ref', chart_height + i + 1, 1, chart_height + i + 1, ncols12 - 1],
                'fill':       {'color': ref_transformation_sector1['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet5.insert_chart('B3', ref_trnsfrm_chart1)

    else:
        pass

    # Create a transformation line chart
    if nrows12 > 0:
        ref_trnsfrm_chart2 = workbook.add_chart({'type': 'line'})
        ref_trnsfrm_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_trnsfrm_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_trnsfrm_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart2.set_y_axis({
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
            
        ref_trnsfrm_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        ref_trnsfrm_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows12):
            ref_trnsfrm_chart2.add_series({
                'name':       [economy + '_trnsfrm_ref', chart_height + i + 1, 0],
                'categories': [economy + '_trnsfrm_ref', chart_height, 1, chart_height, ncols12 - 1],
                'values':     [economy + '_trnsfrm_ref', chart_height + i + 1, 1, chart_height + i + 1, ncols12 - 1],
                'line':       {'color': ref_transformation_sector1['Sector'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet5.insert_chart('J3', ref_trnsfrm_chart2)

    else:
        pass

    # Transformation column

    if nrows13 > 0:
        ref_trnsfrm_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_trnsfrm_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_trnsfrm_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        ref_trnsfrm_chart3.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_trnsfrm_chart3.set_y_axis({
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
            
        ref_trnsfrm_chart3.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        ref_trnsfrm_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows13):
            ref_trnsfrm_chart3.add_series({
                'name':       [economy + '_trnsfrm_ref', chart_height + nrows12 + i + 4, 0],
                'categories': [economy + '_trnsfrm_ref', chart_height + nrows12 + 3, 1, chart_height + nrows12 + 3, ncols13 - 1],
                'values':     [economy + '_trnsfrm_ref', chart_height + nrows12 + i + 4, 1, chart_height + nrows12 + i + 4, ncols13 - 1],
                'fill':       {'color': ref_transformation_sector2['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet5.insert_chart('R3', ref_trnsfrm_chart3)

    else:
        pass

    ###############################################################################
    # Own use charts
    
    # Access the workbook and second sheet
    ref_worksheet6 = writer.sheets[economy + '_own_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet6.set_column(2, ncols14 + 1, None, comma_format)
    ref_worksheet6.set_row(chart_height, None, header_format)
    ref_worksheet6.set_row(chart_height + nrows14 + 3, None, header_format)
    ref_worksheet6.write(0, 0, economy + ' own use and losses', cell_format1)

    # Createn own-use transformation area chart by fuel
    if nrows14 > 0:
        ref_own_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        ref_own_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_own_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        ref_own_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart1.set_y_axis({
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
            
        ref_own_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        ref_own_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows14):
            ref_own_chart1.add_series({
                'name':       [economy + '_own_ref', chart_height + i + 1, 0],
                'categories': [economy + '_own_ref', chart_height, 2, chart_height, ncols14 - 1],
                'values':     [economy + '_own_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols14 - 1],
                'fill':       {'color': ref_ownuse_df1['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet6.insert_chart('B3', ref_own_chart1)

    else:
        pass

    # Createn own-use transformation area chart by fuel
    if nrows14 > 0:
        ref_own_chart2 = workbook.add_chart({'type': 'line'})
        ref_own_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_own_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        ref_own_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart2.set_y_axis({
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
            
        ref_own_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        ref_own_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows14):
            ref_own_chart2.add_series({
                'name':       [economy + '_own_ref', chart_height + i + 1, 0],
                'categories': [economy + '_own_ref', chart_height, 2, chart_height, ncols14 - 1],
                'values':     [economy + '_own_ref', chart_height + i + 1, 2, chart_height + i + 1, ncols14 - 1],
                'line':       {'color': ref_ownuse_df1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet6.insert_chart('J3', ref_own_chart2)

    else:
        pass

    # Transformation column

    if nrows15 > 0:
        ref_own_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        ref_own_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        ref_own_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        ref_own_chart3.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        ref_own_chart3.set_y_axis({
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
            
        ref_own_chart3.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        ref_own_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows15):
            ref_own_chart3.add_series({
                'name':       [economy + '_own_ref', chart_height + nrows14 + i + 4, 0],
                'categories': [economy + '_own_ref', chart_height + nrows14 + 3, 2, chart_height + nrows14 + 3, ncols15 - 1],
                'values':     [economy + '_own_ref', chart_height + nrows14 + i + 4, 2, chart_height + nrows14 + i + 4, ncols15 - 1],
                'fill':       {'color': ref_ownuse_df2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet6.insert_chart('R3', ref_own_chart3)

    else:
        pass

    ########################################################################################################################################

    ################# NET ZERO CHARTS ######################################################################################################

    # Access the workbook and first sheet with data from df1
    netz_worksheet1 = writer.sheets[economy + '_use_fuel_netz']
    
    # Comma format and header format        
    comma_format = workbook.add_format({'num_format': '#,##0'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    netz_worksheet1.set_column(2, ncols21 + 1, None, comma_format)
    netz_worksheet1.set_row(chart_height, None, header_format)
    netz_worksheet1.set_row(chart_height + nrows21 + 3, None, header_format)
    netz_worksheet1.write(0, 0, economy + ' transformation use fuel', cell_format1)

    # Create a use by fuel area chart
    if nrows21 > 0:
        netz_usefuel_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_usefuel_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_usefuel_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_usefuel_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_usefuel_chart1.set_y_axis({
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
            
        netz_usefuel_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_usefuel_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows21):
            netz_usefuel_chart1.add_series({
                'name':       [economy + '_use_fuel_netz', chart_height + i + 1, 0],
                'categories': [economy + '_use_fuel_netz', chart_height, 2, chart_height, ncols21 - 1],
                'values':     [economy + '_use_fuel_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols21 - 1],
                'fill':       {'color': netz_usefuel_df1['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet1.insert_chart('B3', netz_usefuel_chart1)

    else:
        pass

    # Create a use column chart
    if nrows22 > 0:
        netz_usefuel_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_usefuel_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_usefuel_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_usefuel_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_usefuel_chart2.set_y_axis({
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
            
        netz_usefuel_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_usefuel_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.    
        for i in range(nrows22):
            netz_usefuel_chart2.add_series({
                'name':       [economy + '_use_fuel_netz', chart_height + nrows21 + i + 4, 0],
                'categories': [economy + '_use_fuel_netz', chart_height + nrows21 + 3, 2, chart_height + nrows21 + 3, ncols22 - 1],
                'values':     [economy + '_use_fuel_netz', chart_height + nrows21 + i + 4, 2, chart_height + nrows21 + i + 4, ncols22 - 1],
                'fill':       {'color': netz_usefuel_df2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        netz_worksheet1.insert_chart('J3', netz_usefuel_chart2)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    netz_worksheet2 = writer.sheets[economy + '_elec_gen_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet2.set_column(2, ncols23 + 1, None, comma_format)
    netz_worksheet2.set_row(chart_height, None, header_format)
    netz_worksheet2.set_row(chart_height + nrows23 + 3, None, header_format)
    netz_worksheet2.write(0, 0, economy + ' electricity generation by technology', cell_format1)
    
    # Create a electricity production area chart
    if nrows23 > 0:
        netz_prodelec_bytech_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_prodelec_bytech_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_prodelec_bytech_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_prodelec_bytech_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'TWh',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_prodelec_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows23):
            netz_prodelec_bytech_chart1.add_series({
                'name':       [economy + '_elec_gen_netz', chart_height + i + 1, 0],
                'categories': [economy + '_elec_gen_netz', chart_height, 2, chart_height, ncols23 - 1],
                'values':     [economy + '_elec_gen_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols23 - 1],
                'fill':       {'color': netz_prodelec_bytech_df1['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet2.insert_chart('B3', netz_prodelec_bytech_chart1)

    else: 
        pass

    # Create a industry subsector FED chart
    if nrows24 > 0:
        netz_prodelec_bytech_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_prodelec_bytech_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_prodelec_bytech_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_prodelec_bytech_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'TWh',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_prodelec_bytech_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_prodelec_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows24):
            netz_prodelec_bytech_chart2.add_series({
                'name':       [economy + '_elec_gen_netz', chart_height + nrows23 + i + 4, 0],
                'categories': [economy + '_elec_gen_netz', chart_height + nrows23 + 3, 2, chart_height + nrows23 + 3, ncols24 - 1],
                'values':     [economy + '_elec_gen_netz', chart_height + nrows23 + i + 4, 2, chart_height + nrows23 + i + 4, ncols24 - 1],
                'fill':       {'color': netz_prodelec_bytech_df2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet2.insert_chart('J3', netz_prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    netz_worksheet3 = writer.sheets[economy + '_refining_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet3.set_column(2, ncols25 + 1, None, comma_format)
    netz_worksheet3.set_row(chart_height, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows25 + 3, None, header_format)
    netz_worksheet3.set_row(chart_height + nrows25 + nrows26 + 6, None, header_format)
    netz_worksheet3.write(0, 0, economy + ' refining', cell_format1)

    # Create ainput refining line chart
    if nrows25 > 0:
        netz_refinery_chart1 = workbook.add_chart({'type': 'line'})
        netz_refinery_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_refinery_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_refinery_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart1.set_y_axis({
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
            
        netz_refinery_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_refinery_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows25):
            netz_refinery_chart1.add_series({
                'name':       [economy + '_refining_netz', chart_height + i + 1, 0],
                'categories': [economy + '_refining_netz', chart_height, 2, chart_height, ncols25 - 1],
                'values':     [economy + '_refining_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols25 - 1],
                'line':       {'color': netz_refinery_df1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        netz_worksheet3.insert_chart('B3', netz_refinery_chart1)

    else:
        pass

    # Create an output refining line chart
    if nrows26 > 0:
        netz_refinery_chart2 = workbook.add_chart({'type': 'line'})
        netz_refinery_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_refinery_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_refinery_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart2.set_y_axis({
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
            
        netz_refinery_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_refinery_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows26):
            netz_refinery_chart2.add_series({
                'name':       [economy + '_refining_netz', chart_height + nrows25 + i + 4, 0],
                'categories': [economy + '_refining_netz', chart_height + nrows25 + 3, 2, chart_height + nrows25 + 3, ncols26 - 1],
                'values':     [economy + '_refining_netz', chart_height + nrows25 + i + 4, 2, chart_height + nrows25 + i + 4, ncols26 - 1],
                'line':       {'color': netz_refinery_df2['FUEL'].map(colours_dict).loc[i],
                               'width': 1}
            })    
            
        netz_worksheet3.insert_chart('J3', netz_refinery_chart2)

    else: 
        pass

    # Create refinery output column stacked
    if nrows27 > 0:
        netz_refinery_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_refinery_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_refinery_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        netz_refinery_chart3.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_refinery_chart3.set_y_axis({
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
            
        netz_refinery_chart3.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_refinery_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows27):
            netz_refinery_chart3.add_series({
                'name':       [economy + '_refining_netz', chart_height + nrows25 + nrows26 + i + 7, 0],
                'categories': [economy + '_refining_netz', chart_height + nrows25 + nrows26 + 6, 2, chart_height + nrows25 + nrows26 + 6, ncols27 - 1],
                'values':     [economy + '_refining_netz', chart_height + nrows25 + nrows26 + i + 7, 2, chart_height + nrows25 + nrows26 + i + 7, ncols27 - 1],
                'fill':       {'color': netz_refinery_df3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet3.insert_chart('R3', netz_refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    netz_worksheet4 = writer.sheets[economy + '_pow_cap_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet4.set_column(1, ncols28 + 1, None, comma_format)
    netz_worksheet4.set_row(chart_height, None, header_format)
    netz_worksheet4.set_row(chart_height + nrows28 + 3, None, header_format)
    netz_worksheet4.write(0, 0, economy + ' electricity capacity by technology', cell_format1)
    
    # Create a electricity production area chart
    if nrows28 > 0:
        netz_pow_cap_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_pow_cap_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_pow_cap_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_pow_cap_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'GW',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_pow_cap_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows28):
            netz_pow_cap_chart1.add_series({
                'name':       [economy + '_pow_cap_netz', chart_height + i + 1, 0],
                'categories': [economy + '_pow_cap_netz', chart_height, 1, chart_height, ncols28 - 1],
                'values':     [economy + '_pow_cap_netz', chart_height + i + 1, 1, chart_height + i + 1, ncols28 - 1],
                'fill':       {'color': netz_powcap_df1['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet4.insert_chart('B3', netz_pow_cap_chart1)

    else:
        pass

    # Create a industry subsector FED chart
    if nrows29 > 0:
        netz_pow_cap_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_pow_cap_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_pow_cap_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_pow_cap_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart2.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'GW',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
            
        netz_pow_cap_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_pow_cap_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows29):
            netz_pow_cap_chart2.add_series({
                'name':       [economy + '_pow_cap_netz', chart_height + nrows28 + i + 4, 0],
                'categories': [economy + '_pow_cap_netz', chart_height + nrows28 + 3, 1, chart_height + nrows28 + 3, ncols29 - 1],
                'values':     [economy + '_pow_cap_netz', chart_height + nrows28 + i + 4, 1, chart_height + nrows28 + i + 4, ncols29 - 1],
                'fill':       {'color': netz_powcap_df2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet4.insert_chart('J3', netz_pow_cap_chart2)

    else:
        pass

    ############################# Next sheet: Transformation sector ##################################
    
    # Access the workbook and second sheet
    netz_worksheet5 = writer.sheets[economy + '_trnsfrm_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet5.set_column(1, ncols32 + 1, None, comma_format)
    netz_worksheet5.set_row(chart_height, None, header_format)
    netz_worksheet5.set_row(chart_height + nrows32 + 3, None, header_format)
    netz_worksheet5.write(0, 0, economy + ' transformation', cell_format1)

    # Create a transformation area chart
    if nrows32 > 0:
        netz_trnsfrm_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_trnsfrm_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_trnsfrm_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_trnsfrm_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart1.set_y_axis({
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
            
        netz_trnsfrm_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_trnsfrm_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows32):
            netz_trnsfrm_chart1.add_series({
                'name':       [economy + '_trnsfrm_netz', chart_height + i + 1, 0],
                'categories': [economy + '_trnsfrm_netz', chart_height, 1, chart_height, ncols32 - 1],
                'values':     [economy + '_trnsfrm_netz', chart_height + i + 1, 1, chart_height + i + 1, ncols32 - 1],
                'fill':       {'color': netz_transformation_sector1['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet5.insert_chart('B3', netz_trnsfrm_chart1)

    else:
        pass

    # Create a transformation line chart
    if nrows32 > 0:
        netz_trnsfrm_chart2 = workbook.add_chart({'type': 'line'})
        netz_trnsfrm_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_trnsfrm_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_trnsfrm_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart2.set_y_axis({
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
            
        netz_trnsfrm_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_trnsfrm_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows32):
            netz_trnsfrm_chart2.add_series({
                'name':       [economy + '_trnsfrm_netz', chart_height + i + 1, 0],
                'categories': [economy + '_trnsfrm_netz', chart_height, 1, chart_height, ncols32 - 1],
                'values':     [economy + '_trnsfrm_netz', chart_height + i + 1, 1, chart_height + i + 1, ncols32 - 1],
                'line':       {'color': netz_transformation_sector1['Sector'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        netz_worksheet5.insert_chart('J3', netz_trnsfrm_chart2)

    else:
        pass

    # Transformation column

    if nrows33 > 0:
        netz_trnsfrm_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_trnsfrm_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_trnsfrm_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        netz_trnsfrm_chart3.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_trnsfrm_chart3.set_y_axis({
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
            
        netz_trnsfrm_chart3.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_trnsfrm_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows33):
            netz_trnsfrm_chart3.add_series({
                'name':       [economy + '_trnsfrm_netz', chart_height + nrows32 + i + 4, 0],
                'categories': [economy + '_trnsfrm_netz', chart_height + nrows32 + 3, 1, chart_height + nrows32 + 3, ncols33 - 1],
                'values':     [economy + '_trnsfrm_netz', chart_height + nrows32 + i + 4, 1, chart_height + nrows32 + i + 4, ncols33 - 1],
                'fill':       {'color': netz_transformation_sector2['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet5.insert_chart('R3', netz_trnsfrm_chart3)

    else:
        pass

    ###############################################################################
    # Own use charts
    
    # Access the workbook and second sheet
    netz_worksheet6 = writer.sheets[economy + '_own_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet6.set_column(2, ncols34 + 1, None, comma_format)
    netz_worksheet6.set_row(chart_height, None, header_format)
    netz_worksheet6.set_row(chart_height + nrows34 + 3, None, header_format)
    netz_worksheet6.write(0, 0, economy + ' own use and losses', cell_format1)

    # Createn own-use transformation area chart by fuel
    if nrows34 > 0:
        netz_own_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        netz_own_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_own_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        netz_own_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart1.set_y_axis({
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
            
        netz_own_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_own_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows34):
            netz_own_chart1.add_series({
                'name':       [economy + '_own_netz', chart_height + i + 1, 0],
                'categories': [economy + '_own_netz', chart_height, 2, chart_height, ncols34 - 1],
                'values':     [economy + '_own_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols34 - 1],
                'fill':       {'color': netz_ownuse_df1['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet6.insert_chart('B3', netz_own_chart1)

    else:
        pass

    # Createn own-use transformation area chart by fuel
    if nrows34 > 0:
        netz_own_chart2 = workbook.add_chart({'type': 'line'})
        netz_own_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_own_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        netz_own_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart2.set_y_axis({
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
            
        netz_own_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_own_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows34):
            netz_own_chart2.add_series({
                'name':       [economy + '_own_netz', chart_height + i + 1, 0],
                'categories': [economy + '_own_netz', chart_height, 2, chart_height, ncols34 - 1],
                'values':     [economy + '_own_netz', chart_height + i + 1, 2, chart_height + i + 1, ncols34 - 1],
                'line':       {'color': netz_ownuse_df1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        netz_worksheet6.insert_chart('J3', netz_own_chart2)

    else:
        pass

    # Transformation column

    if nrows35 > 0:
        netz_own_chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        netz_own_chart3.set_size({
            'width': 500,
            'height': 300
        })
        
        netz_own_chart3.set_chartarea({
            'border': {'none': True}
        })
        
        netz_own_chart3.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        netz_own_chart3.set_y_axis({
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
            
        netz_own_chart3.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        netz_own_chart3.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(nrows35):
            netz_own_chart3.add_series({
                'name':       [economy + '_own_netz', chart_height + nrows34 + i + 4, 0],
                'categories': [economy + '_own_netz', chart_height + nrows34 + 3, 2, chart_height + nrows34 + 3, ncols35 - 1],
                'values':     [economy + '_own_netz', chart_height + nrows34 + i + 4, 2, chart_height + nrows34 + i + 4, ncols35 - 1],
                'fill':       {'color': netz_ownuse_df2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        netz_worksheet6.insert_chart('R3', netz_own_chart3)

    else:
        pass

    writer.save()

print('Bling blang blaow, you have some Transformation charts now')