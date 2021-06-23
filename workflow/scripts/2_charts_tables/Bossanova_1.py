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
ref_trans_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_transformation_reference.csv').loc[:,:'2050']

netz_power_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_power_netzero.csv').loc[:,:'2050']
netz_refownsup_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_refownsup_netzero.csv').loc[:,:'2050']
netz_pow_capacity_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_powcapacity_netzero.csv').loc[:,:'2050']
netz_trans_df1 = pd.read_csv('./data/4_Joined/OSeMOSYS_transformation_netzero.csv').loc[:,:'2050']

macro_GDP = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'GDP')
macro_GDP.columns = macro_GDP.columns.astype(str) 
macro_GDP['Series'] = 'GDP 2018 USD PPP'
macro_GDP = macro_GDP[['Economy', 'Series'] + list(macro_GDP.loc[:, '2000':'2050'])]

# Change GDP to millions
GDP = macro_GDP.select_dtypes(include=[np.number]) / 1000000 
macro_GDP[GDP.columns] = GDP

macro_GDP_growth = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'GDP_growth')
macro_GDP_growth.columns = macro_GDP_growth.columns.astype(str) 
macro_GDP_growth['Series'] = 'GDP growth'
macro_GDP_growth = macro_GDP_growth[['Economy', 'Series'] + list(macro_GDP_growth.loc[:, '2000':'2050'])]

macro_pop = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'Population')
macro_pop.columns = macro_pop.columns.astype(str) 
macro_pop['Series'] = 'Population'
macro_pop = macro_pop[['Economy', 'Series'] + list(macro_pop.loc[:, '2000':'2050'])]

# Change GDP to millions
pop = macro_pop.select_dtypes(include=[np.number]) / 1000 
macro_pop[pop.columns] = pop

macro_GDPpc = pd.read_excel('./data/2_Mapping_and_other/Key Inputs.xlsx', sheet_name = 'GDP per capita')
macro_GDPpc.columns = macro_GDPpc.columns.astype(str)
macro_GDPpc['Series'] = 'GDP per capita' 
macro_GDPpc = macro_GDPpc[['Economy', 'Series'] + list(macro_GDPpc.loc[:, '2000':'2050'])]

# Define unique values for economy, fuels, and items columns
# only looking at one dataframe which should be sufficient as both have same structure

Economy_codes = EGEDA_years_reference.economy.unique()
Fuels = EGEDA_years_reference.fuel_code.unique()
Items = EGEDA_years_reference.item_code_new.unique()

# Define colour palette

colours_dict = pd.read_csv('./data/2_Mapping_and_other/colours_dict.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

# FED and TPES: vectors for impending df builds

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

required_fuels_heat = ['1_coal', '1_5_lignite', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products', 
                       '8_gas', '9_nuclear', '11_geothermal', '15_solid_biomass', '16_1_biogas', '16_2_industrial_waste',
                       '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable', '16_8_other_liquid_biofuels',
                       '16_9_other_sources', '17_electricity', '18_heat']

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

# FED and TPES: Sectors

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

##############################################################################################################################
# TRANSFORMATION vectors for df builds

# FUEL aggregations for UseByTechnology (input fuels)

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

# Modern renewables

modren_elec_heat = ['POW_Hydro_PP', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP', 'POW_SolarCSP_PP', 
                    'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP', 'POW_WindOff_PP', 'POW_Wind_PP',
                    'POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP', 'POW_Geothermal_PP', 'POW_TIDAL_PP', 
                    'POW_CHP_BIO_PP', 'POW_Solid_Biomass_PP']

# 'POW_Pumped_Hydro'?? in the above

# POW_EXPORT_ELEC_PP need to work this in

prod_agg_tech = ['Coal', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Wind', 'Solar', 'Biomass', 'Geothermal', 'Storage', 'Other', 'Imports']
prod_agg_tech2 = ['Coal', 'Lignite', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Wind', 'Solar', 
                 'Biomass', 'Geothermal', 'Storage', 'Other', 'Imports']

heat_prod_tech = ['Coal', 'Lignite', 'Oil', 'Gas', 'Biomass', 'Waste', 'Heat only', 'Other']

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

# Heat power plants

coal_heat = ['POW_CHP_COAL_PP', 'POW_Ultra_BituCoal_PP', 'POW_Ultra_CHP_PP', 'POW_HEAT_COKE_HP', 'POW_Sub_BituCoal_PP', 'POW_Other_Coal_PP']
lignite_heat = ['POW_Sub_Brown_PP']
gas_heat = ['POW_CCGT_PP', 'POW_CHP_GAS_PP', 'POW_CCGT_CCS_PP']
oil_heat = ['POW_FuelOil_HP', 'POW_Diesel_PP']
bio_heat = ['POW_CHP_BIO_PP', 'POW_Solid_Biomass_PP']
waste_heat = ['POW_WasteToEnergy_PP']
combination_heat = ['POW_HEAT_HP']

# Make space for charts (before data/tables)
chart_height = 18 # number of excel rows before the data is written (can change this)

# Define column chart years
col_chart_years = ['2000', '2010', '2018', '2020', '2030', '2040', '2050']

# Define column chart years for transport
col_chart_years_transport = ['2018', '2020', '2030', '2040', '2050']

# Transformation chart years
trans_col_chart = ['2018', '2020', '2030', '2040', '2050']
gen_col_chart_years = ['2000', '2010', '2018', '2020', '2030', '2040', '2050']

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

EGEDA_data = pd.read_csv('./data/1_EGEDA/EGEDA_2018_years.csv', 
                             names = ['economy', 'fuel_code', 'item_code_new'] + list(range(1980, 2019)),
                             header = 0)

EGEDA_hist_gen_1 = EGEDA_data[(EGEDA_data['item_code_new'] == '18_electricity_output_in_pj') & 
                                (EGEDA_data['fuel_code'].isin(required_fuels_elec))].reset_index(drop = True)

EGEDA_hist_gen_2 = EGEDA_data[(EGEDA_data['fuel_code'] == '17_electricity') & 
                              (EGEDA_data['item_code_new'] == '2_imports')].reset_index(drop = True)

EGEDA_hist_gen = EGEDA_hist_gen_1.append(EGEDA_hist_gen_2).reset_index(drop = True)

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
                                                                '13_tide_wave_ocean': 'Hydro', 
                                                                '14_wind': 'Wind', 
                                                                '15_solid_biomass': 'Biomass', 
                                                                '16_others': 'Other', 
                                                                '17_electricity': 'Imports',
                                                                '18_heat': 'Other'})

EGEDA_hist_gen['Generation'] = 'Electricity'

EGEDA_hist_gen = EGEDA_hist_gen[['economy', 'TECHNOLOGY', 'Generation'] + list(range(2000, 2019))].\
    groupby(['economy', 'TECHNOLOGY', 'Generation']).sum().reset_index()

EGEDA_hist_gen.to_csv('./data/4_Joined/EGEDA_hist_gen.csv', index = False)
EGEDA_hist_gen = pd.read_csv('./data/4_Joined/EGEDA_hist_gen.csv')

########################### Create historical heat dataframe for use later ###########################

EGEDA_hist_heat = EGEDA_data[(EGEDA_data['item_code_new'] == '19_heat_output_in_pj') & 
                             (EGEDA_data['fuel_code'].isin(required_fuels_heat))].reset_index(drop = True)

# China only having data for 1_coal requires workaround to keep lignite data
lignite_alt = EGEDA_hist_heat[EGEDA_hist_heat['fuel_code'] == '1_5_lignite'].copy()\
    .set_index(['economy', 'fuel_code', 'item_code_new']) * -1

lignite_alt = lignite_alt.reset_index()

new_coal = EGEDA_hist_heat[EGEDA_hist_heat['fuel_code'] == '1_coal'].copy().reset_index(drop = True)

lig_coal = new_coal.append(lignite_alt).reset_index(drop = True).groupby(['economy', 'item_code_new']).sum().reset_index()
lig_coal['fuel_code'] = '1_coal'

no_coal = EGEDA_hist_heat[EGEDA_hist_heat['fuel_code'] != '1_coal'].copy().reset_index(drop = True)

EGEDA_hist_heat = no_coal.append(lig_coal).reset_index(drop = True)

EGEDA_hist_heat['TECHNOLOGY'] = EGEDA_hist_heat['fuel_code'].map({'1_coal': 'Coal', 
                                                                '1_5_lignite': 'Lignite', 
                                                                '2_coal_products': 'Coal',
                                                                '6_crude_oil_and_ngl': 'Oil',
                                                                '7_petroleum_products': 'Oil',
                                                                '8_gas': 'Gas', 
                                                                '9_nuclear': 'Other',  
                                                                '11_geothermal': 'Other', 
                                                                '15_solid_biomass': 'Biomass', 
                                                                '16_1_biogas': 'Other',
                                                                '16_2_industrial_waste': 'Waste',
                                                                '16_3_municipal_solid_waste_renewable': 'Waste',
                                                                '16_4_municipal_solid_waste_nonrenewable': 'Waste',
                                                                '16_8_other_liquid_biofuels': 'Other',
                                                                '16_9_other_sources': 'Other',
                                                                '17_electricity': 'Other',
                                                                '18_heat': 'Other'})

EGEDA_hist_heat['Generation'] = 'Heat'

EGEDA_hist_heat = EGEDA_hist_heat[['economy', 'TECHNOLOGY', 'Generation'] + list(range(2000, 2019))].\
    groupby(['economy', 'TECHNOLOGY', 'Generation']).sum().reset_index()

EGEDA_hist_heat.to_csv('./data/4_Joined/EGEDA_hist_heat.csv', index = False)
EGEDA_hist_heat = pd.read_csv('./data/4_Joined/EGEDA_hist_heat.csv')

########################### Create historical elec and heat dataframe for modern renewables ###########################

EGEDA_hist_eh = EGEDA_data[(EGEDA_data['item_code_new'].isin(['18_electricity_output_in_pj', '19_heat_output_in_pj'])) &
                           (EGEDA_data['fuel_code'].isin(Renewables_fuels))].copy().reset_index(drop = True)

EGEDA_hist_eh = EGEDA_hist_eh[['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))].\
    groupby(['economy']).sum().reset_index()

EGEDA_hist_eh['fuel_code'] = 'Modern renewables'
EGEDA_hist_eh['item_code_new'] = 'Electricity and heat'

# Amend Chinese Taipei
CT_amend = EGEDA_data[(EGEDA_data['item_code_new'].isin(['1_indigenous_production', '18_electricity_output_in_pj'])) &
                      (EGEDA_data['fuel_code'] == '10_hydro') &
                      (EGEDA_data['economy'] == '18_CT')].copy().reset_index(drop = True)\
                          [['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))]

new_CT_1 = ['18_CT', 'Modern renewables', 'Electricity and heat'] + list(CT_amend.iloc[0, 3:] - CT_amend.iloc[1, 3:])
new_CT_series1 = pd.Series(new_CT_1, index = CT_amend.columns)

new_CT_2 = ['23_NEA', 'Modern renewables', 'Electricity and heat'] + list(CT_amend.iloc[0, 3:] - CT_amend.iloc[1, 3:])
new_CT_series2 = pd.Series(new_CT_2, index = CT_amend.columns)

new_CT_3 = ['23b_ONEA', 'Modern renewables', 'Electricity and heat'] + list(CT_amend.iloc[0, 3:] - CT_amend.iloc[1, 3:])
new_CT_series3 = pd.Series(new_CT_3, index = CT_amend.columns)

new_CT_4 = ['APEC', 'Modern renewables', 'Electricity and heat'] + list(CT_amend.iloc[0, 3:] - CT_amend.iloc[1, 3:])
new_CT_series4 = pd.Series(new_CT_4, index = CT_amend.columns)

CT_amend = CT_amend.append([new_CT_series1, new_CT_series2, new_CT_series3, new_CT_series4], ignore_index = True)

EGEDA_hist_eh = EGEDA_hist_eh.append(CT_amend.iloc[2:]).reset_index(drop = True)

EGEDA_hist_eh = EGEDA_hist_eh.groupby(['economy', 'fuel_code', 'item_code_new']).sum().reset_index()

#EGEDA_hist_eh = EGEDA_hist_eh[['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))].reset_index(drop = True)

EGEDA_hist_eh.to_csv('./data/4_Joined/EGEDA_hist_eh.csv', index = False)
EGEDA_hist_eh = pd.read_csv('./data/4_Joined/EGEDA_hist_eh.csv')

#########################################################################################################################################

# Now build the subset dataframes for charts and tables

# Fix to do quicker one economy runs
# Economy_codes = ['06_HKC', '10_MAS']

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
    
    renewables = ref_notrad_1[ref_notrad_1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = 'Industry, transport, NE')
    
    others = ref_notrad_1[ref_notrad_1['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Industry, transport, NE')

    # Fed fuel data frame 1

    ref_fedfuel_1 = ref_notrad_1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(ref_notrad_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    # Insert 0 traditional biomass row
    # new_row = ['Biomass', 'Industry, transport, NE'] + [0] * 51
    # new_series = pd.Series(new_row, index = ref_fedfuel_1.columns)

    # ref_fedfuel_1 = ref_fedfuel_1.append(new_series, ignore_index = True)

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
    ref_trn_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                          (EGEDA_years_reference['fuel_code'].isin(Transport_fuels))]
    
    renewables = ref_trn_1[ref_trn_1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '15_transport_sector')
    
    others = ref_trn_1[ref_trn_1['fuel_code'].isin(Other_fuel_trans)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '15_transport_sector')

    trans_gasoline = ref_trn_1[ref_trn_1['fuel_code'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Gasoline', 
                                                            item_code_new = '15_transport_sector')

    trans_jetfuel = ref_trn_1[ref_trn_1['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Jet fuel', 
                                                            item_code_new = '15_transport_sector')
    
    ref_trn_1 = ref_trn_1.append([renewables, trans_gasoline, trans_jetfuel, others])[['fuel_code', 'item_code_new'] + list(ref_trn_1.loc[:, '2000':])].reset_index(drop = True) 

    ref_trn_1.loc[ref_trn_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    ref_trn_1.loc[ref_trn_1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    ref_trn_1.loc[ref_trn_1['fuel_code'] == '7_9_lpg', 'fuel_code'] = 'LPG'
    ref_trn_1.loc[ref_trn_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    ref_trn_1.loc[ref_trn_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_trn_1 = ref_trn_1[ref_trn_1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    ref_trn_1_rows = ref_trn_1.shape[0]
    ref_trn_1_cols = ref_trn_1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    ref_trn_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                               (EGEDA_years_reference['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years_reference['fuel_code'].isin(['19_total']))].copy()
    
    ref_trn_2.loc[ref_trn_2['item_code_new'] == '15_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    ref_trn_2.loc[ref_trn_2['item_code_new'] == '15_2_road', 'item_code_new'] = 'Road'
    ref_trn_2.loc[ref_trn_2['item_code_new'] == '15_3_rail', 'item_code_new'] = 'Rail'
    ref_trn_2.loc[ref_trn_2['item_code_new'] == '15_4_domestic_navigation', 'item_code_new'] = 'Marine'
    ref_trn_2.loc[ref_trn_2['item_code_new'] == '15_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    ref_trn_2.loc[ref_trn_2['item_code_new'] == '15_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    ref_trn_2 = ref_trn_2[ref_trn_2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    ref_trn_2 = ref_trn_2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    ref_trn_2_rows = ref_trn_2.shape[0]
    ref_trn_2_cols = ref_trn_2.shape[1]

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
    
    renewables = netz_notrad_1[netz_notrad_1['fuel_code'].isin(Renewables_fuels_nobiomass)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Other renewables', item_code_new = 'Industry, transport, NE')
    
    others = netz_notrad_1[netz_notrad_1['fuel_code'].isin(Other_fuels_FED)].groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Others', item_code_new = 'Industry, transport, NE')

    # Fed fuel data frame 1 (data frame 6)

    netz_fedfuel_1 = netz_notrad_1.append([coal, oil, renewables, others])\
        [['fuel_code', 'item_code_new'] + list(netz_notrad_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '15_solid_biomass', 'fuel_code'] = 'Biomass'    
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    # Insert 0 traditional biomass row
    # new_row = ['Biomass', 'Industry, transport, NE'] + [0] * 51
    # new_series = pd.Series(new_row, index = netz_fedfuel_1.columns)

    # netz_fedfuel_1 = netz_fedfuel_1.append(new_series, ignore_index = True)

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
    netz_trn_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                          (EGEDA_years_netzero['fuel_code'].isin(Transport_fuels))]
    
    renewables = netz_trn_1[netz_trn_1['fuel_code'].isin(Renew_fuel)].groupby(['economy', 
                                                                                     'item_code_new']).sum().assign(fuel_code = 'Renewables',
                                                                                                                   item_code_new = '15_transport_sector')
    
    others = netz_trn_1[netz_trn_1['fuel_code'].isin(Other_fuel_trans)].groupby(['economy',
                                                                                 'item_code_new']).sum().assign(fuel_code = 'Other', 
                                                                                                                item_code_new = '15_transport_sector')

    trans_gasoline = netz_trn_1[netz_trn_1['fuel_code'].isin(['7_1_motor_gasoline', '7_2_aviation_gasoline'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Gasoline', 
                                                            item_code_new = '15_transport_sector')

    trans_jetfuel = netz_trn_1[netz_trn_1['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['economy', 'item_code_new']).sum().assign(fuel_code = 'Jet fuel', 
                                                            item_code_new = '15_transport_sector')
    
    netz_trn_1 = netz_trn_1.append([renewables, trans_gasoline, trans_jetfuel, others])[['fuel_code', 'item_code_new'] + list(netz_trn_1.loc[:, '2000':])].reset_index(drop = True) 

    netz_trn_1.loc[netz_trn_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Diesel'
    netz_trn_1.loc[netz_trn_1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    netz_trn_1.loc[netz_trn_1['fuel_code'] == '7_9_lpg', 'fuel_code'] = 'LPG'
    netz_trn_1.loc[netz_trn_1['fuel_code'] == '16_9_other_sources', 'fuel_code'] = 'Hydrogen'
    netz_trn_1.loc[netz_trn_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_trn_1 = netz_trn_1[netz_trn_1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index()

    netz_trn_1_rows = netz_trn_1.shape[0]
    netz_trn_1_cols = netz_trn_1.shape[1]
    
    # Second transport data frame that provides a breakdown of the different transport modalities
    netz_trn_2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                               (EGEDA_years_netzero['item_code_new'].isin(Transport_modal)) &
                               (EGEDA_years_netzero['fuel_code'].isin(['19_total']))].copy()
    
    netz_trn_2.loc[netz_trn_2['item_code_new'] == '15_1_domestic_air_transport', 'item_code_new'] = 'Aviation'
    netz_trn_2.loc[netz_trn_2['item_code_new'] == '15_2_road', 'item_code_new'] = 'Road'
    netz_trn_2.loc[netz_trn_2['item_code_new'] == '15_3_rail', 'item_code_new'] = 'Rail'
    netz_trn_2.loc[netz_trn_2['item_code_new'] == '15_4_domestic_navigation', 'item_code_new'] = 'Marine'
    netz_trn_2.loc[netz_trn_2['item_code_new'] == '15_5_pipeline_transport', 'item_code_new'] = 'Pipeline'
    netz_trn_2.loc[netz_trn_2['item_code_new'] == '15_6_nonspecified_transport', 'item_code_new'] = 'Non-specified'

    netz_trn_2 = netz_trn_2[netz_trn_2['item_code_new'].isin(Transport_modal_agg)].set_index(['item_code_new']).loc[Transport_modal_agg].reset_index()

    netz_trn_2 = netz_trn_2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True)

    netz_trn_2_rows = netz_trn_2.shape[0]
    netz_trn_2_cols = netz_trn_2.shape[1]

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

    ref_bunkers_1 = ref_bunkers_1[['fuel_code', 'item_code_new'] + list(ref_bunkers_1.loc[:, '2000':])].reset_index(drop = True)

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

    netz_bunkers_1 = netz_bunkers_1[['fuel_code', 'item_code_new'] + list(netz_bunkers_1.loc[:, '2000':])].reset_index(drop = True)

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

    ################################################################################################################################
    ################################################################################################################################

    # Now, transformation dataframes

    # REFERENCE

    ref_pow_use_1 = ref_power_df1[(ref_power_df1['economy'] == economy) &
                        (ref_power_df1['Sheet_energy'] == 'UseByTechnology') &
                        (ref_power_df1['TECHNOLOGY'] != 'POW_Transmission')].reset_index(drop = True)

    # Now build aggregate variables of the FUELS

    # First level aggregations
    coal = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(coal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    lignite = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(lignite_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Lignite',
                                                                                              TECHNOLOGY = 'Lignite power')                                                                                      

    oil = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(oil_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Oil power')

    gas = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(gas_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Gas',
                                                                                      TECHNOLOGY = 'Gas power')

    nuclear = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(nuclear_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Nuclear',
                                                                                    TECHNOLOGY = 'Nuclear power')

    hydro = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(hydro_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Hydro',
                                                                                    TECHNOLOGY = 'Hydro power')

    solar = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(solar_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Solar',
                                                                                        TECHNOLOGY = 'Solar power')

    wind = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(wind_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Wind',
                                                                                    TECHNOLOGY = 'Wind power')

    geothermal = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(geothermal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Geothermal',
                                                                                    TECHNOLOGY = 'Geothermal power')

    biomass = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(biomass_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Biomass',
                                                                                    TECHNOLOGY = 'Biomass power')

    other_renew = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(other_renew_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other renewables',
                                                                                    TECHNOLOGY = 'Other renewable power')

    other = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(other_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other',
                                                                                        TECHNOLOGY = 'Other power')

    imports = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(imports_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Imports',
                                                                                        TECHNOLOGY = 'Electricity imports')                                                                                         

    # Second level aggregations

    coal2 = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(coal_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    renew2 = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(renewables_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Renewables',
                                                                                      TECHNOLOGY = 'Renewables power')

    # Use by fuel data frame number 1

    ref_pow_use_2 = ref_pow_use_1.append([coal, lignite, oil, gas, nuclear, hydro, solar, wind, geothermal, biomass, other_renew, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + list(ref_pow_use_1.loc[:, '2017':])].reset_index(drop = True)

    ref_pow_use_2 = ref_pow_use_2[ref_pow_use_2['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    ref_pow_use_2 = ref_pow_use_2.groupby('FUEL').sum().reset_index()
    ref_pow_use_2['Transformation'] = 'Input fuel'
    ref_pow_use_2['FUEL'] = pd.Categorical(ref_pow_use_2['FUEL'], use_agg_fuels_1)

    ref_pow_use_2 = ref_pow_use_2.sort_values('FUEL').reset_index(drop = True)

    ref_pow_use_2 = ref_pow_use_2[['FUEL', 'Transformation'] + list(ref_pow_use_2.loc[:, '2017':'2050'])]

    ref_pow_use_2_rows = ref_pow_use_2.shape[0]
    ref_pow_use_2_cols = ref_pow_use_2.shape[1]

    ref_pow_use_3 = ref_pow_use_2[['FUEL', 'Transformation'] + trans_col_chart]

    ref_pow_use_3_rows = ref_pow_use_3.shape[0]
    ref_pow_use_3_cols = ref_pow_use_3.shape[1]

    # Use by fuel data frame number 1

    ref_pow_use_4 = ref_pow_use_1.append([coal2, oil, gas, nuclear, renew2, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + list(ref_pow_use_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_pow_use_4 = ref_pow_use_4[ref_pow_use_4['FUEL'].isin(use_agg_fuels_2)].copy().set_index('FUEL').reset_index() 

    ref_pow_use_4 = ref_pow_use_4.groupby('FUEL').sum().reset_index()
    ref_pow_use_4['Transformation'] = 'Input fuel'
    ref_pow_use_4 = ref_pow_use_4[['FUEL', 'Transformation'] + list(ref_pow_use_4.loc[:, '2017':'2050'])]

    ref_pow_use_4_rows = ref_pow_use_4.shape[0]
    ref_pow_use_4_cols = ref_pow_use_4.shape[1]

    ref_pow_use_5 = ref_pow_use_4[['FUEL', 'Transformation'] + trans_col_chart]

    ref_pow_use_5_rows = ref_pow_use_5.shape[0]
    ref_pow_use_5_cols = ref_pow_use_5.shape[1]

    # Now build production dataframe
    ref_elecgen_1 = ref_power_df1[(ref_power_df1['economy'] == economy) &
                             (ref_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                             (ref_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx']))].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    oil_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    storage_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    # chp_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Cogeneration')
    nuclear_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    other_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    geo_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(geo_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Geothermal')
    misc = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')

    coal_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(thermal_coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(lignite_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    roof_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(solar_roof_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar roof')
    nonroof_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(solar_nr_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_elecgen_2 = ref_elecgen_1.append([coal_pp2, lignite_pp2, oil_pp, gas_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + list(ref_elecgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                                                                    

    ref_elecgen_2['Generation'] = 'Electricity'
    ref_elecgen_2 = ref_elecgen_2[['TECHNOLOGY', 'Generation'] + list(ref_elecgen_2.loc[:, '2017':'2050'])] 

    ref_elecgen_2 = ref_elecgen_2[ref_elecgen_2['TECHNOLOGY'].isin(prod_agg_tech2)].\
        set_index('TECHNOLOGY')

    ref_elecgen_2 = ref_elecgen_2.loc[ref_elecgen_2.index.intersection(prod_agg_tech2)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_gen[EGEDA_hist_gen['economy'] == economy].copy().\
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_gen.loc[:, '2000':'2016'])]

    ref_elecgen_2 = historical_gen.merge(ref_elecgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    ref_elecgen_2['TECHNOLOGY'] = pd.Categorical(ref_elecgen_2['TECHNOLOGY'], prod_agg_tech2)

    ref_elecgen_2 = ref_elecgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = ref_elecgen_2.select_dtypes(include=[np.number]) / 3.6 
    ref_elecgen_2[s.columns] = s

    ref_elecgen_2_rows = ref_elecgen_2.shape[0]
    ref_elecgen_2_cols = ref_elecgen_2.shape[1]

    ref_elecgen_3 = ref_elecgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    ref_elecgen_3_rows = ref_elecgen_3.shape[0]
    ref_elecgen_3_cols = ref_elecgen_3.shape[1]

    ##################################################################################################################################################################

    # Now create some refinery dataframes

    ref_refinery_1 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    ref_refinery_1['Transformation'] = 'Input to refinery'
    ref_refinery_1 = ref_refinery_1[['FUEL', 'Transformation'] + list(ref_refinery_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_refinery_1.loc[ref_refinery_1['FUEL'] == '6_1_crude_oil', 'FUEL'] = 'Crude oil'
    ref_refinery_1.loc[ref_refinery_1['FUEL'] == '6_x_ngls', 'FUEL'] = 'NGLs'

    ref_refinery_1_rows = ref_refinery_1.shape[0]
    ref_refinery_1_cols = ref_refinery_1.shape[1]

    ref_refinery_2 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_new_output))].copy()

    ref_refinery_2['Transformation'] = 'Output from refinery'
    ref_refinery_2 = ref_refinery_2[['FUEL', 'Transformation'] + list(ref_refinery_2.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_1_from_ref', 'FUEL'] = 'Motor gasoline'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_2_from_ref', 'FUEL'] = 'Aviation gasoline'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_3_from_ref', 'FUEL'] = 'Naphtha'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_jet_from_ref', 'FUEL'] = 'Jet fuel'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_6_from_ref', 'FUEL'] = 'Other kerosene'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_7_from_ref', 'FUEL'] = 'Gas diesel oil'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_8_from_ref', 'FUEL'] = 'Fuel oil'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_9_from_ref', 'FUEL'] = 'LPG'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_10_from_ref', 'FUEL'] = 'Refinery gas'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_11_from_ref', 'FUEL'] = 'Ethane'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == '7_other_from_ref', 'FUEL'] = 'Other'

    ref_refinery_2['FUEL'] = pd.Categorical(
        ref_refinery_2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other'], 
        ordered = True)

    ref_refinery_2 = ref_refinery_2.sort_values('FUEL')

    ref_refinery_2_rows = ref_refinery_2.shape[0]
    ref_refinery_2_cols = ref_refinery_2.shape[1]

    ref_refinery_3 = ref_refinery_2[['FUEL', 'Transformation'] + trans_col_chart]

    ref_refinery_3_rows = ref_refinery_3.shape[0]
    ref_refinery_3_cols = ref_refinery_3.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    ref_powcap_1 = ref_pow_capacity_df1[ref_pow_capacity_df1['REGION'] == economy]

    coal_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    oil_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    hydro_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Biomass')
    geo_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(geo_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Geothermal')
    #chp_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(chp_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Cogeneration')
    other_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')

    lignite_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(lignite_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Lignite')
    thermal_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(thermal_coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')

    # Capacity by tech dataframe (with the above aggregations added)

    ref_powcap_1 = ref_powcap_1.append([coal_capacity, gas_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, other_capacity])\
        [['TECHNOLOGY'] + list(ref_powcap_1.loc[:, '2017':'2050'])].reset_index(drop = True) 

    ref_powcap_1 = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    ref_powcap_1['TECHNOLOGY'] = pd.Categorical(ref_powcap_1['TECHNOLOGY'], prod_agg_tech[:-1])

    ref_powcap_1 = ref_powcap_1.sort_values('TECHNOLOGY').reset_index(drop = True)

    ref_powcap_1_rows = ref_powcap_1.shape[0]
    ref_powcap_1_cols = ref_powcap_1.shape[1]

    ref_powcap_2 = ref_powcap_1[['TECHNOLOGY'] + trans_col_chart]

    ref_powcap_2_rows = ref_powcap_2.shape[0]
    ref_powcap_2_cols = ref_powcap_2.shape[1]

    #########################################################################################################################################
    ############ NEW DATAFRAMES #############################################################################################################

    # Refining, supply and own-use, and power
    # SHould this include POW_Transmission?
    ref_trans_1 = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                           (ref_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                           (ref_trans_df1['TECHNOLOGY'] != 'POW_Transmission')]

    ref_transmission1 = ref_trans_df1[(ref_trans_df1['economy'] == economy) &
                                     (ref_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                     (ref_trans_df1['TECHNOLOGY'] == 'POW_Transmission')]

    ref_transmission1 = ref_transmission1.groupby('Sector').sum().copy().reset_index()
    ref_transmission1.loc[ref_transmission1['Sector'] == 'POW', 'Sector'] = 'Transmission'

    ref_trans_2 = ref_trans_1.groupby('Sector').sum().copy().reset_index().append(ref_transmission1)

    ref_trans_2.loc[ref_trans_2['Sector'] == 'OWN', 'Sector'] = 'Own-use'
    ref_trans_2.loc[ref_trans_2['Sector'] == 'POW', 'Sector'] = 'Power'
    ref_trans_2.loc[ref_trans_2['Sector'] == 'REF', 'Sector'] = 'Refining'

    # Gets rid of own-use and Transmission so that the chart is only power and refining
    ref_trans_3 = ref_trans_2[ref_trans_2['Sector'].isin(['Power', 'Refining'])]\
        .reset_index(drop = True)

    ref_trans_3_rows = ref_trans_3.shape[0]
    ref_trans_3_cols = ref_trans_3.shape[1]

    ref_trans_4 = ref_trans_3[['Sector'] + trans_col_chart]

    ref_trans_4_rows = ref_trans_4.shape[0]
    ref_trans_4_cols = ref_trans_4.shape[1]

    # Own-use
    ref_ownuse_1 = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                   (ref_trans_df1['Sector'] == 'OWN')]

    coal_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(coal_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Coal', Sector = 'Own-use and losses')
    oil_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(oil_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Oil', Sector = 'Own-use and losses')
    gas_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(gas_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Gas', Sector = 'Own-use and losses')
    renewables_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(renew_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Renewables', Sector = 'Own-use and losses')
    elec_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(elec_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Electricity', Sector = 'Own-use and losses')
    heat_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(heat_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Heat', Sector = 'Own-use and losses')
    other_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(other_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Other', Sector = 'Own-use and losses')

    ref_ownuse_1 = ref_ownuse_1.append([coal_own, oil_own, gas_own, renewables_own, elec_own, heat_own, other_own])\
        [['FUEL', 'Sector'] + list(ref_ownuse_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_ownuse_1 = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    ref_ownuse_1_rows = ref_ownuse_1.shape[0]
    ref_ownuse_1_cols = ref_ownuse_1.shape[1]

    ref_ownuse_2 = ref_ownuse_1[['FUEL', 'Sector'] + trans_col_chart]

    ref_ownuse_2_rows = ref_ownuse_2.shape[0]
    ref_ownuse_2_cols = ref_ownuse_2.shape[1]

    ###############################################

    # Heat generation dataframes

    ref_heatgen_1 = ref_power_df1[(ref_power_df1['economy'] == economy) &
                             (ref_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                             (ref_power_df1['FUEL'] == '18_heat')].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(coal_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(lignite_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    oil_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(oil_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(gas_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    bio_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(bio_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    waste_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(waste_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Waste')
    comb_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(combination_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Heat only')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_heatgen_2 = ref_heatgen_1.append([coal_hp, lignite_hp, oil_hp, gas_hp, bio_hp, waste_hp, comb_hp])\
        [['TECHNOLOGY'] + list(ref_heatgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                                                                    

    ref_heatgen_2['Generation'] = 'Heat'
    ref_heatgen_2 = ref_heatgen_2[['TECHNOLOGY', 'Generation'] + list(ref_heatgen_2.loc[:, '2017':'2050'])] 

    # Insert 0 other row
    new_row_zero = ['Other', 'Heat'] + [0] * 34
    new_series = pd.Series(new_row_zero, index = ref_heatgen_2.columns)

    ref_heatgen_2 = ref_heatgen_2.append(new_series, ignore_index = True).reset_index(drop = True)

    ref_heatgen_2 = ref_heatgen_2[ref_heatgen_2['TECHNOLOGY'].isin(heat_prod_tech)].\
        set_index('TECHNOLOGY')

    ref_heatgen_2 = ref_heatgen_2.loc[ref_heatgen_2.index.intersection(heat_prod_tech)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_heat[EGEDA_hist_heat['economy'] == economy].copy().\
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_heat.loc[:, '2000':'2016'])]

    ref_heatgen_2 = historical_gen.merge(ref_heatgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    ref_heatgen_2['TECHNOLOGY'] = pd.Categorical(ref_heatgen_2['TECHNOLOGY'], heat_prod_tech)

    ref_heatgen_2 = ref_heatgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    ref_heatgen_2_rows = ref_heatgen_2.shape[0]
    ref_heatgen_2_cols = ref_heatgen_2.shape[1]

    ref_heatgen_3 = ref_heatgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    ref_heatgen_3_rows = ref_heatgen_3.shape[0]
    ref_heatgen_3_cols = ref_heatgen_3.shape[1]

    ######################################################################################################################
    
    # NET-ZERO dataframes

    netz_pow_use_1 = netz_power_df1[(netz_power_df1['economy'] == economy) &
                        (netz_power_df1['Sheet_energy'] == 'UseByTechnology') &
                        (netz_power_df1['TECHNOLOGY'] != 'POW_Transmission')].reset_index(drop = True)

    # Now build aggregate variables of the FUELS

    # First level aggregations
    coal = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(coal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    lignite = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(lignite_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Lignite',
                                                                                              TECHNOLOGY = 'Lignite power')                                                                                      

    oil = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(oil_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Oil power')

    gas = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(gas_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Gas',
                                                                                      TECHNOLOGY = 'Gas power')

    nuclear = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(nuclear_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Nuclear',
                                                                                    TECHNOLOGY = 'Nuclear power')

    hydro = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(hydro_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Hydro',
                                                                                    TECHNOLOGY = 'Hydro power')

    solar = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(solar_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Solar',
                                                                                        TECHNOLOGY = 'Solar power')

    wind = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(wind_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Wind',
                                                                                    TECHNOLOGY = 'Wind power')

    geothermal = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(geothermal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Geothermal',
                                                                                    TECHNOLOGY = 'Geothermal power')

    biomass = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(biomass_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Biomass',
                                                                                    TECHNOLOGY = 'Biomass power')

    other_renew = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(other_renew_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other renewables',
                                                                                    TECHNOLOGY = 'Other renewable power')

    other = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(other_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Other',
                                                                                        TECHNOLOGY = 'Other power')

    imports = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(imports_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Imports',
                                                                                        TECHNOLOGY = 'Electricity imports')                                                                                         

    # Second level aggregations

    coal2 = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(coal_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    renew2 = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(renewables_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Renewables',
                                                                                      TECHNOLOGY = 'Renewables power')

    # Use by fuel data frame number 1

    netz_pow_use_2 = netz_pow_use_1.append([coal, lignite, oil, gas, nuclear, hydro, solar, wind, geothermal, biomass, other_renew, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + list(netz_pow_use_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_pow_use_2 = netz_pow_use_2[netz_pow_use_2['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    netz_pow_use_2 = netz_pow_use_2.groupby('FUEL').sum().reset_index()
    netz_pow_use_2['Transformation'] = 'Input fuel'
    netz_pow_use_2['FUEL'] = pd.Categorical(netz_pow_use_2['FUEL'], use_agg_fuels_1)

    netz_pow_use_2 = netz_pow_use_2.sort_values('FUEL').reset_index(drop = True)

    netz_pow_use_2 = netz_pow_use_2[['FUEL', 'Transformation'] + list(netz_pow_use_2.loc[:,'2017':'2050'])]

    netz_pow_use_2_rows = netz_pow_use_2.shape[0]
    netz_pow_use_2_cols = netz_pow_use_2.shape[1]

    netz_pow_use_3 = netz_pow_use_2[['FUEL', 'Transformation'] + trans_col_chart]

    netz_pow_use_3_rows = netz_pow_use_3.shape[0]
    netz_pow_use_3_cols = netz_pow_use_3.shape[1]

    # Use by fuel data frame number 1

    netz_pow_use_4 = netz_pow_use_1.append([coal2, oil, gas, nuclear, renew2, other, imports])\
        [['FUEL', 'TECHNOLOGY'] + list(netz_pow_use_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_pow_use_4 = netz_pow_use_4[netz_pow_use_4['FUEL'].isin(use_agg_fuels_2)].copy().set_index('FUEL').reset_index() 

    netz_pow_use_4 = netz_pow_use_4.groupby('FUEL').sum().reset_index()
    netz_pow_use_4['Transformation'] = 'Input fuel'
    netz_pow_use_4 = netz_pow_use_4[['FUEL', 'Transformation'] + list(netz_pow_use_4.loc[:,'2017':'2050'])]

    netz_pow_use_4_rows = netz_pow_use_4.shape[0]
    netz_pow_use_4_cols = netz_pow_use_4.shape[1]

    netz_pow_use_5 = netz_pow_use_4[['FUEL', 'Transformation'] + trans_col_chart]

    netz_pow_use_5_rows = netz_pow_use_5.shape[0]
    netz_pow_use_5_cols = netz_pow_use_5.shape[1]

    # Now build production dataframe
    netz_elecgen_1 = netz_power_df1[(netz_power_df1['economy'] == economy) &
                             (netz_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                             (netz_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx']))].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    oil_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    storage_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    # chp_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Cogeneration')
    nuclear_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    other_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    geo_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(geo_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Geothermal')
    misc = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')

    coal_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(thermal_coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(lignite_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    roof_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(solar_roof_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar roof')
    nonroof_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(solar_nr_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_elecgen_2 = netz_elecgen_1.append([coal_pp2, lignite_pp2, oil_pp, gas_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + list(netz_elecgen_1.loc[:,'2017':'2050'])].reset_index(drop = True)                                                                                                    

    netz_elecgen_2['Generation'] = 'Electricity'
    netz_elecgen_2 = netz_elecgen_2[['TECHNOLOGY', 'Generation'] + list(netz_elecgen_2.loc[:,'2017':'2050'])] 

    netz_elecgen_2 = netz_elecgen_2[netz_elecgen_2['TECHNOLOGY'].isin(prod_agg_tech2)].\
        set_index('TECHNOLOGY')

    netz_elecgen_2 = netz_elecgen_2.loc[netz_elecgen_2.index.intersection(prod_agg_tech2)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_gen[EGEDA_hist_gen['economy'] == economy].copy().\
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_gen.loc[:,'2000':'2016'])]

    netz_elecgen_2 = historical_gen.merge(netz_elecgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    netz_elecgen_2['TECHNOLOGY'] = pd.Categorical(netz_elecgen_2['TECHNOLOGY'], prod_agg_tech2)

    netz_elecgen_2 = netz_elecgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = netz_elecgen_2.select_dtypes(include=[np.number]) / 3.6 
    netz_elecgen_2[s.columns] = s

    netz_elecgen_2_rows = netz_elecgen_2.shape[0]
    netz_elecgen_2_cols = netz_elecgen_2.shape[1]

    netz_elecgen_3 = netz_elecgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    netz_elecgen_3_rows = netz_elecgen_3.shape[0]
    netz_elecgen_3_cols = netz_elecgen_3.shape[1]

    ##################################################################################################################################################################

    # Now create some refinery dataframes

    netz_refinery_1 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                 (netz_refownsup_df1['Sector'] == 'REF') & 
                                 (netz_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    netz_refinery_1['Transformation'] = 'Input to refinery'
    netz_refinery_1 = netz_refinery_1[['FUEL', 'Transformation'] + list(netz_refinery_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_refinery_1.loc[netz_refinery_1['FUEL'] == '6_1_crude_oil', 'FUEL'] = 'Crude oil'
    netz_refinery_1.loc[netz_refinery_1['FUEL'] == '6_x_ngls', 'FUEL'] = 'NGLs'

    netz_refinery_1_rows = netz_refinery_1.shape[0]
    netz_refinery_1_cols = netz_refinery_1.shape[1]

    netz_refinery_2 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                 (netz_refownsup_df1['Sector'] == 'REF') & 
                                 (netz_refownsup_df1['FUEL'].isin(refinery_new_output))].copy()

    netz_refinery_2['Transformation'] = 'Output from refinery'
    netz_refinery_2 = netz_refinery_2[['FUEL', 'Transformation'] + list(netz_refinery_2.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_1_from_ref', 'FUEL'] = 'Motor gasoline'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_2_from_ref', 'FUEL'] = 'Aviation gasoline'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_3_from_ref', 'FUEL'] = 'Naphtha'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_jet_from_ref', 'FUEL'] = 'Jet fuel'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_6_from_ref', 'FUEL'] = 'Other kerosene'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_7_from_ref', 'FUEL'] = 'Gas diesel oil'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_8_from_ref', 'FUEL'] = 'Fuel oil'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_9_from_ref', 'FUEL'] = 'LPG'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_10_from_ref', 'FUEL'] = 'Refinery gas'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_11_from_ref', 'FUEL'] = 'Ethane'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == '7_other_from_ref', 'FUEL'] = 'Other'

    netz_refinery_2['FUEL'] = pd.Categorical(
        netz_refinery_2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other'], 
        ordered = True)

    netz_refinery_2 = netz_refinery_2.sort_values('FUEL')

    netz_refinery_2_rows = netz_refinery_2.shape[0]
    netz_refinery_2_cols = netz_refinery_2.shape[1]

    netz_refinery_3 = netz_refinery_2[['FUEL', 'Transformation'] + trans_col_chart]

    netz_refinery_3_rows = netz_refinery_3.shape[0]
    netz_refinery_3_cols = netz_refinery_3.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    netz_powcap_1 = netz_pow_capacity_df1[netz_pow_capacity_df1['REGION'] == economy]

    coal_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    oil_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    hydro_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Biomass')
    geo_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(geo_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Geothermal')
    #chp_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(chp_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Cogeneration')
    other_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')

    lignite_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(lignite_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Lignite')
    thermal_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(thermal_coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')

    # Capacity by tech dataframe (with the above aggregations added)

    netz_powcap_1 = netz_powcap_1.append([coal_capacity, gas_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, other_capacity])\
        [['TECHNOLOGY'] + list(netz_powcap_1.loc[:,'2017':'2050'])].reset_index(drop = True) 

    netz_powcap_1 = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(pow_capacity_agg)].reset_index(drop = True)

    netz_powcap_1['TECHNOLOGY'] = pd.Categorical(netz_powcap_1['TECHNOLOGY'], prod_agg_tech[:-1])

    netz_powcap_1 = netz_powcap_1.sort_values('TECHNOLOGY').reset_index(drop = True)

    netz_powcap_1_rows = netz_powcap_1.shape[0]
    netz_powcap_1_cols = netz_powcap_1.shape[1]

    netz_powcap_2 = netz_powcap_1[['TECHNOLOGY'] + trans_col_chart]

    netz_powcap_2_rows = netz_powcap_2.shape[0]
    netz_powcap_2_cols = netz_powcap_2.shape[1]

    #########################################################################################################################################
    ############ NEW DATAFRAMES #############################################################################################################

    # Refining, supply and own-use, and power
    # SHould this include POW_Transmission?
    netz_trans_1 = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                           (netz_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                           (netz_trans_df1['TECHNOLOGY'] != 'POW_Transmission')]

    netz_transmission1 = netz_trans_df1[(netz_trans_df1['economy'] == economy) &
                                     (netz_trans_df1['Sheet_energy'] == 'UseByTechnology') &
                                     (netz_trans_df1['TECHNOLOGY'] == 'POW_Transmission')]

    netz_transmission1 = netz_transmission1.groupby('Sector').sum().copy().reset_index()
    netz_transmission1.loc[netz_transmission1['Sector'] == 'POW', 'Sector'] = 'Transmission'

    netz_trans_2 = netz_trans_1.groupby('Sector').sum().copy().reset_index().append(netz_transmission1)

    netz_trans_2.loc[netz_trans_2['Sector'] == 'OWN', 'Sector'] = 'Own-use'
    netz_trans_2.loc[netz_trans_2['Sector'] == 'POW', 'Sector'] = 'Power'
    netz_trans_2.loc[netz_trans_2['Sector'] == 'REF', 'Sector'] = 'Refining'

    # Gets rid of own-use and Transmission so that the chart is only power and refining
    netz_trans_3 = netz_trans_2[netz_trans_2['Sector'].isin(['Power', 'Refining'])]\
        .reset_index(drop = True)

    netz_trans_3_rows = netz_trans_3.shape[0]
    netz_trans_3_cols = netz_trans_3.shape[1]

    netz_trans_4 = netz_trans_3[['Sector'] + trans_col_chart]

    netz_trans_4_rows = netz_trans_4.shape[0]
    netz_trans_4_cols = netz_trans_4.shape[1]

    # Own-use
    netz_ownuse_1 = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                   (netz_trans_df1['Sector'] == 'OWN')]

    coal_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(coal_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Coal', Sector = 'Own-use and losses')
    oil_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(oil_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Oil', Sector = 'Own-use and losses')
    gas_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(gas_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Gas', Sector = 'Own-use and losses')
    renewables_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(renew_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Renewables', Sector = 'Own-use and losses')
    elec_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(elec_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Electricity', Sector = 'Own-use and losses')
    heat_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(heat_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Heat', Sector = 'Own-use and losses')
    other_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(other_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Other', Sector = 'Own-use and losses')

    netz_ownuse_1 = netz_ownuse_1.append([coal_own, oil_own, gas_own, renewables_own, elec_own, heat_own, other_own])\
        [['FUEL', 'Sector'] + list(netz_ownuse_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_ownuse_1 = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    netz_ownuse_1_rows = netz_ownuse_1.shape[0]
    netz_ownuse_1_cols = netz_ownuse_1.shape[1]

    netz_ownuse_2 = netz_ownuse_1[['FUEL', 'Sector'] + trans_col_chart]

    netz_ownuse_2_rows = netz_ownuse_2.shape[0]
    netz_ownuse_2_cols = netz_ownuse_2.shape[1]

    ###############################################

    # Heat generation dataframes

    netz_heatgen_1 = netz_power_df1[(netz_power_df1['economy'] == economy) &
                             (netz_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                             (netz_power_df1['FUEL'] == '18_heat')].reset_index(drop = True)

    # Now build the aggregations of technology (power plants)

    coal_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(coal_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(lignite_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    oil_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(oil_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(gas_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    bio_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(bio_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    waste_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(waste_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Waste')
    comb_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(combination_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Heat only')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_heatgen_2 = netz_heatgen_1.append([coal_hp, lignite_hp, oil_hp, gas_hp, bio_hp, waste_hp, comb_hp])\
        [['TECHNOLOGY'] + list(netz_heatgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                              

    netz_heatgen_2['Generation'] = 'Heat'
    netz_heatgen_2 = netz_heatgen_2[['TECHNOLOGY', 'Generation'] + list(netz_heatgen_2.loc[:, '2017':'2050'])]

    # Insert 0 other row
    new_row_zero = ['Other', 'Heat'] + [0] * 34
    new_series = pd.Series(new_row_zero, index = netz_heatgen_2.columns)

    netz_heatgen_2 = netz_heatgen_2.append(new_series, ignore_index = True).reset_index(drop = True)

    netz_heatgen_2 = netz_heatgen_2[netz_heatgen_2['TECHNOLOGY'].isin(heat_prod_tech)].\
        set_index('TECHNOLOGY')

    netz_heatgen_2 = netz_heatgen_2.loc[netz_heatgen_2.index.intersection(heat_prod_tech)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_heat[EGEDA_hist_heat['economy'] == economy].copy().\
        iloc[:,:-2][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_heat.loc[:, '2000':'2016'])]

    netz_heatgen_2 = historical_gen.merge(netz_heatgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    netz_heatgen_2['TECHNOLOGY'] = pd.Categorical(netz_heatgen_2['TECHNOLOGY'], heat_prod_tech)

    netz_heatgen_2 = netz_heatgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    netz_heatgen_2_rows = netz_heatgen_2.shape[0]
    netz_heatgen_2_cols = netz_heatgen_2.shape[1]

    netz_heatgen_3 = netz_heatgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    netz_heatgen_3_rows = netz_heatgen_3.shape[0]
    netz_heatgen_3_cols = netz_heatgen_3.shape[1]

    #############################################################################################################################

    # REFERENCE: Modern renewables
    # Agriculture, buildings, industry, transport, non-specified others

    ref_ag_temp = ref_ag_1[ref_ag_1['fuel_code'] == 'Other renewables'].copy().groupby(['item_code_new']).sum().reset_index()
    ref_bld_temp = ref_bld_2[ref_bld_2['fuel_code'] == 'Other renewables'].copy().groupby(['item_code_new']).sum().reset_index()
    ref_trn_temp = ref_trn_1[ref_trn_1['fuel_code'] == 'Renewables'].copy().groupby(['item_code_new']).sum().reset_index()
    
    ref_ind_temp = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                                         (EGEDA_years_reference['fuel_code'].isin(Renewables_fuels)) &
                                         (EGEDA_years_reference['item_code_new'] == '14_industry_sector')]\
                                             [['fuel_code', 'item_code_new'] + list(EGEDA_years_reference.loc[:, '2000':'2050'])]\
                                             .copy().replace(np.nan, 0).groupby(['item_code_new']).sum().reset_index()

    ref_nonspec_temp = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                                         (EGEDA_years_reference['fuel_code'].isin(Renewables_fuels_nobiomass)) &
                                         (EGEDA_years_reference['item_code_new'] == '16_5_nonspecified_others')]\
                                             [['fuel_code', 'item_code_new'] + list(EGEDA_years_reference.loc[:, '2000':'2050'])]\
                                             .copy().replace(np.nan, 0).groupby(['item_code_new']).sum().reset_index()

    ref_modren_1 = ref_ag_temp.append([ref_bld_temp, ref_trn_temp, ref_ind_temp, ref_nonspec_temp]).reset_index(drop = True)
    ref_modren_1['fuel_code'] = 'Modern renewables'

    ref_modren_1.loc[ref_modren_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    ref_modren_1.loc[ref_modren_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_modren_1.loc[ref_modren_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_modren_1.loc[ref_modren_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified others'

    ref_modren_1 = ref_modren_1[['fuel_code', 'item_code_new'] + list(ref_modren_1.loc[:, '2000':'2050'])]

    # Electricity and heat renewables

    ref_modren_elecheat = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                    (ref_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                                    (ref_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx', '18_heat'])) &
                                    (ref_power_df1['TECHNOLOGY'].isin(modren_elec_heat))].copy().groupby(['economy'])\
                                        .sum().reset_index(drop = True)

    ref_modren_elecheat['fuel_code'] = 'Modern renewables'
    ref_modren_elecheat['item_code_new'] = 'Electricity and heat'

    # Grab historical
    historical_eh = EGEDA_hist_eh[EGEDA_hist_eh['economy'] == economy].copy().iloc[:,1:-2]

    ref_modren_elecheat = historical_eh.merge(ref_modren_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_modren_elecheat = ref_modren_elecheat[['fuel_code', 'item_code_new'] + list(ref_modren_elecheat\
        .loc[:, '2000':'2050'])]

    ref_modren_2 = ref_modren_1.append(ref_modren_elecheat).reset_index(drop = True)
    ref_modren_2 = ref_modren_2.append(ref_modren_2.sum(numeric_only = True), ignore_index = True) 

    ref_modren_2.iloc[6, 0] = 'Modern renewables'
    ref_modren_2.iloc[6, 1] = 'Total'

    ref_modren_3 = ref_modren_2.append(ref_tfec_1).reset_index(drop = True)

    modren_prop1 = ['Modern renewables', 'Reference'] + list(ref_modren_3.iloc[6, 2:] / ref_modren_3.iloc[7, 2:])
    modren_prop_series1 = pd.Series(modren_prop1, index = ref_modren_3.columns)

    ref_modren_4 = ref_modren_3.append(modren_prop_series1, ignore_index = True).reset_index(drop = True)

    ref_modren_4 = ref_modren_4[ref_modren_4['item_code_new'].isin(['Total', 'TFEC', 'Reference'])].copy().reset_index(drop = True)

    ref_modren_4_rows = ref_modren_4.shape[0]
    ref_modren_4_cols = ref_modren_4.shape[1]

    # NET-ZERO: Modern renewables
    # Agriculture, buildings, industry, transport, non-specified others

    netz_ag_temp = netz_ag_1[netz_ag_1['fuel_code'] == 'Other renewables'].copy().groupby(['item_code_new']).sum().reset_index()
    netz_bld_temp = netz_bld_2[netz_bld_2['fuel_code'] == 'Other renewables'].copy().groupby(['item_code_new']).sum().reset_index()
    netz_trn_temp = netz_trn_1[netz_trn_1['fuel_code'] == 'Renewables'].copy().groupby(['item_code_new']).sum().reset_index()
    
    netz_ind_temp = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                                         (EGEDA_years_netzero['fuel_code'].isin(Renewables_fuels)) &
                                         (EGEDA_years_netzero['item_code_new'] == '14_industry_sector')]\
                                             [['fuel_code', 'item_code_new'] + list(EGEDA_years_netzero.loc[:, '2000':'2050'])]\
                                             .copy().replace(np.nan, 0).groupby(['item_code_new']).sum().reset_index()

    netz_nonspec_temp = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                                         (EGEDA_years_netzero['fuel_code'].isin(Renewables_fuels_nobiomass)) &
                                         (EGEDA_years_netzero['item_code_new'] == '16_5_nonspecified_others')]\
                                             [['fuel_code', 'item_code_new'] + list(EGEDA_years_netzero.loc[:, '2000':'2050'])]\
                                             .copy().replace(np.nan, 0).groupby(['item_code_new']).sum().reset_index()

    netz_modren_1 = netz_ag_temp.append([netz_bld_temp, netz_trn_temp, netz_ind_temp, netz_nonspec_temp]).reset_index(drop = True)
    netz_modren_1['fuel_code'] = 'Modern renewables'

    netz_modren_1.loc[netz_modren_1['item_code_new'] == '16_x_buildings', 'item_code_new'] = 'Buildings'
    netz_modren_1.loc[netz_modren_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_modren_1.loc[netz_modren_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_modren_1.loc[netz_modren_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified others'

    netz_modren_1 = netz_modren_1[['fuel_code', 'item_code_new'] + list(netz_modren_1.loc[:, '2000':'2050'])]

    # Electricity and heat renewables

    netz_modren_elecheat = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                    (netz_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                                    (netz_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx', '18_heat'])) &
                                    (netz_power_df1['TECHNOLOGY'].isin(modren_elec_heat))].copy().groupby(['economy'])\
                                        .sum().reset_index(drop = True)

    netz_modren_elecheat['fuel_code'] = 'Modern renewables'
    netz_modren_elecheat['item_code_new'] = 'Electricity and heat'

    # Grab historical
    historical_eh = EGEDA_hist_eh[EGEDA_hist_eh['economy'] == economy].copy().iloc[:,1:-2]

    netz_modren_elecheat = historical_eh.merge(netz_modren_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_modren_elecheat = netz_modren_elecheat[['fuel_code', 'item_code_new'] + list(netz_modren_elecheat\
        .loc[:, '2000':'2050'])]

    netz_modren_2 = netz_modren_1.append(netz_modren_elecheat).reset_index(drop = True)
    netz_modren_2 = netz_modren_2.append(netz_modren_2.sum(numeric_only = True), ignore_index = True) 

    netz_modren_2.iloc[6, 0] = 'Modern renewables'
    netz_modren_2.iloc[6, 1] = 'Total'

    netz_modren_3 = netz_modren_2.append(netz_tfec_1).reset_index(drop = True)

    modren_prop1 = ['Modern renewables', 'Net-zero'] + list(netz_modren_3.iloc[6, 2:] / netz_modren_3.iloc[7, 2:])
    modren_prop_series1 = pd.Series(modren_prop1, index = netz_modren_3.columns)

    netz_modren_4 = netz_modren_3.append(modren_prop_series1, ignore_index = True).reset_index(drop = True)

    netz_modren_4 = netz_modren_4[netz_modren_4['item_code_new'].isin(['Total', 'TFEC', 'Net-zero'])].copy().reset_index(drop = True)

    netz_modren_4_rows = netz_modren_4.shape[0]
    netz_modren_4_cols = netz_modren_4.shape[1]

    # Macro dataframe

    if any(economy in s for s in list(macro_GDP['Economy'])):

        macro_1 = macro_GDP[macro_GDP['Economy'] == economy].copy().\
            append(macro_GDP_growth[macro_GDP_growth['Economy'] == economy].copy()).\
                append(macro_pop[macro_pop['Economy'] == economy].copy()).\
                    append(macro_GDPpc[macro_GDPpc['Economy'] == economy].copy()).reset_index(drop = True)    

        macro_1_rows = macro_1.shape[0]
        macro_1_cols = macro_1.shape[1]

    else:
        macro_1 = pd.DataFrame()
        macro_1_rows = macro_1.shape[0]
        macro_1_cols = macro_1.shape[1]

    #############################################################################################################

    # Energy intensity
    # REFERENCE

    if any(economy in s for s in list(macro_GDP['Economy'])):

        ref_enint_1 = ref_tfec_1.copy()
        ref_enint_1['Economy'] = economy
        ref_enint_1['Series'] = 'TFEC'

        ref_enint_1 = ref_enint_1.append(macro_1[macro_1['Series'] == 'GDP 2018 USD PPP']).copy().reset_index(drop = True)

        ref_enint_1 = ref_enint_1[['Economy', 'Series'] + list(ref_enint_1.loc[:,'2000':'2050'])]

        ref_ei_calc1 = [economy, 'TFEC energy intensity'] + list(ref_enint_1.iloc[0, 2:] / ref_enint_1.iloc[1, 2:])
        ref_ei_series1 = pd.Series(ref_ei_calc1, index = ref_enint_1.columns)

        ref_enint_2 = ref_enint_1.append(ref_ei_series1, ignore_index = True).reset_index(drop = True)

        ref_ei_calc2 = [economy, 'Reference'] + list(ref_enint_2.iloc[2, 2:] / ref_enint_2.iloc[2, 7] * 100)
        ref_ei_series2 = pd.Series(ref_ei_calc2, index = ref_enint_2.columns)

        ref_enint_3 = ref_enint_2.append(ref_ei_series2, ignore_index = True).reset_index(drop = True)

        ref_enint_3_rows = ref_enint_3.shape[0]
        ref_enint_3_cols = ref_enint_3.shape[1]

        # NET-ZERO

        netz_enint_1 = netz_tfec_1.copy()
        netz_enint_1['Economy'] = economy
        netz_enint_1['Series'] = 'TFEC'

        netz_enint_1 = netz_enint_1.append(macro_1[macro_1['Series'] == 'GDP 2018 USD PPP']).copy().reset_index(drop = True)

        netz_enint_1 = netz_enint_1[['Economy', 'Series'] + list(netz_enint_1.loc[:,'2000':'2050'])]

        netz_ei_calc1 = [economy, 'TFEC energy intensity'] + list(netz_enint_1.iloc[0, 2:] / netz_enint_1.iloc[1, 2:])
        netz_ei_series1 = pd.Series(netz_ei_calc1, index = netz_enint_1.columns)

        netz_enint_2 = netz_enint_1.append(netz_ei_series1, ignore_index = True).reset_index(drop = True)

        netz_ei_calc2 = [economy, 'Net-zero'] + list(netz_enint_2.iloc[2, 2:] / netz_enint_2.iloc[2, 7] * 100)
        netz_ei_series2 = pd.Series(netz_ei_calc2, index = netz_enint_2.columns)

        netz_enint_3 = netz_enint_2.append(netz_ei_series2, ignore_index = True).reset_index(drop = True)

        netz_enint_3_rows = netz_enint_3.shape[0]
        netz_enint_3_cols = netz_enint_3.shape[1]

    else:
        ref_enint_3 = pd.DataFrame()
        ref_enint_3_rows = ref_enint_3.shape[0]
        ref_enint_r_cols = ref_enint_3.shape[1]

        netz_enint_3 = pd.DataFrame()
        netz_enint_3_rows = netz_enint_3.shape[0]
        netz_enint_r_cols = netz_enint_3.shape[1]

    # Df builds are complete

    ##############################################################################################################################
    
    # Define directory to save charts and tables workbook
    script_dir = './results/'
    results_dir = os.path.join(script_dir, economy)
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_charts.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Insert the various dataframes into different sheets of the workbook
    # REFERENCE and NETZERO
    # FED
    ref_fedfuel_1.to_excel(writer, sheet_name = economy + '_FED_fuel', index = False, startrow = chart_height)
    netz_fedfuel_1.to_excel(writer, sheet_name = economy + '_FED_fuel', index = False, startrow = (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6)
    ref_fedfuel_2.to_excel(writer, sheet_name = economy + '_FED_fuel', index = False, startrow = chart_height + ref_fedfuel_1_rows + 3)
    netz_fedfuel_2.to_excel(writer, sheet_name = economy + '_FED_fuel', index = False, startrow = (2 * chart_height) + ref_fedfuel_1_rows + netz_fedfuel_1_rows + ref_fedfuel_2_rows + 9)
    ref_fedsector_2.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = chart_height)
    netz_fedsector_2.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9)
    ref_fedsector_3.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = chart_height + ref_fedsector_2_rows + 3)
    netz_fedsector_3.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12)
    ref_tfec_1.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + 6)
    netz_tfec_1.to_excel(writer, sheet_name = economy + '_FED_sector', index = False, startrow = (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + netz_fedsector_3_rows + 15)
    ref_bld_2.to_excel(writer, sheet_name = economy + '_FED_bld', index = False, startrow = chart_height)
    netz_bld_2.to_excel(writer, sheet_name = economy + '_FED_bld', index = False, startrow = (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6)
    ref_bld_3.to_excel(writer, sheet_name = economy + '_FED_bld', index = False, startrow = chart_height + ref_bld_2_rows + 3)
    netz_bld_3.to_excel(writer, sheet_name = economy + '_FED_bld', index = False, startrow = (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9)
    ref_ind_1.to_excel(writer, sheet_name = economy + '_FED_ind', index = False, startrow = chart_height)
    netz_ind_1.to_excel(writer, sheet_name = economy + '_FED_ind', index = False, startrow = (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6)
    ref_ind_2.to_excel(writer, sheet_name = economy + '_FED_ind', index = False, startrow = chart_height + ref_ind_1_rows + 3)
    netz_ind_2.to_excel(writer, sheet_name = economy + '_FED_ind', index = False, startrow = (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9)
    ref_trn_1.to_excel(writer, sheet_name = economy + '_FED_trn', index = False, startrow = chart_height)
    netz_trn_1.to_excel(writer, sheet_name = economy + '_FED_trn', index = False, startrow = (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6)
    ref_trn_2.to_excel(writer, sheet_name = economy + '_FED_trn', index = False, startrow = chart_height + ref_trn_1_rows + 3)
    netz_trn_2.to_excel(writer, sheet_name = economy + '_FED_trn', index = False, startrow = (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9)
    ref_ag_1.to_excel(writer, sheet_name = economy + '_FED_agr', index = False, startrow = chart_height)
    netz_ag_1.to_excel(writer, sheet_name = economy + '_FED_agr', index = False, startrow = (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6)
    ref_ag_2.to_excel(writer, sheet_name = economy + '_FED_agr', index = False, startrow = chart_height + ref_ag_1_rows + 3)
    netz_ag_2.to_excel(writer, sheet_name = economy + '_FED_agr', index = False, startrow = (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9)
    ref_hyd_1.to_excel(writer, sheet_name = economy + '_FED_hyd', index = False, startrow = chart_height)
    netz_hyd_1.to_excel(writer, sheet_name = economy + '_FED_hyd', index = False, startrow = chart_height + ref_hyd_1_rows + 3)

    # TPES
    ref_tpes_1.to_excel(writer, sheet_name = economy + '_TPES', index = False, startrow = chart_height)
    netz_tpes_1.to_excel(writer, sheet_name = economy + '_TPES', index = False, startrow = (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6)
    ref_tpes_2.to_excel(writer, sheet_name = economy + '_TPES', index = False, startrow = chart_height + ref_tpes_1_rows + 3)
    netz_tpes_2.to_excel(writer, sheet_name = economy + '_TPES', index = False, startrow = (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + 9)
    ref_prod_1.to_excel(writer, sheet_name = economy + '_prod', index = False, startrow = chart_height)
    netz_prod_1.to_excel(writer, sheet_name = economy + '_prod', index = False, startrow = (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6)
    ref_prod_2.to_excel(writer, sheet_name = economy + '_prod', index = False, startrow = chart_height + ref_prod_1_rows + 3)
    netz_prod_2.to_excel(writer, sheet_name = economy + '_prod', index = False, startrow = (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9)
    ref_tpes_comp_1.to_excel(writer, sheet_name = economy + '_TPES_comp_ref', index = False, startrow = chart_height)
    netz_tpes_comp_1.to_excel(writer, sheet_name = economy + '_TPES_comp_netz', index = False, startrow = chart_height)
    ref_imports_1.to_excel(writer, sheet_name = economy + '_TPES_comp_ref', index = False, startrow = chart_height + ref_tpes_comp_1_rows + 3)
    netz_imports_1.to_excel(writer, sheet_name = economy + '_TPES_comp_netz', index = False, startrow = chart_height + netz_tpes_comp_1_rows + 3)
    ref_imports_2.to_excel(writer, sheet_name = economy + '_TPES_comp_ref', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6)
    netz_imports_2.to_excel(writer, sheet_name = economy + '_TPES_comp_netz', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6)
    ref_exports_1.to_excel(writer, sheet_name = economy + '_TPES_comp_ref', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9)
    netz_exports_1.to_excel(writer, sheet_name = economy + '_TPES_comp_netz', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9)
    ref_exports_2.to_excel(writer, sheet_name = economy + '_TPES_comp_ref', index = False, startrow = chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12)
    netz_exports_2.to_excel(writer, sheet_name = economy + '_TPES_comp_netz', index = False, startrow = chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12)
    ref_bunkers_1.to_excel(writer, sheet_name = economy + '_TPES_bunkers', index = False, startrow = chart_height)
    netz_bunkers_1.to_excel(writer, sheet_name = economy + '_TPES_bunkers', index = False, startrow = (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6)
    ref_bunkers_2.to_excel(writer, sheet_name = economy + '_TPES_bunkers', index = False, startrow = chart_height + ref_bunkers_1_rows + 3)
    netz_bunkers_2.to_excel(writer, sheet_name = economy + '_TPES_bunkers', index = False, startrow = (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9)

    # Transformation
    ref_pow_use_2.to_excel(writer, sheet_name = economy + '_pow_input', index = False, startrow = chart_height)
    netz_pow_use_2.to_excel(writer, sheet_name = economy + '_pow_input', index = False, startrow = (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6)
    ref_pow_use_3.to_excel(writer, sheet_name = economy + '_pow_input', index = False, startrow = chart_height + ref_pow_use_2_rows + 3)
    netz_pow_use_3.to_excel(writer, sheet_name = economy + '_pow_input', index = False, startrow = (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9)
    ref_elecgen_2.to_excel(writer, sheet_name = economy + '_elec_gen', index = False, startrow = chart_height)
    netz_elecgen_2.to_excel(writer, sheet_name = economy + '_elec_gen', index = False, startrow = (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6)
    ref_elecgen_3.to_excel(writer, sheet_name = economy + '_elec_gen', index = False, startrow = chart_height + ref_elecgen_2_rows + 3)
    netz_elecgen_3.to_excel(writer, sheet_name = economy + '_elec_gen', index = False, startrow = (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9)
    ref_powcap_1.to_excel(writer, sheet_name = economy + '_pow_cap', index = False, startrow = chart_height)
    netz_powcap_1.to_excel(writer, sheet_name = economy + '_pow_cap', index = False, startrow = (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6)
    ref_powcap_2.to_excel(writer, sheet_name = economy + '_pow_cap', index = False, startrow = chart_height + ref_powcap_1_rows + 3)
    netz_powcap_2.to_excel(writer, sheet_name = economy + '_pow_cap', index = False, startrow = (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9)
    ref_refinery_1.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = chart_height)
    netz_refinery_1.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9)
    ref_refinery_2.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = chart_height + ref_refinery_1_rows + 3)
    netz_refinery_2.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12)
    ref_refinery_3.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6)
    netz_refinery_3.to_excel(writer, sheet_name = economy + '_refining', index = False, startrow = (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15)
    ref_trans_3.to_excel(writer, sheet_name = economy + '_trnsfrm', index = False, startrow = chart_height)
    netz_trans_3.to_excel(writer, sheet_name = economy + '_trnsfrm', index = False, startrow = (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6)
    ref_trans_4.to_excel(writer, sheet_name = economy + '_trnsfrm', index = False, startrow = chart_height + ref_trans_3_rows + 3)
    netz_trans_4.to_excel(writer, sheet_name = economy + '_trnsfrm', index = False, startrow = (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + 9)
    ref_ownuse_1.to_excel(writer, sheet_name = economy + '_ownuse', index = False, startrow = chart_height)
    netz_ownuse_1.to_excel(writer, sheet_name = economy + '_ownuse', index = False, startrow = (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6)
    ref_ownuse_2.to_excel(writer, sheet_name = economy + '_ownuse', index = False, startrow = chart_height + ref_ownuse_1_rows + 3)
    netz_ownuse_2.to_excel(writer, sheet_name = economy + '_ownuse', index = False, startrow = (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9)
    ref_heatgen_2.to_excel(writer, sheet_name = economy + '_heat_gen', index = False, startrow = chart_height)
    netz_heatgen_2.to_excel(writer, sheet_name = economy + '_heat_gen', index = False, startrow = (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6)
    ref_heatgen_3.to_excel(writer, sheet_name = economy + '_heat_gen', index = False, startrow = chart_height + ref_heatgen_2_rows + 3)
    netz_heatgen_3.to_excel(writer, sheet_name = economy + '_heat_gen', index = False, startrow = (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9) 

    # Miscellaneous 
    ref_modren_4.to_excel(writer, sheet_name = economy + '_mod_renew', index = False, startrow = chart_height)
    netz_modren_4.to_excel(writer, sheet_name = economy + '_mod_renew', index = False, startrow = chart_height + ref_modren_4_rows + 3)
    ref_enint_3.to_excel(writer, sheet_name = economy + '_eintensity', index = False, startrow = chart_height)
    netz_enint_3.to_excel(writer, sheet_name = economy + '_eintensity', index = False, startrow = chart_height + ref_enint_3_rows + 3)
    macro_1.to_excel(writer, sheet_name = economy + '_macro', index = False, startrow = chart_height)
    
    ################################################################################################################################

    # CHARTS
    # REFERENCE

    # Access the workbook and first sheet with data from df1
    ref_worksheet1 = writer.sheets[economy + '_FED_fuel']
    
    # Comma format and header format        
    space_format = workbook.add_format({'num_format': '# ### ### ##0.0;-# ### ### ##0.0;-'})
    percentage_format = workbook.add_format({'num_format': '0.0%'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})

        
    # Apply comma format and header format to relevant data rows
    ref_worksheet1.set_column(1, ref_fedfuel_1_cols + 1, None, space_format)
    ref_worksheet1.set_row(chart_height, None, header_format)
    ref_worksheet1.set_row(chart_height + ref_fedfuel_1_rows + 3, None, header_format)
    ref_worksheet1.set_row((2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, None, header_format)
    ref_worksheet1.set_row((2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + 9, None, header_format)
    ref_worksheet1.write(0, 0, economy + ' FED fuel reference', cell_format1)
    ref_worksheet1.write(42, 0, economy + ' FED fuel net-zero', cell_format1)

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
    for i in range(ref_fedfuel_1_rows):
        ref_fedfuel_chart1.add_series({
            'name':       [economy + '_FED_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel', chart_height, 2, chart_height, ref_fedfuel_1_cols - 1],
            'values':     [economy + '_FED_fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_fedfuel_1_cols - 1],
            'fill':       {'color': ref_fedfuel_1['fuel_code'].map(colours_dict).loc[i]},
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
        i = ref_fedfuel_2[ref_fedfuel_2['fuel_code'] == component].index[0]
        ref_fedfuel_chart2.add_series({
            'name':       [economy + '_FED_fuel', chart_height + ref_fedfuel_1_rows + i + 4, 0],
            'categories': [economy + '_FED_fuel', chart_height + ref_fedfuel_1_rows + 3, 2, chart_height + ref_fedfuel_1_rows + 3, ref_fedfuel_2_cols - 1],
            'values':     [economy + '_FED_fuel', chart_height + ref_fedfuel_1_rows + i + 4, 2, chart_height + ref_fedfuel_1_rows + i + 4, ref_fedfuel_2_cols - 1],
            'fill':       {'color': ref_fedfuel_2['fuel_code'].map(colours_dict).loc[i]},
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
    for i in range(ref_fedfuel_1_rows):
        ref_fedfuel_chart3.add_series({
            'name':       [economy + '_FED_fuel', chart_height + i + 1, 0],
            'categories': [economy + '_FED_fuel', chart_height, 2, chart_height, ref_fedfuel_1_cols - 1],
            'values':     [economy + '_FED_fuel', chart_height + i + 1, 2, chart_height + i + 1, ref_fedfuel_1_cols - 1],
            'line':       {'color': ref_fedfuel_1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
        })    
        
    ref_worksheet1.insert_chart('R3', ref_fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    ref_worksheet2 = writer.sheets[economy + '_FED_sector']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet2.set_column(1, ref_fedsector_2_cols + 1, None, space_format)
    ref_worksheet2.set_row(chart_height, None, header_format)
    ref_worksheet2.set_row(chart_height + ref_fedsector_2_rows + 3, None, header_format)
    ref_worksheet2.set_row(chart_height + ref_fedsector_2_rows + ref_fedsector_3_rows + 6, None, header_format)
    ref_worksheet2.set_row((2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, None, header_format)
    ref_worksheet2.set_row((2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12, None, header_format)
    ref_worksheet2.set_row((2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + netz_fedsector_3_rows + 15, None, header_format)
    ref_worksheet2.write(0, 0, economy + ' FED sector reference', cell_format1)
    ref_worksheet2.write(40, 0, economy + ' FED sector net-zero', cell_format1)

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
    for i in range(ref_fedsector_2_rows):
        ref_fedsector_chart3.add_series({
            'name':       [economy + '_FED_sector', chart_height + i + 1, 1],
            'categories': [economy + '_FED_sector', chart_height, 2, chart_height, ref_fedsector_2_cols - 1],
            'values':     [economy + '_FED_sector', chart_height + i + 1, 2, chart_height + i + 1, ref_fedsector_2_cols - 1],
            'fill':       {'color': ref_fedsector_2['item_code_new'].map(colours_dict).loc[i]},
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
        i = ref_fedsector_3[ref_fedsector_3['item_code_new'] == component].index[0]
        ref_fedsector_chart4.add_series({
            'name':       [economy + '_FED_sector', chart_height + ref_fedsector_2_rows + i + 4, 1],
            'categories': [economy + '_FED_sector', chart_height + ref_fedsector_2_rows + 3, 2, chart_height + ref_fedsector_2_rows + 3, ref_fedsector_3_cols - 1],
            'values':     [economy + '_FED_sector', chart_height + ref_fedsector_2_rows + i + 4, 2, chart_height + ref_fedsector_2_rows + i + 4, ref_fedsector_3_cols - 1],
            'fill':       {'color': ref_fedsector_3['item_code_new'].map(colours_dict).loc[i]},
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
    for i in range(ref_fedsector_2_rows):
        ref_fedsector_chart5.add_series({
            'name':       [economy + '_FED_sector', chart_height + i + 1, 1],
            'categories': [economy + '_FED_sector', chart_height, 2, chart_height, ref_fedsector_2_cols - 1],
            'values':     [economy + '_FED_sector', chart_height + i + 1, 2, chart_height + i + 1, ref_fedsector_2_cols - 1],
            'line':       {'color': ref_fedsector_2['item_code_new'].map(colours_dict).loc[i], 'width': 1.25}
        })    
        
    ref_worksheet2.insert_chart('R3', ref_fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    ref_worksheet3 = writer.sheets[economy + '_FED_bld']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet3.set_column(2, ref_bld_2_cols + 1, None, space_format)
    ref_worksheet3.set_row(chart_height, None, header_format)
    ref_worksheet3.set_row(chart_height + ref_bld_2_rows + 3, None, header_format)
    ref_worksheet3.set_row((2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6, None, header_format)
    ref_worksheet3.set_row((2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9, None, header_format)
    ref_worksheet3.write(0, 0, economy + ' buildings reference', cell_format1)
    ref_worksheet3.write(35, 0, economy + ' buildings net-zero', cell_format1)
    
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
        i = ref_bld_2[ref_bld_2['fuel_code'] == component].index[0]
        ref_fed_bld_chart1.add_series({
            'name':       [economy + '_FED_bld', chart_height + i + 1, 0],
            'categories': [economy + '_FED_bld', chart_height, 2, chart_height, ref_bld_2_cols - 1],
            'values':     [economy + '_FED_bld', chart_height + i + 1, 2, chart_height + i + 1, ref_bld_2_cols - 1],
            'fill':       {'color': ref_bld_2['fuel_code'].map(colours_dict).loc[i]},
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
    for bld_sect in ['Services', 'Residential']:
        i = ref_bld_3[ref_bld_3['item_code_new'] == bld_sect].index[0]
        ref_fed_bld_chart2.add_series({
            'name':       [economy + '_FED_bld', chart_height + ref_bld_2_rows + 4 + i, 1],
            'categories': [economy + '_FED_bld', chart_height + ref_bld_2_rows + 3, 2, chart_height + ref_bld_2_rows + 3, ref_bld_3_cols - 1],
            'values':     [economy + '_FED_bld', chart_height + ref_bld_2_rows + 4 + i, 2, chart_height + ref_bld_2_rows + 4 + i, ref_bld_3_cols - 1],
            'fill':       {'color': ref_bld_3['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('J3', ref_fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # Access the workbook and fourth sheet with data from bld_df1
    ref_worksheet4 = writer.sheets[economy + '_FED_ind']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet4.set_column(2, ref_ind_1_cols + 1, None, space_format)
    ref_worksheet4.set_row(chart_height, None, header_format)
    ref_worksheet4.set_row(chart_height + ref_ind_1_rows + 3, None, header_format)
    ref_worksheet4.set_row((2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6, None, header_format)
    ref_worksheet4.set_row((2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9, None, header_format)
    ref_worksheet4.write(0, 0, economy + ' industry reference', cell_format1)
    ref_worksheet4.write(40, 0, economy + ' industry net-zero', cell_format1)
    
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
    for i in range(ref_ind_1_rows):
        ref_fed_ind_chart1.add_series({
            'name':       [economy + '_FED_ind', chart_height + i + 1, 1],
            'categories': [economy + '_FED_ind', chart_height, 2, chart_height, ref_ind_1_cols - 1],
            'values':     [economy + '_FED_ind', chart_height + i + 1, 2, chart_height + i + 1, ref_ind_1_cols - 1],
            'fill':       {'color': ref_ind_1['item_code_new'].map(colours_dict).loc[i]},
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
        j = ref_ind_2[ref_ind_2['fuel_code'] == fuel_agg].index[0]
        ref_fed_ind_chart2.add_series({
            'name':       [economy + '_FED_ind', chart_height + ref_ind_1_rows + j + 4, 0],
            'categories': [economy + '_FED_ind', chart_height + ref_ind_1_rows + 3, 2, chart_height + ref_ind_1_rows + 3, ref_ind_2_cols - 1],
            'values':     [economy + '_FED_ind', chart_height + ref_ind_1_rows + j + 4, 2, chart_height + ref_ind_1_rows + j + 4, ref_ind_2_cols - 1],
            'fill':       {'color': ref_ind_2['fuel_code'].map(colours_dict).loc[j]},
            'border':     {'none': True}
        })
    
    ref_worksheet4.insert_chart('J3', ref_fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    ref_worksheet5 = writer.sheets[economy + '_FED_trn']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet5.set_column(2, ref_trn_1_cols + 1, None, space_format)
    ref_worksheet5.set_row(chart_height, None, header_format)
    ref_worksheet5.set_row(chart_height + ref_trn_1_rows + 3, None, header_format)
    ref_worksheet5.set_row((2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6, None, header_format)
    ref_worksheet5.set_row((2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9, None, header_format)
    ref_worksheet5.write(0, 0, economy + ' FED transport reference', cell_format1)
    ref_worksheet5.write(39, 0, economy + ' FED transport net-zero', cell_format1)
    
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
        j = ref_trn_1[ref_trn_1['fuel_code'] == fuel_agg].index[0]
        ref_transport_chart1.add_series({
            'name':       [economy + '_FED_trn', chart_height + j + 1, 0],
            'categories': [economy + '_FED_trn', chart_height, 2, chart_height, ref_trn_1_cols - 1],
            'values':     [economy + '_FED_trn', chart_height + j + 1, 2, chart_height + j + 1, ref_trn_1_cols - 1],
            'fill':       {'color': ref_trn_1['fuel_code'].map(colours_dict).loc[j]},
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
        j = ref_trn_2[ref_trn_2['item_code_new'] == modality].index[0]
        ref_transport_chart2.add_series({
            'name':       [economy + '_FED_trn', chart_height + ref_trn_1_rows + j + 4, 1],
            'categories': [economy + '_FED_trn', chart_height + ref_trn_1_rows + 3, 2, chart_height + ref_trn_1_rows + 3, ref_trn_2_cols - 1],
            'values':     [economy + '_FED_trn', chart_height + ref_trn_1_rows + j + 4, 2, chart_height + ref_trn_1_rows + j + 4, ref_trn_2_cols - 1],
            'fill':       {'color': ref_trn_2['item_code_new'].map(colours_dict).loc[j]},
            'border':     {'none': True}
        })
    
    ref_worksheet5.insert_chart('J3', ref_transport_chart2)

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    ref_worksheet6 = writer.sheets[economy + '_FED_agr']
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet6.set_column(2, ref_ag_1_cols + 1, None, space_format)
    ref_worksheet6.set_row(chart_height, None, header_format)
    ref_worksheet6.set_row(chart_height + ref_ag_1_rows + 3, None, header_format)
    ref_worksheet6.set_row((2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, None, header_format)
    ref_worksheet6.set_row((2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9, None, header_format)
    ref_worksheet6.write(0, 0, economy + ' FED agriculture reference', cell_format1)
    ref_worksheet6.write(42, 0, economy + ' FED agriculture net-zero', cell_format1)

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
    for i in range(ref_ag_1_rows):
        ref_ag_chart1.add_series({
            'name':       [economy + '_FED_agr', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agr', chart_height, 2, chart_height, ref_ag_1_cols - 1],
            'values':     [economy + '_FED_agr', chart_height + i + 1, 2, chart_height + i + 1, ref_ag_1_cols - 1],
            'line':       {'color': ref_ag_1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
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
    for i in range(ref_ag_1_rows):
        ref_ag_chart2.add_series({
            'name':       [economy + '_FED_agr', chart_height + i + 1, 0],
            'categories': [economy + '_FED_agr', chart_height, 2, chart_height, ref_ag_1_cols - 1],
            'values':     [economy + '_FED_agr', chart_height + i + 1, 2, chart_height + i + 1, ref_ag_1_cols - 1],
            'fill':       {'color': ref_ag_1['fuel_code'].map(colours_dict).loc[i]},
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
    for i in range(ref_ag_2_rows):
        ref_ag_chart3.add_series({
            'name':       [economy + '_FED_agr', chart_height + ref_ag_1_rows + i + 4, 0],
            'categories': [economy + '_FED_agr', chart_height + ref_ag_1_rows + 3, 2, chart_height + ref_ag_1_rows + 3, ref_ag_2_cols - 1],
            'values':     [economy + '_FED_agr', chart_height + ref_ag_1_rows + i + 4, 2, chart_height + ref_ag_1_rows + i + 4, ref_ag_2_cols - 1],
            'fill':       {'color': ref_ag_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet6.insert_chart('R3', ref_ag_chart3)

    # HYDROGEN CHARTS

    # Access the workbook and first sheet with data from df1
    hyd_worksheet1 = writer.sheets[economy + '_FED_hyd']
    
    # Comma format and header format        
    # space_format = workbook.add_format({'num_format': '#,##0'})
    # header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    # cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    hyd_worksheet1.set_column(1, ref_hyd_1_cols + 1, None, space_format)
    hyd_worksheet1.set_row(chart_height, None, header_format)
    hyd_worksheet1.set_row(chart_height, None, header_format)
    hyd_worksheet1.set_row(chart_height + ref_hyd_1_rows + 3, None, header_format)
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
    for i in range(ref_hyd_1_rows):
        ref_hyd_chart1.add_series({
            'name':       [economy + '_FED_hyd', chart_height + i + 1, 1],
            'categories': [economy + '_FED_hyd', chart_height, 2, chart_height, ref_hyd_1_cols - 1],
            'values':     [economy + '_FED_hyd', chart_height + i + 1, 2, chart_height + i + 1, ref_hyd_1_cols - 1],
            'fill':       {'color': ref_hyd_1['item_code_new'].map(colours_dict).loc[i]},
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
    for i in range(netz_hyd_1_rows):
        netz_hyd_chart1.add_series({
            'name':       [economy + '_FED_hyd', chart_height + ref_hyd_1_rows + i + 4, 1],
            'categories': [economy + '_FED_hyd', chart_height + ref_hyd_1_rows + 3, 2, chart_height + ref_hyd_1_rows + 3, netz_hyd_1_cols - 1],
            'values':     [economy + '_FED_hyd', chart_height + ref_hyd_1_rows + i + 4, 2, chart_height + ref_hyd_1_rows + i + 4, netz_hyd_1_cols - 1],
            'fill':       {'color': netz_hyd_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    hyd_worksheet1.insert_chart('J3', netz_hyd_chart1)

    ##############################################################################################################################

    ##############################################################################################################################

    ##############################################################################################################################

    # CHARTS
    # NET ZERO

    # Access the workbook and first sheet with data from df1
    # netz_worksheet1 = writer.sheets[economy + '_FED_fuel']
    
    # Comma format and header format        
    # space_format = workbook.add_format({'num_format': '#,##0'})
    # header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    # cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    # netz_worksheet1.set_column(1, netz_fedfuel_1_cols + 1, None, space_format)
    # netz_worksheet1.set_row(chart_height, None, header_format)
    # netz_worksheet1.set_row(chart_height, None, header_format)
    # netz_worksheet1.set_row(chart_height + netz_fedfuel_1_rows + 3, None, header_format)
    # netz_worksheet1.write(0, 0, economy + ' FED fuel', cell_format1)

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
    for i in range(netz_fedfuel_1_rows):
        netz_fedfuel_chart1.add_series({
            'name':       [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 0],
            'categories': [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, 2,\
                 (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, netz_fedfuel_1_cols - 1],
            'values':     [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 2,\
                 (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, netz_fedfuel_1_cols - 1],
            'fill':       {'color': netz_fedfuel_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet1.insert_chart('B45', netz_fedfuel_chart1)

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
        i = netz_fedfuel_2[netz_fedfuel_2['fuel_code'] == component].index[0]
        netz_fedfuel_chart2.add_series({
            'name':       [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + i + 10, 0],
            'categories': [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + 9,\
                 2, (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + 9, netz_fedfuel_2_cols - 1],
            'values':     [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + i + 10,\
                 2, (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + netz_fedfuel_1_rows + i + 10, netz_fedfuel_2_cols - 1],
            'fill':       {'color': netz_fedfuel_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet1.insert_chart('J45', netz_fedfuel_chart2)

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
    for i in range(netz_fedfuel_1_rows):
        netz_fedfuel_chart3.add_series({
            'name':       [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 0],
            'categories': [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, 2,\
                 (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + 6, netz_fedfuel_1_cols - 1],
            'values':     [economy + '_FED_fuel', (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, 2,\
                 (2 * chart_height) + ref_fedfuel_1_rows + ref_fedfuel_2_rows + i + 7, netz_fedfuel_1_cols - 1],
            'line':       {'color': netz_fedfuel_1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
        })    
        
    ref_worksheet1.insert_chart('R45', netz_fedfuel_chart3)

    ############################## Next sheet: FED (TFC) by sector ##############################
    
    # Access the workbook and second sheet with data from df2
    # netz_worksheet2 = writer.sheets[economy + '_FED_sector']
        
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet2.set_column(1, netz_fedsector_1_cols + 1, None, space_format)
    # netz_worksheet2.set_row(chart_height, None, header_format)
    # netz_worksheet2.set_row(chart_height + netz_fedsector_2_rows + 3, None, header_format)
    # netz_worksheet2.set_row(chart_height + netz_fedsector_2_rows + netz_fedsector_3_rows + 6, None, header_format)
    # netz_worksheet2.write(0, 0, economy + ' FED sector', cell_format1)

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
    for i in range(netz_fedsector_2_rows):
        netz_fedsector_chart3.add_series({
            'name':       [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 1],
            'categories': [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, 2, (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, netz_fedsector_2_cols - 1],
            'values':     [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 2, (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, netz_fedsector_2_cols - 1],
            'fill':       {'color': netz_fedsector_2['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet2.insert_chart('B43', netz_fedsector_chart3)

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
        i = netz_fedsector_3[netz_fedsector_3['item_code_new'] == component].index[0]
        netz_fedsector_chart4.add_series({
            'name':       [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + i + 13, 1],
            'categories': [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12, 2,\
                (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + 12, netz_fedsector_3_cols - 1],
            'values':     [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + i + 13, 2,\
                (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + netz_fedsector_2_rows + i + 13, netz_fedsector_3_cols - 1],
            'fill':       {'color': netz_fedsector_3['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet2.insert_chart('J43', netz_fedsector_chart4)

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
    for i in range(netz_fedsector_2_rows):
        netz_fedsector_chart5.add_series({
            'name':       [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 1],
            'categories': [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, 2,\
                (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + 9, netz_fedsector_2_cols - 1],
            'values':     [economy + '_FED_sector', (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, 2,\
                (2 * chart_height) + ref_fedsector_2_rows + ref_fedsector_3_rows + ref_tfec_1_rows + i + 10, netz_fedsector_2_cols - 1],
            'line':       {'color': netz_fedsector_2['item_code_new'].map(colours_dict).loc[i], 'width': 1.25}
        })    
        
    ref_worksheet2.insert_chart('R43', netz_fedsector_chart5)
    
    ############################# Next sheet: FED (TFC) for building sector ##################################
    
    # Access the workbook and third sheet with data from bld_df1
    # netz_worksheet3 = writer.sheets[economy + '_FED_bld']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet3.set_column(2, netz_bld_2_cols + 1, None, space_format)
    # netz_worksheet3.set_row(chart_height, None, header_format)
    # netz_worksheet3.set_row(chart_height + netz_bld_2_rows + 3, None, header_format)
    # netz_worksheet3.write(0, 0, economy + ' buildings', cell_format1)
    
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
        i = netz_bld_2[netz_bld_2['fuel_code'] == component].index[0]
        netz_fed_bld_chart1.add_series({
            'name':       [economy + '_FED_bld', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + i + 7, 0],
            'categories': [economy + '_FED_bld', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6, 2,\
                (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + 6, netz_bld_2_cols - 1],
            'values':     [economy + '_FED_bld', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + i + 7, 2,\
                (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + i + 7, netz_bld_2_cols - 1],
            'fill':       {'color': netz_bld_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })

    ref_worksheet3.insert_chart('B38', netz_fed_bld_chart1)
    
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
    for bld_sect in ['Services', 'Residential']:
        i = netz_bld_3[netz_bld_3['item_code_new'] == bld_sect].index[0]
        netz_fed_bld_chart2.add_series({
            'name':       [economy + '_FED_bld', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 10 + i, 1],
            'categories': [economy + '_FED_bld', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9, 2,\
                (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 9, netz_bld_3_cols - 1],
            'values':     [economy + '_FED_bld', (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 10 + i, 2,\
                (2 * chart_height) + ref_bld_2_rows + ref_bld_3_rows + netz_bld_2_rows + 10 + i, netz_bld_3_cols - 1],
            'fill':       {'color': netz_bld_3['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet3.insert_chart('J38', netz_fed_bld_chart2)
    
    ############################# Next sheet: FED (TFC) for industry ##################################
    
    # # Access the workbook and fourth sheet with data from bld_df1
    # netz_worksheet4 = writer.sheets[economy + '_FED_ind']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet4.set_column(2, netz_ind_1_cols + 1, None, space_format)
    # netz_worksheet4.set_row(chart_height, None, header_format)
    # netz_worksheet4.set_row(chart_height + netz_ind_1_rows + 2, None, header_format)
    # netz_worksheet4.write(0, 0, economy + ' industry', cell_format1)
    
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
    for i in range(netz_ind_1_rows):
        netz_fed_ind_chart1.add_series({
            'name':       [economy + '_FED_ind', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + i + 7, 1],
            'categories': [economy + '_FED_ind', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6, 2,\
                (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + 6, netz_ind_1_cols - 1],
            'values':     [economy + '_FED_ind', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + i + 7, netz_ind_1_cols - 1],
            'fill':       {'color': netz_ind_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet4.insert_chart('B43', netz_fed_ind_chart1)
    
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
        j = netz_ind_2[netz_ind_2['fuel_code'] == fuel_agg].index[0]
        netz_fed_ind_chart2.add_series({
            'name':       [economy + '_FED_ind', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + j + 10, 0],
            'categories': [economy + '_FED_ind', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9, 2,\
                (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + 9, netz_ind_2_cols - 1],
            'values':     [economy + '_FED_ind', (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + j + 10, 2,\
                (2 * chart_height) + ref_ind_1_rows + ref_ind_2_rows + netz_ind_1_rows + j + 10, netz_ind_2_cols - 1],
            'fill':       {'color': netz_ind_2['fuel_code'].map(colours_dict).loc[j]},
            'border':     {'none': True}
        })
    
    ref_worksheet4.insert_chart('J43', netz_fed_ind_chart2)

    ################################# NEXT SHEET: TRANSPORT FED ################################################################

    # Access the workbook and first sheet with data from df1
    # netz_worksheet5 = writer.sheets[economy + '_FED_trn']
        
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet5.set_column(2, netz_trn_1_cols + 1, None, space_format)
    # netz_worksheet5.set_row(chart_height, None, header_format)
    # netz_worksheet5.set_row(chart_height + netz_trn_1_rows + 3, None, header_format)
    # netz_worksheet5.write(0, 0, economy + ' FED transport', cell_format1)
    
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
        j = netz_trn_1[netz_trn_1['fuel_code'] == fuel_agg].index[0]
        netz_transport_chart1.add_series({
            'name':       [economy + '_FED_trn', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + j + 7, 0],
            'categories': [economy + '_FED_trn', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6, 2,\
                (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + 6, netz_trn_1_cols - 1],
            'values':     [economy + '_FED_trn', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + j + 7, 2,\
                (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + j + 7, netz_trn_1_cols - 1],
            'fill':       {'color': netz_trn_1['fuel_code'].map(colours_dict).loc[j]},
            'border':     {'none': True} 
        })
    
    ref_worksheet5.insert_chart('B42', netz_transport_chart1)
            
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
        j = netz_trn_2[netz_trn_2['item_code_new'] == modality].index[0]
        netz_transport_chart2.add_series({
            'name':       [economy + '_FED_trn', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + j + 10, 1],
            'categories': [economy + '_FED_trn', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9, 2,\
                (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + 9, netz_trn_2_cols - 1],
            'values':     [economy + '_FED_trn', (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + j + 10, 2,\
                (2 * chart_height) + ref_trn_1_rows + ref_trn_2_rows + netz_trn_1_rows + j + 10, netz_trn_2_cols - 1],
            'fill':       {'color': netz_trn_2['item_code_new'].map(colours_dict).loc[j]},
            'border':     {'none': True}
        })
    
    ref_worksheet5.insert_chart('J42', netz_transport_chart2)

    ################################# NEXT SHEET: AGRICULTURE FED ################################################################

    # Access the workbook and first sheet with data from df1
    # netz_worksheet6 = writer.sheets[economy + '_FED_agr']
        
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet6.set_column(2, netz_ag_1_cols + 1, None, space_format)
    # netz_worksheet6.set_row(chart_height, None, header_format)
    # netz_worksheet6.set_row(chart_height + netz_ag_1_rows + 3, None, header_format)
    # netz_worksheet6.write(0, 0, economy + ' FED agriculture', cell_format1)

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
    for i in range(netz_ag_1_rows):
        netz_ag_chart1.add_series({
            'name':       [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 0],
            'categories': [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, 2,\
                (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, netz_ag_1_cols - 1],
            'values':     [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, netz_ag_1_cols - 1],
            'line':       {'color': netz_ag_1['fuel_code'].map(colours_dict).loc[i], 'width': 1.25}
        })    
        
    ref_worksheet6.insert_chart('B45', netz_ag_chart1)

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
    for i in range(netz_ag_1_rows):
        netz_ag_chart2.add_series({
            'name':       [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 0],
            'categories': [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, 2,\
                (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + 6, netz_ag_1_cols - 1],
            'values':     [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + i + 7, netz_ag_1_cols - 1],
            'fill':       {'color': netz_ag_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet6.insert_chart('J45', netz_ag_chart2)

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
    for i in range(netz_ag_2_rows):
        netz_ag_chart3.add_series({
            'name':       [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + i + 10, 0],
            'categories': [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9, 2,\
                (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + 9, netz_ag_2_cols - 1],
            'values':     [economy + '_FED_agr', (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + i + 10,\
                2, (2 * chart_height) + ref_ag_1_rows + ref_ag_2_rows + netz_ag_1_rows + i + 10, netz_ag_2_cols - 1],
            'fill':       {'color': netz_ag_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet6.insert_chart('R45', netz_ag_chart3)

    ############################################################################################################################

    # TPES charts

    ################################################################### CHARTS ###################################################################
    # REFERENCE
    # Access the sheet created using writer above
    ref_worksheet11 = writer.sheets[economy + '_TPES']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet11.set_column(2, ref_tpes_1_cols + 1, None, space_format)
    ref_worksheet11.set_row(chart_height, None, header_format)
    ref_worksheet11.set_row(chart_height + ref_tpes_1_rows + 3, None, header_format)
    ref_worksheet11.set_row((2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, None, header_format)
    ref_worksheet11.set_row((2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + ref_tpes_1_rows + 9, None, header_format)
    ref_worksheet11.write(0, 0, economy + ' TPES fuel reference', cell_format1)
    ref_worksheet11.write(36, 0, economy + ' TPES fuel net-zero', cell_format1)

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
    for i in range(ref_tpes_1_rows):
        ref_tpes_chart2.add_series({
            'name':       [economy + '_TPES', chart_height + i + 1, 0],
            'categories': [economy + '_TPES', chart_height, 2, chart_height, ref_tpes_1_cols - 1],
            'values':     [economy + '_TPES', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_1_cols - 1],
            'fill':       {'color': ref_tpes_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet11.insert_chart('B3', ref_tpes_chart2)

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
    for i in range(ref_tpes_1_rows):
        ref_tpes_chart4.add_series({
            'name':       [economy + '_TPES', chart_height + i + 1, 0],
            'categories': [economy + '_TPES', chart_height, 2, chart_height, ref_tpes_1_cols - 1],
            'values':     [economy + '_TPES', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_1_cols - 1],
            'line':       {'color': ref_tpes_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1}
        })    
        
    ref_worksheet11.insert_chart('R3', ref_tpes_chart4)

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
        i = ref_tpes_2[ref_tpes_2['fuel_code'] == component].index[0]
        ref_tpes_chart3.add_series({
            'name':       [economy + '_TPES', chart_height + ref_tpes_1_rows + i + 4, 0],
            'categories': [economy + '_TPES', chart_height + ref_tpes_1_rows + 3, 2, chart_height + ref_tpes_1_rows + 3, ref_tpes_2_cols - 1],
            'values':     [economy + '_TPES', chart_height + ref_tpes_1_rows + i + 4, 2, chart_height + ref_tpes_1_rows + i + 4, ref_tpes_2_cols - 1],
            'fill':       {'color': ref_tpes_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet11.insert_chart('J3', ref_tpes_chart3)

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    ref_worksheet12 = writer.sheets[economy + '_prod']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet12.set_column(2, ref_prod_1_cols + 1, None, space_format)
    ref_worksheet12.set_row(chart_height, None, header_format)
    ref_worksheet12.set_row(chart_height + ref_prod_1_rows + 3, None, header_format)
    ref_worksheet12.set_row((2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, None, header_format)
    ref_worksheet12.set_row((2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9, None, header_format)
    ref_worksheet12.write(0, 0, economy + ' prod fuel reference', cell_format1)
    ref_worksheet12.write(36, 0, economy + ' prod fuel net-zero', cell_format1)

    (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows

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
    for i in range(ref_prod_1_rows):
        ref_prod_chart2.add_series({
            'name':       [economy + '_prod', chart_height + i + 1, 0],
            'categories': [economy + '_prod', chart_height, 2, chart_height, ref_prod_1_cols - 1],
            'values':     [economy + '_prod', chart_height + i + 1, 2, chart_height + i + 1, ref_prod_1_cols - 1],
            'fill':       {'color': ref_prod_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet12.insert_chart('B3', ref_prod_chart2)

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
    for i in range(ref_prod_1_rows):
        ref_prod_chart2.add_series({
            'name':       [economy + '_prod', chart_height + i + 1, 0],
            'categories': [economy + '_prod', chart_height, 2, chart_height, ref_prod_1_cols - 1],
            'values':     [economy + '_prod', chart_height + i + 1, 2, chart_height + i + 1, ref_prod_1_cols - 1],
            'line':       {'color': ref_prod_1['fuel_code'].map(colours_dict).loc[i],
                           'width': 1} 
        })    
        
    ref_worksheet12.insert_chart('R3', ref_prod_chart2)

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
        i = ref_prod_2[ref_prod_2['fuel_code'] == component].index[0]
        ref_prod_chart3.add_series({
            'name':       [economy + '_prod', chart_height + ref_prod_1_rows + i + 4, 0],
            'categories': [economy + '_prod', chart_height + ref_prod_1_rows + 3, 2, chart_height + ref_prod_1_rows + 3, ref_prod_2_cols - 1],
            'values':     [economy + '_prod', chart_height + ref_prod_1_rows + i + 4, 2, chart_height + ref_prod_1_rows + i + 4, ref_prod_2_cols - 1],
            'fill':       {'color': ref_prod_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet12.insert_chart('J3', ref_prod_chart3)
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    ref_worksheet13 = writer.sheets[economy + '_TPES_comp_ref']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet13.set_column(2, ref_imports_1_cols + 1, None, space_format)
    ref_worksheet13.set_row(chart_height, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + 3, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9, None, header_format)
    ref_worksheet13.set_row(chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12, None, header_format)
    ref_worksheet13.write(0, 0, economy + ' TPES components reference', cell_format1)
    
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
        i = ref_tpes_comp_1[ref_tpes_comp_1['item_code_new'] == component].index[0]
        ref_tpes_comp_chart1.add_series({
            'name':       [economy + '_TPES_comp_ref', chart_height + i + 1, 1],
            'categories': [economy + '_TPES_comp_ref', chart_height, 2, chart_height, ref_tpes_comp_1_cols - 1],
            'values':     [economy + '_TPES_comp_ref', chart_height + i + 1, 2, chart_height + i + 1, ref_tpes_comp_1_cols - 1],
            'fill':       {'color': ref_tpes_comp_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet13.insert_chart('B3', ref_tpes_comp_chart1)

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
        i = ref_imports_1[ref_imports_1['fuel_code'] == fuel].index[0]
        ref_imports_line.add_series({
            'name':       [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + i + 4, 0],
            'categories': [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + 3, 2, chart_height + ref_tpes_comp_1_rows + 3, ref_imports_1_cols - 1],
            'values':     [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + i + 4, 2, chart_height + ref_tpes_comp_1_rows + i + 4, ref_imports_1_cols - 1],
            'line':       {'color': ref_imports_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    ref_worksheet13.insert_chart('J3', ref_imports_line)

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
    for i in range(ref_imports_2_rows):
        ref_imports_column.add_series({
            'name':       [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + i + 7, 0],
            'categories': [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + 6, ref_imports_2_cols - 1],
            'values':     [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + i + 7, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + i + 7, ref_imports_2_cols - 1],
            'fill':       {'color': ref_imports_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet13.insert_chart('R3', ref_imports_column)

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
        i = ref_exports_1[ref_exports_1['fuel_code'] == fuel].index[0]
        ref_exports_line.add_series({
            'name':       [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + i + 10, 0],
            'categories': [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + 9, ref_imports_1_cols - 1],
            'values':     [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + i + 10, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + i + 10, ref_imports_1_cols - 1],
            'line':       {'color': ref_exports_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    ref_worksheet13.insert_chart('Z3', ref_exports_line)

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
    for i in range(ref_exports_2_rows):
        ref_exports_column.add_series({
            'name':       [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + i + 13, 0],
            'categories': [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + 12, ref_exports_2_cols - 1],
            'values':     [economy + '_TPES_comp_ref', chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + i + 13, 2, chart_height + ref_tpes_comp_1_rows + ref_imports_1_rows + ref_imports_2_rows + ref_exports_1_rows + i + 13, ref_exports_2_cols - 1],
            'fill':       {'color': ref_exports_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet13.insert_chart('AH3', ref_exports_column)

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    ref_worksheet14 = writer.sheets[economy + '_TPES_bunkers']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet14.set_column(2, ref_bunkers_1_cols + 1, None, space_format)
    ref_worksheet14.set_row(chart_height, None, header_format)
    ref_worksheet14.set_row(chart_height + ref_bunkers_1_rows + 3, None, header_format)
    ref_worksheet14.set_row((2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, None, header_format)
    ref_worksheet14.set_row((2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9, None, header_format)
    ref_worksheet14.write(0, 0, economy + ' TPES bunkers reference', cell_format1)
    ref_worksheet14.write(28, 0, economy + ' TPES bunkers net-zero', cell_format1)
    
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
    for i in range(ref_bunkers_1_rows):
        ref_marine_line.add_series({
            'name':       [economy + '_TPES_bunkers', chart_height + i + 1, 0],
            'categories': [economy + '_TPES_bunkers', chart_height, 2, chart_height, ref_bunkers_1_cols - 1],
            'values':     [economy + '_TPES_bunkers', chart_height + i + 1, 2, chart_height + i + 1, ref_bunkers_1_cols - 1],
            'line':       {'color': ref_bunkers_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    ref_worksheet14.insert_chart('B3', ref_marine_line)

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
    for i in range(ref_bunkers_2_rows):
        ref_aviation_line.add_series({
            'name':       [economy + '_TPES_bunkers', chart_height + ref_bunkers_1_rows + i + 4, 0],
            'categories': [economy + '_TPES_bunkers', chart_height + ref_bunkers_1_rows + 3, 2, chart_height + ref_bunkers_1_rows + 3, ref_bunkers_2_cols - 1],
            'values':     [economy + '_TPES_bunkers', chart_height + ref_bunkers_1_rows + i + 4, 2, chart_height + ref_bunkers_1_rows + i + 4, ref_bunkers_2_cols - 1],
            'line':       {'color': ref_bunkers_2['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    ref_worksheet14.insert_chart('J3', ref_aviation_line)

    ###############################################################################################################
    # Net-zero
    # Access the sheet created using writer above
    # netz_worksheet11 = writer.sheets[economy + '_TPES']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet11.set_column(2, netz_tpes_1_cols + 1, None, space_format)
    # netz_worksheet11.set_row(chart_height, None, header_format)
    # netz_worksheet11.set_row(chart_height + netz_tpes_1_rows + 3, None, header_format)
    # netz_worksheet11.write(0, 0, economy + ' TPES fuel net-zero', cell_format1)

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
    for i in range(netz_tpes_1_rows):
        netz_tpes_chart2.add_series({
            'name':       [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 0],
            'categories': [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, 2,\
                (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, netz_tpes_1_cols - 1],
            'values':     [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, netz_tpes_1_cols - 1],
            'fill':       {'color': netz_tpes_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet11.insert_chart('B39', netz_tpes_chart2)

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
    for i in range(netz_tpes_1_rows):
        netz_tpes_chart4.add_series({
            'name':       [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 0],
            'categories': [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, 2,\
                (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + 6, netz_tpes_1_cols - 1],
            'values':     [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + i + 7, netz_tpes_1_cols - 1],
            'line':       {'color': netz_tpes_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1}
        })    
        
    ref_worksheet11.insert_chart('R39', netz_tpes_chart4)

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
        i = netz_tpes_2[netz_tpes_2['fuel_code'] == component].index[0]
        netz_tpes_chart3.add_series({
            'name':       [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + i + 10, 0],
            'categories': [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + 9, 2,\
                (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + 9, netz_tpes_2_cols - 1],
            'values':     [economy + '_TPES', (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + i + 10, 2,\
                (2 * chart_height) + ref_tpes_1_rows + ref_tpes_2_rows + netz_tpes_1_rows + i + 10, netz_tpes_2_cols - 1],
            'fill':       {'color': netz_tpes_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet11.insert_chart('J39', netz_tpes_chart3)

    ########################################### PRODUCTION CHARTS #############################################
    
    # access the sheet for production created above
    # netz_worksheet12 = writer.sheets[economy + '_prod']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet12.set_column(2, netz_prod_1_cols + 1, None, space_format)
    # netz_worksheet12.set_row(chart_height, None, header_format)
    # netz_worksheet12.set_row(chart_height + netz_prod_1_rows + 3, None, header_format)
    # netz_worksheet12.write(0, 0, economy + ' prod fuel net-zero', cell_format1)

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
    for i in range(netz_prod_1_rows):
        netz_prod_chart2.add_series({
            'name':       [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 0],
            'categories': [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, 2,\
                (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, netz_prod_1_cols - 1],
            'values':     [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, netz_prod_1_cols - 1],
            'fill':       {'color': netz_prod_1['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })    
        
    ref_worksheet12.insert_chart('B39', netz_prod_chart2)

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
    for i in range(netz_prod_1_rows):
        netz_prod_chart2.add_series({
            'name':       [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 0],
            'categories': [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, 2,\
                (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + 6, netz_prod_1_cols - 1],
            'values':     [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + i + 7, netz_prod_1_cols - 1],
            'line':       {'color': netz_prod_1['fuel_code'].map(colours_dict).loc[i],
                           'width': 1} 
        })    
        
    ref_worksheet12.insert_chart('R39', netz_prod_chart2)

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
        i = netz_prod_2[netz_prod_2['fuel_code'] == component].index[0]
        netz_prod_chart3.add_series({
            'name':       [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + i + 10, 0],
            'categories': [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9, 2,\
                (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + 9, netz_prod_2_cols - 1],
            'values':     [economy + '_prod', (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + i + 10, 2,\
                (2 * chart_height) + ref_prod_1_rows + ref_prod_2_rows + netz_prod_1_rows + i + 10, netz_prod_2_cols - 1],
            'fill':       {'color': netz_prod_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    ref_worksheet12.insert_chart('J39', netz_prod_chart3)
    
    ###################################### TPES components I ###########################################
    
    # access the sheet for production created above
    netz_worksheet13 = writer.sheets[economy + '_TPES_comp_netz']
    
    # Apply comma format and header format to relevant data rows
    netz_worksheet13.set_column(2, netz_imports_1_cols + 1, None, space_format)
    netz_worksheet13.set_row(chart_height, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + 3, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9, None, header_format)
    netz_worksheet13.set_row(chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12, None, header_format)
    netz_worksheet13.write(0, 0, economy + ' TPES components net-zero', cell_format1)
    
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
        i = netz_tpes_comp_1[netz_tpes_comp_1['item_code_new'] == component].index[0]
        netz_tpes_comp_chart1.add_series({
            'name':       [economy + '_TPES_comp_netz', chart_height + i + 1, 1],
            'categories': [economy + '_TPES_comp_netz', chart_height, 2, chart_height, netz_tpes_comp_1_cols - 1],
            'values':     [economy + '_TPES_comp_netz', chart_height + i + 1, 2, chart_height + i + 1, netz_tpes_comp_1_cols - 1],
            'fill':       {'color': netz_tpes_comp_1['item_code_new'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet13.insert_chart('B3', netz_tpes_comp_chart1)

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
        i = netz_imports_1[netz_imports_1['fuel_code'] == fuel].index[0]
        netz_imports_line.add_series({
            'name':       [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + i + 4, 0],
            'categories': [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + 3, 2, chart_height + netz_tpes_comp_1_rows + 3, netz_imports_1_cols - 1],
            'values':     [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + i + 4, 2, chart_height + netz_tpes_comp_1_rows + i + 4, netz_imports_1_cols - 1],
            'line':       {'color': netz_imports_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    netz_worksheet13.insert_chart('J3', netz_imports_line)

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
    for i in range(netz_imports_2_rows):
        netz_imports_column.add_series({
            'name':       [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + i + 7, 0],
            'categories': [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + 6, netz_imports_2_cols - 1],
            'values':     [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + i + 7, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + i + 7, netz_imports_2_cols - 1],
            'fill':       {'color': netz_imports_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet13.insert_chart('R3', netz_imports_column)

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
        i = netz_exports_1[netz_exports_1['fuel_code'] == fuel].index[0]
        netz_exports_line.add_series({
            'name':       [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + i + 10, 0],
            'categories': [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + 9, netz_imports_1_cols - 1],
            'values':     [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + i + 10, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + i + 10, netz_imports_1_cols - 1],
            'line':       {'color': netz_exports_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    netz_worksheet13.insert_chart('Z3', netz_exports_line)

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
    for i in range(netz_exports_2_rows):
        netz_exports_column.add_series({
            'name':       [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + i + 13, 0],
            'categories': [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + 12, netz_exports_2_cols - 1],
            'values':     [economy + '_TPES_comp_netz', chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + i + 13, 2, chart_height + netz_tpes_comp_1_rows + netz_imports_1_rows + netz_imports_2_rows + netz_exports_1_rows + i + 13, netz_exports_2_cols - 1],
            'fill':       {'color': netz_exports_2['fuel_code'].map(colours_dict).loc[i]},
            'border':     {'none': True}
        })
    
    netz_worksheet13.insert_chart('AH3', netz_exports_column)

    ###################################### TPES components II ###########################################
    
    # access the sheet for production created above
    # netz_worksheet14 = writer.sheets[economy + '_TPES_bunkers']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet14.set_column(2, netz_bunkers_1_cols + 1, None, space_format)
    # netz_worksheet14.set_row(chart_height, None, header_format)
    # netz_worksheet14.set_row(chart_height + netz_bunkers_1_rows + 3, None, header_format)
    # netz_worksheet14.write(0, 0, economy + ' TPES components II net-zero', cell_format1)
    
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
    for i in range(netz_bunkers_1_rows):
        netz_marine_line.add_series({
            'name':       [economy + '_TPES_bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + i + 7, 0],
            'categories': [economy + '_TPES_bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, 2,\
                (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + 6, netz_bunkers_1_cols - 1],
            'values':     [economy + '_TPES_bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + i + 7, 2,\
                (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + i + 7, netz_bunkers_1_cols - 1],
            'line':       {'color': netz_bunkers_1['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    ref_worksheet14.insert_chart('B31', netz_marine_line)

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
    for i in range(netz_bunkers_2_rows):
        netz_aviation_line.add_series({
            'name':       [economy + '_TPES_bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + i + 10, 0],
            'categories': [economy + '_TPES_bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9, 2,\
                (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + 9, netz_bunkers_2_cols - 1],
            'values':     [economy + '_TPES_bunkers', (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + i + 10, 2,\
                (2 * chart_height) + ref_bunkers_1_rows + ref_bunkers_2_rows + netz_bunkers_1_rows + i + 10, netz_bunkers_2_cols - 1],
            'line':       {'color': netz_bunkers_2['fuel_code'].map(colours_dict).loc[i], 
                           'width': 1.25},
        })    
        
    ref_worksheet14.insert_chart('J31', netz_aviation_line)

    #########################################################################################################################

    # TRANSFORMATION CHARTS

    # Access the workbook and first sheet with data from df1 
    ref_worksheet21 = writer.sheets[economy + '_pow_input']
    
    # Comma format and header format        
    # space_format = workbook.add_format({'num_format': '#,##0'})
    # header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    # cell_format1 = workbook.add_format({'bold': True})
        
    # Apply comma format and header format to relevant data rows
    ref_worksheet21.set_column(2, ref_pow_use_2_cols + 1, None, space_format)
    ref_worksheet21.set_row(chart_height, None, header_format)
    ref_worksheet21.set_row(chart_height + ref_pow_use_2_rows + 3, None, header_format)
    ref_worksheet21.set_row((2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, None, header_format)
    ref_worksheet21.set_row((2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9, None, header_format)
    ref_worksheet21.write(0, 0, economy + ' power input fuel reference (NOTE: THIS IS NOT ELECTRICITY GENERATION)', cell_format1)
    ref_worksheet21.write(48, 0, economy + ' power input fuel net_zero (NOTE: THIS IS NOT ELECTRICITY GENERATION)', cell_format1)

    # Create a use by fuel area chart
    if ref_pow_use_2_rows > 0:
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
        for i in range(ref_pow_use_2_rows):
            usefuel_chart1.add_series({
                'name':       [economy + '_pow_input', chart_height + i + 1, 0],
                'categories': [economy + '_pow_input', chart_height, 2, chart_height, ref_pow_use_2_cols - 1],
                'values':     [economy + '_pow_input', chart_height + i + 1, 2, chart_height + i + 1, ref_pow_use_2_cols - 1],
                'fill':       {'color': ref_pow_use_2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet21.insert_chart('B3', usefuel_chart1)

    else:
        pass

    # Create a use column chart
    if ref_pow_use_3_rows > 0:
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
        for i in range(ref_pow_use_3_rows):
            usefuel_chart2.add_series({
                'name':       [economy + '_pow_input', chart_height + ref_pow_use_2_rows + i + 4, 0],
                'categories': [economy + '_pow_input', chart_height + ref_pow_use_2_rows + 3, 2, chart_height + ref_pow_use_2_rows + 3, ref_pow_use_3_cols - 1],
                'values':     [economy + '_pow_input', chart_height + ref_pow_use_2_rows + i + 4, 2, chart_height + ref_pow_use_2_rows + i + 4, ref_pow_use_3_cols - 1],
                'fill':       {'color': ref_pow_use_3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        ref_worksheet21.insert_chart('J3', usefuel_chart2)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    ref_worksheet22 = writer.sheets[economy + '_elec_gen']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet22.set_column(2, ref_elecgen_2_cols + 1, None, space_format)
    ref_worksheet22.set_row(chart_height, None, header_format)
    ref_worksheet22.set_row(chart_height + ref_elecgen_2_rows + 3, None, header_format)
    ref_worksheet22.set_row((2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, None, header_format)
    ref_worksheet22.set_row((2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, None, header_format)
    ref_worksheet22.write(0, 0, economy + ' electricity generation reference', cell_format1)
    ref_worksheet22.write(50, 0, economy + ' electricity generation net-zero', cell_format1)
    
    # Create a electricity production area chart
    if ref_elecgen_2_rows > 0:
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
        for i in range(ref_elecgen_2_rows):
            prodelec_bytech_chart1.add_series({
                'name':       [economy + '_elec_gen', chart_height + i + 1, 0],
                'categories': [economy + '_elec_gen', chart_height, 2, chart_height, ref_elecgen_2_cols - 1],
                'values':     [economy + '_elec_gen', chart_height + i + 1, 2, chart_height + i + 1, ref_elecgen_2_cols - 1],
                'fill':       {'color': ref_elecgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet22.insert_chart('B3', prodelec_bytech_chart1)

    else: 
        pass

    # Create a chart
    if ref_elecgen_3_rows > 0:
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
        for i in range(ref_elecgen_3_rows):
            prodelec_bytech_chart2.add_series({
                'name':       [economy + '_elec_gen', chart_height + ref_elecgen_2_rows + i + 4, 0],
                'categories': [economy + '_elec_gen', chart_height + ref_elecgen_2_rows + 3, 2, chart_height + ref_elecgen_2_rows + 3, ref_elecgen_3_cols - 1],
                'values':     [economy + '_elec_gen', chart_height + ref_elecgen_2_rows + i + 4, 2, chart_height + ref_elecgen_2_rows + i + 4, ref_elecgen_3_cols - 1],
                'fill':       {'color': ref_elecgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet22.insert_chart('J3', prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    ref_worksheet23 = writer.sheets[economy + '_refining']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet23.set_column(2, ref_refinery_1_cols + 1, None, space_format)
    ref_worksheet23.set_row(chart_height, None, header_format)
    ref_worksheet23.set_row(chart_height + ref_refinery_1_rows + 3, None, header_format)
    ref_worksheet23.set_row(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6, None, header_format)
    ref_worksheet23.set_row((2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, None, header_format)
    ref_worksheet23.set_row((2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12, None, header_format)
    ref_worksheet23.set_row((2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15, None, header_format)
    ref_worksheet23.write(0, 0, economy + ' refining reference', cell_format1)
    ref_worksheet23.write(chart_height + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, 0, economy + ' refining net-zero', cell_format1)

    # Create ainput refining line chart
    if ref_refinery_1_rows > 0:
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
        for i in range(ref_refinery_1_rows):
            refinery_chart1.add_series({
                'name':       [economy + '_refining', chart_height + i + 1, 0],
                'categories': [economy + '_refining', chart_height, 2, chart_height, ref_refinery_1_cols - 1],
                'values':     [economy + '_refining', chart_height + i + 1, 2, chart_height + i + 1, ref_refinery_1_cols - 1],
                'line':       {'color': ref_refinery_1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet23.insert_chart('B3', refinery_chart1)

    else:
        pass

    # Create an output refining line chart
    if ref_refinery_2_rows > 0:
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
        for i in range(ref_refinery_2_rows):
            refinery_chart2.add_series({
                'name':       [economy + '_refining', chart_height + ref_refinery_1_rows + i + 4, 0],
                'categories': [economy + '_refining', chart_height + ref_refinery_1_rows + 3, 2, chart_height + ref_refinery_1_rows + 3, ref_refinery_2_cols - 1],
                'values':     [economy + '_refining', chart_height + ref_refinery_1_rows + i + 4, 2, chart_height + ref_refinery_1_rows + i + 4, ref_refinery_2_cols - 1],
                'line':       {'color': ref_refinery_2['FUEL'].map(colours_dict).loc[i],
                               'width': 1}
            })    
            
        ref_worksheet23.insert_chart('J3', refinery_chart2)

    else: 
        pass

    # Create refinery output column stacked
    if ref_refinery_3_rows > 0:
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
        for i in range(ref_refinery_3_rows):
            refinery_chart3.add_series({
                'name':       [economy + '_refining', chart_height + ref_refinery_1_rows + ref_refinery_2_rows + i + 7, 0],
                'categories': [economy + '_refining', chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6, 2, chart_height + ref_refinery_1_rows + ref_refinery_2_rows + 6, ref_refinery_3_cols - 1],
                'values':     [economy + '_refining', chart_height + ref_refinery_1_rows + ref_refinery_2_rows + i + 7, 2, chart_height + ref_refinery_1_rows + ref_refinery_2_rows + i + 7, ref_refinery_3_cols - 1],
                'fill':       {'color': ref_refinery_3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet23.insert_chart('R3', refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    ref_worksheet24 = writer.sheets[economy + '_pow_cap']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet24.set_column(1, ref_powcap_1_cols + 1, None, space_format)
    ref_worksheet24.set_row(chart_height, None, header_format)
    ref_worksheet24.set_row(chart_height + ref_powcap_1_rows + 3, None, header_format)
    ref_worksheet24.set_row((2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, None, header_format)
    ref_worksheet24.set_row((2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, None, header_format)
    ref_worksheet24.write(0, 0, economy + ' power capacity reference', cell_format1)
    ref_worksheet24.write(chart_height + ref_powcap_1_rows + ref_powcap_2_rows + 6, 0, economy + ' power capacity net-zero', cell_format1)
    
    # Create a electricity production area chart
    if ref_powcap_1_rows > 0:
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
        for i in range(ref_powcap_1_rows):
            pow_cap_chart1.add_series({
                'name':       [economy + '_pow_cap', chart_height + i + 1, 0],
                'categories': [economy + '_pow_cap', chart_height, 1, chart_height, ref_powcap_1_cols - 1],
                'values':     [economy + '_pow_cap', chart_height + i + 1, 1, chart_height + i + 1, ref_powcap_1_cols - 1],
                'fill':       {'color': ref_powcap_1['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet24.insert_chart('B3', pow_cap_chart1)

    else:
        pass

    # Create a industry subsector FED chart
    if ref_powcap_2_rows > 0:
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
        for i in range(ref_powcap_2_rows):
            pow_cap_chart2.add_series({
                'name':       [economy + '_pow_cap', chart_height + ref_powcap_1_rows + i + 4, 0],
                'categories': [economy + '_pow_cap', chart_height + ref_powcap_1_rows + 3, 1, chart_height + ref_powcap_1_rows + 3, ref_powcap_2_cols - 1],
                'values':     [economy + '_pow_cap', chart_height + ref_powcap_1_rows + i + 4, 1, chart_height + ref_powcap_1_rows + i + 4, ref_powcap_2_cols - 1],
                'fill':       {'color': ref_powcap_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet24.insert_chart('J3', pow_cap_chart2)

    else:
        pass

    ############################# Next sheet: Transformation sector ##################################
    
    # Access the workbook and second sheet
    ref_worksheet25 = writer.sheets[economy + '_trnsfrm']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet25.set_column(1, ref_trans_3_cols + 1, None, space_format)
    ref_worksheet25.set_row(chart_height, None, header_format)
    ref_worksheet25.set_row(chart_height + ref_trans_3_rows + 3, None, header_format)
    ref_worksheet25.set_row((2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, None, header_format)
    ref_worksheet25.set_row((2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + ref_trans_3_rows + 9, None, header_format)
    ref_worksheet25.write(0, 0, economy + ' transformation reference', cell_format1)
    ref_worksheet25.write(chart_height + ref_trans_3_rows + ref_trans_4_rows + 6, 0, economy + ' transformation net-zero', cell_format1)

    # Create a transformation area chart
    if ref_trans_3_rows > 0:
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
        for i in range(ref_trans_3_rows):
            ref_trnsfrm_chart1.add_series({
                'name':       [economy + '_trnsfrm', chart_height + i + 1, 0],
                'categories': [economy + '_trnsfrm', chart_height, 1, chart_height, ref_trans_3_cols - 1],
                'values':     [economy + '_trnsfrm', chart_height + i + 1, 1, chart_height + i + 1, ref_trans_3_cols - 1],
                'fill':       {'color': ref_trans_3['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet25.insert_chart('B3', ref_trnsfrm_chart1)

    else:
        pass

    # Create a transformation line chart
    if ref_trans_3_rows > 0:
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
        for i in range(ref_trans_3_rows):
            ref_trnsfrm_chart2.add_series({
                'name':       [economy + '_trnsfrm', chart_height + i + 1, 0],
                'categories': [economy + '_trnsfrm', chart_height, 1, chart_height, ref_trans_3_cols - 1],
                'values':     [economy + '_trnsfrm', chart_height + i + 1, 1, chart_height + i + 1, ref_trans_3_cols - 1],
                'line':       {'color': ref_trans_3['Sector'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet25.insert_chart('J3', ref_trnsfrm_chart2)

    else:
        pass

    # Transformation column

    if ref_trans_4_rows > 0:
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
        for i in range(ref_trans_4_rows):
            ref_trnsfrm_chart3.add_series({
                'name':       [economy + '_trnsfrm', chart_height + ref_trans_3_rows + i + 4, 0],
                'categories': [economy + '_trnsfrm', chart_height + ref_trans_3_rows + 3, 1, chart_height + ref_trans_3_rows + 3, ref_trans_4_cols - 1],
                'values':     [economy + '_trnsfrm', chart_height + ref_trans_3_rows + i + 4, 1, chart_height + ref_trans_3_rows + i + 4, ref_trans_4_cols - 1],
                'fill':       {'color': ref_trans_4['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet25.insert_chart('R3', ref_trnsfrm_chart3)

    else:
        pass

    ###############################################################################
    # Own use charts
    
    # Access the workbook and second sheet
    ref_worksheet26 = writer.sheets[economy + '_ownuse']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet26.set_column(2, ref_ownuse_1_cols + 1, None, space_format)
    ref_worksheet26.set_row(chart_height, None, header_format)
    ref_worksheet26.set_row(chart_height + ref_ownuse_1_rows + 3, None, header_format)
    ref_worksheet26.set_row((2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, None, header_format)
    ref_worksheet26.set_row((2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9, None, header_format)
    ref_worksheet26.write(0, 0, economy + ' own use and losses reference', cell_format1)
    ref_worksheet26.write(38, 0, economy + ' own use and losses net-zero', cell_format1)

    # Createn own-use transformation area chart by fuel
    if ref_ownuse_1_rows > 0:
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
        for i in range(ref_ownuse_1_rows):
            ref_own_chart1.add_series({
                'name':       [economy + '_ownuse', chart_height + i + 1, 0],
                'categories': [economy + '_ownuse', chart_height, 2, chart_height, ref_ownuse_1_cols - 1],
                'values':     [economy + '_ownuse', chart_height + i + 1, 2, chart_height + i + 1, ref_ownuse_1_cols - 1],
                'fill':       {'color': ref_ownuse_1['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet26.insert_chart('B3', ref_own_chart1)

    else:
        pass

    # Createn own-use transformation area chart by fuel
    if ref_ownuse_1_rows > 0:
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
        for i in range(ref_ownuse_1_rows):
            ref_own_chart2.add_series({
                'name':       [economy + '_ownuse', chart_height + i + 1, 0],
                'categories': [economy + '_ownuse', chart_height, 2, chart_height, ref_ownuse_1_cols - 1],
                'values':     [economy + '_ownuse', chart_height + i + 1, 2, chart_height + i + 1, ref_ownuse_1_cols - 1],
                'line':       {'color': ref_ownuse_1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet26.insert_chart('J3', ref_own_chart2)

    else:
        pass

    # Transformation column

    if ref_ownuse_2_rows > 0:
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
        for i in range(ref_ownuse_2_rows):
            ref_own_chart3.add_series({
                'name':       [economy + '_ownuse', chart_height + ref_ownuse_1_rows + i + 4, 0],
                'categories': [economy + '_ownuse', chart_height + ref_ownuse_1_rows + 3, 2, chart_height + ref_ownuse_1_rows + 3, ref_ownuse_2_cols - 1],
                'values':     [economy + '_ownuse', chart_height + ref_ownuse_1_rows + i + 4, 2, chart_height + ref_ownuse_1_rows + i + 4, ref_ownuse_2_cols - 1],
                'fill':       {'color': ref_ownuse_2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet26.insert_chart('R3', ref_own_chart3)

    else:
        pass

    ############## HEAT Charts #########################################

    # Access the workbook and second sheet
    ref_worksheet27 = writer.sheets[economy + '_heat_gen']
    
    # Apply comma format and header format to relevant data rows
    ref_worksheet27.set_column(2, ref_heatgen_2_cols + 1, None, space_format)
    ref_worksheet27.set_row(chart_height, None, header_format)
    ref_worksheet27.set_row(chart_height + ref_heatgen_2_rows + 3, None, header_format)
    ref_worksheet27.set_row((2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, None, header_format)
    ref_worksheet27.set_row((2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, None, header_format)
    ref_worksheet27.write(0, 0, economy + ' heat generation reference', cell_format1)
    ref_worksheet27.write(40, 0, economy + ' heat generation net-zero', cell_format1)
    
    # Create a electricity production area chart
    if ref_heatgen_2_rows > 0:
        heatgen_bytech_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        heatgen_bytech_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        heatgen_bytech_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        heatgen_bytech_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart1.set_y_axis({
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
            
        heatgen_bytech_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        heatgen_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_heatgen_2_rows):
            heatgen_bytech_chart1.add_series({
                'name':       [economy + '_heat_gen', chart_height + i + 1, 0],
                'categories': [economy + '_heat_gen', chart_height, 2, chart_height, ref_heatgen_2_cols - 1],
                'values':     [economy + '_heat_gen', chart_height + i + 1, 2, chart_height + i + 1, ref_heatgen_2_cols - 1],
                'fill':       {'color': ref_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet27.insert_chart('B3', heatgen_bytech_chart1)

    else: 
        pass

    # Create a chart
    if ref_heatgen_3_rows > 0:
        heatgen_bytech_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        heatgen_bytech_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        heatgen_bytech_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        heatgen_bytech_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart2.set_y_axis({
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
            
        heatgen_bytech_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        heatgen_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(ref_heatgen_3_rows):
            heatgen_bytech_chart2.add_series({
                'name':       [economy + '_heat_gen', chart_height + ref_heatgen_2_rows + i + 4, 0],
                'categories': [economy + '_heat_gen', chart_height + ref_heatgen_2_rows + 3, 2, chart_height + ref_heatgen_2_rows + 3, ref_heatgen_3_cols - 1],
                'values':     [economy + '_heat_gen', chart_height + ref_heatgen_2_rows + i + 4, 2, chart_height + ref_heatgen_2_rows + i + 4, ref_heatgen_3_cols - 1],
                'fill':       {'color': ref_heatgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet27.insert_chart('J3', heatgen_bytech_chart2)
    
    else:
        pass

    ########################################################################################################################################

    ################# NET ZERO CHARTS ######################################################################################################

    # Access the workbook and first sheet with data from df1
    # netz_worksheet21 = writer.sheets[economy + '_pow_input']
    
    # # Comma format and header format        
    # # space_format = workbook.add_format({'num_format': '#,##0'})
    # # header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    # # cell_format1 = workbook.add_format({'bold': True})
        
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet21.set_column(2, netz_pow_use_2_cols + 1, None, space_format)
    # netz_worksheet21.set_row(chart_height, None, header_format)
    # netz_worksheet21.set_row(chart_height + netz_pow_use_2_rows + 3, None, header_format)
    # netz_worksheet21.write(0, 0, economy + ' transformation use fuel', cell_format1)

    # Create a use by fuel area chart
    if netz_pow_use_2_rows > 0:
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
        for i in range(netz_pow_use_2_rows):
            netz_usefuel_chart1.add_series({
                'name':       [economy + '_pow_input', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, 0],
                'categories': [economy + '_pow_input', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, 2,\
                    (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + 6, netz_pow_use_2_cols - 1],
                'values':     [economy + '_pow_input', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, 2,\
                    (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + i + 7, netz_pow_use_2_cols - 1],
                'fill':       {'color': netz_pow_use_2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet21.insert_chart('B51', netz_usefuel_chart1)

    else:
        pass

    # Create a use column chart
    if netz_pow_use_3_rows > 0:
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
        for i in range(netz_pow_use_3_rows):
            netz_usefuel_chart2.add_series({
                'name':       [economy + '_pow_input', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + i + 10, 0],
                'categories': [economy + '_pow_input', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9, 2,\
                    (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + 9, netz_pow_use_3_cols - 1],
                'values':     [economy + '_pow_input', (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + i + 10, 2,\
                    (2 * chart_height) + ref_pow_use_2_rows + ref_pow_use_3_rows + netz_pow_use_2_rows + i + 10, netz_pow_use_3_cols - 1],
                'fill':       {'color': netz_pow_use_3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })

        ref_worksheet21.insert_chart('J51', netz_usefuel_chart2)

    else:
        pass

    ############################# Next sheet: Production of electricity by technology ##################################
    
    # Access the workbook and second sheet
    # netz_worksheet22 = writer.sheets[economy + '_elec_gen']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet22.set_column(2, netz_elecgen_2_cols + 1, None, space_format)
    # netz_worksheet22.set_row(chart_height, None, header_format)
    # netz_worksheet22.set_row(chart_height + netz_elecgen_2_rows + 3, None, header_format)
    # netz_worksheet22.write(0, 0, economy + ' electricity generation by technology', cell_format1)
    
    # Create a electricity production area chart
    if netz_elecgen_2_rows > 0:
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
        for i in range(netz_elecgen_2_rows):
            netz_prodelec_bytech_chart1.add_series({
                'name':       [economy + '_elec_gen', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, 0],
                'categories': [economy + '_elec_gen', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, 2,\
                    (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + 6, netz_elecgen_2_cols - 1],
                'values':     [economy + '_elec_gen', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, 2,\
                    (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + i + 7, netz_elecgen_2_cols - 1],
                'fill':       {'color': netz_elecgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet22.insert_chart('B53', netz_prodelec_bytech_chart1)

    else: 
        pass

    # Create a industry subsector FED chart
    if netz_elecgen_3_rows > 0:
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
        for i in range(netz_elecgen_3_rows):
            netz_prodelec_bytech_chart2.add_series({
                'name':       [economy + '_elec_gen', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, 0],
                'categories': [economy + '_elec_gen', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, 2,\
                    (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + 9, netz_elecgen_3_cols - 1],
                'values':     [economy + '_elec_gen', (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, 2,\
                    (2 * chart_height) + ref_elecgen_2_rows + ref_elecgen_3_rows + netz_elecgen_2_rows + i + 10, netz_elecgen_3_cols - 1],
                'fill':       {'color': netz_elecgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet22.insert_chart('J53', netz_prodelec_bytech_chart2)
    
    else:
        pass

    #################################################################################################################################################

    ## Refining sheet

    # Access the workbook and second sheet
    # netz_worksheet23 = writer.sheets[economy + '_refining']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet23.set_column(2, netz_refinery_1_cols + 1, None, space_format)
    # netz_worksheet23.set_row(chart_height, None, header_format)
    # netz_worksheet23.set_row(chart_height + netz_refinery_1_rows + 3, None, header_format)
    # netz_worksheet23.set_row(chart_height + netz_refinery_1_rows + netz_refinery_2_rows + 6, None, header_format)
    # netz_worksheet23.write(0, 0, economy + ' refining', cell_format1)

    # Create ainput refining line chart
    if netz_refinery_1_rows > 0:
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
        for i in range(netz_refinery_1_rows):
            netz_refinery_chart1.add_series({
                'name':       [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + i + 10, 0],
                'categories': [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, 2,\
                    (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + 9, netz_refinery_1_cols - 1],
                'values':     [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + i + 10, 2,\
                    (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + i + 10, netz_refinery_1_cols - 1],
                'line':       {'color': netz_refinery_1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet23.insert_chart('B54', netz_refinery_chart1)

    else:
        pass

    # Create an output refining line chart
    if netz_refinery_2_rows > 0:
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
        for i in range(netz_refinery_2_rows):
            netz_refinery_chart2.add_series({
                'name':       [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + i + 13, 0],
                'categories': [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12, 2,\
                    (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + 12, netz_refinery_2_cols - 1],
                'values':     [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + i + 13, 2,\
                    (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + i + 13, netz_refinery_2_cols - 1],
                'line':       {'color': netz_refinery_2['FUEL'].map(colours_dict).loc[i],
                               'width': 1}
            })    
            
        ref_worksheet23.insert_chart('J54', netz_refinery_chart2)

    else: 
        pass

    # Create refinery output column stacked
    if netz_refinery_3_rows > 0:
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
        for i in range(netz_refinery_3_rows):
            netz_refinery_chart3.add_series({
                'name':       [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + i + 16, 0],
                'categories': [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15, 2,\
                    (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + 15, netz_refinery_3_cols - 1],
                'values':     [economy + '_refining', (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + i + 16, 2,\
                    (2 * chart_height) + ref_refinery_1_rows + ref_refinery_2_rows + ref_refinery_3_rows + netz_refinery_1_rows + netz_refinery_2_rows + i + 16, netz_refinery_3_cols - 1],
                'fill':       {'color': netz_refinery_3['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet23.insert_chart('R54', netz_refinery_chart3)

    else:
        pass

    ############################# Next sheet: Power capacity ##################################
    
    # Access the workbook and second sheet
    # netz_worksheet24 = writer.sheets[economy + '_pow_cap']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet24.set_column(1, netz_powcap_1_cols + 1, None, space_format)
    # netz_worksheet24.set_row(chart_height, None, header_format)
    # netz_worksheet24.set_row(chart_height + netz_powcap_1_rows + 3, None, header_format)
    # netz_worksheet24.write(0, 0, economy + ' electricity capacity by technology', cell_format1)
    
    # Create a electricity production area chart
    if netz_powcap_1_rows > 0:
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
        for i in range(netz_powcap_1_rows):
            netz_pow_cap_chart1.add_series({
                'name':       [economy + '_pow_cap', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, 0],
                'categories': [economy + '_pow_cap', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, 1,\
                    (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + 6, netz_powcap_1_cols - 1],
                'values':     [economy + '_pow_cap', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, 1,\
                    (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + i + 7, netz_powcap_1_cols - 1],
                'fill':       {'color': netz_powcap_1['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet24.insert_chart('B45', netz_pow_cap_chart1)

    else:
        pass

    # Create a industry subsector FED chart
    if netz_powcap_2_rows > 0:
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
        for i in range(netz_powcap_2_rows):
            netz_pow_cap_chart2.add_series({
                'name':       [economy + '_pow_cap', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, 0],
                'categories': [economy + '_pow_cap', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, 1,\
                    (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + 9, netz_powcap_2_cols - 1],
                'values':     [economy + '_pow_cap', (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, 1,\
                    (2 * chart_height) + ref_powcap_1_rows + ref_powcap_2_rows + netz_powcap_1_rows + i + 10, netz_powcap_2_cols - 1],
                'fill':       {'color': netz_powcap_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet24.insert_chart('J45', netz_pow_cap_chart2)

    else:
        pass

    ############################# Next sheet: Transformation sector ##################################
    
    # Access the workbook and second sheet
    # netz_worksheet25 = writer.sheets[economy + '_trnsfrm']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet25.set_column(1, netz_trans_3_cols + 1, None, space_format)
    # netz_worksheet25.set_row(chart_height, None, header_format)
    # netz_worksheet25.set_row(chart_height + netz_trans_3_rows + 3, None, header_format)
    # netz_worksheet25.write(0, 0, economy + ' transformation', cell_format1)

    # Create a transformation area chart
    if netz_trans_3_rows > 0:
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
        for i in range(netz_trans_3_rows):
            netz_trnsfrm_chart1.add_series({
                'name':       [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 0],
                'categories': [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, netz_trans_3_cols - 1],
                'values':     [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, netz_trans_3_cols - 1],
                'fill':       {'color': netz_trans_3['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet25.insert_chart('B31', netz_trnsfrm_chart1)

    else:
        pass

    # Create a transformation line chart
    if netz_trans_3_rows > 0:
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
        for i in range(netz_trans_3_rows):
            netz_trnsfrm_chart2.add_series({
                'name':       [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 0],
                'categories': [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + 6, netz_trans_3_cols - 1],
                'values':     [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + i + 7, netz_trans_3_cols - 1],
                'line':       {'color': netz_trans_3['Sector'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet25.insert_chart('J31', netz_trnsfrm_chart2)

    else:
        pass

    # Transformation column

    if netz_trans_4_rows > 0:
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
        for i in range(netz_trans_4_rows):
            netz_trnsfrm_chart3.add_series({
                'name':       [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + i + 10, 0],
                'categories': [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + 9, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + 9, netz_trans_4_cols - 1],
                'values':     [economy + '_trnsfrm', (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + i + 10, 1,\
                    (2 * chart_height) + ref_trans_3_rows + ref_trans_4_rows + netz_trans_3_rows + i + 10, netz_trans_4_cols - 1],
                'fill':       {'color': netz_trans_4['Sector'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet25.insert_chart('R31', netz_trnsfrm_chart3)

    else:
        pass

    ###############################################################################
    # Own use charts
    
    # Access the workbook and second sheet
    # netz_worksheet26 = writer.sheets[economy + '_ownuse']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet26.set_column(2, netz_ownuse_1_cols + 1, None, space_format)
    # netz_worksheet26.set_row(chart_height, None, header_format)
    # netz_worksheet26.set_row(chart_height + netz_ownuse_1_rows + 3, None, header_format)
    # netz_worksheet26.write(0, 0, economy + ' own use and losses', cell_format1)

    # Createn own-use transformation area chart by fuel
    if netz_ownuse_1_rows > 0:
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
        for i in range(netz_ownuse_1_rows):
            netz_own_chart1.add_series({
                'name':       [economy + '_ownuse', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + i + 7, 0],
                'categories': [economy + '_ownuse', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, 2,\
                    (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + 6, netz_ownuse_1_cols - 1],
                'values':     [economy + '_ownuse', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + i + 7, 2,\
                    (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + i + 7, netz_ownuse_1_cols - 1],
                'fill':       {'color': netz_ownuse_1['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet26.insert_chart('B41', netz_own_chart1)

    else:
        pass

    # Createn own-use transformation area chart by fuel
    if netz_ownuse_1_rows > 0:
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
        for i in range(netz_ownuse_1_rows):
            netz_own_chart2.add_series({
                'name':       [economy + '_ownuse', chart_height + i + 1, 0],
                'categories': [economy + '_ownuse', chart_height, 2, chart_height, netz_ownuse_1_cols - 1],
                'values':     [economy + '_ownuse', chart_height + i + 1, 2, chart_height + i + 1, netz_ownuse_1_cols - 1],
                'line':       {'color': netz_ownuse_1['FUEL'].map(colours_dict).loc[i],
                               'width': 1.25}
            })    
            
        ref_worksheet26.insert_chart('J41', netz_own_chart2)

    else:
        pass

    # Transformation column

    if netz_ownuse_2_rows > 0:
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
        for i in range(netz_ownuse_2_rows):
            netz_own_chart3.add_series({
                'name':       [economy + '_ownuse', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + i + 10, 0],
                'categories': [economy + '_ownuse', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9, 2,\
                    (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + 9, netz_ownuse_2_cols - 1],
                'values':     [economy + '_ownuse', (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + i + 10, 2,\
                    (2 * chart_height) + ref_ownuse_1_rows + ref_ownuse_2_rows + netz_ownuse_1_rows + i + 10, netz_ownuse_2_cols - 1],
                'fill':       {'color': netz_ownuse_2['FUEL'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet26.insert_chart('R41', netz_own_chart3)

    else:
        pass

    ##################### HEAT Charts ###########################################

    # Access the workbook and second sheet
    # netz_worksheet27 = writer.sheets[economy + '_heat_gen']
    
    # # Apply comma format and header format to relevant data rows
    # netz_worksheet27.set_column(2, netz_heatgen_2_cols + 1, None, space_format)
    # netz_worksheet27.set_row(chart_height, None, header_format)
    # netz_worksheet27.set_row(chart_height + netz_heatgen_2_rows + 3, None, header_format)
    # netz_worksheet27.write(0, 0, economy + ' heat generation by technology', cell_format1)
    
    # Create a electricity production area chart
    if netz_heatgen_2_rows > 0:
        heatgen_bytech_chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        heatgen_bytech_chart1.set_size({
            'width': 500,
            'height': 300
        })
        
        heatgen_bytech_chart1.set_chartarea({
            'border': {'none': True}
        })
        
        heatgen_bytech_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart1.set_y_axis({
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
            
        heatgen_bytech_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        heatgen_bytech_chart1.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_heatgen_2_rows):
            heatgen_bytech_chart1.add_series({
                'name':       [economy + '_heat_gen', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, 0],
                'categories': [economy + '_heat_gen', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, 2,\
                    (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + 6, netz_heatgen_2_cols - 1],
                'values':     [economy + '_heat_gen', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, 2,\
                    (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + i + 7, netz_heatgen_2_cols - 1],
                'fill':       {'color': netz_heatgen_2['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet27.insert_chart('B43', heatgen_bytech_chart1)

    else: 
        pass

    # Create a chart
    if netz_heatgen_3_rows > 0:
        heatgen_bytech_chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
        heatgen_bytech_chart2.set_size({
            'width': 500,
            'height': 300
        })
        
        heatgen_bytech_chart2.set_chartarea({
            'border': {'none': True}
        })
        
        heatgen_bytech_chart2.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'line': {'color': '#bebebe'}
        })
            
        heatgen_bytech_chart2.set_y_axis({
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
            
        heatgen_bytech_chart2.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
            
        heatgen_bytech_chart2.set_title({
            'none': True
        })
        
        # Configure the series of the chart from the dataframe data.
        for i in range(netz_heatgen_3_rows):
            heatgen_bytech_chart2.add_series({
                'name':       [economy + '_heat_gen', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, 0],
                'categories': [economy + '_heat_gen', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, 2,\
                    (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + 9, netz_heatgen_3_cols - 1],
                'values':     [economy + '_heat_gen', (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, 2,\
                    (2 * chart_height) + ref_heatgen_2_rows + ref_heatgen_3_rows + netz_heatgen_2_rows + i + 10, netz_heatgen_3_cols - 1],
                'fill':       {'color': netz_heatgen_3['TECHNOLOGY'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })    
            
        ref_worksheet27.insert_chart('J43', heatgen_bytech_chart2)
    
    else:
        pass

    # Miscellaneous

    # Access the workbook and second sheet
    both_worksheet31 = writer.sheets[economy + '_mod_renew']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet31.set_column(2, ref_modren_4_cols + 1, None, space_format)
    both_worksheet31.set_row(chart_height, None, header_format)
    both_worksheet31.set_row(chart_height + ref_modren_4_rows + 3, None, header_format)
    both_worksheet31.set_row(chart_height + 3, None, percentage_format)
    both_worksheet31.set_row(chart_height + ref_modren_4_rows + 6, None, percentage_format)
    both_worksheet31.write(0, 0, economy + ' modern renewables', cell_format1)

    # line chart
    if ref_modren_4_rows > 0 & netz_modren_4_rows > 0:
        modren_chart1 = workbook.add_chart({'type': 'line'})
        modren_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        modren_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        modren_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
                
        modren_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Proportion',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
                
        modren_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
                
        modren_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = ref_modren_4[ref_modren_4['item_code_new'] == 'Reference'].index[0]
        modren_chart1.add_series({
            'name':       [economy + '_mod_renew', chart_height + i + 1, 1],
            'categories': [economy + '_mod_renew', chart_height, 2, chart_height, ref_modren_4_cols - 1],
            'values':     [economy + '_mod_renew', chart_height + i + 1, 2, chart_height + i + 1, ref_modren_4_cols - 1],
            'line':       {'color': ref_modren_4['item_code_new'].map(colours_dict).loc[i],
                            'width': 1.5}
        })
        j = netz_modren_4[netz_modren_4['item_code_new'] == 'Net-zero'].index[0]
        modren_chart1.add_series({
            'name':       [economy + '_mod_renew', chart_height + ref_modren_4_rows + j + 4, 1],
            'categories': [economy + '_mod_renew', chart_height + ref_modren_4_rows + 3, 2, chart_height + ref_modren_4_rows + 3, netz_modren_4_cols - 1],
            'values':     [economy + '_mod_renew', chart_height + ref_modren_4_rows + j + 4, 2, chart_height + ref_modren_4_rows + j + 4, netz_modren_4_cols - 1],
            'line':       {'color': netz_modren_4['item_code_new'].map(colours_dict).loc[j],
                            'width': 1.5}
        })    
                
        both_worksheet31.insert_chart('B3', modren_chart1)
    
    else:
        pass

    ##############################################################
    # Energy intensity chart

    # Access the workbook and second sheet
    both_worksheet33 = writer.sheets[economy + '_eintensity']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet33.set_column(2, ref_enint_3_cols + 1, None, space_format)
    both_worksheet33.set_row(chart_height, None, header_format)
    both_worksheet33.set_row(chart_height + ref_enint_3_rows + 3, None, header_format)
    both_worksheet33.write(0, 0, economy + ' energy intensity', cell_format1)

    # line chart
    if ref_enint_3_rows > 0 & netz_enint_3_rows > 0:
        enint_chart1 = workbook.add_chart({'type': 'line'})
        enint_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        enint_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        enint_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
                
        enint_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'TFEC energy intensity (2005 = 100)',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
                
        enint_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10}
            #'none': True
        })
                
        enint_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        i = ref_enint_3[ref_enint_3['Series'] == 'Reference'].index[0]
        enint_chart1.add_series({
            'name':       [economy + '_eintensity', chart_height + i + 1, 1],
            'categories': [economy + '_eintensity', chart_height, 2, chart_height, ref_enint_3_cols - 1],
            'values':     [economy + '_eintensity', chart_height + i + 1, 2, chart_height + i + 1, ref_enint_3_cols - 1],
            'line':       {'color': ref_enint_3['Series'].map(colours_dict).loc[i],
                            'width': 1.5}
        })
        j = netz_enint_3[netz_enint_3['Series'] == 'Net-zero'].index[0]
        enint_chart1.add_series({
            'name':       [economy + '_eintensity', chart_height + ref_enint_3_rows + j + 4, 1],
            'categories': [economy + '_eintensity', chart_height + ref_enint_3_rows + 3, 2, chart_height + ref_enint_3_rows + 3, netz_enint_3_cols - 1],
            'values':     [economy + '_eintensity', chart_height + ref_enint_3_rows + j + 4, 2, chart_height + ref_enint_3_rows + j + 4, netz_enint_3_cols - 1],
            'line':       {'color': netz_enint_3['Series'].map(colours_dict).loc[j],
                            'width': 1.5}
        })    
                
        both_worksheet33.insert_chart('B3', enint_chart1)

    else:
        pass

    ################################################
    # Macro charts

    # Access the workbook and second sheet
    both_worksheet32 = writer.sheets[economy + '_macro']
    
    # Apply comma format and header format to relevant data rows
    both_worksheet32.set_column(2, macro_1_cols + 1, None, space_format)
    both_worksheet32.set_row(chart_height, None, header_format)
    both_worksheet32.set_row(chart_height + 2, None, percentage_format)
    both_worksheet32.write(0, 0, economy + ' macro assumptions', cell_format1)

    # line chart
    if macro_1_rows > 0:
        GDP_chart1 = workbook.add_chart({'type': 'line'})
        GDP_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        GDP_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        GDP_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
                
        GDP_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'GDP (millions 2018 USD PPP 2018)',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
                
        GDP_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10},
            'none': True
        })
                
        GDP_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = macro_1[macro_1['Series'] == 'GDP 2018 USD PPP'].index[0]
        GDP_chart1.add_series({
            'name':       [economy + '_macro', chart_height + i + 1, 1],
            'categories': [economy + '_macro', chart_height, 2, chart_height, macro_1_cols - 1],
            'values':     [economy + '_macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
            'line':       {'color': macro_1['Series'].map(colours_dict).loc[i],
                            'width': 1.5}
        })

        both_worksheet32.insert_chart('B3', GDP_chart1)

        # column chart
        if any('GDP growth' in s for s in list(macro_1['Series'])):
            GDP_chart2 = workbook.add_chart({'type': 'column'})
            GDP_chart2.set_size({
                'width': 500,            
                'height': 300
            })
                
            GDP_chart2.set_chartarea({
                'border': {'none': True}
            })
                
            GDP_chart2.set_x_axis({
                'name': 'Year',
                'label_position': 'low',
                'major_tick_mark': 'none',
                'minor_tick_mark': 'none',
                'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
                #'position_axis': 'on_tick',
                'interval_unit': 4,
                'line': {'color': '#bebebe'}
            })
                    
            GDP_chart2.set_y_axis({
                'major_tick_mark': 'none', 
                'minor_tick_mark': 'none',
                'name': 'GDP growth',
                'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': '#bebebe'}
                },
                'line': {'color': '#bebebe'}
            })
                    
            GDP_chart2.set_legend({
                'font': {'font': 'Segoe UI', 'size': 10},
                'none': True
            })
                    
            GDP_chart2.set_title({
                'none': True
            })
                
            # Configure the series of the chart from the dataframe data.
            i = macro_1[macro_1['Series'] == 'GDP growth'].index[0]
            GDP_chart2.add_series({
                'name':       [economy + '_macro', chart_height + i + 1, 1],
                'categories': [economy + '_macro', chart_height, 2, chart_height, macro_1_cols - 1],
                'values':     [economy + '_macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
                'fill':       {'color': macro_1['Series'].map(colours_dict).loc[i]},
                'border':     {'none': True}
            })
                    
            both_worksheet32.insert_chart('J3', GDP_chart2)

        else:
            pass

        # Population line chart
        pop_chart1 = workbook.add_chart({'type': 'line'})
        pop_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        pop_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        pop_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
                
        pop_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'Population (millions)',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
                
        pop_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10},
            'none': True
        })
                
        pop_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = macro_1[macro_1['Series'] == 'Population'].index[0]
        pop_chart1.add_series({
            'name':       [economy + '_macro', chart_height + i + 1, 1],
            'categories': [economy + '_macro', chart_height, 2, chart_height, macro_1_cols - 1],
            'values':     [economy + '_macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
            'line':       {'color': macro_1['Series'].map(colours_dict).loc[i],
                            'width': 1.5}
        })

        both_worksheet32.insert_chart('R3', pop_chart1)  

        # GDP pc line chart
        GDPpc_chart1 = workbook.add_chart({'type': 'line'})
        GDPpc_chart1.set_size({
            'width': 500,            
            'height': 300
        })
            
        GDPpc_chart1.set_chartarea({
            'border': {'none': True}
        })
            
        GDPpc_chart1.set_x_axis({
            'name': 'Year',
            'label_position': 'low',
            'major_tick_mark': 'none',
            'minor_tick_mark': 'none',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232', 'rotation': -45},
            'position_axis': 'on_tick',
            'interval_unit': 4,
            'line': {'color': '#bebebe'}
        })
                
        GDPpc_chart1.set_y_axis({
            'major_tick_mark': 'none', 
            'minor_tick_mark': 'none',
            'name': 'GDP per capita (2018 USD PPP)',
            'num_font': {'font': 'Segoe UI', 'size': 10, 'color': '#323232'},
            'major_gridlines': {
                'visible': True,
                'line': {'color': '#bebebe'}
            },
            'line': {'color': '#bebebe'}
        })
                
        GDPpc_chart1.set_legend({
            'font': {'font': 'Segoe UI', 'size': 10},
            'none': True
        })
                
        GDPpc_chart1.set_title({
            'none': True
        })
            
        # Configure the series of the chart from the dataframe data.
        # for component in ['Reference proportion', 'Net-zero proportion']:
        i = macro_1[macro_1['Series'] == 'GDP per capita'].index[0]
        GDPpc_chart1.add_series({
            'name':       [economy + '_macro', chart_height + i + 1, 1],
            'categories': [economy + '_macro', chart_height, 2, chart_height, macro_1_cols - 1],
            'values':     [economy + '_macro', chart_height + i + 1, 2, chart_height + i + 1, macro_1_cols - 1],
            'line':       {'color': macro_1['Series'].map(colours_dict).loc[i],
                            'width': 1.5}
        })

        both_worksheet32.insert_chart('Z3', GDPpc_chart1)

    else:
        pass   

    writer.save()

print('Bling blang blaow, you have some charts now')

