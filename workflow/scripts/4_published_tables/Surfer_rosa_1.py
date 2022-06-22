# Please run /1_historical_to_projected/OSeMOSYS_to_EGEDA_2018_actuals.py
# All portions of FED, Supply and Transformation tables in one script

# import dependencies

from numpy.core.numeric import NaN
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

# Gas trade
ref_gastrade_df1 = pd.read_csv('./data/4_Joined/lngpipe_reference.csv').loc[:,:'2050']
netz_gastrade_df1 = pd.read_csv('./data/4_Joined/lngpipe_netzero.csv').loc[:,:'2050']

# Captured emissions dataframes

ref_capemiss_df1 = pd.read_csv('./data/4_Joined/captured_ref.csv').loc[:,:'2050']
netz_capemiss_df1 = pd.read_csv('./data/4_Joined/captured_cn.csv').loc[:,:'2050']

# Emissions dataframe 

EGEDA_emissions_reference = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_emissions_2018_reference.csv')
EGEDA_emissions_netzero = pd.read_csv('./data/4_Joined/OSeMOSYS_to_EGEDA_emissions_2018_netzero.csv')

# OSeMOSYS only

ref_osemo_1 = pd.read_csv('./data/4_Joined/OSeMOSYS_only_reference.csv').loc[:,:'2050']
netz_osemo_1 = pd.read_csv('./data/4_Joined/OSeMOSYS_only_netzero.csv').loc[:,:'2050']

# Define month and year to create folder for saving charts/tables

day_month_year = pd.to_datetime('today').strftime('%Y-%m-%d-%H%M')

# Macro

macro_GDP = pd.read_excel('./data/2_Mapping_and_other/macro_APEC.xlsx', sheet_name = 'GDP')
macro_GDP.columns = macro_GDP.columns.astype(str) 
macro_GDP['Series'] = 'GDP 2018 USD PPP'
macro_GDP = macro_GDP[['Economy', 'Series'] + list(macro_GDP.loc[:, '2000':'2050'])]

# Change GDP to millions
GDP = macro_GDP.select_dtypes(include=[np.number]) / 1000000000 
macro_GDP[GDP.columns] = GDP

macro_GDP_growth = pd.read_excel('./data/2_Mapping_and_other/macro_APEC.xlsx', sheet_name = 'GDP_growth')
macro_GDP_growth.columns = macro_GDP_growth.columns.astype(str) 
macro_GDP_growth['Series'] = 'GDP growth'
macro_GDP_growth = macro_GDP_growth[['Economy', 'Series'] + list(macro_GDP_growth.loc[:, '2000':'2050'])]

macro_pop = pd.read_excel('./data/2_Mapping_and_other/macro_APEC.xlsx', sheet_name = 'Population')
macro_pop.columns = macro_pop.columns.astype(str) 
macro_pop['Series'] = 'Population'
macro_pop = macro_pop[['Economy', 'Series'] + list(macro_pop.loc[:, '2000':'2050'])]

# Change population to millions
pop = macro_pop.select_dtypes(include=[np.number]) / 1000 
macro_pop[pop.columns] = pop

macro_GDPpc = pd.read_excel('./data/2_Mapping_and_other/macro_APEC.xlsx', sheet_name = 'GDP per capita')
macro_GDPpc.columns = macro_GDPpc.columns.astype(str)
macro_GDPpc['Series'] = 'GDP per capita' 
macro_GDPpc = macro_GDPpc[['Economy', 'Series'] + list(macro_GDPpc.loc[:, '2000':'2050'])]

# Define unique values for economy, fuels, and items columns
# only looking at one dataframe which should be sufficient as both have same structure

Economy_codes = EGEDA_years_reference.economy.unique()
Economy_codes = np.delete(Economy_codes, [23, 25])
Fuels = EGEDA_years_reference.fuel_code.unique()
Items = EGEDA_years_reference.item_code_new.unique()

# Define colour palette

colours_dict = pd.read_csv('./data/2_Mapping_and_other/colours_dict.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

APEC_economies = pd.read_csv('./data/2_Mapping_and_other/APEC_economies.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

# FED and TPES: vectors for impending df builds

# Fuelsa

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

required_power_input = ['1_coal', '1_5_lignite', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products', 
                        '8_gas', '9_nuclear', '10_hydro', '11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', 
                        '16_1_biogas', '16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable', 
                        '16_6_biodiesel', '16_8_other_liquid_biofuels', '16_9_other_sources']

required_ol_input = ['1_coal', '1_5_lignite', '2_coal_products', '3_peat', '4_peat_products', '6_crude_oil_and_ngl', 
                     '7_petroleum_products', '8_gas', '15_solid_biomass', '16_1_biogas', '16_2_industrial_waste',  
                     '16_6_biodiesel', '17_electricity', '18_heat']

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

marine_bunker_fuels = ['7_7_gas_diesel_oil', '7_8_fuel_oil', '8_1_natural_gas', '16_x_hydrogen', '16_6_biodiesel']
aviation_bunker_fuels = ['7_x_jet_fuel', '16_x_hydrogen', '16_7_bio_jet_kerosene', '7_2_aviation_gasoline']

### Transport fuel vectors

Transport_fuels = ['1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products', '7_1_motor_gasoline', '7_2_aviation_gasoline',
                   '7_x_jet_fuel', '7_7_gas_diesel_oil', '7_8_fuel_oil', '7_9_lpg',
                   '7_x_other_petroleum_products', '8_1_natural_gas', '16_5_biogasoline', '16_6_biodiesel',
                   '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels', '16_x_hydrogen', '17_electricity'] 

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

# Single fuel vectors

fuel_vector_1 = ['1_indigenous_production', '2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers',
                 '6_stock_change', '7_total_primary_energy_supply']

fuel_vector_ref = ['2_imports', '3_exports', '4_international_marine_bunkers', '5_international_aviation_bunkers',
                   '6_stock_change', '7_total_primary_energy_supply']

fuel_final_nobunk = ['Production', 'Imports', 'Exports', 'Stock change', 'Total primary energy supply']
fuel_final_bunk = ['Production', 'Imports', 'Exports', 'Bunkers', 'Stock change', 'Total primary energy supply']
fuel_final_ref = ['Domestic refining', 'Imports', 'Exports', 'Bunkers', 'Stock change', 'Total primary energy supply']

fuel_vector_3 = ['9_1_main_activity_producer', '9_2_autoproducers', '10_losses_and_own_use', '14_industry_sector',
                 '15_transport_sector', '16_1_commercial_and_public_services', '16_2_residential', '16_3_agriculture',
                 '16_4_fishing', '16_5_nonspecified_others', '17_nonenergy_use']

##################################################################################
# Emissions

# Subsets for impending emissions df builds

First_level_emiss = ['1_coal', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products',
                     '8_gas', '16_others', '17_electricity', '18_heat', '19_total']

Required_emiss = ['1_coal', '2_coal_products', '6_crude_oil_and_ngl', '7_petroleum_products',
                  '8_gas', '16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources',
                  '17_electricity', '18_heat', '19_total']

Coal_emiss = ['1_coal', '2_coal_products', '3_peat', '4_peat_products']

Oil_emiss = ['6_crude_oil_and_ngl', '7_petroleum_products']

Heat_others_emiss = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable', '16_9_other_sources', '18_heat']

# Emissions sectors (DEMANDS)

Sectors_emiss = ['9_x_power', '10_losses_and_own_use', 
               '14_industry_sector', '15_transport_sector', '16_1_commercial_and_public_services', '16_2_residential',
               '16_3_agriculture', '16_4_fishing', '16_5_nonspecified_others', '17_nonenergy_use']

Buildings_emiss = ['16_1_commercial_and_public_services', '16_2_residential']

Ag_emiss = ['16_3_agriculture', '16_4_fishing']

# FED aggregate fuels

Emissions_agg_fuels = ['Coal', 'Oil', 'Gas', 'Electricity', 'Heat & others']

Emissions_agg_sectors = ['Power', 'Own use', 'Industry', 'Transport', 'Buildings', 'Agriculture', 'Non-specified']

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
other_fuel_1 = ['16_4_municipal_solid_waste_nonrenewable', '16_x_hydrogen', '16_2_industrial_waste']

# '17_electricity', '18_heat'

imports_fuel_1 = ['17_electricity_export']

# Second aggreagtion: Oil, Gas, Nuclear, Imports, Other from above and below two new aggregations (7 fuels)
coal_fuel_2 = ['1_x_coal_thermal', '1_5_lignite', '2_coal_products']
renewables_fuel_2 = ['10_hydro', '11_geothermal', '12_1_of_which_photovoltaics', '13_tide_wave_ocean', '14_wind', '15_1_fuelwood_and_woodwaste', 
                     '15_2_bagasse', '15_4_black_liquor', '15_5_other_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable']

# For heat
waste_fuel = ['16_2_industrial_waste', '16_3_municipal_solid_waste_renewable', '16_4_municipal_solid_waste_nonrenewable']

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
waste_ou = ['16_2_industrial_waste', '16_4_municipal_solid_waste_nonrenewable']

own_use_fuels = ['Coal', 'Oil', 'Gas', 'Renewables', 'Electricity', 'Heat', 'Waste', 'Total']

# Note, 12_1_of_which_photovoltaics is a subset of 12_solar so including will lead to double counting

use_agg_fuels_1 = ['Coal', 'Lignite', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Solar', 'Wind', 
                   'Biomass', 'Geothermal', 'Other renewables', 'Other', 'Total']
use_agg_fuels_2 = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Other']

heat_agg_fuels = ['Coal', 'Lignite', 'Oil', 'Gas', 'Biomass', 'Waste', 'Total']

# TECHNOLOGY aggregations for ProductionByTechnology

coal_tech = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']
coal_ccs_tech = ['POW_COAL_CCS_PP']
oil_tech = ['POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP', 'POW_PetCoke_PP']
gas_tech = ['POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP']
gas_ccs_tech = ['POW_CCGT_CCS_PP']
nuclear_tech = ['POW_Nuclear_PP', 'POW_IMP_Nuclear_PP']
hydro_tech = ['POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP']
solar_tech = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP']
wind_tech = ['POW_WindOff_PP', 'POW_Wind_PP']
bio_tech = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP']
geo_tech = ['POW_Geothermal_PP']
storage_tech = ['POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage']
waste_tech = ['POW_WasteToEnergy_PP']
other_tech = ['POW_IPP_PP', 'POW_TIDAL_PP', 'POW_CHP_PP']
# chp_tech = ['POW_CHP_PP']
im_tech = ['POW_IMPORTS_PP', 'POW_IMPORT_ELEC_PP']

lignite_tech = ['POW_Sub_Brown_PP']
thermal_coal_tech = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']
solar_roof_tech = ['POW_SolarRoofPV_PP']
solar_nr_tech = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP']

# Another aggregation of other from Alex
other_higheragg_tech = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP', 'POW_Geothermal_PP', 
                        'POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage', 'POW_WasteToEnergy_PP',
                        'POW_IPP_PP', 'POW_TIDAL_PP', 'POW_CHP_PP']

# Modern renewables

modren_elec_heat = ['POW_Hydro_PP', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP', 'POW_SolarCSP_PP', 
                    'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP', 'POW_WindOff_PP', 'POW_Wind_PP',
                    'POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP', 'POW_Geothermal_PP', 'POW_TIDAL_PP', 
                    'POW_CHP_BIO_PP', 'POW_Solid_Biomass_PP']

all_elec_heat = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_Ultra_BituCoal_PP', 
                 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP', 'POW_COAL_CCS_PP', 'POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_FuelOil_HP', 'POW_OilProducts_PP', 'POW_PetCoke_PP',
                 'POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP', 'POW_CCGT_CCS_PP', 'POW_Nuclear_PP', 'POW_IMP_Nuclear_PP',
                 'POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP', 'POW_SolarCSP_PP', 
                 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP', 'POW_WindOff_PP', 'POW_Wind_PP', 'POW_Solid_Biomass_PP', 
                 'POW_CHP_BIO_PP', 'POW_Biogas_PP', 'POW_Geothermal_PP', 'POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage',
                 'POW_IPP_PP', 'POW_TIDAL_PP', 'POW_WasteToEnergy_PP', 'POW_WasteToHeat_HP', 'POW_CHP_PP', 'POW_HEAT_HP', 'YYY_18_heat']

# 'POW_Pumped_Hydro'?? in the above

# POW_EXPORT_ELEC_PP need to work this in

prod_agg_tech = ['Coal', 'Coal CCS', 'Oil', 'Gas', 'Gas CCS', 'Hydro', 'Nuclear', 'Wind', 'Solar', 'Bio', 
                 'Geothermal', 'Waste', 'Storage', 'Other', 'Imports', 'Total']
prod_agg_tech2 = ['Coal', 'Coal CCS', 'Lignite', 'Oil', 'Gas', 'Gas CCS', 'Hydro', 'Nuclear', 'Wind', 'Solar', 
                 'Bio', 'Geothermal', 'Waste', 'Storage', 'Other', 'Imports', 'Total']
prod_agg_tech3 = ['Coal', 'Coal CCS', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Wind', 'Solar', 'Other', 'Imports', 'Total']

heat_prod_tech = ['Coal', 'Lignite', 'Oil', 'Gas', 'Gas CCS', 'Nuclear', 'Biomass', 'Waste', 'Non-specified', 'Other', 'Total']

# Power input fuel categories

powinput_fuel = ['Coal', 'Lignite', 'Oil', 'Gas', 'Hydro', 'Nuclear', 'Wind', 'Solar', 'Biomass', 'Geothermal',
                 'Other renewables', 'Other']

# Refinery vectors

refinery_input = ['d_ref_6_1_crude_oil', 'd_ref_6_x_ngls']
refinery_output = ['d_ref_7_1_motor_gasoline_refine', 'd_ref_7_2_aviation_gasoline_refine', 'd_ref_7_3_naphtha_refine', 
                   'd_ref_7_x_jet_fuel_refine', 'd_ref_7_6_kerosene_refine', 'd_ref_7_7_gas_diesel_oil_refine', 
                   'd_ref_7_8_fuel_oil_refine', 'd_ref_7_9_lpg_refine', 'd_ref_7_10_refinery_gas_not_liquefied_refine', 
                   'd_ref_7_11_ethane_refine', 'd_ref_7_x_other_petroleum_products_refine']

refinery_new_output = ['7_1_from_ref', '7_2_from_ref', '7_3_from_ref', '7_jet_from_ref', '7_6_from_ref', '7_7_from_ref',
                       '7_8_from_ref', '7_9_from_ref', '7_10_from_ref', '7_11_from_ref', '7_other_from_ref']

# Capacity vectors
    
coal_cap = ['POW_Black_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Sub_Brown_PP', 'POW_CHP_COAL_PP', 'POW_Other_Coal_PP', 'POW_Ultra_BituCoal_PP', 'POW_Ultra_CHP_PP']
coal_ccs_cap = ['POW_COAL_CCS_PP']
gas_cap = ['POW_CCGT_PP', 'POW_OCGT_PP', 'POW_CHP_GAS_PP']
gas_ccs_cap = ['POW_CCGT_CCS_PP']
oil_cap = ['POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP', 'POW_PetCoke_PP']
nuclear_cap = ['POW_Nuclear_PP', 'POW_IMP_Nuclear_PP']
hydro_cap = ['POW_Hydro_PP', 'POW_Pumped_Hydro', 'POW_Storage_Hydro_PP', 'POW_IMP_Hydro_PP']
bio_cap = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP']
wind_cap = ['POW_Wind_PP', 'POW_WindOff_PP']
solar_cap = ['POW_SolarCSP_PP', 'POW_SolarFloatPV_PP', 'POW_SolarPV_PP', 'POW_SolarRoofPV_PP']
geo_cap = ['POW_Geothermal_PP']
storage_cap = ['POW_AggregatedEnergy_Storage_VPP', 'POW_EmbeddedBattery_Storage']
waste_cap = ['POW_WasteToEnergy_PP']
other_cap = ['POW_IPP_PP', 'POW_TIDAL_PP', 'POW_CHP_PP']
# chp_cap = ['POW_CHP_PP']
# 'POW_HEAT_HP' not in electricity capacity
transmission_cap = ['POW_Transmission']

lignite_cap = ['POW_Sub_Brown_PP']
thermal_coal_cap = ['POW_Black_Coal_PP', 'POW_Other_Coal_PP', 'POW_Sub_BituCoal_PP', 'POW_Ultra_BituCoal_PP', 'POW_CHP_COAL_PP', 'POW_Ultra_CHP_PP']

# Other cap from Alex
other_higheragg_cap = ['POW_Solid_Biomass_PP', 'POW_CHP_BIO_PP', 'POW_Biogas_PP', 'POW_Geothermal_PP', 'POW_AggregatedEnergy_Storage_VPP', 
                       'POW_EmbeddedBattery_Storage', 'POW_WasteToEnergy_PP', 'POW_IPP_PP', 'POW_TIDAL_PP', 'POW_CHP_PP']

pow_capacity_agg = ['Coal', 'Coal CCS', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Bio', 'Wind', 'Solar', 'Geothermal', 'Waste', 'Storage', 'Other']
pow_capacity_agg2 = ['Coal', 'Coal CCS', 'Lignite', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Bio', 'Wind', 
                     'Solar', 'Geothermal', 'Waste', 'Storage', 'Other']

pow_capacity_agg3 = ['Coal', 'Coal CCS', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Wind', 'Solar', 'Other', 'Total']

# Heat power plants

coal_heat = ['POW_CHP_COAL_PP', 'POW_Ultra_BituCoal_PP', 'POW_Ultra_CHP_PP', 'POW_HEAT_COKE_HP', 'POW_Sub_BituCoal_PP', 'POW_Other_Coal_PP']
lignite_heat = ['POW_Sub_Brown_PP']
gas_heat = ['POW_CCGT_PP', 'POW_CHP_GAS_PP']
gas_ccs_heat = ['POW_CCGT_CCS_PP']
oil_heat = ['POW_FuelOil_HP', 'POW_Diesel_PP', 'POW_FuelOil_PP', 'POW_OilProducts_PP']
bio_heat = ['POW_CHP_BIO_PP', 'POW_Solid_Biomass_PP', 'POW_Biogas_PP']
nuke_heat = ['POW_Nuclear_PP']
waste_heat = ['POW_WasteToEnergy_PP', 'POW_WasteToHeat_HP']
combination_heat = ['POW_HEAT_HP', 'YYY_18_heat']
nons_heat = ['POW_CHP_PP']

# Heat only power plants

heat_only = ['POW_FuelOil_HP', 'POW_HEAT_HP', 'POW_WasteToHeat_HP', 'POW_HEAT_COKE_HP', 'YYY_18_heat']

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

TPES_agg_fuels1 = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Electricity', 'Hydrogen', 'Other fuels']
TPES_agg_fuels2 = ['Coal', 'Oil', 'Gas', 'Nuclear', 'Renewables', 'Other fuels']
TPES_agg_trade = ['Coal', 'Crude oil & NGL', 'Petroleum products', 'Gas', 
                  'Renewables', 'Electricity', 'Hydrogen', 'Other fuels']
avi_bunker = ['Aviation gasoline', 'Jet fuel', 'Biojet kerosene', 'Hydrogen']

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

# Create a copy for alternative historical with different aggregations
EGEDA_hist_gen2 = EGEDA_hist_gen.copy()

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
                                                                '15_solid_biomass': 'Bio', 
                                                                '16_others': 'Other', 
                                                                '17_electricity': 'Imports',
                                                                '18_heat': 'Other'})

EGEDA_hist_gen['Generation'] = 'Electricity'

EGEDA_hist_gen = EGEDA_hist_gen[['economy', 'TECHNOLOGY', 'Generation'] + list(range(2000, 2019))].\
    groupby(['economy', 'TECHNOLOGY', 'Generation']).sum().reset_index()

EGEDA_hist_gen.to_csv('./data/4_Joined/EGEDA_hist_gen.csv', index = False)
EGEDA_hist_gen = pd.read_csv('./data/4_Joined/EGEDA_hist_gen.csv')

# Same historical generation with different aggregations
EGEDA_hist_gen2['TECHNOLOGY'] = EGEDA_hist_gen2['fuel_code'].map({'1_coal': 'Coal', 
                                                                  '1_5_lignite': 'Coal', 
                                                                  '2_coal_products': 'Coal',
                                                                  '6_crude_oil_and_ngl': 'Oil',
                                                                  '7_petroleum_products': 'Oil',
                                                                  '8_gas': 'Gas', 
                                                                  '9_nuclear': 'Nuclear', 
                                                                  '10_hydro': 'Hydro', 
                                                                  '11_geothermal': 'Other', 
                                                                  '12_solar': 'Solar', 
                                                                  '13_tide_wave_ocean': 'Hydro', 
                                                                  '14_wind': 'Wind', 
                                                                  '15_solid_biomass': 'Other', 
                                                                  '16_others': 'Other', 
                                                                  '17_electricity': 'Imports',
                                                                  '18_heat': 'Other'})

EGEDA_hist_gen2['Generation'] = 'Electricity'

EGEDA_hist_gen2 = EGEDA_hist_gen2[['economy', 'TECHNOLOGY', 'Generation'] + list(range(2000, 2019))].\
    groupby(['economy', 'TECHNOLOGY', 'Generation']).sum().reset_index()

EGEDA_hist_gen2.to_csv('./data/4_Joined/EGEDA_hist_gen2.csv', index = False)
EGEDA_hist_gen2 = pd.read_csv('./data/4_Joined/EGEDA_hist_gen2.csv')

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

#########################################################################################################################

EGEDA_hist_eh2 = EGEDA_data[(EGEDA_data['item_code_new'].isin(['18_electricity_output_in_pj', '19_heat_output_in_pj'])) &
                           (EGEDA_data['fuel_code'] == '19_total')].copy().reset_index(drop = True)

EGEDA_hist_eh2 = EGEDA_hist_eh2[['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))].\
    groupby(['economy']).sum().reset_index()

EGEDA_hist_eh2['fuel_code'] = 'Total'
EGEDA_hist_eh2['item_code_new'] = 'Electricity and heat'

EGEDA_hist_eh2 = EGEDA_hist_eh2[['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))].reset_index(drop = True)

EGEDA_hist_eh2.to_csv('./data/4_Joined/EGEDA_hist_eh2.csv', index = False)
EGEDA_hist_eh2 = pd.read_csv('./data/4_Joined/EGEDA_hist_eh2.csv')

######################################################################

# Create histrocial i) own use and losses and ii) power consumption for use later
# WORK IN PROGRESS

EGEDA_hist_power = EGEDA_data[(EGEDA_data['item_code_new'].isin(['9_1_main_activity_producer', '9_2_autoproducers'])) &
                              (EGEDA_data['fuel_code'].isin(required_power_input))].copy().reset_index(drop = True)

# China only having data for 1_coal requires workaround to keep lignite data
lignite_alt = EGEDA_hist_power[EGEDA_hist_power['fuel_code'] == '1_5_lignite'].copy()\
    .set_index(['economy', 'fuel_code', 'item_code_new']) * -1

lignite_alt = lignite_alt.reset_index()

new_coal = EGEDA_hist_power[EGEDA_hist_power['fuel_code'] == '1_coal'].copy().reset_index(drop = True)

lig_coal = new_coal.append(lignite_alt).reset_index(drop = True).groupby(['economy', 'item_code_new']).sum().reset_index()
lig_coal['fuel_code'] = '1_coal'

no_coal = EGEDA_hist_power[EGEDA_hist_power['fuel_code'] != '1_coal'].copy().reset_index(drop = True)

EGEDA_hist_power = no_coal.append(lig_coal).reset_index(drop = True)

EGEDA_hist_power['FUEL'] = EGEDA_hist_power['fuel_code'].map({'1_coal': 'Coal', 
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
                                                              '16_1_biogas': 'Other renewables',
                                                              '16_2_industrial_waste': 'Other',
                                                              '16_3_municipal_solid_waste_renewable': 'Other renewables',
                                                              '16_4_municipal_solid_waste_nonrenewable': 'Other',
                                                              '16_6_biodiesel': 'Other renewables',
                                                              '16_8_other_liquid_biofuels': 'Other renewables',
                                                              '16_9_other_sources': 'Other'})

EGEDA_hist_power['Transformation'] = 'Input fuel'

EGEDA_hist_power = EGEDA_hist_power[['economy', 'FUEL', 'Transformation'] + list(range(2000, 2019))].copy()\
    .groupby(['economy', 'FUEL', 'Transformation']).sum().reset_index()

neg_to_pos = EGEDA_hist_power.select_dtypes(include=[np.number]) * -1  
EGEDA_hist_power[neg_to_pos.columns] = neg_to_pos

EGEDA_hist_power.to_csv('./data/4_Joined/EGEDA_hist_power.csv', index = False)
EGEDA_hist_power = pd.read_csv('./data/4_Joined/EGEDA_hist_power.csv')

### Extra grab for consumption chart

EGEDA_histpower_oil = EGEDA_data[(EGEDA_data['item_code_new'].isin(['9_1_main_activity_producer', '9_2_autoproducers'])) &
                                 (EGEDA_data['fuel_code'].isin(['6_crude_oil_and_ngl', '7_petroleum_products']))]\
                                    .copy().reset_index(drop = True)

EGEDA_histpower_oil['FUEL'] = EGEDA_histpower_oil['fuel_code'].map({'6_crude_oil_and_ngl': 'Crude oil & NGL',
                                                                    '7_petroleum_products': 'Petroleum products'})

EGEDA_histpower_oil = EGEDA_histpower_oil.groupby(['economy', 'FUEL']).sum().reset_index().assign(item_code_new = 'Power')\
    [['economy', 'FUEL', 'item_code_new'] + list(range(2000, 2019))]\

neg_to_pos = EGEDA_histpower_oil.select_dtypes(include = [np.number]) * -1
EGEDA_histpower_oil[neg_to_pos.columns] = neg_to_pos

EGEDA_histpower_oil.to_csv('./data/4_Joined/EGEDA_histpower_oil.csv', index = False)
EGEDA_histpower_oil = pd.read_csv('./data/4_Joined/EGEDA_histpower_oil.csv')

### liquid and solid renewables

EGEDA_histpower_renew = EGEDA_data[(EGEDA_data['item_code_new'].isin(['9_1_main_activity_producer', '9_2_autoproducers'])) &
                                 (EGEDA_data['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))]\
                                    .copy().reset_index(drop = True)

EGEDA_histpower_renew['FUEL'] = EGEDA_histpower_renew['fuel_code'].map({'15_solid_biomass': 'Liquid and solid renewables', 
                                                                        '16_1_biogas': 'Liquid and solid renewables', 
                                                                        '16_3_municipal_solid_waste_renewable': 'Liquid and solid renewables',
                                                                        '16_5_biogasoline': 'Liquid and solid renewables', 
                                                                        '16_6_biodiesel': 'Liquid and solid renewables', 
                                                                        '16_7_bio_jet_kerosene': 'Liquid and solid renewables', 
                                                                        '16_8_other_liquid_biofuels': 'Liquid and solid renewables'})

EGEDA_histpower_renew = EGEDA_histpower_renew.groupby(['economy', 'FUEL']).sum().reset_index().assign(item_code_new = 'Power')\
    [['economy', 'FUEL', 'item_code_new'] + list(range(2000, 2019))]\

neg_to_pos = EGEDA_histpower_renew.select_dtypes(include = [np.number]) * -1
EGEDA_histpower_renew[neg_to_pos.columns] = neg_to_pos

EGEDA_histpower_renew.to_csv('./data/4_Joined/EGEDA_histpower_renew.csv', index = False)
EGEDA_histpower_renew = pd.read_csv('./data/4_Joined/EGEDA_histpower_renew.csv')


################################################################################

# Own use and losses historical

EGEDA_hist_own = EGEDA_data[(EGEDA_data['item_code_new'].isin(['10_losses_and_own_use'])) &
                              (EGEDA_data['fuel_code'].isin(required_ol_input))].copy().reset_index(drop = True)

# China only having data for 1_coal requires workaround to keep lignite data
lignite_alt = EGEDA_hist_own[EGEDA_hist_own['fuel_code'] == '1_5_lignite'].copy()\
    .set_index(['economy', 'fuel_code', 'item_code_new']) * -1

lignite_alt = lignite_alt.reset_index()

new_coal = EGEDA_hist_own[EGEDA_hist_own['fuel_code'] == '1_coal'].copy().reset_index(drop = True)

lig_coal = new_coal.append(lignite_alt).reset_index(drop = True).groupby(['economy', 'item_code_new']).sum().reset_index()
lig_coal['fuel_code'] = '1_coal'

no_coal = EGEDA_hist_own[EGEDA_hist_own['fuel_code'] != '1_coal'].copy().reset_index(drop = True)

EGEDA_hist_own = no_coal.append(lig_coal).reset_index(drop = True)

# Special grab for coal report ##########
EGEDA_hist_owncoal = EGEDA_hist_own[EGEDA_hist_own['fuel_code'].isin(['1_coal', '1_5_lignite', '2_coal_products'])].copy()

EGEDA_hist_own['FUEL'] = EGEDA_hist_own['fuel_code'].map({'1_coal': 'Coal', 
                                                          '1_5_lignite': 'Coal', 
                                                          '2_coal_products': 'Coal',
                                                          '3_peat': 'Coal',
                                                          '4_peat_products': 'Coal',
                                                          '6_crude_oil_and_ngl': 'Oil',
                                                          '7_petroleum_products': 'Oil',
                                                          '8_gas': 'Gas',  
                                                          '15_solid_biomass': 'Renewables', 
                                                          '16_1_biogas': 'Renewables',
                                                          '16_2_industrial_waste': 'Other',
                                                          '16_6_biodiesel': 'Renewables',
                                                          '17_electricity': 'Electricity',
                                                          '18_heat': 'Heat'})

EGEDA_hist_own['Sector'] = 'Own-use and losses'

EGEDA_hist_own = EGEDA_hist_own[['economy', 'FUEL', 'Sector'] + list(range(2000, 2019))].copy()\
    .groupby(['economy', 'FUEL', 'Sector']).sum().reset_index()

neg_to_pos = EGEDA_hist_own.select_dtypes(include=[np.number]) * -1  
EGEDA_hist_own[neg_to_pos.columns] = neg_to_pos

EGEDA_hist_own.to_csv('./data/4_Joined/EGEDA_hist_own.csv', index = False)
EGEDA_hist_own = pd.read_csv('./data/4_Joined/EGEDA_hist_own.csv')

# Special grab for coal report continued
EGEDA_hist_owncoal['FUEL'] = EGEDA_hist_owncoal['fuel_code'].map({'1_coal': 'Thermal coal',
                                                                  '1_5_lignite': 'Lignite',
                                                                  '2_coal_products': 'Metallurgical coal'})

EGEDA_hist_owncoal['Sector'] = 'Own-use and losses'

EGEDA_hist_owncoal = EGEDA_hist_owncoal[['economy', 'FUEL', 'Sector'] + list(range(2000, 2019))].copy()\
    .groupby(['economy', 'FUEL', 'Sector']).sum().reset_index()

neg_to_pos = EGEDA_hist_owncoal.select_dtypes(include=[np.number]) * -1  
EGEDA_hist_owncoal[neg_to_pos.columns] = neg_to_pos

EGEDA_hist_owncoal.to_csv('./data/4_Joined/EGEDA_hist_owncoal.csv', index = False)
EGEDA_hist_owncoal = pd.read_csv('./data/4_Joined/EGEDA_hist_owncoal.csv')

### Extra grab for consumption chart

EGEDA_hist_own_oil = EGEDA_data[(EGEDA_data['item_code_new'].isin(['10_losses_and_own_use'])) &
                                (EGEDA_data['fuel_code'].isin(['6_crude_oil_and_ngl', '7_petroleum_products']))]\
                                  .copy().reset_index(drop = True)

EGEDA_hist_own_oil['FUEL'] = EGEDA_hist_own_oil['fuel_code'].map({'6_crude_oil_and_ngl': 'Crude oil & NGL',
                                                                  '7_petroleum_products': 'Petroleum products'})

EGEDA_hist_own_oil = EGEDA_hist_own_oil[['economy', 'FUEL', 'item_code_new'] + list(range(2000, 2019))]\
                        .copy().reset_index(drop = True)

neg_to_pos = EGEDA_hist_own_oil.select_dtypes(include = [np.number]) * -1
EGEDA_hist_own_oil[neg_to_pos.columns] = neg_to_pos

EGEDA_hist_own_oil.to_csv('./data/4_Joined/EGEDA_hist_own_oil.csv', index = False)
EGEDA_hist_own_oil = pd.read_csv('./data/4_Joined/EGEDA_hist_own_oil.csv')

# Refining historical

EGEDA_hist_refining = EGEDA_data[(EGEDA_data['item_code_new'].isin(['9_4_oil_refineries'])) &
                                 (EGEDA_data['fuel_code'].isin(['6_crude_oil_and_ngl']))].copy().reset_index(drop = True)

EGEDA_hist_refining = EGEDA_hist_refining[['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))]\
                        .copy().reset_index(drop = True)

neg_to_pos = EGEDA_hist_refining.select_dtypes(include = [np.number]) * -1
EGEDA_hist_refining[neg_to_pos.columns] = neg_to_pos

EGEDA_hist_refining.to_csv('./data/4_Joined/EGEDA_hist_refining.csv', index = False)
EGEDA_hist_refining = pd.read_csv('./data/4_Joined/EGEDA_hist_refining.csv')

# Refinery Output historical
EGEDA_hist_refiningout = EGEDA_data[(EGEDA_data['item_code_new'].isin(['9_4_oil_refineries'])) &
                                    (EGEDA_data['fuel_code'].isin(['7_petroleum_products']))].copy().reset_index(drop = True)

EGEDA_hist_refiningout = EGEDA_hist_refiningout[['economy', 'fuel_code', 'item_code_new'] + list(range(2000, 2019))]\
                        .copy().reset_index(drop = True)

EGEDA_hist_refiningout.to_csv('./data/4_Joined/EGEDA_hist_refiningout.csv', index = False)
EGEDA_hist_refiningout = pd.read_csv('./data/4_Joined/EGEDA_hist_refiningout.csv')

# liquid and solid renewables historical

EGEDA_hist_own_renew = EGEDA_data[(EGEDA_data['item_code_new'].isin(['10_losses_and_own_use'])) &
                                (EGEDA_data['fuel_code'].isin(['15_solid_biomass', '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                                                    '16_6_biodiesel', '16_7_bio_jet_kerosene', 
                                                                                    '16_8_other_liquid_biofuels']))]\
                                  .copy().reset_index(drop = True)

EGEDA_hist_own_renew['FUEL'] = EGEDA_hist_own_renew['fuel_code'].map({'15_solid_biomass': 'Liquid and solid renewables', 
                                                                      '16_1_biogas': 'Liquid and solid renewables', 
                                                                      '16_3_municipal_solid_waste_renewable': 'Liquid and solid renewables',
                                                                      '16_5_biogasoline': 'Liquid and solid renewables', 
                                                                      '16_6_biodiesel': 'Liquid and solid renewables', 
                                                                      '16_7_bio_jet_kerosene': 'Liquid and solid renewables', 
                                                                      '16_8_other_liquid_biofuels': 'Liquid and solid renewables'})

EGEDA_hist_own_renew = EGEDA_hist_own_renew[['economy', 'FUEL', 'item_code_new'] + list(range(2000, 2019))]\
                        .copy().reset_index(drop = True)

neg_to_pos = EGEDA_hist_own_renew.select_dtypes(include = [np.number]) * -1
EGEDA_hist_own_renew[neg_to_pos.columns] = neg_to_pos

EGEDA_hist_own_renew.to_csv('./data/4_Joined/EGEDA_hist_own_renew.csv', index = False)
EGEDA_hist_own_renew = pd.read_csv('./data/4_Joined/EGEDA_hist_own_renew.csv')

#########################################################################################################################################

# OSeMOSYS demand reults dataframes

# Read heavyind mapping file
heavyind_mapping = pd.read_csv('./data/2_Mapping_and_other/heavyind_mapping.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

# Need a couple of strings for steel
steel_ind = ['IND_steel', 'IND_hysteel']

# REFERENCE

# Heavy industry dataframes

ref_steel_1 = ref_osemo_1[ref_osemo_1['TECHNOLOGY'].str.contains('|'.join(steel_ind))].copy()
ref_steel_1['tech_mix'] = ref_osemo_1['TECHNOLOGY'].map(heavyind_mapping)

ref_chem_1 = ref_osemo_1[ref_osemo_1['TECHNOLOGY'].str.contains('IND_chem')].copy()
ref_chem_1['tech_mix'] = ref_osemo_1['TECHNOLOGY'].map(heavyind_mapping)

ref_cement_1 = ref_osemo_1[ref_osemo_1['TECHNOLOGY'].str.contains('IND_cem')].copy()
ref_cement_1['tech_mix'] = ref_osemo_1['TECHNOLOGY'].map(heavyind_mapping)

ref_steel_2 = ref_steel_1.groupby(['REGION', 'tech_mix']).sum().reset_index()
ref_steel_2['Industry'] = 'Steel'
ref_steel_2 = ref_steel_2[['REGION', 'Industry', 'tech_mix'] + list(ref_steel_2.loc[:,'2018':'2050'])]

ref_chem_2 = ref_chem_1.groupby(['REGION', 'tech_mix']).sum().reset_index()
ref_chem_2['Industry'] = 'Chemical'
ref_chem_2 = ref_chem_2[['REGION', 'Industry', 'tech_mix'] + list(ref_chem_2.loc[:,'2018':'2050'])]

ref_cement_2 = ref_cement_1.groupby(['REGION', 'tech_mix']).sum().reset_index()
ref_cement_2['Industry'] = 'Cement'
ref_cement_2 = ref_cement_2[['REGION', 'Industry', 'tech_mix'] + list(ref_cement_2.loc[:,'2018':'2050'])]

# CARBON NEUTRALITY

# Heavy industry dataframes

netz_steel_1 = netz_osemo_1[netz_osemo_1['TECHNOLOGY'].str.contains('|'.join(steel_ind))].copy()
netz_steel_1['tech_mix'] = netz_osemo_1['TECHNOLOGY'].map(heavyind_mapping)

netz_chem_1 = netz_osemo_1[netz_osemo_1['TECHNOLOGY'].str.contains('IND_chem')].copy()
netz_chem_1['tech_mix'] = netz_osemo_1['TECHNOLOGY'].map(heavyind_mapping)

netz_cement_1 = netz_osemo_1[netz_osemo_1['TECHNOLOGY'].str.contains('IND_cem')].copy()
netz_cement_1['tech_mix'] = netz_osemo_1['TECHNOLOGY'].map(heavyind_mapping)

netz_steel_2 = netz_steel_1.groupby(['REGION', 'tech_mix']).sum().reset_index()
netz_steel_2['Industry'] = 'Steel'
netz_steel_2 = netz_steel_2[['REGION', 'Industry', 'tech_mix'] + list(netz_steel_2.loc[:,'2018':'2050'])]

netz_chem_2 = netz_chem_1.groupby(['REGION', 'tech_mix']).sum().reset_index()
netz_chem_2['Industry'] = 'Chemical'
netz_chem_2 = netz_chem_2[['REGION', 'Industry', 'tech_mix'] + list(netz_chem_2.loc[:,'2018':'2050'])]

netz_cement_2 = netz_cement_1.groupby(['REGION', 'tech_mix']).sum().reset_index()
netz_cement_2['Industry'] = 'Cement'
netz_cement_2 = netz_cement_2[['REGION', 'Industry', 'tech_mix'] + list(netz_cement_2.loc[:,'2018':'2050'])]

# Read heavyind mapping file
trn_mapping_2 = pd.read_csv('./data/2_Mapping_and_other/trn_mapping_2.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

trn_mapping_3 = pd.read_csv('./data/2_Mapping_and_other/trn_mapping_3.csv',\
    header = None, index_col = 0, squeeze = True).to_dict()

# Transport OSeMOSYS only

# REFERENCE
ref_roadmodal_1 = ref_osemo_1[ref_osemo_1['TECHNOLOGY'].str.contains('TRN_')].copy()
ref_roadmodal_1['modality'] = ref_osemo_1['TECHNOLOGY'].map(trn_mapping_2)

ref_roadmodal_2 = ref_roadmodal_1.groupby(['REGION', 'modality']).sum().reset_index()
ref_roadmodal_2['Transport'] = 'Road'
ref_roadmodal_2 = ref_roadmodal_2[['REGION', 'Transport', 'modality'] + list(ref_roadmodal_2.loc[:,'2018':'2050'])]

ref_roadfuel_1 = ref_osemo_1[ref_osemo_1['TECHNOLOGY'].str.contains('TRN_')].copy()
ref_roadfuel_1['modality'] = ref_osemo_1['TECHNOLOGY'].map(trn_mapping_3)

ref_roadfuel_2 = ref_roadfuel_1.groupby(['REGION', 'modality']).sum().reset_index()
ref_roadfuel_2['Transport'] = 'Road'
ref_roadfuel_2 = ref_roadfuel_2[['REGION', 'Transport', 'modality'] + list(ref_roadfuel_2.loc[:,'2018':'2050'])]

# CARBON NEUTRALITY
netz_roadmodal_1 = netz_osemo_1[netz_osemo_1['TECHNOLOGY'].str.contains('TRN_')].copy()
netz_roadmodal_1['modality'] = netz_osemo_1['TECHNOLOGY'].map(trn_mapping_2)

netz_roadmodal_2 = netz_roadmodal_1.groupby(['REGION', 'modality']).sum().reset_index()
netz_roadmodal_2['Transport'] = 'Road'
netz_roadmodal_2 = netz_roadmodal_2[['REGION', 'Transport', 'modality'] + list(netz_roadmodal_2.loc[:,'2018':'2050'])]

netz_roadfuel_1 = netz_osemo_1[netz_osemo_1['TECHNOLOGY'].str.contains('TRN_')].copy()
netz_roadfuel_1['modality'] = netz_osemo_1['TECHNOLOGY'].map(trn_mapping_3)

netz_roadfuel_2 = netz_roadfuel_1.groupby(['REGION', 'modality']).sum().reset_index()
netz_roadfuel_2['Transport'] = 'Road'
netz_roadfuel_2 = netz_roadfuel_2[['REGION', 'Transport', 'modality'] + list(netz_roadfuel_2.loc[:,'2018':'2050'])]

# Now build the subset dataframes for charts and tables

# Fix to do quicker one economy runs
# Economy_codes = ['01_AUS']

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
    ref_fedfuel_1.loc[ref_fedfuel_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
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
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_tradbio_2.loc[ref_tradbio_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_tradbio_2 = ref_tradbio_2[ref_tradbio_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    ref_fedfuel_1 = ref_fedfuel_1.append(ref_tradbio_2)

    # Combine the two dataframes that account for Modern renewables
    ref_fedfuel_1 = ref_fedfuel_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = '12_total_final_consumption')\
        .reset_index()[['fuel_code', 'item_code_new'] + list(ref_fedfuel_1.loc[:,'2000':'2050'])]\
            .set_index('fuel_code').loc[FED_agg_fuels].reset_index().replace(np.nan, 0)

    ref_fedfuel_1.loc['Total'] = ref_fedfuel_1.sum(numeric_only = True)

    ref_fedfuel_1.loc['Total', 'fuel_code'] = 'Total'
    ref_fedfuel_1.loc['Total', 'item_code_new'] = '12_total_final_consumption'

    # Get rid of zero rows
    # non_zero = (ref_fedfuel_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_fedfuel_1 = ref_fedfuel_1.loc[non_zero].reset_index(drop = True)

    ref_fedfuel_1_rows = ref_fedfuel_1.shape[0]
    ref_fedfuel_1_cols = ref_fedfuel_1.shape[1]

    ref_fedfuel_2 = ref_fedfuel_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_fedfuel_2_rows = ref_fedfuel_2.shape[0]
    ref_fedfuel_2_cols = ref_fedfuel_2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    ref_fedsector_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                        (EGEDA_years_reference['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years_reference['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True).replace(np.nan, 0)

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
    ref_fedsector_2 = ref_fedsector_2[['fuel_code', 'item_code_new'] + list(ref_fedsector_2.loc[:, '2000':])].replace(np.nan, 0)

    ref_fedsector_2.loc['Total'] = ref_fedsector_2.sum(numeric_only = True)

    ref_fedsector_2.loc['Total', 'fuel_code'] = '19_total'
    ref_fedsector_2.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    # non_zero = (ref_fedsector_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_fedsector_2 = ref_fedsector_2.loc[non_zero].reset_index(drop = True)

    ref_fedsector_2_rows = ref_fedsector_2.shape[0]
    ref_fedsector_2_cols = ref_fedsector_2.shape[1]

    ref_fedsector_3 = ref_fedsector_2[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_fedsector_3_rows = ref_fedsector_3.shape[0]
    ref_fedsector_3_cols = ref_fedsector_3.shape[1]

    # New FED by sector (not including non-energy)

    ref_tfec_1 = ref_fedsector_2[~ref_fedsector_2['item_code_new'].isin(['Non-energy', 'Total'])].copy().groupby(['fuel_code'])\
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
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_bld_2.loc[ref_bld_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_bld_2 = ref_bld_2[ref_bld_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code')\
        .loc[FED_agg_fuels].reset_index().replace(np.nan, 0)

    ref_bld_2.loc['Total'] = ref_bld_2.sum(numeric_only = True)

    ref_bld_2.loc['Total', 'fuel_code'] = 'Total'
    ref_bld_2.loc['Total', 'item_code_new'] = '16_x_buildings'

    # Get rid of zero rows
    # non_zero = (ref_bld_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_bld_2 = ref_bld_2.loc[non_zero].reset_index(drop = True)

    ref_bld_2_rows = ref_bld_2.shape[0]
    ref_bld_2_cols = ref_bld_2.shape[1]

    ref_bld_3 = ref_bld_1[(ref_bld_1['fuel_code'] == '19_total') &
                      (ref_bld_1['item_code_new'].isin(Buildings_items))].copy().reset_index(drop = True).replace(np.nan, 0)

    ref_bld_3.loc[ref_bld_3['item_code_new'] == '16_1_commercial_and_public_services', 'item_code_new'] = 'Services' 
    ref_bld_3.loc[ref_bld_3['item_code_new'] == '16_2_residential', 'item_code_new'] = 'Residential'

    ref_bld_3.loc['Total'] = ref_bld_3.sum(numeric_only = True)

    ref_bld_3.loc['Total', 'fuel_code'] = '19_total'
    ref_bld_3.loc['Total', 'item_code_new'] = 'Buildings'

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

    ref_ind_1 = ref_ind_1[['fuel_code', 'item_code_new'] + list(ref_ind_1.loc[:, '2000':])].replace(np.nan, 0)

    ref_ind_1.loc['Total'] = ref_ind_1.sum(numeric_only = True)

    ref_ind_1.loc['Total', 'fuel_code'] = '19_total'
    ref_ind_1.loc['Total', 'item_code_new'] = 'Industry'

    # Get rid of zero rows
    non_zero = (ref_ind_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_ind_1 = ref_ind_1.loc[non_zero].reset_index(drop = True)

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
    ref_ind_2.loc[ref_ind_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    ref_ind_2.loc[ref_ind_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_ind_2.loc[ref_ind_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_ind_2 = ref_ind_2[ref_ind_2['fuel_code'].isin(FED_agg_fuels_ind)].set_index('fuel_code').loc[FED_agg_fuels_ind].reset_index().replace(np.nan, 0)

    ref_ind_2.loc['Total'] = ref_ind_2.sum(numeric_only = True)

    ref_ind_2.loc['Total', 'fuel_code'] = 'Total'
    ref_ind_2.loc['Total', 'item_code_new'] = '14_industry_sector'

    # Get rid of zero rows
    # non_zero = (ref_ind_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_ind_2 = ref_ind_2.loc[non_zero].reset_index(drop = True)
    
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
    ref_trn_1.loc[ref_trn_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    ref_trn_1.loc[ref_trn_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_trn_1 = ref_trn_1[ref_trn_1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index().replace(np.nan, 0)

    ref_trn_1.loc['Total'] = ref_trn_1.sum(numeric_only = True)

    ref_trn_1.loc['Total', 'fuel_code'] = 'Total'
    ref_trn_1.loc['Total', 'item_code_new'] = '15_transport_sector'

    # Get rid of zero rows
    # non_zero = (ref_trn_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_trn_1 = ref_trn_1.loc[non_zero].reset_index(drop = True)

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

    ref_trn_2 = ref_trn_2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True).replace(np.nan, 0)

    ref_trn_2.loc['Total'] = ref_trn_2.sum(numeric_only = True)

    ref_trn_2.loc['Total', 'fuel_code'] = '19_total'
    ref_trn_2.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    non_zero = (ref_trn_2.loc[:,'2018':] != 0).any(axis = 1)
    ref_trn_2 = ref_trn_2.loc[non_zero].reset_index(drop = True)

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
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_ag_1.loc[ref_ag_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    ref_ag_1 = ref_ag_1[ref_ag_1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index().replace(np.nan, 0)

    ref_ag_1.loc['Total'] = ref_ag_1.sum(numeric_only = True)

    ref_ag_1.loc['Total', 'fuel_code'] = 'Total'
    ref_ag_1.loc['Total', 'item_code_new'] = 'Agriculture'

    # # Get rid of zero rows
    # non_zero = (ref_ag_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_ag_1 = ref_ag_1.loc[non_zero].reset_index(drop = True)
    
    ref_ag_1_rows = ref_ag_1.shape[0]
    ref_ag_1_cols = ref_ag_1.shape[1]

    ref_ag_2 = ref_ag_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_ag_2_rows = ref_ag_2.shape[0]
    ref_ag_2_cols = ref_ag_2.shape[1]

    # Hydrogen data frame reference

    ref_hyd_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                                        (EGEDA_years_reference['item_code_new'].isin(Sectors_tfc)) &
                                        (EGEDA_years_reference['fuel_code'] == '16_x_hydrogen')].groupby('item_code_new').sum().assign(fuel_code = 'Hydrogen').reset_index()

    buildings_hy = ref_hyd_1[ref_hyd_1['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Buildings', fuel_code = 'Hydrogen')

    ag_hy = ref_hyd_1[ref_hyd_1['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Agriculture', fuel_code = 'Hydrogen')

    ref_hyd_1 = ref_hyd_1.append([buildings_hy, ag_hy])\
        [['fuel_code', 'item_code_new'] + list(ref_hyd_1.loc[:, '2018':'2050'])].reset_index(drop = True)

    ref_hyd_1.loc[ref_hyd_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_hyd_1.loc[ref_hyd_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'

    ref_hyd_1 = ref_hyd_1[ref_hyd_1['item_code_new'].isin(['Agriculture', 'Buildings', 'Industry', 'Transport'])]\
        .copy().reset_index(drop = True).replace(np.nan, 0)

    ref_hyd_1.loc['Total'] = ref_hyd_1.sum(numeric_only = True)

    ref_hyd_1.loc['Total', 'fuel_code'] = 'Hydrogen'
    ref_hyd_1.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    non_zero = (ref_hyd_1.loc[:,'2018':] != 0).any(axis = 1)
    ref_hyd_1 = ref_hyd_1.loc[non_zero].reset_index(drop = True)

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
    netz_fedfuel_1.loc[netz_fedfuel_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
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
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_tradbio_2.loc[netz_tradbio_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_tradbio_2 = netz_tradbio_2[netz_tradbio_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index()

    netz_fedfuel_1 = netz_fedfuel_1.append(netz_tradbio_2)

    # Combine the two dataframes that account for Modern renewables
    netz_fedfuel_1 = netz_fedfuel_1.copy().groupby(['fuel_code']).sum().assign(item_code_new = '12_total_final_consumption')\
        .reset_index()[['fuel_code', 'item_code_new'] + list(netz_fedfuel_1.loc[:,'2000':'2050'])]\
            .set_index('fuel_code').loc[FED_agg_fuels].reset_index().replace(np.nan, 0)

    netz_fedfuel_1.loc['Total'] = netz_fedfuel_1.sum(numeric_only = True)

    netz_fedfuel_1.loc['Total', 'fuel_code'] = 'Total'
    netz_fedfuel_1.loc['Total', 'item_code_new'] = '12_total_final_consumption'

    # Get rid of zero rows
    # non_zero = (netz_fedfuel_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_fedfuel_1 = netz_fedfuel_1.loc[non_zero].reset_index(drop = True)

    netz_fedfuel_1_rows = netz_fedfuel_1.shape[0]
    netz_fedfuel_1_cols = netz_fedfuel_1.shape[1]

    netz_fedfuel_2 = netz_fedfuel_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_fedfuel_2_rows = netz_fedfuel_2.shape[0]
    netz_fedfuel_2_cols = netz_fedfuel_2.shape[1]                                                                          
    
    # Second data frame construction: FED by sectors
    netz_fedsector_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                        (EGEDA_years_netzero['item_code_new'].isin(Sectors_tfc)) &
                        (EGEDA_years_netzero['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    netz_fedsector_1 = netz_fedsector_1[['fuel_code', 'item_code_new'] + list(netz_fedsector_1.loc[:,'2000':])].replace(np.nan, 0)
    
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
    netz_fedsector_2 = netz_fedsector_2[['fuel_code', 'item_code_new'] + list(netz_fedsector_2.loc[:, '2000':])].replace(np.nan, 0)

    netz_fedsector_2.loc['Total'] = netz_fedsector_2.sum(numeric_only = True)

    netz_fedsector_2.loc['Total', 'fuel_code'] = '19_total'
    netz_fedsector_2.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    # non_zero = (netz_fedsector_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_fedsector_2 = netz_fedsector_2.loc[non_zero].reset_index(drop = True)

    netz_fedsector_2_rows = netz_fedsector_2.shape[0]
    netz_fedsector_2_cols = netz_fedsector_2.shape[1]

    netz_fedsector_3 = netz_fedsector_2[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_fedsector_3_rows = netz_fedsector_3.shape[0]
    netz_fedsector_3_cols = netz_fedsector_3.shape[1]

    # New FED by sector (not including non-energy)

    netz_tfec_1 = netz_fedsector_2[~netz_fedsector_2['item_code_new'].isin(['Non-energy', 'Total'])].copy().groupby(['fuel_code'])\
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
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_bld_2.loc[netz_bld_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_bld_2 = netz_bld_2[netz_bld_2['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code')\
        .loc[FED_agg_fuels].reset_index().replace(np.nan, 0)

    netz_bld_2.loc['Total'] = netz_bld_2.sum(numeric_only = True)

    netz_bld_2.loc['Total', 'fuel_code'] = 'Total'
    netz_bld_2.loc['Total', 'item_code_new'] = '16_x_buildings'

    # Get rid of zero rows
    # non_zero = (netz_bld_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_bld_2 = netz_bld_2.loc[non_zero].reset_index(drop = True)

    netz_bld_2_rows = netz_bld_2.shape[0]
    netz_bld_2_cols = netz_bld_2.shape[1]

    netz_bld_3 = netz_bld_1[(netz_bld_1['fuel_code'] == '19_total') &
                      (netz_bld_1['item_code_new'].isin(Buildings_items))].copy().reset_index(drop = True).replace(np.nan, 0)

    netz_bld_3.loc[netz_bld_3['item_code_new'] == '16_1_commercial_and_public_services', 'item_code_new'] = 'Services' 
    netz_bld_3.loc[netz_bld_3['item_code_new'] == '16_2_residential', 'item_code_new'] = 'Residential'

    netz_bld_3.loc['Total'] = netz_bld_3.sum(numeric_only = True)

    netz_bld_3.loc['Total', 'fuel_code'] = '19_total'
    netz_bld_3.loc['Total', 'item_code_new'] = 'Buildings'

    netz_bld_3 = netz_bld_3.copy().reset_index(drop = True)

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

    netz_ind_1 = netz_ind_1[['fuel_code', 'item_code_new'] + list(netz_ind_1.loc[:, '2000':])].replace(np.nan, 0)

    netz_ind_1.loc['Total'] = netz_ind_1.sum(numeric_only = True)

    netz_ind_1.loc['Total', 'fuel_code'] = '19_total'
    netz_ind_1.loc['Total', 'item_code_new'] = 'Industry'

    # Get rid of zero rows
    # non_zero = (netz_ind_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_ind_1 = netz_ind_1.loc[non_zero].reset_index(drop = True)

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
    netz_ind_2.loc[netz_ind_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    netz_ind_2.loc[netz_ind_2['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_ind_2.loc[netz_ind_2['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_ind_2 = netz_ind_2[netz_ind_2['fuel_code'].isin(FED_agg_fuels_ind)].set_index('fuel_code').loc[FED_agg_fuels_ind].reset_index().replace(np.nan, 0)

    netz_ind_2.loc['Total'] = netz_ind_2.sum(numeric_only = True)

    netz_ind_2.loc['Total', 'fuel_code'] = 'Total'
    netz_ind_2.loc['Total', 'item_code_new'] = '14_industry_sector'
    
    # Get rid of zero rows
    # non_zero = (netz_ind_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_ind_2 = netz_ind_2.loc[non_zero].reset_index(drop = True)

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
    netz_trn_1.loc[netz_trn_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    netz_trn_1.loc[netz_trn_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_trn_1 = netz_trn_1[netz_trn_1['fuel_code'].isin(Transport_fuels_agg)].set_index('fuel_code').loc[Transport_fuels_agg].reset_index().replace(np.nan, 0)

    netz_trn_1.loc['Total'] = netz_trn_1.sum(numeric_only = True)

    netz_trn_1.loc['Total', 'fuel_code'] = 'Total'
    netz_trn_1.loc['Total', 'item_code_new'] = '15_transport_sector'

    # Get rid of zero rows
    # non_zero = (netz_trn_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_trn_1 = netz_trn_1.loc[non_zero].reset_index(drop = True)

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

    netz_trn_2 = netz_trn_2[['fuel_code', 'item_code_new'] + col_chart_years_transport].reset_index(drop = True).replace(np.nan, 0)

    netz_trn_2.loc['Total'] = netz_trn_2.sum(numeric_only = True)

    netz_trn_2.loc['Total', 'fuel_code'] = '19_total'
    netz_trn_2.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    # non_zero = (netz_trn_2.loc[:,'2018':] != 0).any(axis = 1)
    # netz_trn_2 = netz_trn_2.loc[non_zero].reset_index(drop = True)

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
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_ag_1.loc[netz_ag_1['fuel_code'] == '18_heat', 'fuel_code'] = 'Heat'

    netz_ag_1 = netz_ag_1[netz_ag_1['fuel_code'].isin(FED_agg_fuels)].set_index('fuel_code').loc[FED_agg_fuels].reset_index().replace(np.nan, 0)

    netz_ag_1.loc['Total'] = netz_ag_1.sum(numeric_only = True)

    netz_ag_1.loc['Total', 'fuel_code'] = 'Total'
    netz_ag_1.loc['Total', 'item_code_new'] = 'Agriculture'
    
    # Get rid of zero rows
    # non_zero = (netz_ag_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_ag_1 = netz_ag_1.loc[non_zero].reset_index(drop = True)

    netz_ag_1_rows = netz_ag_1.shape[0]
    netz_ag_1_cols = netz_ag_1.shape[1]

    netz_ag_2 = netz_ag_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_ag_2_rows = netz_ag_2.shape[0]
    netz_ag_2_cols = netz_ag_2.shape[1]

    # Hydrogen data frame net zero

    netz_hyd_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                                        (EGEDA_years_netzero['item_code_new'].isin(Sectors_tfc)) &
                                        (EGEDA_years_netzero['fuel_code'] == '16_x_hydrogen')].groupby('item_code_new').sum().assign(fuel_code = 'Hydrogen').reset_index()

    buildings_hy = netz_hyd_1[netz_hyd_1['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Buildings', fuel_code = 'Hydrogen')

    ag_hy = netz_hyd_1[netz_hyd_1['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])].groupby('fuel_code')\
        .sum().assign(item_code_new = 'Agriculture', fuel_code = 'Hydrogen')

    netz_hyd_1 = netz_hyd_1.append([buildings_hy, ag_hy])\
        [['fuel_code', 'item_code_new'] + list(netz_hyd_1.loc[:, '2018':'2050'])].reset_index(drop = True)

    netz_hyd_1.loc[netz_hyd_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_hyd_1.loc[netz_hyd_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'

    netz_hyd_1 = netz_hyd_1[netz_hyd_1['item_code_new'].isin(['Agriculture', 'Buildings', 'Industry', 'Transport'])]\
        .copy().reset_index(drop = True).replace(np.nan, 0)

    netz_hyd_1.loc['Total'] = netz_hyd_1.sum(numeric_only = True)

    netz_hyd_1.loc['Total', 'fuel_code'] = 'Hydrogen'
    netz_hyd_1.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    # non_zero = (netz_hyd_1.loc[:,'2018':] != 0).any(axis = 1)
    # netz_hyd_1 = netz_hyd_1.loc[non_zero].reset_index(drop = True)

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
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    ref_tpes_1 = ref_tpes_1[ref_tpes_1['fuel_code'].isin(TPES_agg_fuels1)].set_index('fuel_code').loc[TPES_agg_fuels1].reset_index().replace(np.nan, 0)

    ref_tpes_1.loc['Total'] = ref_tpes_1.sum(numeric_only = True)

    ref_tpes_1.loc['Total', 'fuel_code'] = 'Total'
    ref_tpes_1.loc['Total', 'item_code_new'] = '7_total_primary_energy_supply'

    # Get rid of zero rows
    non_zero = (ref_tpes_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_tpes_1 = ref_tpes_1.loc[non_zero].reset_index(drop = True)

    ref_tpes_1_rows = ref_tpes_1.shape[0]
    ref_tpes_1_cols = ref_tpes_1.shape[1]

    ref_tpes_2 = ref_tpes_1[['fuel_code', 'item_code_new'] + col_chart_years]
    # ref_tpes_2 = ref_tpes_2[ref_tpes_2['fuel_code'] != 'Total']

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

    ref_prod_1 = ref_prod_1[ref_prod_1['fuel_code'].isin(TPES_agg_fuels2)].set_index('fuel_code').loc[TPES_agg_fuels2].reset_index().replace(np.nan, 0)

    ref_prod_1.loc['Total'] = ref_prod_1.sum(numeric_only = True)

    ref_prod_1.loc['Total', 'fuel_code'] = 'Total'
    ref_prod_1.loc['Total', 'item_code_new'] = '1_indigenous_production'

    # Get rid of zero rows
    non_zero = (ref_prod_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_prod_1 = ref_prod_1.loc[non_zero].reset_index(drop = True)

    ref_prod_1_rows = ref_prod_1.shape[0]
    ref_prod_1_cols = ref_prod_1.shape[1]

    ref_prod_2 = ref_prod_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_prod_2_rows = ref_prod_2.shape[0]
    ref_prod_2_cols = ref_prod_2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    ref_tpes_comp_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                            (EGEDA_years_reference['item_code_new'].isin(tpes_items)) &
                                            (EGEDA_years_reference['fuel_code'] == '19_total')]
    
    net_trade = ref_tpes_comp_1[ref_tpes_comp_1['item_code_new'].isin(['2_imports', '3_exports'])]\
        .groupby(['economy']).sum().assign(fuel_code = '19_total', item_code_new = 'Net trade')
                           
    bunkers = ref_tpes_comp_1[ref_tpes_comp_1['item_code_new'].isin(['4_international_marine_bunkers', '5_international_aviation_bunkers'])]\
        .groupby(['economy', 'fuel_code']).sum().assign(fuel_code = '19_total', item_code_new = 'Bunkers')
    
    ref_tpes_comp_1 = ref_tpes_comp_1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    ref_tpes_comp_1.loc[ref_tpes_comp_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    ref_tpes_comp_1.loc[ref_tpes_comp_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock changes'
    
    ref_tpes_comp_1 = ref_tpes_comp_1.loc[ref_tpes_comp_1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True).replace(np.nan, 0)

    # Get rid of zero rows
    non_zero = (ref_tpes_comp_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_tpes_comp_1 = ref_tpes_comp_1.loc[non_zero].reset_index(drop = True)
    
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
    ref_imports_1.loc[ref_imports_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_imports_1.loc[ref_imports_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    ref_imports_1 = ref_imports_1[ref_imports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_imports_1.loc[:, '2000':])].replace(np.nan, 0)

    ref_imports_1.loc['Total'] = ref_imports_1.sum(numeric_only = True)

    ref_imports_1.loc['Total', 'fuel_code'] = 'Total'
    ref_imports_1.loc['Total', 'item_code_new'] = '2_imports'

    # Get rid of zero rows
    # non_zero = (ref_imports_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_imports_1 = ref_imports_1.loc[non_zero].reset_index(drop = True)

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
    ref_exports_1.loc[ref_exports_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_exports_1.loc[ref_exports_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    ref_exports_1 = ref_exports_1[ref_exports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_exports_1.loc[:, '2000':])].replace(np.nan, 0)

    ref_exports_1.loc['Total'] = ref_exports_1.sum(numeric_only = True)

    ref_exports_1.loc['Total', 'fuel_code'] = 'Total'
    ref_exports_1.loc['Total', 'item_code_new'] = '3_exports'

    # Get rid of zero rows
    # non_zero = (ref_exports_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_exports_1 = ref_exports_1.loc[non_zero].reset_index(drop = True)

    ref_exports_1_rows = ref_exports_1.shape[0]
    ref_exports_1_cols = ref_exports_1.shape[1]

    ref_exports_2 = ref_exports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_exports_2_rows = ref_exports_2.shape[0]
    ref_exports_2_cols = ref_exports_2.shape[1]

    # Temporary exports file to get net trade dataframe

    ref_exports_temp1 = ref_exports_1.copy().select_dtypes(include = [np.number]) * -1
    ref_exports_temp2 = ref_exports_1.copy()
    ref_exports_temp2[ref_exports_temp1.columns] = ref_exports_temp1

    # Net trade

    ref_nettrade_1 = ref_imports_1.copy().append(ref_exports_temp2).groupby('fuel_code').sum()\
        .assign(item_code_new = 'Net trade').reset_index()

    ref_nettrade_1 = ref_nettrade_1[['fuel_code', 'item_code_new'] + list(ref_nettrade_1.loc[:, '2000': '2050'])]

    ref_nettrade_1.loc[ref_nettrade_1['fuel_code'] == 'Total', 'fuel_code'] = 'Trade balance'

    ref_nettrade_1_rows = ref_nettrade_1.shape[0]
    ref_nettrade_1_cols = ref_nettrade_1.shape[1]

    # Electricity trade

    ref_electrade_1 = ref_imports_2[ref_imports_2['fuel_code'] == 'Electricity'].copy()\
        .append(ref_exports_2[ref_exports_2['fuel_code'] == 'Electricity'].copy()).reset_index(drop = True)

    # Change exports back to negative
    ref_electrade_1.loc[ref_electrade_1['item_code_new'] == '3_exports', list(ref_electrade_1.columns[2:])]\
        = ref_electrade_1.loc[ref_electrade_1['item_code_new'] == '3_exports', list(ref_electrade_1.columns[2:])]\
            .apply(lambda x: x * -1)

    ref_electrade_1.loc[ref_electrade_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    ref_electrade_1.loc[ref_electrade_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'

    ref_electrade_1_rows = ref_electrade_1.shape[0]
    ref_electrade_1_cols = ref_electrade_1.shape[1]

    # Bunkers dataframes

    ref_bunkers_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '4_international_marine_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(marine_bunker_fuels))]

    ref_bunkers_1 = ref_bunkers_1[['fuel_code', 'item_code_new'] + list(ref_bunkers_1.loc[:, '2000':])].reset_index(drop = True)\
        .replace(np.nan, 0)

    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '7_8_fuel_oil', 'fuel_code'] = 'Fuel oil'
    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '16_6_biodiesel', 'fuel_code'] = 'Biodiesel'
    ref_bunkers_1.loc[ref_bunkers_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'


    # Make bunkers data non-negative
    ref_bunkers_1.update(ref_bunkers_1.select_dtypes(include = [np.number]).abs())

    # Get rid of zero rows
    # non_zero = (ref_bunkers_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_bunkers_1 = ref_bunkers_1.loc[non_zero].reset_index(drop = True)

    ref_bunkers_1_rows = ref_bunkers_1.shape[0]
    ref_bunkers_1_cols = ref_bunkers_1.shape[1]

    ref_bunkers_2 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                              (EGEDA_years_reference['item_code_new'] == '5_international_aviation_bunkers') & 
                              (EGEDA_years_reference['fuel_code'].isin(aviation_bunker_fuels))]

    jetfuel = ref_bunkers_2[ref_bunkers_2['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Jet fuel',
                                                 item_code_new = '5_international_aviation_bunkers')
    
    ref_bunkers_2 = ref_bunkers_2.append([jetfuel]).reset_index(drop = True)

    ref_bunkers_2 = ref_bunkers_2[['fuel_code', 'item_code_new'] + list(ref_bunkers_2.loc[:, '2000':])]

    ref_bunkers_2.loc[ref_bunkers_2['fuel_code'] == '7_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'
    ref_bunkers_2.loc[ref_bunkers_2['fuel_code'] == '16_7_bio_jet_kerosene', 'fuel_code'] = 'Biojet kerosene'
    ref_bunkers_2.loc[ref_bunkers_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    ref_bunkers_2 = ref_bunkers_2[ref_bunkers_2['fuel_code'].isin(avi_bunker)]\
        .set_index('fuel_code').loc[avi_bunker].reset_index()\
            [['fuel_code', 'item_code_new'] + list(ref_bunkers_2.loc[:, '2000':])].replace(np.nan, 0)

    # Make bunkers data non-negative
    ref_bunkers_2.update(ref_bunkers_2.select_dtypes(include = [np.number]).abs())

    # Get rid of zero rows
    # non_zero = (ref_bunkers_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_bunkers_2 = ref_bunkers_2.loc[non_zero].reset_index(drop = True)

    ref_bunkers_2_rows = ref_bunkers_2.shape[0]
    ref_bunkers_2_cols = ref_bunkers_2.shape[1]

    ######################################################################################################################

    # TPES CARBON NEUTRALITY DATA FRAMES
    # First data frame: TPES by fuels (and also fourth and sixth dataframe with slight tweaks)
    netz_tpes_df = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

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
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    netz_tpes_1 = netz_tpes_1[netz_tpes_1['fuel_code'].isin(TPES_agg_fuels1)].set_index('fuel_code').loc[TPES_agg_fuels1].reset_index().replace(np.nan, 0)

    netz_tpes_1.loc['Total'] = netz_tpes_1.sum(numeric_only = True)

    netz_tpes_1.loc['Total', 'fuel_code'] = 'Total'
    netz_tpes_1.loc['Total', 'item_code_new'] = '7_total_primary_energy_supply'

    # Get rid of zero rows
    # non_zero = (netz_tpes_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_tpes_1 = netz_tpes_1.loc[non_zero].reset_index(drop = True)

    netz_tpes_1_rows = netz_tpes_1.shape[0]
    netz_tpes_1_cols = netz_tpes_1.shape[1]

    netz_tpes_2 = netz_tpes_1[['fuel_code', 'item_code_new'] + col_chart_years]
    # netz_tpes_2 = netz_tpes_2[netz_tpes_2['fuel_code'] != 'Total']

    netz_tpes_2_rows = netz_tpes_2.shape[0]
    netz_tpes_2_cols = netz_tpes_2.shape[1]
    
    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    netz_prod_df = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

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

    netz_prod_1 = netz_prod_1[netz_prod_1['fuel_code'].isin(TPES_agg_fuels2)].set_index('fuel_code').loc[TPES_agg_fuels2].reset_index().replace(np.nan, 0)

    netz_prod_1.loc['Total'] = netz_prod_1.sum(numeric_only = True)

    netz_prod_1.loc['Total', 'fuel_code'] = 'Total'
    netz_prod_1.loc['Total', 'item_code_new'] = '1_indigenous_production'

    # Get rid of zero rows
    # non_zero = (netz_prod_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_prod_1 = netz_prod_1.loc[non_zero].reset_index(drop = True)

    netz_prod_1_rows = netz_prod_1.shape[0]
    netz_prod_1_cols = netz_prod_1.shape[1]

    netz_prod_2 = netz_prod_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_prod_2_rows = netz_prod_2.shape[0]
    netz_prod_2_cols = netz_prod_2.shape[1]
    
    # Third data frame: production; net exports; bunkers; stock changes
    
    netz_tpes_comp_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                           (EGEDA_years_netzero['item_code_new'].isin(tpes_items)) &
                           (EGEDA_years_netzero['fuel_code'] == '19_total')]
    
    net_trade = netz_tpes_comp_1[netz_tpes_comp_1['item_code_new'].isin(['2_imports', '3_exports'])]\
        .groupby(['economy']).sum().assign(fuel_code = '19_total', item_code_new = 'Net trade')
                           
    bunkers = netz_tpes_comp_1[netz_tpes_comp_1['item_code_new'].isin(['4_international_marine_bunkers', '5_international_aviation_bunkers'])]\
        .groupby(['economy', 'fuel_code']).sum().assign(fuel_code = '19_total', item_code_new = 'Bunkers')
    
    netz_tpes_comp_1 = netz_tpes_comp_1.append([net_trade, bunkers])[['fuel_code', 'item_code_new'] + col_chart_years].reset_index(drop = True)
    
    netz_tpes_comp_1.loc[netz_tpes_comp_1['item_code_new'] == '1_indigenous_production', 'item_code_new'] = 'Production'
    netz_tpes_comp_1.loc[netz_tpes_comp_1['item_code_new'] == '6_stock_change', 'item_code_new'] = 'Stock changes'
    
    netz_tpes_comp_1 = netz_tpes_comp_1.loc[netz_tpes_comp_1['item_code_new'].isin(['Production',
                                                                           'Net trade',
                                                                           'Bunkers',
                                                                           'Stock changes'])].reset_index(drop = True).replace(np.nan, 0)
    
    # Get rid of zero rows
    # non_zero = (netz_tpes_comp_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_tpes_comp_1 = netz_tpes_comp_1.loc[non_zero].reset_index(drop = True)

    netz_tpes_comp_1_rows = netz_tpes_comp_1.shape[0]
    netz_tpes_comp_1_cols = netz_tpes_comp_1.shape[1]

    # Imports/exports data frame

    netz_imports_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '2_imports') & 
                              (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))]

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
    netz_imports_1.loc[netz_imports_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_imports_1.loc[netz_imports_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    netz_imports_1 = netz_imports_1[netz_imports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_imports_1.loc[:, '2000':])].replace(np.nan, 0)

    netz_imports_1.loc['Total'] = netz_imports_1.sum(numeric_only = True)

    netz_imports_1.loc['Total', 'fuel_code'] = 'Total'
    netz_imports_1.loc['Total', 'item_code_new'] = '2_imports'

    # Get rid of zero rows
    # non_zero = (netz_imports_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_imports_1 = netz_imports_1.loc[non_zero].reset_index(drop = True)

    netz_imports_1_rows = netz_imports_1.shape[0]
    netz_imports_1_cols = netz_imports_1.shape[1] 

    netz_imports_2 = netz_imports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_imports_2_rows = netz_imports_2.shape[0]
    netz_imports_2_cols = netz_imports_2.shape[1]                             

    netz_exports_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '3_exports') & 
                              (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].copy()

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
    netz_exports_1.loc[netz_exports_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_exports_1.loc[netz_exports_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    netz_exports_1 = netz_exports_1[netz_exports_1['fuel_code'].isin(TPES_agg_trade)]\
        .set_index('fuel_code').loc[TPES_agg_trade].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_exports_1.loc[:, '2000':])].replace(np.nan, 0)

    netz_exports_1.loc['Total'] = netz_exports_1.sum(numeric_only = True)

    netz_exports_1.loc['Total', 'fuel_code'] = 'Total'
    netz_exports_1.loc['Total', 'item_code_new'] = '3_exports'

    # Get rid of zero rows
    # non_zero = (netz_exports_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_exports_1 = netz_exports_1.loc[non_zero].reset_index(drop = True)

    netz_exports_1_rows = netz_exports_1.shape[0]
    netz_exports_1_cols = netz_exports_1.shape[1]

    netz_exports_2 = netz_exports_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_exports_2_rows = netz_exports_2.shape[0]
    netz_exports_2_cols = netz_exports_2.shape[1]

    # Temporary exports file to get net trade dataframe

    netz_exports_temp1 = netz_exports_1.copy().select_dtypes(include = [np.number]) * -1
    netz_exports_temp2 = netz_exports_1.copy()
    netz_exports_temp2[netz_exports_temp1.columns] = netz_exports_temp1

    # Net trade

    netz_nettrade_1 = netz_imports_1.copy().append(netz_exports_temp2).groupby('fuel_code').sum()\
        .assign(item_code_new = 'Net trade').reset_index()

    netz_nettrade_1 = netz_nettrade_1[['fuel_code', 'item_code_new'] + list(netz_nettrade_1.loc[:,'2000': '2050'])]

    netz_nettrade_1.loc[netz_nettrade_1['fuel_code'] == 'Total', 'fuel_code'] = 'Trade balance'

    netz_nettrade_1_rows = netz_nettrade_1.shape[0]
    netz_nettrade_1_cols = netz_nettrade_1.shape[1] 

    # Electricity trade

    netz_electrade_1 = netz_imports_2[netz_imports_2['fuel_code'] == 'Electricity'].copy()\
        .append(netz_exports_2[netz_exports_2['fuel_code'] == 'Electricity'].copy()).reset_index(drop = True)

    # Change exports back to negative
    netz_electrade_1.loc[netz_electrade_1['item_code_new'] == '3_exports', list(netz_electrade_1.columns[2:])]\
        = netz_electrade_1.loc[netz_electrade_1['item_code_new'] == '3_exports', list(netz_electrade_1.columns[2:])]\
            .apply(lambda x: x * -1)

    netz_electrade_1.loc[netz_electrade_1['item_code_new'] == '2_imports', 'item_code_new'] = 'Imports'
    netz_electrade_1.loc[netz_electrade_1['item_code_new'] == '3_exports', 'item_code_new'] = 'Exports'

    netz_electrade_1_rows = netz_electrade_1.shape[0]
    netz_electrade_1_cols = netz_electrade_1.shape[1]

    # Bunkers dataframes

    netz_bunkers_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '4_international_marine_bunkers') & 
                              (EGEDA_years_netzero['fuel_code'].isin(marine_bunker_fuels))]

    netz_bunkers_1 = netz_bunkers_1[['fuel_code', 'item_code_new'] + list(netz_bunkers_1.loc[:, '2000':])].reset_index(drop = True)\
        .replace(np.nan, 0)

    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '7_7_gas_diesel_oil', 'fuel_code'] = 'Gas diesel oil'
    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '7_8_fuel_oil', 'fuel_code'] = 'Fuel oil'
    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '8_1_natural_gas', 'fuel_code'] = 'Gas'
    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '16_6_biodiesel', 'fuel_code'] = 'Biodiesel'
    netz_bunkers_1.loc[netz_bunkers_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    # Make bunkers data non-negative
    netz_bunkers_1.update(netz_bunkers_1.select_dtypes(include = [np.number]).abs())

    # Get rid of zero rows
    # non_zero = (netz_bunkers_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_bunkers_1 = netz_bunkers_1.loc[non_zero].reset_index(drop = True)

    netz_bunkers_1_rows = netz_bunkers_1.shape[0]
    netz_bunkers_1_cols = netz_bunkers_1.shape[1]

    netz_bunkers_2 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                              (EGEDA_years_netzero['item_code_new'] == '5_international_aviation_bunkers') & 
                              (EGEDA_years_netzero['fuel_code'].isin(aviation_bunker_fuels))]

    jetfuel = netz_bunkers_2[netz_bunkers_2['fuel_code'].isin(['7_x_jet_fuel'])]\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Jet fuel',
                                                 item_code_new = '5_international_aviation_bunkers')
    
    netz_bunkers_2 = netz_bunkers_2.append([jetfuel]).reset_index(drop = True)

    netz_bunkers_2 = netz_bunkers_2[['fuel_code', 'item_code_new'] + list(netz_bunkers_2.loc[:, '2000':])]

    netz_bunkers_2.loc[netz_bunkers_2['fuel_code'] == '7_2_aviation_gasoline', 'fuel_code'] = 'Aviation gasoline'
    netz_bunkers_2.loc[netz_bunkers_2['fuel_code'] == '16_7_bio_jet_kerosene', 'fuel_code'] = 'Biojet kerosene'
    netz_bunkers_2.loc[netz_bunkers_2['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    netz_bunkers_2 = netz_bunkers_2[netz_bunkers_2['fuel_code'].isin(avi_bunker)]\
        .set_index('fuel_code').loc[avi_bunker].reset_index()\
            [['fuel_code', 'item_code_new'] + list(netz_bunkers_2.loc[:, '2000':])].replace(np.nan, 0)

    # Make bunkers data non-negative
    netz_bunkers_2.update(netz_bunkers_2.select_dtypes(include = [np.number]).abs())

    # Get rid of zero rows
    # non_zero = (netz_bunkers_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_bunkers_2 = netz_bunkers_2.loc[non_zero].reset_index(drop = True)

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

    # imports = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(imports_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Imports',
    #                                                                                     TECHNOLOGY = 'Electricity imports')                                                                                         

    # Second level aggregations

    coal2 = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(coal_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    renew2 = ref_pow_use_1[ref_pow_use_1['FUEL'].isin(renewables_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Renewables',
                                                                                      TECHNOLOGY = 'Renewables power')

    # Use by fuel data frame number 1

    ref_pow_use_2 = ref_pow_use_1.append([coal, lignite, oil, gas, nuclear, hydro, solar, wind, geothermal, biomass, other_renew, other])\
        [['FUEL', 'TECHNOLOGY'] + list(ref_pow_use_1.loc[:, '2019':])].reset_index(drop = True)

    ref_pow_use_2 = ref_pow_use_2[ref_pow_use_2['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    ref_pow_use_2 = ref_pow_use_2.groupby('FUEL').sum().reset_index()
    ref_pow_use_2['Transformation'] = 'Input fuel'

    #################################################################################
    historical_input = EGEDA_hist_power[EGEDA_hist_power['economy'] == economy].copy().\
        iloc[:,:][['FUEL', 'Transformation'] + list(EGEDA_hist_power.loc[:, '2000':'2018'])]

    ref_pow_use_2 = historical_input.merge(ref_pow_use_2, how = 'right', on = ['FUEL', 'Transformation']).replace(np.nan, 0)

    ref_pow_use_2 = ref_pow_use_2[['FUEL', 'Transformation'] + list(ref_pow_use_2.loc[:, '2000':'2050'])]

    ref_pow_use_2.loc['Total'] = ref_pow_use_2.sum(numeric_only = True)

    ref_pow_use_2.loc['Total', 'FUEL'] = 'Total'
    ref_pow_use_2.loc['Total', 'Transformation'] = 'Input fuel'

    ref_pow_use_2['FUEL'] = pd.Categorical(ref_pow_use_2['FUEL'], use_agg_fuels_1)

    ref_pow_use_2 = ref_pow_use_2.sort_values('FUEL').reset_index(drop = True)

    # Get rid of zero rows
    # non_zero = (ref_pow_use_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_pow_use_2 = ref_pow_use_2.loc[non_zero].reset_index(drop = True)

    ref_pow_use_2_rows = ref_pow_use_2.shape[0]
    ref_pow_use_2_cols = ref_pow_use_2.shape[1]

    ref_pow_use_3 = ref_pow_use_2[['FUEL', 'Transformation'] + col_chart_years]

    ref_pow_use_3_rows = ref_pow_use_3.shape[0]
    ref_pow_use_3_cols = ref_pow_use_3.shape[1]

    # Use by fuel data frame number 1

    ref_pow_use_4 = ref_pow_use_1.append([coal2, oil, gas, nuclear, renew2, other])\
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
    coal_ccs_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(coal_ccs_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal CCS')
    oil_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    gas_ccs_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(gas_ccs_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas CCS')
    storage_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    # chp_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Cogeneration')
    nuclear_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Bio')
    other_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    geo_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(geo_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Geothermal')
    misc = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')
    waste_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(waste_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Waste')

    coal_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(thermal_coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(lignite_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    roof_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(solar_roof_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar roof')
    nonroof_pp = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(solar_nr_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')

    # New aggregations from ALEX
    other_pp2 = ref_elecgen_1[ref_elecgen_1['TECHNOLOGY'].isin(other_higheragg_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_elecgen_2 = ref_elecgen_1.append([coal_pp2, coal_ccs_pp, lignite_pp2, oil_pp, gas_pp, gas_ccs_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, waste_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + list(ref_elecgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                                                                    

    ref_elecgen_2['Generation'] = 'Electricity'
    ref_elecgen_2 = ref_elecgen_2[['TECHNOLOGY', 'Generation'] + list(ref_elecgen_2.loc[:, '2019':'2050'])] 

    ref_elecgen_2 = ref_elecgen_2[ref_elecgen_2['TECHNOLOGY'].isin(prod_agg_tech2)].\
        set_index('TECHNOLOGY')

    ref_elecgen_2 = ref_elecgen_2.loc[ref_elecgen_2.index.intersection(prod_agg_tech2)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    ref_elecgen_NEW = ref_elecgen_2[ref_elecgen_2['TECHNOLOGY'].isin(['Coal CCS', 'Gas', 'Gas CCS',\
        'Oil', 'Nuclear', 'Hydro', 'Wind', 'Solar', 'Imports'])].copy().append([coal_pp, other_pp2]).reset_index(drop = True)

    ref_elecgen_NEW = ref_elecgen_NEW[['TECHNOLOGY', 'Generation'] + list(ref_elecgen_NEW.loc[:, '2019':'2050'])]

    ref_elecgen_NEW.loc[ref_elecgen_NEW['TECHNOLOGY'] == 'Coal', 'Generation'] = 'Electricity'
    ref_elecgen_NEW.loc[ref_elecgen_NEW['TECHNOLOGY'] == 'Other', 'Generation'] = 'Electricity'

    #################################################################################
    historical_gen = EGEDA_hist_gen[EGEDA_hist_gen['economy'] == economy].copy().\
        iloc[:,:][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_gen.loc[:, '2000':'2018'])]

    ref_elecgen_2 = historical_gen.merge(ref_elecgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    ref_elecgen_2['TECHNOLOGY'] = pd.Categorical(ref_elecgen_2['TECHNOLOGY'], prod_agg_tech2)

    ref_elecgen_2 = ref_elecgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = ref_elecgen_2.select_dtypes(include=[np.number]) / 3.6 
    ref_elecgen_2[s.columns] = s

    ref_elecgen_2.loc['Total'] = ref_elecgen_2.sum(numeric_only = True)

    ref_elecgen_2.loc['Total', 'TECHNOLOGY'] = 'Total'
    ref_elecgen_2.loc['Total', 'Generation'] = 'Electricity'

    # Get rid of zero rows
    # non_zero = (ref_elecgen_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_elecgen_2 = ref_elecgen_2.loc[non_zero].reset_index(drop = True)

    ref_elecgen_2_rows = ref_elecgen_2.shape[0]
    ref_elecgen_2_cols = ref_elecgen_2.shape[1]

    ref_elecgen_3 = ref_elecgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    ref_elecgen_3_rows = ref_elecgen_3.shape[0]
    ref_elecgen_3_cols = ref_elecgen_3.shape[1]

    # And now for new aggregations
    #################################################################################
    historical_gen2 = EGEDA_hist_gen2[EGEDA_hist_gen2['economy'] == economy].copy().\
        iloc[:,:][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_gen2.loc[:, '2000':'2018'])]

    ref_elecgen_4 = historical_gen2.merge(ref_elecgen_NEW, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    ref_elecgen_4['TECHNOLOGY'] = pd.Categorical(ref_elecgen_4['TECHNOLOGY'], prod_agg_tech3)

    ref_elecgen_4 = ref_elecgen_4.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = ref_elecgen_4.select_dtypes(include=[np.number]) / 3.6 
    ref_elecgen_4[s.columns] = s

    ref_elecgen_4.loc['Total'] = ref_elecgen_4.sum(numeric_only = True)

    ref_elecgen_4.loc['Total', 'TECHNOLOGY'] = 'Total'
    ref_elecgen_4.loc['Total', 'Generation'] = 'Electricity'

    # Get rid of zero rows
    non_zero = (ref_elecgen_4.loc[:,'2000':] != 0).any(axis = 1)
    ref_elecgen_4 = ref_elecgen_4.loc[non_zero].reset_index(drop = True)

    ref_elecgen_4_rows = ref_elecgen_4.shape[0]
    ref_elecgen_4_cols = ref_elecgen_4.shape[1]

    ref_elecgen_5 = ref_elecgen_4[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    ref_elecgen_5_rows = ref_elecgen_5.shape[0]
    ref_elecgen_5_cols = ref_elecgen_5.shape[1]


    ##################################################################################################################################################################

    # Now create some refinery dataframes

    ref_refinery_1 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    ref_refinery_1['Transformation'] = 'Input to refinery'
    ref_refinery_1 = ref_refinery_1[['FUEL', 'Transformation'] + list(ref_refinery_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_refinery_1.loc[ref_refinery_1['FUEL'] == 'd_ref_6_1_crude_oil', 'FUEL'] = 'Crude oil'
    ref_refinery_1.loc[ref_refinery_1['FUEL'] == 'd_ref_6_x_ngls', 'FUEL'] = 'NGLs'

    ref_refinery_1.loc['Total'] = ref_refinery_1.sum(numeric_only = True)

    ref_refinery_1.loc['Total', 'FUEL'] = 'Total'
    ref_refinery_1.loc['Total', 'Transformation'] = 'Input to refinery'

    # # Get rid of zero rows
    # non_zero = (ref_refinery_1.loc[:,'2017':] != 0).any(axis = 1)
    # ref_refinery_1 = ref_refinery_1.loc[non_zero].reset_index(drop = True)

    ref_refinery_1_rows = ref_refinery_1.shape[0]
    ref_refinery_1_cols = ref_refinery_1.shape[1]

    ref_refinery_2 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                 (ref_refownsup_df1['Sector'] == 'REF') & 
                                 (ref_refownsup_df1['FUEL'].isin(refinery_output))].copy()

    ref_refinery_2['Transformation'] = 'Output from refinery'
    ref_refinery_2 = ref_refinery_2[['FUEL', 'Transformation'] + list(ref_refinery_2.loc[:, '2017':'2050'])].reset_index(drop = True)

    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_1_motor_gasoline_refine', 'FUEL'] = 'Motor gasoline'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_2_aviation_gasoline_refine', 'FUEL'] = 'Aviation gasoline'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_3_naphtha_refine', 'FUEL'] = 'Naphtha'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_x_jet_fuel_refine', 'FUEL'] = 'Jet fuel'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_6_kerosene_refine', 'FUEL'] = 'Other kerosene'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_7_gas_diesel_oil_refine', 'FUEL'] = 'Gas diesel oil'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_8_fuel_oil_refine', 'FUEL'] = 'Fuel oil'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_9_lpg_refine', 'FUEL'] = 'LPG'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_10_refinery_gas_not_liquefied_refine', 'FUEL'] = 'Refinery gas'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_11_ethane_refine', 'FUEL'] = 'Ethane'
    ref_refinery_2.loc[ref_refinery_2['FUEL'] == 'd_ref_7_x_other_petroleum_products_refine', 'FUEL'] = 'Other'

    ref_refinery_2['FUEL'] = pd.Categorical(
        ref_refinery_2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 
                      'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other', 'Total'], 
        ordered = True)

    ref_refinery_2 = ref_refinery_2.sort_values('FUEL')

    ref_refinery_2.loc['Total'] = ref_refinery_2.sum(numeric_only = True)

    ref_refinery_2.loc['Total', 'FUEL'] = 'Total'
    ref_refinery_2.loc['Total', 'Transformation'] = 'Output from refinery'

    # Get rid of zero rows
    non_zero = (ref_refinery_2.loc[:,'2017':] != 0).any(axis = 1)
    ref_refinery_2 = ref_refinery_2.loc[non_zero].reset_index(drop = True)

    ref_refinery_2_rows = ref_refinery_2.shape[0]
    ref_refinery_2_cols = ref_refinery_2.shape[1]

    ref_refinery_3 = ref_refinery_2[['FUEL', 'Transformation'] + trans_col_chart]

    ref_refinery_3_rows = ref_refinery_3.shape[0]
    ref_refinery_3_cols = ref_refinery_3.shape[1]

    ##################################################################
    # Hydrogen output (similar to refinery output)

    ref_hydrogen_1 = ref_refownsup_df1[(ref_refownsup_df1['economy'] == economy) &
                                         (ref_refownsup_df1['Sector'] == 'HYD') & 
                                         (ref_refownsup_df1['FUEL'].isin(['16_x_hydrogen', '16_x_hydrogen_exports']))].copy()

    ref_hydrogen_1 = ref_hydrogen_1[['FUEL', 'TECHNOLOGY'] + list(ref_hydrogen_1.loc[:, '2018':'2050'])]\
        .rename(columns = {'FUEL': 'Fuel', 'TECHNOLOGY': 'Technology'}).reset_index(drop = True)

    ref_hydrogen_1.loc[ref_hydrogen_1['Fuel'] == '16_x_hydrogen', 'Fuel'] = 'Hydrogen'
    ref_hydrogen_1.loc[ref_hydrogen_1['Fuel'] == '16_x_hydrogen_exports', 'Fuel'] = 'Hydrogen'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_ng_smr', 'Technology'] = 'Steam methane reforming'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_ng_smr_ccs', 'Technology'] = 'Steam methane reforming CCS'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_coal_gas_ccs', 'Technology'] = 'Coal gasification CCS'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_pem_elyzer', 'Technology'] = 'Electrolysis'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_ng_smr_export', 'Technology'] = 'Steam methane reforming'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_ng_smr_ccs_export', 'Technology'] = 'Steam methane reforming CCS'
    ref_hydrogen_1.loc[ref_hydrogen_1['Technology'] == 'HYD_pem_elyzer_export', 'Technology'] = 'Electrolysis'

    ref_hydrogen_1 = ref_hydrogen_1.groupby(['Fuel', 'Technology']).sum().reset_index()

    # Hydrogen trade
    ref_hydrogen_trade_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                                (EGEDA_years_reference['fuel_code'] == '16_x_hydrogen') &
                                                (EGEDA_years_reference['item_code_new'].isin(['2_imports', '3_exports',\
                                                    '4_international_marine_bunkers', '5_international_aviation_bunkers']))]\
                                                        .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_hydrogen_trade_1 = ref_hydrogen_trade_1[['fuel_code', 'item_code_new'] + list(ref_hydrogen_trade_1.loc[:, '2018': '2050'])]\
        .rename(columns = {'fuel_code': 'Fuel', 'item_code_new': 'Technology'}).reset_index(drop = True)

    ref_hydrogen_trade_1.loc[ref_hydrogen_trade_1['Fuel'] == '16_x_hydrogen', 'Fuel'] = 'Hydrogen'
    ref_hydrogen_trade_1.loc[ref_hydrogen_trade_1['Technology'] == '2_imports', 'Technology'] = 'Imports'
    ref_hydrogen_trade_1.loc[ref_hydrogen_trade_1['Technology'] == '3_exports', 'Technology'] = 'Exports'
    ref_hydrogen_trade_1.loc[ref_hydrogen_trade_1['Technology'] == '4_international_marine_bunkers', 'Technology'] = 'Bunkers'
    ref_hydrogen_trade_1.loc[ref_hydrogen_trade_1['Technology'] == '5_international_aviation_bunkers', 'Technology'] = 'Bunkers'

    ref_hydrogen_trade_1 = ref_hydrogen_trade_1.copy().groupby(['Fuel', 'Technology']).sum().reset_index()

    ref_hydrogen_2 = ref_hydrogen_1.append(ref_hydrogen_trade_1).copy().reset_index(drop = True)

    ref_hydrogen_2['Technology'] = pd.Categorical(
        ref_hydrogen_2['Technology'], 
        categories = ['Steam methane reforming', 'Steam methane reforming CCS', 'Coal gasification CCS', 'Electrolysis', 'Imports', 'Exports', 'Bunkers'], 
        ordered = True)

    ref_hydrogen_2 = ref_hydrogen_2.sort_values('Technology')

    # Get rid of zero rows
    non_zero = (ref_hydrogen_2.loc[:,'2018':] != 0).any(axis = 1)
    ref_hydrogen_2 = ref_hydrogen_2.loc[non_zero].reset_index(drop = True)

    ref_hydrogen_2_rows = ref_hydrogen_2.shape[0]
    ref_hydrogen_2_cols = ref_hydrogen_2.shape[1]

    ref_hydrogen_3 = ref_hydrogen_2[['Fuel', 'Technology'] + trans_col_chart].reset_index(drop = True)

    ref_hydrogen_3_rows = ref_hydrogen_3.shape[0]
    ref_hydrogen_3_cols = ref_hydrogen_3.shape[1]

    #######################################

    # Reference hydrogen use

    ref_hyd_use_1 = ref_osemo_1[(ref_osemo_1['REGION'] == economy) &
                                (ref_osemo_1['TECHNOLOGY'].str.startswith('HYD'))].copy().reset_index(drop = True)

    hyd_coal = ref_hyd_use_1[ref_hyd_use_1['FUEL'].isin(['1_1_coking_coal'])].groupby(['REGION'])\
        .sum().assign(TECHNOLOGY = 'Input fuel', FUEL = 'Coal')
    
    hyd_gas = ref_hyd_use_1[ref_hyd_use_1['FUEL'].isin(['8_1_natural_gas'])].groupby(['REGION'])\
        .sum().assign(TECHNOLOGY = 'Input fuel', FUEL = 'Gas')
    
    hyd_elec = ref_hyd_use_1[ref_hyd_use_1['FUEL'].isin(['17_electricity_h2', '17_electricity_green'])]\
        .groupby(['REGION']).sum().assign(TECHNOLOGY = 'Input fuel', FUEL = 'Electricity')

    # Now append coal, gas and electricity to dataframe    
    ref_hyd_use_1 = ref_hyd_use_1.append([hyd_coal, hyd_gas, hyd_elec])[['FUEL', 'TECHNOLOGY'] + list(ref_hyd_use_1.loc[:,'2018':'2050'])]\
        .reset_index(drop = True)

    ref_hyd_use_1 = ref_hyd_use_1[ref_hyd_use_1['FUEL'].isin(['Coal', 'Gas', 'Electricity'])].reset_index(drop = True)

    ref_hyd_use_1.loc['Total'] = ref_hyd_use_1.sum(numeric_only = True)

    ref_hyd_use_1.loc['Total', 'FUEL'] = 'Total'
    ref_hyd_use_1.loc['Total', 'TECHNOLOGY'] = 'Input fuel'

    # Get rid of zero rows
    non_zero = (ref_hyd_use_1.loc[:,'2018':] != 0).any(axis = 1)
    ref_hyd_use_1 = ref_hyd_use_1.loc[non_zero].reset_index(drop = True)

    ref_hyd_use_1_rows = ref_hyd_use_1.shape[0]
    ref_hyd_use_1_cols = ref_hyd_use_1.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    ref_powcap_1 = ref_pow_capacity_df1[ref_pow_capacity_df1['REGION'] == economy]

    coal_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    coal_ccs_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(coal_ccs_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal CCS')
    oil_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    gas_ccs_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(gas_ccs_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas CCS')
    hydro_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Bio')
    geo_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(geo_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Geothermal')
    #chp_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(chp_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Cogeneration')
    other_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')
    waste_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(waste_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Waste')

    lignite_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(lignite_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Lignite')
    thermal_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(thermal_coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')

    other2_capacity = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(other_higheragg_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')

    # Capacity by tech dataframe (with the above aggregations added)

    ref_powcap_1 = ref_powcap_1.append([thermal_capacity, coal_ccs_capacity, lignite_capacity, gas_capacity, gas_ccs_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, waste_capacity, other_capacity])\
        [['TECHNOLOGY'] + list(ref_powcap_1.loc[:, '2018':'2050'])].reset_index(drop = True) 

    ref_powcap_1 = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(pow_capacity_agg2)].reset_index(drop = True)

    ref_powcap_1['TECHNOLOGY'] = pd.Categorical(ref_powcap_1['TECHNOLOGY'], pow_capacity_agg2)

    ref_powcap_1 = ref_powcap_1.sort_values('TECHNOLOGY').reset_index(drop = True)

    ref_powcap_1.loc['Total'] = ref_powcap_1.sum(numeric_only = True)

    ref_powcap_1.loc['Total', 'TECHNOLOGY'] = 'Total'

    ref_powcap_NEW = ref_powcap_1[ref_powcap_1['TECHNOLOGY'].isin(['Coal', 'Coal CCS', 'Oil', 'Gas', 'Gas CCS',\
        'Nuclear', 'Hydro', 'Wind', 'Solar'])].copy().append([other2_capacity]).reset_index(drop = True)

    ref_powcap_NEW = ref_powcap_NEW[['TECHNOLOGY'] + list(ref_powcap_NEW.loc[:, '2018':'2050'])]

    ref_powcap_NEW.loc['Total'] = ref_powcap_NEW.sum(numeric_only = True)

    ref_powcap_NEW.loc['Total', 'TECHNOLOGY'] = 'Total'

    ref_powcap_NEW['TECHNOLOGY'] = pd.Categorical(ref_powcap_NEW['TECHNOLOGY'], pow_capacity_agg3)

    ref_powcap_3 = ref_powcap_NEW.sort_values('TECHNOLOGY').reset_index(drop = True)

    # Get rid of zero rows
    # non_zero = (ref_powcap_1.loc[:,'2018':] != 0).any(axis = 1)
    # ref_powcap_1 = ref_powcap_1.loc[non_zero].reset_index(drop = True)

    ref_powcap_1_rows = ref_powcap_1.shape[0]
    ref_powcap_1_cols = ref_powcap_1.shape[1]

    ref_powcap_2 = ref_powcap_1[['TECHNOLOGY'] + trans_col_chart]

    ref_powcap_2_rows = ref_powcap_2.shape[0]
    ref_powcap_2_cols = ref_powcap_2.shape[1]

    # Get rid of zero rows
    non_zero = (ref_powcap_3.loc[:,'2018':] != 0).any(axis = 1)
    ref_powcap_3 = ref_powcap_3.loc[non_zero].reset_index(drop = True)

    ref_powcap_3_rows = ref_powcap_3.shape[0]
    ref_powcap_3_cols = ref_powcap_3.shape[1]

    ref_powcap_4 = ref_powcap_3[['TECHNOLOGY'] + trans_col_chart]

    ref_powcap_4_rows = ref_powcap_4.shape[0]
    ref_powcap_4_cols = ref_powcap_4.shape[1]


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

    # Get rid of zero rows
    non_zero = (ref_trans_3.loc[:,'2017':] != 0).any(axis = 1)
    ref_trans_3 = ref_trans_3.loc[non_zero].reset_index(drop = True)

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
    waste_own = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(waste_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Waste', Sector = 'Own-use and losses')

    ref_ownuse_1 = ref_ownuse_1.append([coal_own, oil_own, gas_own, renewables_own, elec_own, heat_own, waste_own])\
        [['FUEL', 'Sector'] + list(ref_ownuse_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    # SPECIAL GRAB: Own-use for coal report #########################

    ref_owncoal_1 = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(['1_1_coking_coal', '1_5_lignite', '1_x_coal_thermal', '2_coal_products'])]\
        .copy().reset_index(drop = True)

    ref_owncoal_1.loc[ref_owncoal_1['FUEL'] == '1_1_coking_coal', 'FUEL'] = 'Metallurgical coal'
    ref_owncoal_1.loc[ref_owncoal_1['FUEL'] == '1_5_lignite', 'FUEL'] = 'Lignite'
    ref_owncoal_1.loc[ref_owncoal_1['FUEL'] == '1_x_coal_thermal', 'FUEL'] = 'Thermal coal'
    ref_owncoal_1.loc[ref_owncoal_1['FUEL'] == '2_coal_products', 'FUEL'] = 'Metallurgical coal'

    ref_owncoal_1 = ref_owncoal_1.copy().groupby(['FUEL', 'Sector']).sum().reset_index()\
        [['FUEL', 'Sector'] + list(ref_owncoal_1.loc[:,'2019':])]

    hist_owncoal = EGEDA_hist_owncoal[EGEDA_hist_owncoal['economy'] == economy].copy()\
        .iloc[:,:][['FUEL', 'Sector'] + list(EGEDA_hist_owncoal.loc[:, '2000':'2018'])]

    ref_owncoal_1 = hist_owncoal.merge(ref_owncoal_1, how = 'right', on = ['FUEL', 'Sector']).replace(np.nan, 0)

    ref_owncoal_1 = ref_owncoal_1[['FUEL', 'Sector'] + list(ref_owncoal_1.loc[:, '2000':'2050'])]

    # Get rid of zero rows
    non_zero = (ref_owncoal_1.loc[:,'2000':] != 0).any(axis = 1)
    ref_owncoal_1 = ref_owncoal_1.loc[non_zero].reset_index(drop = True)

    ref_owncoal_1_rows = ref_owncoal_1.shape[0]
    ref_owncoal_1_cols = ref_owncoal_1.shape[1]

    #################################################################

    ref_ownuse_1 = ref_ownuse_1[ref_ownuse_1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)\
        [['FUEL', 'Sector'] + list(ref_ownuse_1.loc[:,'2019':])]

    #################################################################################
    historical_input = EGEDA_hist_own[EGEDA_hist_own['economy'] == economy].copy().\
        iloc[:,:][['FUEL', 'Sector'] + list(EGEDA_hist_own.loc[:, '2000':'2018'])]

    ref_ownuse_1 = historical_input.merge(ref_ownuse_1, how = 'right', on = ['FUEL', 'Sector']).replace(np.nan, 0)

    ref_ownuse_1['FUEL'] = pd.Categorical(ref_ownuse_1['FUEL'], own_use_fuels)

    ref_ownuse_1 = ref_ownuse_1.sort_values('FUEL').reset_index(drop = True)

    ref_ownuse_1 = ref_ownuse_1[['FUEL', 'Sector'] + list(ref_ownuse_1.loc[:, '2000':'2050'])]

    ref_ownuse_1.loc['Total'] = ref_ownuse_1.sum(numeric_only = True)

    ref_ownuse_1.loc['Total', 'FUEL'] = 'Total'
    ref_ownuse_1.loc['Total', 'Sector'] = 'Own-use and losses'

    # Get rid of zero rows
    # non_zero = (ref_ownuse_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_ownuse_1 = ref_ownuse_1.loc[non_zero].reset_index(drop = True)

    ref_ownuse_1_rows = ref_ownuse_1.shape[0]
    ref_ownuse_1_cols = ref_ownuse_1.shape[1]

    ref_ownuse_2 = ref_ownuse_1[['FUEL', 'Sector'] + col_chart_years]

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
    gas_ccs_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(gas_ccs_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas CCS')
    nuclear_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(nuke_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(bio_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    waste_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(waste_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Waste')
    comb_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(combination_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    nons_hp = ref_heatgen_1[ref_heatgen_1['TECHNOLOGY'].isin(nons_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Non-specified')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    ref_heatgen_2 = ref_heatgen_1.append([coal_hp, lignite_hp, oil_hp, gas_hp, gas_ccs_hp, nuclear_hp, bio_hp, waste_hp, comb_hp, nons_hp])\
        [['TECHNOLOGY'] + list(ref_heatgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                                                                    

    ref_heatgen_2['Generation'] = 'Heat'
    ref_heatgen_2 = ref_heatgen_2[['TECHNOLOGY', 'Generation'] + list(ref_heatgen_2.loc[:, '2019':'2050'])] 

    # # Insert 0 other row
    # new_row_zero = ['Gas CCS', 'Heat'] + [0] * 34
    # new_series = pd.Series(new_row_zero, index = ref_heatgen_2.columns)

    # ref_heatgen_2 = ref_heatgen_2.append(new_series, ignore_index = True).reset_index(drop = True)

    ref_heatgen_2 = ref_heatgen_2[ref_heatgen_2['TECHNOLOGY'].isin(heat_prod_tech)].\
        set_index('TECHNOLOGY')

    ref_heatgen_2 = ref_heatgen_2.loc[ref_heatgen_2.index.intersection(heat_prod_tech)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_heat[EGEDA_hist_heat['economy'] == economy].copy().\
        iloc[:,:][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_heat.loc[:, '2000':'2018'])]

    ref_heatgen_2 = historical_gen.merge(ref_heatgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    ref_heatgen_2['TECHNOLOGY'] = pd.Categorical(ref_heatgen_2['TECHNOLOGY'], heat_prod_tech)

    ref_heatgen_2 = ref_heatgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    ref_heatgen_2.loc['Total'] = ref_heatgen_2.sum(numeric_only = True)

    ref_heatgen_2.loc['Total', 'TECHNOLOGY'] = 'Total'
    ref_heatgen_2.loc['Total', 'Generation'] = 'Heat'

    # # Get rid of zero rows
    # non_zero = (ref_heatgen_2.loc[:,'2000':] != 0).any(axis = 1)
    # ref_heatgen_2 = ref_heatgen_2.loc[non_zero].reset_index(drop = True)

    ref_heatgen_2_rows = ref_heatgen_2.shape[0]
    ref_heatgen_2_cols = ref_heatgen_2.shape[1]

    ref_heatgen_3 = ref_heatgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    ref_heatgen_3_rows = ref_heatgen_3.shape[0]
    ref_heatgen_3_cols = ref_heatgen_3.shape[1]

    ################################################################################

    # Heat use dataframes

    # REFERENCE

    ref_heat_use_1 = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                   (ref_power_df1['Sheet_energy'] == 'UseByTechnology') &
                                   (ref_power_df1['TECHNOLOGY'].isin(heat_only))].reset_index(drop = True)

    coal = ref_heat_use_1[ref_heat_use_1['FUEL'].isin(coal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                                      TECHNOLOGY = 'Coal heat')

    lignite = ref_heat_use_1[ref_heat_use_1['FUEL'].isin(lignite_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Lignite',
                                                                                              TECHNOLOGY = 'Lignite heat')                                                                                      

    oil = ref_heat_use_1[ref_heat_use_1['FUEL'].isin(oil_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Oil heat')

    gas = ref_heat_use_1[ref_heat_use_1['FUEL'].isin(gas_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Gas',
                                                                                      TECHNOLOGY = 'Gas heat')

    biomass = ref_heat_use_1[ref_heat_use_1['FUEL'].isin(biomass_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Biomass',
                                                                                                        TECHNOLOGY = 'Biomass heat')

    waste = ref_heat_use_1[ref_heat_use_1['FUEL'].isin(waste_fuel)].groupby(['economy']).sum().assign(FUEL = 'Waste',
                                                                                               TECHNOLOGY = 'Waste heat')

    ref_heat_use_2 = ref_heat_use_1.append([coal, lignite, oil, gas, biomass, waste])\
        [['FUEL', 'TECHNOLOGY'] + list(ref_heat_use_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    ref_heat_use_2 = ref_heat_use_2[ref_heat_use_2['FUEL'].isin(heat_agg_fuels)].copy().set_index('FUEL').reset_index()

    ref_heat_use_2 = ref_heat_use_2.groupby('FUEL').sum().reset_index()
    ref_heat_use_2['Transformation'] = 'Heat plant input fuel'
    ref_heat_use_2['FUEL'] = pd.Categorical(ref_heat_use_2['FUEL'], heat_agg_fuels)

    ref_heat_use_2 = ref_heat_use_2.sort_values('FUEL').reset_index(drop = True)

    ref_heat_use_2 = ref_heat_use_2[['FUEL', 'Transformation'] + list(ref_heat_use_2.loc[:,'2017':'2050'])]

    ref_heat_use_2.loc['Total'] = ref_heat_use_2.sum(numeric_only = True)

    ref_heat_use_2.loc['Total', 'FUEL'] = 'Total'
    ref_heat_use_2.loc['Total', 'Transformation'] = 'Heat plant input fuel'

    # Get rid of zero rows
    non_zero = (ref_heat_use_2.loc[:,'2017':] != 0).any(axis = 1)
    ref_heat_use_2 = ref_heat_use_2.loc[non_zero].reset_index(drop = True)

    ref_heat_use_2_rows = ref_heat_use_2.shape[0]
    ref_heat_use_2_cols = ref_heat_use_2.shape[1]

    ref_heat_use_3 = ref_heat_use_2[['FUEL', 'Transformation'] + trans_col_chart]

    ref_heat_use_3_rows = ref_heat_use_3.shape[0]
    ref_heat_use_3_cols = ref_heat_use_3.shape[1]

    ######################################################################################################################
    
    # CARBON NEUTRALITY dataframes

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

    # imports = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(imports_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Imports',
    #                                                                                     TECHNOLOGY = 'Electricity imports')                                                                                         

    # Second level aggregations

    coal2 = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(coal_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                      TECHNOLOGY = 'Coal power')

    renew2 = netz_pow_use_1[netz_pow_use_1['FUEL'].isin(renewables_fuel_2)].groupby(['economy']).sum().assign(FUEL = 'Renewables',
                                                                                      TECHNOLOGY = 'Renewables power')

    # Use by fuel data frame number 1

    netz_pow_use_2 = netz_pow_use_1.append([coal, lignite, oil, gas, nuclear, hydro, solar, wind, geothermal, biomass, other_renew, other])\
        [['FUEL', 'TECHNOLOGY'] + list(netz_pow_use_1.loc[:,'2019':'2050'])].reset_index(drop = True)

    netz_pow_use_2 = netz_pow_use_2[netz_pow_use_2['FUEL'].isin(use_agg_fuels_1)].copy().set_index('FUEL').reset_index()

    netz_pow_use_2 = netz_pow_use_2.groupby('FUEL').sum().reset_index()
    netz_pow_use_2['Transformation'] = 'Input fuel'
    
    #################################################################################
    historical_input = EGEDA_hist_power[EGEDA_hist_power['economy'] == economy].copy().\
        iloc[:,:][['FUEL', 'Transformation'] + list(EGEDA_hist_power.loc[:, '2000':'2018'])]

    netz_pow_use_2 = historical_input.merge(netz_pow_use_2, how = 'right', on = ['FUEL', 'Transformation']).replace(np.nan, 0)

    netz_pow_use_2['FUEL'] = pd.Categorical(netz_pow_use_2['FUEL'], use_agg_fuels_1)

    netz_pow_use_2 = netz_pow_use_2.sort_values('FUEL').reset_index(drop = True)

    netz_pow_use_2 = netz_pow_use_2[['FUEL', 'Transformation'] + list(netz_pow_use_2.loc[:, '2000':'2050'])]

    netz_pow_use_2.loc['Total'] = netz_pow_use_2.sum(numeric_only = True)

    netz_pow_use_2.loc['Total', 'FUEL'] = 'Total'
    netz_pow_use_2.loc['Total', 'Transformation'] = 'Input fuel'

    # Get rid of zero rows
    # non_zero = (netz_pow_use_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_pow_use_2 = netz_pow_use_2.loc[non_zero].reset_index(drop = True)

    netz_pow_use_2_rows = netz_pow_use_2.shape[0]
    netz_pow_use_2_cols = netz_pow_use_2.shape[1]

    netz_pow_use_3 = netz_pow_use_2[['FUEL', 'Transformation'] + col_chart_years]

    netz_pow_use_3_rows = netz_pow_use_3.shape[0]
    netz_pow_use_3_cols = netz_pow_use_3.shape[1]

    # Use by fuel data frame number 1

    netz_pow_use_4 = netz_pow_use_1.append([coal2, oil, gas, nuclear, renew2, other])\
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
    coal_ccs_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(coal_ccs_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal CCS')
    oil_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(oil_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Oil')
    gas_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(gas_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas')
    gas_ccs_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(gas_ccs_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas CCS')
    storage_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(storage_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Storage')
    # chp_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(chp_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Cogeneration')
    nuclear_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(nuclear_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(bio_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Bio')
    other_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(other_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    hydro_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(hydro_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Hydro')
    geo_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(geo_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Geothermal')
    misc = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(im_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Imports')
    solar_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(solar_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')
    wind_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(wind_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Wind')
    waste_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(waste_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Waste')

    coal_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(thermal_coal_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Coal')
    lignite_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(lignite_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Lignite')
    roof_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(solar_roof_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar roof')
    nonroof_pp = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(solar_nr_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Solar')

    # New aggregations from ALEX
    other_pp2 = netz_elecgen_1[netz_elecgen_1['TECHNOLOGY'].isin(other_higheragg_tech)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_elecgen_2 = netz_elecgen_1.append([coal_pp2, coal_ccs_pp, lignite_pp2, oil_pp, gas_pp, gas_ccs_pp, storage_pp, nuclear_pp,\
        bio_pp, geo_pp, waste_pp, other_pp, hydro_pp, misc, solar_pp, wind_pp])\
        [['TECHNOLOGY'] + list(netz_elecgen_1.loc[:,'2017':'2050'])].reset_index(drop = True)                                                                                                    

    netz_elecgen_2['Generation'] = 'Electricity'
    netz_elecgen_2 = netz_elecgen_2[['TECHNOLOGY', 'Generation'] + list(netz_elecgen_2.loc[:,'2019':'2050'])] 

    netz_elecgen_2 = netz_elecgen_2[netz_elecgen_2['TECHNOLOGY'].isin(prod_agg_tech2)].\
        set_index('TECHNOLOGY')

    netz_elecgen_2 = netz_elecgen_2.loc[netz_elecgen_2.index.intersection(prod_agg_tech2)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    netz_elecgen_NEW = netz_elecgen_2[netz_elecgen_2['TECHNOLOGY'].isin(['Coal CCS', 'Gas', 'Gas CCS',\
        'Oil', 'Nuclear', 'Hydro', 'Wind', 'Solar', 'Imports'])].copy().append([coal_pp, other_pp2]).reset_index(drop = True)

    netz_elecgen_NEW = netz_elecgen_NEW[['TECHNOLOGY', 'Generation'] + list(netz_elecgen_NEW.loc[:, '2019':'2050'])]

    netz_elecgen_NEW.loc[netz_elecgen_NEW['TECHNOLOGY'] == 'Coal', 'Generation'] = 'Electricity'
    netz_elecgen_NEW.loc[netz_elecgen_NEW['TECHNOLOGY'] == 'Other', 'Generation'] = 'Electricity'

    #################################################################################
    historical_gen = EGEDA_hist_gen[EGEDA_hist_gen['economy'] == economy].copy().\
        iloc[:,:][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_gen.loc[:,'2000':'2018'])]

    netz_elecgen_2 = historical_gen.merge(netz_elecgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    netz_elecgen_2['TECHNOLOGY'] = pd.Categorical(netz_elecgen_2['TECHNOLOGY'], prod_agg_tech2)

    netz_elecgen_2 = netz_elecgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = netz_elecgen_2.select_dtypes(include=[np.number]) / 3.6 
    netz_elecgen_2[s.columns] = s

    netz_elecgen_2.loc['Total'] = netz_elecgen_2.sum(numeric_only = True)

    netz_elecgen_2.loc['Total', 'TECHNOLOGY'] = 'Total'
    netz_elecgen_2.loc['Total', 'Generation'] = 'Electricity'

    # Get rid of zero rows
    # non_zero = (netz_elecgen_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_elecgen_2 = netz_elecgen_2.loc[non_zero].reset_index(drop = True)

    netz_elecgen_2_rows = netz_elecgen_2.shape[0]
    netz_elecgen_2_cols = netz_elecgen_2.shape[1]

    netz_elecgen_3 = netz_elecgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    netz_elecgen_3_rows = netz_elecgen_3.shape[0]
    netz_elecgen_3_cols = netz_elecgen_3.shape[1]

    # And now for new aggregations
    #################################################################################
    historical_gen2 = EGEDA_hist_gen2[EGEDA_hist_gen2['economy'] == economy].copy().\
        iloc[:,:][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_gen2.loc[:, '2000':'2018'])]

    netz_elecgen_4 = historical_gen2.merge(netz_elecgen_NEW, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    netz_elecgen_4['TECHNOLOGY'] = pd.Categorical(netz_elecgen_4['TECHNOLOGY'], prod_agg_tech3)

    netz_elecgen_4 = netz_elecgen_4.sort_values('TECHNOLOGY').reset_index(drop = True)

    # CHange to TWh from Petajoules

    s = netz_elecgen_4.select_dtypes(include=[np.number]) / 3.6 
    netz_elecgen_4[s.columns] = s

    netz_elecgen_4.loc['Total'] = netz_elecgen_4.sum(numeric_only = True)

    netz_elecgen_4.loc['Total', 'TECHNOLOGY'] = 'Total'
    netz_elecgen_4.loc['Total', 'Generation'] = 'Electricity'

    # Get rid of zero rows
    # non_zero = (netz_elecgen_4.loc[:,'2000':] != 0).any(axis = 1)
    # netz_elecgen_4 = netz_elecgen_4.loc[non_zero].reset_index(drop = True)

    netz_elecgen_4_rows = netz_elecgen_4.shape[0]
    netz_elecgen_4_cols = netz_elecgen_4.shape[1]

    netz_elecgen_5 = netz_elecgen_4[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    netz_elecgen_5_rows = netz_elecgen_5.shape[0]
    netz_elecgen_5_cols = netz_elecgen_5.shape[1]

    ##################################################################################################################################################################

    # Now create some refinery dataframes

    netz_refinery_1 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                 (netz_refownsup_df1['Sector'] == 'REF') & 
                                 (netz_refownsup_df1['FUEL'].isin(refinery_input))].copy()

    netz_refinery_1['Transformation'] = 'Input to refinery'
    netz_refinery_1 = netz_refinery_1[['FUEL', 'Transformation'] + list(netz_refinery_1.loc[:, '2017':'2050'])].reset_index(drop = True)

    netz_refinery_1.loc[netz_refinery_1['FUEL'] == 'd_ref_6_1_crude_oil', 'FUEL'] = 'Crude oil'
    netz_refinery_1.loc[netz_refinery_1['FUEL'] == 'd_ref_6_x_ngls', 'FUEL'] = 'NGLs'

    netz_refinery_1.loc['Total'] = netz_refinery_1.sum(numeric_only = True)

    netz_refinery_1.loc['Total', 'FUEL'] = 'Total'
    netz_refinery_1.loc['Total', 'Transformation'] = 'Input to refinery'

    # # Get rid of zero rows
    # non_zero = (netz_refinery_1.loc[:,'2017':] != 0).any(axis = 1)
    # netz_refinery_1 = netz_refinery_1.loc[non_zero].reset_index(drop = True)

    netz_refinery_1_rows = netz_refinery_1.shape[0]
    netz_refinery_1_cols = netz_refinery_1.shape[1]

    netz_refinery_2 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                 (netz_refownsup_df1['Sector'] == 'REF') & 
                                 (netz_refownsup_df1['FUEL'].isin(refinery_output))].copy()

    netz_refinery_2['Transformation'] = 'Output from refinery'
    netz_refinery_2 = netz_refinery_2[['FUEL', 'Transformation'] + list(netz_refinery_2.loc[:, '2017':'2050'])].reset_index(drop = True)

    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_1_motor_gasoline_refine', 'FUEL'] = 'Motor gasoline'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_2_aviation_gasoline_refine', 'FUEL'] = 'Aviation gasoline'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_3_naphtha_refine', 'FUEL'] = 'Naphtha'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_x_jet_fuel_refine', 'FUEL'] = 'Jet fuel'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_6_kerosene_refine', 'FUEL'] = 'Other kerosene'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_7_gas_diesel_oil_refine', 'FUEL'] = 'Gas diesel oil'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_8_fuel_oil_refine', 'FUEL'] = 'Fuel oil'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_9_lpg_refine', 'FUEL'] = 'LPG'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_10_refinery_gas_not_liquefied_refine', 'FUEL'] = 'Refinery gas'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_11_ethane_refine', 'FUEL'] = 'Ethane'
    netz_refinery_2.loc[netz_refinery_2['FUEL'] == 'd_ref_7_x_other_petroleum_products_refine', 'FUEL'] = 'Other'

    netz_refinery_2['FUEL'] = pd.Categorical(
        netz_refinery_2['FUEL'], 
        categories = ['Motor gasoline', 'Aviation gasoline', 'Naphtha', 'Jet fuel', 'Other kerosene', 
                      'Gas diesel oil', 'Fuel oil', 'LPG', 'Refinery gas', 'Ethane', 'Other', 'Total'], 
        ordered = True)

    netz_refinery_2 = netz_refinery_2.sort_values('FUEL')

    netz_refinery_2.loc['Total'] = netz_refinery_2.sum(numeric_only = True)

    netz_refinery_2.loc['Total', 'FUEL'] = 'Total'
    netz_refinery_2.loc['Total', 'Transformation'] = 'Output from refinery'

    # # Get rid of zero rows
    # non_zero = (netz_refinery_2.loc[:,'2017':] != 0).any(axis = 1)
    # netz_refinery_2 = netz_refinery_2.loc[non_zero].reset_index(drop = True)

    netz_refinery_2_rows = netz_refinery_2.shape[0]
    netz_refinery_2_cols = netz_refinery_2.shape[1]

    netz_refinery_3 = netz_refinery_2[['FUEL', 'Transformation'] + trans_col_chart]

    netz_refinery_3_rows = netz_refinery_3.shape[0]
    netz_refinery_3_cols = netz_refinery_3.shape[1]

    #############################################################

    # Hydrogen output (similar to refinery output)

    netz_hydrogen_1 = netz_refownsup_df1[(netz_refownsup_df1['economy'] == economy) &
                                         (netz_refownsup_df1['Sector'] == 'HYD') & 
                                         (netz_refownsup_df1['FUEL'].isin(['16_x_hydrogen', '16_x_hydrogen_exports']))].copy()

    netz_hydrogen_1 = netz_hydrogen_1[['FUEL', 'TECHNOLOGY'] + list(netz_hydrogen_1.loc[:, '2018':'2050'])]\
        .rename(columns = {'FUEL': 'Fuel', 'TECHNOLOGY': 'Technology'}).reset_index(drop = True)

    netz_hydrogen_1.loc[netz_hydrogen_1['Fuel'] == '16_x_hydrogen', 'Fuel'] = 'Hydrogen'
    netz_hydrogen_1.loc[netz_hydrogen_1['Fuel'] == '16_x_hydrogen_exports', 'Fuel'] = 'Hydrogen'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_ng_smr', 'Technology'] = 'Steam methane reforming'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_ng_smr_ccs', 'Technology'] = 'Steam methane reforming CCS'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_coal_gas_ccs', 'Technology'] = 'Coal gasification CCS'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_pem_elyzer', 'Technology'] = 'Electrolysis'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_ng_smr_export', 'Technology'] = 'Steam methane reforming'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_ng_smr_ccs_export', 'Technology'] = 'Steam methane reforming CCS'
    netz_hydrogen_1.loc[netz_hydrogen_1['Technology'] == 'HYD_pem_elyzer_export', 'Technology'] = 'Electrolysis'

    netz_hydrogen_1 = netz_hydrogen_1.groupby(['Fuel', 'Technology']).sum().reset_index()

    # Hydrogen trade
    netz_hydrogen_trade_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                                (EGEDA_years_netzero['fuel_code'] == '16_x_hydrogen') &
                                                (EGEDA_years_netzero['item_code_new'].isin(['2_imports', '3_exports',\
                                                    '4_international_marine_bunkers', '5_international_aviation_bunkers']))]\
                                                        .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_hydrogen_trade_1 = netz_hydrogen_trade_1[['fuel_code', 'item_code_new'] + list(netz_hydrogen_trade_1.loc[:, '2018': '2050'])]\
        .rename(columns = {'fuel_code': 'Fuel', 'item_code_new': 'Technology'}).reset_index(drop = True)

    netz_hydrogen_trade_1.loc[netz_hydrogen_trade_1['Fuel'] == '16_x_hydrogen', 'Fuel'] = 'Hydrogen'
    netz_hydrogen_trade_1.loc[netz_hydrogen_trade_1['Technology'] == '2_imports', 'Technology'] = 'Imports'
    netz_hydrogen_trade_1.loc[netz_hydrogen_trade_1['Technology'] == '3_exports', 'Technology'] = 'Exports'
    netz_hydrogen_trade_1.loc[netz_hydrogen_trade_1['Technology'] == '4_international_marine_bunkers', 'Technology'] = 'Bunkers'
    netz_hydrogen_trade_1.loc[netz_hydrogen_trade_1['Technology'] == '5_international_aviation_bunkers', 'Technology'] = 'Bunkers'

    netz_hydrogen_trade_1 = netz_hydrogen_trade_1.copy().groupby(['Fuel', 'Technology']).sum().reset_index()

    netz_hydrogen_2 = netz_hydrogen_1.append(netz_hydrogen_trade_1).copy().reset_index(drop = True)

    netz_hydrogen_2['Technology'] = pd.Categorical(
        netz_hydrogen_2['Technology'], 
        categories = ['Steam methane reforming', 'Steam methane reforming CCS', 'Coal gasification CCS', 'Electrolysis', 'Imports', 'Exports', 'Bunkers'], 
        ordered = True)

    netz_hydrogen_2 = netz_hydrogen_2.sort_values('Technology')

    # # Get rid of zero rows
    # non_zero = (netz_hydrogen_2.loc[:,'2018':] != 0).any(axis = 1)
    # netz_hydrogen_2 = netz_hydrogen_2.loc[non_zero].reset_index(drop = True)

    netz_hydrogen_2_rows = netz_hydrogen_2.shape[0]
    netz_hydrogen_2_cols = netz_hydrogen_2.shape[1]

    netz_hydrogen_3 = netz_hydrogen_2[['Fuel', 'Technology'] + trans_col_chart].reset_index(drop = True)

    netz_hydrogen_3_rows = netz_hydrogen_3.shape[0]
    netz_hydrogen_3_cols = netz_hydrogen_3.shape[1]

    # CARBON NEUTRALITY hydrogen use

    netz_hyd_use_1 = netz_osemo_1[(netz_osemo_1['REGION'] == economy) &
                                (netz_osemo_1['TECHNOLOGY'].str.startswith('HYD'))].copy().reset_index(drop = True)

    hyd_coal = netz_hyd_use_1[netz_hyd_use_1['FUEL'].isin(['1_1_coking_coal'])].groupby(['REGION'])\
        .sum().assign(TECHNOLOGY = 'Input fuel', FUEL = 'Coal')
    
    hyd_gas = netz_hyd_use_1[netz_hyd_use_1['FUEL'].isin(['8_1_natural_gas'])].groupby(['REGION'])\
        .sum().assign(TECHNOLOGY = 'Input fuel', FUEL = 'Gas')
    
    hyd_elec = netz_hyd_use_1[netz_hyd_use_1['FUEL'].isin(['17_electricity_h2', '17_electricity_green'])]\
        .groupby(['REGION']).sum().assign(TECHNOLOGY = 'Input fuel', FUEL = 'Electricity')

    # Now append coal, gas and electricity to dataframe    
    netz_hyd_use_1 = netz_hyd_use_1.append([hyd_coal, hyd_gas, hyd_elec])[['FUEL', 'TECHNOLOGY'] + list(netz_hyd_use_1.loc[:,'2018':'2050'])]\
        .reset_index(drop = True)

    netz_hyd_use_1 = netz_hyd_use_1[netz_hyd_use_1['FUEL'].isin(['Coal', 'Gas', 'Electricity'])].reset_index(drop = True)

    # # Get rid of zero rows
    # non_zero = (netz_hyd_use_1.loc[:,'2018':] != 0).any(axis = 1)
    # netz_hyd_use_1 = netz_hyd_use_1.loc[non_zero].reset_index(drop = True)

    netz_hyd_use_1_rows = netz_hyd_use_1.shape[0]
    netz_hyd_use_1_cols = netz_hyd_use_1.shape[1]

    #####################################################################################################################################################################

    # Create some power capacity dataframes

    netz_powcap_1 = netz_pow_capacity_df1[netz_pow_capacity_df1['REGION'] == economy]

    coal_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')
    coal_ccs_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(coal_ccs_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal CCS')
    oil_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(oil_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Oil')
    wind_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(wind_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Wind')
    storage_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(storage_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Storage')
    gas_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(gas_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas')
    gas_ccs_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(gas_ccs_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Gas CCS')
    hydro_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(hydro_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Hydro')
    solar_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(solar_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Solar')
    nuclear_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(nuclear_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(bio_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Bio')
    geo_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(geo_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Geothermal')
    #chp_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(chp_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Cogeneration')
    other_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(other_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')
    transmission = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(transmission_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Transmission')
    waste_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(waste_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Waste')

    lignite_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(lignite_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Lignite')
    thermal_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(thermal_coal_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Coal')

    other2_capacity = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(other_higheragg_cap)].groupby(['REGION']).sum().assign(TECHNOLOGY = 'Other')

    # Capacity by tech dataframe (with the above aggregations added)

    netz_powcap_1 = netz_powcap_1.append([thermal_capacity, coal_ccs_capacity, lignite_capacity, gas_capacity, gas_ccs_capacity, oil_capacity, nuclear_capacity,
                                            hydro_capacity, bio_capacity, wind_capacity, solar_capacity, 
                                            storage_capacity, geo_capacity, waste_capacity, other_capacity])\
        [['TECHNOLOGY'] + list(netz_powcap_1.loc[:,'2018':'2050'])].reset_index(drop = True) 

    netz_powcap_1 = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(pow_capacity_agg2)].reset_index(drop = True)

    netz_powcap_1['TECHNOLOGY'] = pd.Categorical(netz_powcap_1['TECHNOLOGY'], pow_capacity_agg2)

    netz_powcap_1 = netz_powcap_1.sort_values('TECHNOLOGY').reset_index(drop = True)

    netz_powcap_1.loc['Total'] = netz_powcap_1.sum(numeric_only = True)

    netz_powcap_1.loc['Total', 'TECHNOLOGY'] = 'Total'

    netz_powcap_NEW = netz_powcap_1[netz_powcap_1['TECHNOLOGY'].isin(['Coal', 'Coal CCS', 'Oil', 'Gas', 'Gas CCS',\
        'Nuclear', 'Hydro', 'Wind', 'Solar'])].copy().append([other2_capacity]).reset_index(drop = True)

    netz_powcap_NEW = netz_powcap_NEW[['TECHNOLOGY'] + list(netz_powcap_NEW.loc[:, '2018':'2050'])]

    netz_powcap_NEW.loc['Total'] = netz_powcap_NEW.sum(numeric_only = True)

    netz_powcap_NEW.loc['Total', 'TECHNOLOGY'] = 'Total'

    netz_powcap_NEW['TECHNOLOGY'] = pd.Categorical(netz_powcap_NEW['TECHNOLOGY'], pow_capacity_agg3)

    netz_powcap_3 = netz_powcap_NEW.sort_values('TECHNOLOGY').reset_index(drop = True)

    # Get rid of zero rows
    # non_zero = (netz_powcap_1.loc[:,'2018':] != 0).any(axis = 1)
    # netz_powcap_1 = netz_powcap_1.loc[non_zero].reset_index(drop = True)

    netz_powcap_1_rows = netz_powcap_1.shape[0]
    netz_powcap_1_cols = netz_powcap_1.shape[1]

    netz_powcap_2 = netz_powcap_1[['TECHNOLOGY'] + trans_col_chart]

    netz_powcap_2_rows = netz_powcap_2.shape[0]
    netz_powcap_2_cols = netz_powcap_2.shape[1]

    # # Get rid of zero rows
    # non_zero = (netz_powcap_3.loc[:,'2018':] != 0).any(axis = 1)
    # netz_powcap_3 = netz_powcap_3.loc[non_zero].reset_index(drop = True)

    netz_powcap_3_rows = netz_powcap_3.shape[0]
    netz_powcap_3_cols = netz_powcap_3.shape[1]

    netz_powcap_4 = netz_powcap_3[['TECHNOLOGY'] + trans_col_chart]

    netz_powcap_4_rows = netz_powcap_4.shape[0]
    netz_powcap_4_cols = netz_powcap_4.shape[1]


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

    # # Get rid of zero rows
    # non_zero = (netz_trans_3.loc[:,'2017':] != 0).any(axis = 1)
    # netz_trans_3 = netz_trans_3.loc[non_zero].reset_index(drop = True)

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
    waste_own = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(waste_ou)].groupby(['economy']).\
        sum().assign(FUEL = 'Waste', Sector = 'Own-use and losses')

    netz_ownuse_1 = netz_ownuse_1.append([coal_own, oil_own, gas_own, renewables_own, elec_own, heat_own, waste_own])\
        [['FUEL', 'Sector'] + list(netz_ownuse_1.loc[:,'2019':'2050'])].reset_index(drop = True)

    netz_ownuse_1 = netz_ownuse_1[netz_ownuse_1['FUEL'].isin(own_use_fuels)].reset_index(drop = True)

    #################################################################################
    historical_input = EGEDA_hist_own[EGEDA_hist_own['economy'] == economy].copy().\
        iloc[:,:][['FUEL', 'Sector'] + list(EGEDA_hist_own.loc[:, '2000':'2018'])]

    netz_ownuse_1 = historical_input.merge(netz_ownuse_1, how = 'right', on = ['FUEL', 'Sector']).replace(np.nan, 0)

    netz_ownuse_1['FUEL'] = pd.Categorical(netz_ownuse_1['FUEL'], own_use_fuels)

    netz_ownuse_1 = netz_ownuse_1.sort_values('FUEL').reset_index(drop = True)

    netz_ownuse_1 = netz_ownuse_1[['FUEL', 'Sector'] + list(netz_ownuse_1.loc[:, '2000':'2050'])]

    netz_ownuse_1.loc['Total'] = netz_ownuse_1.sum(numeric_only = True)

    netz_ownuse_1.loc['Total', 'FUEL'] = 'Total'
    netz_ownuse_1.loc['Total', 'Sector'] = 'Own-use and losses'

    # Get rid of zero rows
    # non_zero = (netz_ownuse_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_ownuse_1 = netz_ownuse_1.loc[non_zero].reset_index(drop = True)

    netz_ownuse_1_rows = netz_ownuse_1.shape[0]
    netz_ownuse_1_cols = netz_ownuse_1.shape[1]

    netz_ownuse_2 = netz_ownuse_1[['FUEL', 'Sector'] + col_chart_years]

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
    gas_ccs_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(gas_ccs_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Gas CCS')
    nuclear_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(nuke_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Nuclear')
    bio_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(bio_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Biomass')
    waste_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(waste_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Waste')
    comb_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(combination_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Other')
    nons_hp = netz_heatgen_1[netz_heatgen_1['TECHNOLOGY'].isin(nons_heat)].groupby(['economy']).sum().assign(TECHNOLOGY = 'Non-specified')

    # Generation of electricity by tech dataframe (with the above aggregations added)

    netz_heatgen_2 = netz_heatgen_1.append([coal_hp, lignite_hp, oil_hp, gas_hp, gas_ccs_hp, nuclear_hp, bio_hp, waste_hp, comb_hp, nons_hp])\
        [['TECHNOLOGY'] + list(netz_heatgen_1.loc[:, '2017':'2050'])].reset_index(drop = True)                                                              

    netz_heatgen_2['Generation'] = 'Heat'
    netz_heatgen_2 = netz_heatgen_2[['TECHNOLOGY', 'Generation'] + list(netz_heatgen_2.loc[:, '2019':'2050'])]

    # # Insert 0 other row
    # new_row_zero = ['Gas CCS', 'Heat'] + [0] * 34
    # new_series = pd.Series(new_row_zero, index = netz_heatgen_2.columns)

    # netz_heatgen_2 = netz_heatgen_2.append(new_series, ignore_index = True).reset_index(drop = True)

    netz_heatgen_2 = netz_heatgen_2[netz_heatgen_2['TECHNOLOGY'].isin(heat_prod_tech)].\
        set_index('TECHNOLOGY')

    netz_heatgen_2 = netz_heatgen_2.loc[netz_heatgen_2.index.intersection(heat_prod_tech)].reset_index()\
        .rename(columns = {'index': 'TECHNOLOGY'})

    #################################################################################
    historical_gen = EGEDA_hist_heat[EGEDA_hist_heat['economy'] == economy].copy().\
        iloc[:,:][['TECHNOLOGY', 'Generation'] + list(EGEDA_hist_heat.loc[:, '2000':'2018'])]

    netz_heatgen_2 = historical_gen.merge(netz_heatgen_2, how = 'right', on = ['TECHNOLOGY', 'Generation']).replace(np.nan, 0)

    netz_heatgen_2['TECHNOLOGY'] = pd.Categorical(netz_heatgen_2['TECHNOLOGY'], heat_prod_tech)

    netz_heatgen_2 = netz_heatgen_2.sort_values('TECHNOLOGY').reset_index(drop = True)

    netz_heatgen_2.loc['Total'] = netz_heatgen_2.sum(numeric_only = True)

    netz_heatgen_2.loc['Total', 'TECHNOLOGY'] = 'Total'
    netz_heatgen_2.loc['Total', 'Generation'] = 'Heat'

    # # Get rid of zero rows
    # non_zero = (netz_heatgen_2.loc[:,'2000':] != 0).any(axis = 1)
    # netz_heatgen_2 = netz_heatgen_2.loc[non_zero].reset_index(drop = True)

    netz_heatgen_2_rows = netz_heatgen_2.shape[0]
    netz_heatgen_2_cols = netz_heatgen_2.shape[1]

    netz_heatgen_3 = netz_heatgen_2[['TECHNOLOGY', 'Generation'] + gen_col_chart_years]

    netz_heatgen_3_rows = netz_heatgen_3.shape[0]
    netz_heatgen_3_cols = netz_heatgen_3.shape[1]

    #######################################################################################
     
    # Heat use dataframes

    # CARBON NEUTRALITY

    netz_heat_use_1 = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                   (netz_power_df1['Sheet_energy'] == 'UseByTechnology') &
                                   (netz_power_df1['TECHNOLOGY'].isin(heat_only))].reset_index(drop = True)

    coal = netz_heat_use_1[netz_heat_use_1['FUEL'].isin(coal_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Coal',
                                                                                                      TECHNOLOGY = 'Coal heat')

    lignite = netz_heat_use_1[netz_heat_use_1['FUEL'].isin(lignite_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Lignite',
                                                                                              TECHNOLOGY = 'Lignite heat')                                                                                      

    oil = netz_heat_use_1[netz_heat_use_1['FUEL'].isin(oil_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Oil',
                                                                                    TECHNOLOGY = 'Oil heat')

    gas = netz_heat_use_1[netz_heat_use_1['FUEL'].isin(gas_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Gas',
                                                                                      TECHNOLOGY = 'Gas heat')

    biomass = netz_heat_use_1[netz_heat_use_1['FUEL'].isin(biomass_fuel_1)].groupby(['economy']).sum().assign(FUEL = 'Biomass',
                                                                                                        TECHNOLOGY = 'Biomass heat')

    waste = netz_heat_use_1[netz_heat_use_1['FUEL'].isin(waste_fuel)].groupby(['economy']).sum().assign(FUEL = 'Waste',
                                                                                               TECHNOLOGY = 'Waste heat')

    netz_heat_use_2 = netz_heat_use_1.append([coal, lignite, oil, gas, biomass, waste])\
        [['FUEL', 'TECHNOLOGY'] + list(netz_heat_use_1.loc[:,'2017':'2050'])].reset_index(drop = True)

    netz_heat_use_2 = netz_heat_use_2[netz_heat_use_2['FUEL'].isin(heat_agg_fuels)].copy().set_index('FUEL').reset_index()

    netz_heat_use_2 = netz_heat_use_2.groupby('FUEL').sum().reset_index()
    netz_heat_use_2['Transformation'] = 'Heat plant input fuel'
    netz_heat_use_2['FUEL'] = pd.Categorical(netz_heat_use_2['FUEL'], heat_agg_fuels)

    netz_heat_use_2 = netz_heat_use_2.sort_values('FUEL').reset_index(drop = True)

    netz_heat_use_2 = netz_heat_use_2[['FUEL', 'Transformation'] + list(netz_heat_use_2.loc[:,'2017':'2050'])]

    netz_heat_use_2.loc['Total'] = netz_heat_use_2.sum(numeric_only = True)

    netz_heat_use_2.loc['Total', 'FUEL'] = 'Total'
    netz_heat_use_2.loc['Total', 'Transformation'] = 'Heat plant input fuel'

    # # Get rid of zero rows
    # non_zero = (netz_heat_use_2.loc[:,'2017':] != 0).any(axis = 1)
    # netz_heat_use_2 = netz_heat_use_2.loc[non_zero].reset_index(drop = True)

    netz_heat_use_2_rows = netz_heat_use_2.shape[0]
    netz_heat_use_2_cols = netz_heat_use_2.shape[1]

    netz_heat_use_3 = netz_heat_use_2[['FUEL', 'Transformation'] + trans_col_chart]

    netz_heat_use_3_rows = netz_heat_use_3.shape[0]
    netz_heat_use_3_cols = netz_heat_use_3.shape[1]

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

    ref_elecheat = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                 (ref_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                                 (ref_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx', '18_heat'])) &
                                 (ref_power_df1['TECHNOLOGY'].isin(all_elec_heat))].copy().groupby(['economy'])\
                                     .sum().reset_index(drop = True)

    ref_elecheat['fuel_code'] = 'Total'
    ref_elecheat['item_code_new'] = 'Electricity and heat'

    # Grab historical for modern renewables
    historical_eh = EGEDA_hist_eh[EGEDA_hist_eh['economy'] == economy].copy().iloc[:,1:]

    ref_modren_elecheat = historical_eh.merge(ref_modren_elecheat[['fuel_code', 'item_code_new'] + list(ref_modren_elecheat.loc[:,'2019': '2050'])],\
        how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_modren_elecheat = ref_modren_elecheat[['fuel_code', 'item_code_new'] + list(ref_modren_elecheat\
        .loc[:, '2000':'2050'])]

    ref_modren_2 = ref_modren_1.append(ref_modren_elecheat).reset_index(drop = True)

    # NEW EDIT: First slot in 'Total' electricity and heat (including losses and own use)

    # Grab historical for all electricity and heat
    historical_eh2 = EGEDA_hist_eh2[EGEDA_hist_eh2['economy'] == economy].copy().iloc[:, 1:]

    ref_all_elecheat = historical_eh2.merge(ref_elecheat[['fuel_code', 'item_code_new'] + list(ref_elecheat.loc[:,'2019': '2050'])],\
        how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_all_elecheat = ref_all_elecheat[['fuel_code', 'item_code_new'] + list(ref_all_elecheat.loc[:, '2000':'2050'])]

    ref_modren_2 = ref_modren_2.append(ref_all_elecheat).reset_index(drop = True)

    # Now a new line (ratio of 2nd last line to last line)
    ref_rengen_ratio = ['Renewable generation ratio', 'incl losses']\
        + list(ref_modren_2.iloc[ref_modren_2.shape[0] - 2, 2:] / ref_modren_2.iloc[ref_modren_2.shape[0] - 1, 2:])
    ref_rengen_series1 = pd.Series(ref_rengen_ratio, index = ref_modren_2.columns)

    # Now electricity and heat in TFEC
    ref_eh_tfec1 = ref_fedfuel_1[ref_fedfuel_1['fuel_code'].isin(['Electricity', 'Heat'])].copy()

    ref_eh_tfec1.loc['Total'] = ref_eh_tfec1.sum(numeric_only = True)

    ref_eh_tfec1.loc['Total', 'fuel_code'] = 'Total'
    ref_eh_tfec1.loc['Total', 'item_code_new'] = 'Electricity and heat TFEC'    

    ref_eh_tfec1 = ref_eh_tfec1[ref_eh_tfec1['fuel_code'] == 'Total'].copy().reset_index(drop = True)

    ref_modren_2 = ref_modren_2.append(ref_rengen_series1, ignore_index = True).reset_index(drop = True)
    ref_modren_2 = ref_modren_2.append(ref_eh_tfec1).reset_index(drop = True)

    # Another new line that is ratio multiplied by elec and heat tfec
    ref_eh_tfec2 = ['Modern renewables', 'Electricity and heat TFEC']\
        + list(ref_modren_2.iloc[ref_modren_2.shape[0] - 2, 2:] * ref_modren_2.iloc[ref_modren_2.shape[0] - 1, 2:])
    ref_eh_series2 = pd.Series(ref_eh_tfec2, index = ref_modren_2.columns)

    ref_modren_2 = ref_modren_2.append(ref_eh_series2, ignore_index = True).reset_index(drop = True)

    ref_modren_2temp = ref_modren_2[ref_modren_2['item_code_new'] == 'Electricity and heat'].copy().reset_index(drop = True)
    
    ref_modren_2 = ref_modren_2[(ref_modren_2['item_code_new']\
        .isin(['Agriculture', 'Buildings', 'Transport', 'Industry', 'Non-specified others', 'Electricity and heat TFEC'])) &
        (ref_modren_2['fuel_code'] == 'Modern renewables')]\
            .copy().reset_index(drop = True)

    ####################

    ref_modren_2 = ref_modren_2.append(ref_modren_2.sum(numeric_only = True), ignore_index = True) 

    ref_modren_2.iloc[ref_modren_2.shape[0] - 1, 0] = 'Modern renewables'
    ref_modren_2.iloc[ref_modren_2.shape[0] - 1, 1] = 'Total'

    ref_modren_2 = ref_modren_2.append(ref_modren_2temp).reset_index(drop = True)

    # Grab historical for all electricity and heat
    # historical_eh2 = EGEDA_hist_eh2[EGEDA_hist_eh2['economy'] == economy].copy().iloc[:, 1:-2]

    # ref_all_elecheat = historical_eh2.merge(ref_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    # ref_all_elecheat = ref_all_elecheat[['fuel_code', 'item_code_new'] + list(ref_all_elecheat.loc[:, '2000':'2050'])]

    ref_modren_3 = ref_modren_2.append([ref_tfec_1]).reset_index(drop = True)

    non_ren_eh1 = ['Non modern renewables', 'Electricity and heat'] + list(ref_modren_3.iloc[ref_modren_3.shape[0] - 2, 2:] - ref_modren_3.iloc[ref_modren_3.shape[0] - 3, 2:])
    non_ren_series1 = pd.Series(non_ren_eh1, index = ref_modren_3.columns)

    modren_prop1 = ['Modern renewables', 'Reference'] + list(ref_modren_3.iloc[ref_modren_3.shape[0] - 4, 2:] / ref_modren_3.iloc[ref_modren_3.shape[0] - 1, 2:])
    modren_prop_series1 = pd.Series(modren_prop1, index = ref_modren_3.columns)

    ref_modren_4 = ref_modren_3.append([non_ren_series1, modren_prop_series1], ignore_index = True).reset_index(drop = True)

    #ref_modren_4 = ref_modren_4[ref_modren_4['item_code_new'].isin(['Total', 'TFEC', 'Reference'])].copy().reset_index(drop = True)

    ref_modren_4_rows = ref_modren_4.shape[0]
    ref_modren_4_cols = ref_modren_4.shape[1]

    # CARBON NEUTRALITY: Modern renewables
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

    netz_elecheat = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                 (netz_power_df1['Sheet_energy'] == 'ProductionByTechnology') &
                                 (netz_power_df1['FUEL'].isin(['17_electricity', '17_electricity_Dx', '18_heat'])) &
                                 (netz_power_df1['TECHNOLOGY'].isin(all_elec_heat))].copy().groupby(['economy'])\
                                     .sum().reset_index(drop = True)

    netz_elecheat['fuel_code'] = 'Total'
    netz_elecheat['item_code_new'] = 'Electricity and heat'

    # Grab historical for modern renewables
    historical_eh = EGEDA_hist_eh[EGEDA_hist_eh['economy'] == economy].copy().iloc[:,1:]

    netz_modren_elecheat = historical_eh.merge(netz_modren_elecheat[['fuel_code', 'item_code_new'] + list(netz_modren_elecheat.loc[:,'2019': '2050'])],\
        how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_modren_elecheat = netz_modren_elecheat[['fuel_code', 'item_code_new'] + list(netz_modren_elecheat\
        .loc[:, '2000':'2050'])]

    netz_modren_2 = netz_modren_1.append(netz_modren_elecheat).reset_index(drop = True)
    # NEW EDIT: First slot in 'Total' electricity and heat (including losses and own use)

    # Grab historical for all electricity and heat
    historical_eh2 = EGEDA_hist_eh2[EGEDA_hist_eh2['economy'] == economy].copy().iloc[:, 1:]

    netz_all_elecheat = historical_eh2.merge(netz_elecheat[['fuel_code', 'item_code_new'] + list(netz_elecheat.loc[:,'2019': '2050'])],\
        how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_all_elecheat = netz_all_elecheat[['fuel_code', 'item_code_new'] + list(netz_all_elecheat.loc[:, '2000':'2050'])]

    netz_modren_2 = netz_modren_2.append(netz_all_elecheat).reset_index(drop = True)

    # Now a new line (ratio of 2nd last line to last line)
    netz_rengen_ratio = ['Renewable generation ratio', 'incl losses']\
        + list(netz_modren_2.iloc[netz_modren_2.shape[0] - 2, 2:] / netz_modren_2.iloc[netz_modren_2.shape[0] - 1, 2:])
    netz_rengen_series1 = pd.Series(netz_rengen_ratio, index = netz_modren_2.columns)

    # Now electricity and heat in TFEC
    netz_eh_tfec1 = netz_fedfuel_1[netz_fedfuel_1['fuel_code'].isin(['Electricity', 'Heat'])].copy()

    netz_eh_tfec1.loc['Total'] = netz_eh_tfec1.sum(numeric_only = True)

    netz_eh_tfec1.loc['Total', 'fuel_code'] = 'Total'
    netz_eh_tfec1.loc['Total', 'item_code_new'] = 'Electricity and heat TFEC'    

    netz_eh_tfec1 = netz_eh_tfec1[netz_eh_tfec1['fuel_code'] == 'Total'].copy().reset_index(drop = True)

    netz_modren_2 = netz_modren_2.append(netz_rengen_series1, ignore_index = True).reset_index(drop = True)
    netz_modren_2 = netz_modren_2.append(netz_eh_tfec1).reset_index(drop = True)

    # Another new line that is ratio multiplied by elec and heat tfec
    netz_eh_tfec2 = ['Modern renewables', 'Electricity and heat TFEC']\
        + list(netz_modren_2.iloc[netz_modren_2.shape[0] - 2, 2:] * netz_modren_2.iloc[netz_modren_2.shape[0] - 1, 2:])
    netz_eh_series2 = pd.Series(netz_eh_tfec2, index = netz_modren_2.columns)

    netz_modren_2 = netz_modren_2.append(netz_eh_series2, ignore_index = True).reset_index(drop = True)

    netz_modren_2temp = netz_modren_2[netz_modren_2['item_code_new'] == 'Electricity and heat'].copy().reset_index(drop = True)
    
    netz_modren_2 = netz_modren_2[(netz_modren_2['item_code_new']\
        .isin(['Agriculture', 'Buildings', 'Transport', 'Industry', 'Non-specified others', 'Electricity and heat TFEC'])) &
        (netz_modren_2['fuel_code'] == 'Modern renewables')]\
            .copy().reset_index(drop = True)

    ####################

    netz_modren_2 = netz_modren_2.append(netz_modren_2.sum(numeric_only = True), ignore_index = True) 

    netz_modren_2.iloc[netz_modren_2.shape[0] - 1, 0] = 'Modern renewables'
    netz_modren_2.iloc[netz_modren_2.shape[0] - 1, 1] = 'Total'

    netz_modren_2 = netz_modren_2.append(netz_modren_2temp).reset_index(drop = True)

    # Grab historical for all electricity and heat
    # historical_eh2 = EGEDA_hist_eh2[EGEDA_hist_eh2['economy'] == economy].copy().iloc[:, 1:-2]

    # netz_all_elecheat = historical_eh2.merge(netz_elecheat, how = 'left', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    # netz_all_elecheat = netz_all_elecheat[['fuel_code', 'item_code_new'] + list(netz_all_elecheat.loc[:, '2000':'2050'])]

    netz_modren_3 = netz_modren_2.append([netz_tfec_1]).reset_index(drop = True)

    non_ren_eh1 = ['Non modern renewables', 'Electricity and heat'] + list(netz_modren_3.iloc[netz_modren_3.shape[0] - 2, 2:] - netz_modren_3.iloc[netz_modren_3.shape[0] - 3, 2:])
    non_ren_series1 = pd.Series(non_ren_eh1, index = netz_modren_3.columns)

    modren_prop1 = ['Modern renewables', 'Carbon Neutrality'] + list(netz_modren_3.iloc[netz_modren_3.shape[0] - 4, 2:] / netz_modren_3.iloc[netz_modren_3.shape[0] - 1, 2:])
    modren_prop_series1 = pd.Series(modren_prop1, index = netz_modren_3.columns)

    netz_modren_4 = netz_modren_3.append([non_ren_series1, modren_prop_series1], ignore_index = True).reset_index(drop = True)

    # Remove historical from CN
    netz_modren_4.loc[netz_modren_4['item_code_new'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

    #netz_modren_4 = netz_modren_4[netz_modren_4['item_code_new'].isin(['Total', 'TFEC', 'Net-zero'])].copy().reset_index(drop = True)

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

    # # Energy intensity removed


    ##############################################################################################################

    # # OSeMOSYS datafrane builds removed (transport modality and heavy industry)


    ######################################################################

    # EMISSIONS dataframes

    # First data frame construction: Emissions by fuels
    ref_emiss_1 = EGEDA_emissions_reference[(EGEDA_emissions_reference['economy'] == economy) & 
                          (EGEDA_emissions_reference['item_code_new'].isin(['13_x_dem_pow_own_hyd'])) &
                          (EGEDA_emissions_reference['fuel_code'].isin(Required_emiss))].loc[:, 'fuel_code':].reset_index(drop = True)

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = ref_emiss_1[ref_emiss_1['fuel_code'].isin(Coal_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_x_dem_pow_own_hyd')
    
    oil = ref_emiss_1[ref_emiss_1['fuel_code'].isin(Oil_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_x_dem_pow_own_hyd')
    
    heat_others = ref_emiss_1[ref_emiss_1['fuel_code'].isin(Heat_others_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_x_dem_pow_own_hyd')

    # EMISSIONS fuel data frame 1 (data frame 6)

    ref_emiss_fuel_1 = ref_emiss_1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(ref_emiss_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_emiss_fuel_1.loc[ref_emiss_fuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_emiss_fuel_1.loc[ref_emiss_fuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    ref_emiss_fuel_1 = ref_emiss_fuel_1[ref_emiss_fuel_1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()\
        .replace(np.nan, 0)

    ref_emiss_fuel_1.loc['Total'] = ref_emiss_fuel_1.sum(numeric_only = True)

    ref_emiss_fuel_1.loc['Total', 'fuel_code'] = 'Total'
    ref_emiss_fuel_1.loc['Total', 'item_code_new'] = 'Emissions'

    # Get rid of zero rows
    # non_zero = (ref_emiss_fuel_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_emiss_fuel_1 = ref_emiss_fuel_1.loc[non_zero].reset_index(drop = True)

    ref_emiss_fuel_1_rows = ref_emiss_fuel_1.shape[0]
    ref_emiss_fuel_1_cols = ref_emiss_fuel_1.shape[1]

    ref_emiss_fuel_2 = ref_emiss_fuel_1[['fuel_code', 'item_code_new'] + col_chart_years]
    # ref_emiss_fuel_2 = ref_emiss_fuel_2[ref_emiss_fuel_2['fuel_code'] != 'Total']

    ref_emiss_fuel_2_rows = ref_emiss_fuel_2.shape[0]
    ref_emiss_fuel_2_cols = ref_emiss_fuel_2.shape[1]

    # Second data frame construction: FED by sectors
    ref_emiss_2 = EGEDA_emissions_reference[(EGEDA_emissions_reference['economy'] == economy) &
                               (EGEDA_emissions_reference['item_code_new'].isin(Sectors_emiss)) &
                               (EGEDA_emissions_reference['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    ref_emiss_2 = ref_emiss_2[['fuel_code', 'item_code_new'] + list(ref_emiss_2.loc[:,'2000':'2050'])]
    
    ref_emiss_2_rows = ref_emiss_2.shape[0]
    ref_emiss_2_cols = ref_emiss_2.shape[1]

    # Now build aggregate sector variables
    
    buildings = ref_emiss_2[ref_emiss_2['item_code_new'].isin(Buildings_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = ref_emiss_2[ref_emiss_2['item_code_new'].isin(Ag_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    ref_emiss_sector_1 = ref_emiss_2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(ref_emiss_2.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '9_x_power', 'item_code_new'] = 'Power'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own use'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_emiss_sector_1.loc[ref_emiss_sector_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    ref_emiss_sector_1 = ref_emiss_sector_1[ref_emiss_sector_1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()\
        .replace(np.nan, 0)
    ref_emiss_sector_1 = ref_emiss_sector_1[['fuel_code', 'item_code_new'] + list(ref_emiss_sector_1.loc[:, '2000':'2050'])]

    ref_emiss_sector_1.loc['Total'] = ref_emiss_sector_1.sum(numeric_only = True)

    ref_emiss_sector_1.loc['Total', 'fuel_code'] = '19_total'
    ref_emiss_sector_1.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    # non_zero = (ref_emiss_sector_1.loc[:,'2000':] != 0).any(axis = 1)
    # ref_emiss_sector_1 = ref_emiss_sector_1.loc[non_zero].reset_index(drop = True)

    ref_emiss_sector_1_rows = ref_emiss_sector_1.shape[0]
    ref_emiss_sector_1_cols = ref_emiss_sector_1.shape[1]

    ref_emiss_sector_2 = ref_emiss_sector_1[['fuel_code', 'item_code_new'] + col_chart_years]

    ref_emiss_sector_2_rows = ref_emiss_sector_2.shape[0]
    ref_emiss_sector_2_cols = ref_emiss_sector_2.shape[1]

    ##################################################################################################################################
    # NET ZERO DATA FRAMES
    # First data frame construction: Emissions by fuels
    netz_emiss_1 = EGEDA_emissions_netzero[(EGEDA_emissions_netzero['economy'] == economy) & 
                          (EGEDA_emissions_netzero['item_code_new'].isin(['13_x_dem_pow_own_hyd'])) &
                          (EGEDA_emissions_netzero['fuel_code'].isin(Required_emiss))].loc[:, 'fuel_code':].reset_index(drop = True)

    # Now build aggregate variables of the first level fuels in EGEDA

    coal = netz_emiss_1[netz_emiss_1['fuel_code'].isin(Coal_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                    item_code_new = '13_x_dem_pow_own_hyd')
    
    oil = netz_emiss_1[netz_emiss_1['fuel_code'].isin(Oil_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                  item_code_new = '13_x_dem_pow_own_hyd')
    
    heat_others = netz_emiss_1[netz_emiss_1['fuel_code'].isin(Heat_others_emiss)].groupby(['item_code_new']).sum().assign(fuel_code = 'Heat & others',
                                                                                                                  item_code_new = '13_x_dem_pow_own_hyd')

    # EMISSIONS fuel data frame 1 (data frame 6)

    netz_emiss_fuel_1 = netz_emiss_1.append([coal, oil, heat_others])[['fuel_code',
                                                             'item_code_new'] + list(netz_emiss_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_emiss_fuel_1.loc[netz_emiss_fuel_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_emiss_fuel_1.loc[netz_emiss_fuel_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'

    netz_emiss_fuel_1 = netz_emiss_fuel_1[netz_emiss_fuel_1['fuel_code'].isin(Emissions_agg_fuels)].set_index('fuel_code').loc[Emissions_agg_fuels].reset_index()\
        .replace(np.nan, 0)

    netz_emiss_fuel_1.loc['Total'] = netz_emiss_fuel_1.sum(numeric_only = True)

    netz_emiss_fuel_1.loc['Total', 'fuel_code'] = 'Total'
    netz_emiss_fuel_1.loc['Total', 'item_code_new'] = 'Emissions'

    # Get rid of zero rows
    # non_zero = (netz_emiss_fuel_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_emiss_fuel_1 = netz_emiss_fuel_1.loc[non_zero].reset_index(drop = True)

    netz_emiss_fuel_1_rows = netz_emiss_fuel_1.shape[0]
    netz_emiss_fuel_1_cols = netz_emiss_fuel_1.shape[1]

    netz_emiss_fuel_2 = netz_emiss_fuel_1[['fuel_code', 'item_code_new'] + col_chart_years]
    netz_emiss_fuel_2 = netz_emiss_fuel_2[netz_emiss_fuel_2['fuel_code'] != 'Total']

    netz_emiss_fuel_2_rows = netz_emiss_fuel_2.shape[0]
    netz_emiss_fuel_2_cols = netz_emiss_fuel_2.shape[1]

    # Second data frame construction: FED by sectors
    netz_emiss_2 = EGEDA_emissions_netzero[(EGEDA_emissions_netzero['economy'] == economy) &
                               (EGEDA_emissions_netzero['item_code_new'].isin(Sectors_emiss)) &
                               (EGEDA_emissions_netzero['fuel_code'].isin(['19_total']))].loc[:,'fuel_code':].reset_index(drop = True)

    netz_emiss_2 = netz_emiss_2[['fuel_code', 'item_code_new'] + list(netz_emiss_2.loc[:,'2000':'2050'])]
    
    netz_emiss_2_rows = netz_emiss_2.shape[0]
    netz_emiss_2_cols = netz_emiss_2.shape[1]

    # Now build aggregate sector variables
    
    buildings = netz_emiss_2[netz_emiss_2['item_code_new'].isin(Buildings_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                              item_code_new = 'Buildings')
    
    agriculture = netz_emiss_2[netz_emiss_2['item_code_new'].isin(Ag_emiss)].groupby(['fuel_code']).sum().assign(fuel_code = '19_total',
                                                                                                         item_code_new = 'Agriculture')

    # Build aggregate data frame of FED sector

    netz_emiss_sector_1 = netz_emiss_2.append([buildings, agriculture])[['fuel_code', 'item_code_new'] + list(netz_emiss_2.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '9_x_power', 'item_code_new'] = 'Power'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own use'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    #emissions_sector_df1.loc[emissions_sector_df1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_emiss_sector_1.loc[netz_emiss_sector_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'

    netz_emiss_sector_1 = netz_emiss_sector_1[netz_emiss_sector_1['item_code_new'].isin(Emissions_agg_sectors)].set_index('item_code_new').loc[Emissions_agg_sectors].reset_index()\
        .replace(np.nan, 0)
    netz_emiss_sector_1 = netz_emiss_sector_1[['fuel_code', 'item_code_new'] + list(netz_emiss_sector_1.loc[:, '2000':'2050'])]

    netz_emiss_sector_1.loc['Total'] = netz_emiss_sector_1.sum(numeric_only = True)

    netz_emiss_sector_1.loc['Total', 'fuel_code'] = '19_total'
    netz_emiss_sector_1.loc['Total', 'item_code_new'] = 'Total'

    # Get rid of zero rows
    # non_zero = (netz_emiss_sector_1.loc[:,'2000':] != 0).any(axis = 1)
    # netz_emiss_sector_1 = netz_emiss_sector_1.loc[non_zero].reset_index(drop = True)

    netz_emiss_sector_1_rows = netz_emiss_sector_1.shape[0]
    netz_emiss_sector_1_cols = netz_emiss_sector_1.shape[1]

    netz_emiss_sector_2 = netz_emiss_sector_1[['fuel_code', 'item_code_new'] + col_chart_years]

    netz_emiss_sector_2_rows = netz_emiss_sector_2.shape[0]
    netz_emiss_sector_2_cols = netz_emiss_sector_2.shape[1]

    # Total emissions dataframe

    emiss_total_1 = ref_emiss_fuel_1[ref_emiss_fuel_1['fuel_code'] == 'Total'].copy()

    emiss_total_1.loc[emiss_total_1['fuel_code'] == 'Total', 'fuel_code'] = 'Reference'
        
    emiss_total_1 = emiss_total_1.append(netz_emiss_fuel_1[netz_emiss_fuel_1['fuel_code'] == 'Total'].copy())\
        .reset_index(drop = True) 

    emiss_total_1.loc[emiss_total_1['fuel_code'] == 'Total', 'fuel_code'] = 'Carbon Neutrality'

    # Remove historical from carbon neutrality
    emiss_total_1.loc[emiss_total_1['fuel_code'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

    emiss_total_1_rows = emiss_total_1.shape[0]
    emiss_total_1_cols = emiss_total_1.shape[1]

    # Wedge chart removed

    ##################################################################################################

    # Removed fuel supply

    ###########################################################################################

    # Fuel consummption data frame builds

    # Crude oil

    ref_crude_ind = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                         (EGEDA_years_reference['item_code_new'].isin(['14_industry_sector'])) &
                                         (EGEDA_years_reference['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                             .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_crude_ind = ref_crude_ind[['fuel_code', 'item_code_new'] + list(ref_crude_ind.loc[:, '2000':'2050'])]
    ref_crude_ind.loc[ref_crude_ind['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    ref_crude_bld = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                          (EGEDA_years_reference['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                          (EGEDA_years_reference['fuel_code'].isin(['6_crude_oil_and_ngl']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Buildings')

    ref_crude_bld = ref_crude_bld[['fuel_code', 'item_code_new'] + list(ref_crude_bld.loc[:, '2000':'2050'])]

    ref_crude_ag = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                          (EGEDA_years_reference['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])) &
                                          (EGEDA_years_reference['fuel_code'].isin(['6_crude_oil_and_ngl']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Agriculture')

    ref_crude_ag = ref_crude_ag[['fuel_code', 'item_code_new'] + list(ref_crude_ag.loc[:, '2000':'2050'])]

    ref_crude_trn = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                          (EGEDA_years_reference['item_code_new'].isin(['15_transport_sector'])) &
                                          (EGEDA_years_reference['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                              .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_crude_trn = ref_crude_trn[['fuel_code', 'item_code_new'] + list(ref_crude_trn.loc[:, '2000':'2050'])]
    ref_crude_trn.loc[ref_crude_trn['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    ref_crude_ne = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['17_nonenergy_use'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                            .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_crude_ne = ref_crude_ne[['fuel_code', 'item_code_new'] + list(ref_crude_ne.loc[:, '2000':'2050'])]
    ref_crude_ne.loc[ref_crude_ne['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    ref_crude_ns = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                            .copy().replace(np.nan, 0).reset_index(drop = True)

    ref_crude_ns = ref_crude_ns[['fuel_code', 'item_code_new'] + list(ref_crude_ns.loc[:, '2000':'2050'])]
    ref_crude_ns.loc[ref_crude_ns['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    # Own-use
    ref_crude_own = ref_trans_df1[(ref_trans_df1['economy'] == economy) & 
                                  (ref_trans_df1['Sector'] == 'OWN') &
                                  (ref_trans_df1['FUEL'].isin(['6_1_crude_oil', '6_x_ngls']))]\
                                      .copy().reset_index(drop = True)

    ref_crude_own = ref_crude_own.groupby(['economy']).sum().copy().reset_index(drop = True)\
                        .assign(fuel_code = 'Crude oil & NGL', item_code_new = '10_losses_and_own_use')

    #################################################################################
    hist_ownoil = EGEDA_hist_own_oil[(EGEDA_hist_own_oil['economy'] == economy) &
                                     (EGEDA_hist_own_oil['FUEL'] == 'Crude oil & NGL')].copy().\
                                        iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_crude_own = hist_ownoil.merge(ref_crude_own[['fuel_code', 'item_code_new'] + list(ref_crude_own.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_crude_own = ref_crude_own[['fuel_code', 'item_code_new'] + list(ref_crude_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    ref_crude_power = ref_power_df1[(ref_power_df1['economy'] == economy) &
                                    (ref_power_df1['FUEL'].isin(['6_1_crude_oil', '6_x_ngls']))].copy().reset_index(drop = True)

    ref_crude_power = ref_crude_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                          .assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Power')

    #################################################################################
    hist_poweroil = EGEDA_histpower_oil[(EGEDA_histpower_oil['economy'] == economy) &
                                        (EGEDA_histpower_oil['FUEL'] == 'Crude oil & NGL')].copy()\
                                            .iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    ref_crude_power = hist_poweroil.merge(ref_crude_power[['fuel_code', 'item_code_new'] + list(ref_crude_power.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_crude_power = ref_crude_power[['fuel_code', 'item_code_new'] + list(ref_crude_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)
    
    # Refining
    ref_crude_refinery = ref_refinery_1[ref_refinery_1['FUEL'].isin(['Crude oil', 'NGLs'])]\
                                .copy().groupby(['Transformation']).sum().reset_index(drop = True)\
                                    .assign(fuel_code = '6_crude_oil_and_ngl', item_code_new = '9_4_oil_refineries')

    hist_refine = EGEDA_hist_refining[EGEDA_hist_refining['economy'] == economy].copy()\
                     .iloc[:,:][['fuel_code', 'item_code_new'] + list(EGEDA_hist_refining.loc[:, '2000':'2018'])]\
                     .reset_index(drop = True)


    ref_crude_refinery = hist_refine.merge(ref_crude_refinery[['fuel_code', 'item_code_new'] + list(ref_crude_refinery.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    ref_crude_refinery = ref_crude_refinery[['fuel_code', 'item_code_new'] + list(ref_crude_refinery.loc[:, '2000': '2050'])]

    ref_crude_refinery.loc[ref_crude_refinery['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    ref_crudecons_1 = ref_crude_ind.append([ref_crude_bld, ref_crude_ag, ref_crude_trn, ref_crude_ne, 
                                            ref_crude_ns, ref_crude_own, ref_crude_power, ref_crude_refinery])\
                                               .copy().reset_index(drop = True)

    ref_crudecons_1.loc[ref_crudecons_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'
    ref_crudecons_1.loc[ref_crudecons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    ref_crudecons_1.loc[ref_crudecons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    ref_crudecons_1.loc[ref_crudecons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    ref_crudecons_1.loc[ref_crudecons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'
    ref_crudecons_1.loc[ref_crudecons_1['item_code_new'] == '9_4_oil_refineries', 'item_code_new'] = 'Refining'

    ref_crudecons_1.loc['Total'] = ref_crudecons_1.sum(numeric_only = True)

    ref_crudecons_1.loc['Total', 'fuel_code'] = 'Crude oil & NGL'
    ref_crudecons_1.loc['Total', 'item_code_new'] = 'Total'

    ref_crudecons_1 = ref_crudecons_1.copy().reset_index(drop = True)

    ref_crudecons_1_rows = ref_crudecons_1.shape[0]
    ref_crudecons_1_cols = ref_crudecons_1.shape[1]

    ## CARBON NEUTRALITY
    # Crude oil

    netz_crude_ind = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                         (EGEDA_years_netzero['item_code_new'].isin(['14_industry_sector'])) &
                                         (EGEDA_years_netzero['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                             .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_crude_ind = netz_crude_ind[['fuel_code', 'item_code_new'] + list(netz_crude_ind.loc[:, '2000':'2050'])]
    netz_crude_ind.loc[netz_crude_ind['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    netz_crude_bld = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                          (EGEDA_years_netzero['item_code_new'].isin(['16_1_commercial_and_public_services', '16_2_residential'])) &
                                          (EGEDA_years_netzero['fuel_code'].isin(['6_crude_oil_and_ngl']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Buildings')

    netz_crude_bld = netz_crude_bld[['fuel_code', 'item_code_new'] + list(netz_crude_bld.loc[:, '2000':'2050'])]

    netz_crude_ag = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                          (EGEDA_years_netzero['item_code_new'].isin(['16_3_agriculture', '16_4_fishing'])) &
                                          (EGEDA_years_netzero['fuel_code'].isin(['6_crude_oil_and_ngl']))].copy().replace(np.nan, 0).groupby(['fuel_code'])\
                                                .sum().reset_index(drop = True).assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Agriculture')

    netz_crude_ag = netz_crude_ag[['fuel_code', 'item_code_new'] + list(netz_crude_ag.loc[:, '2000':'2050'])]

    netz_crude_trn = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                          (EGEDA_years_netzero['item_code_new'].isin(['15_transport_sector'])) &
                                          (EGEDA_years_netzero['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                              .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_crude_trn = netz_crude_trn[['fuel_code', 'item_code_new'] + list(netz_crude_trn.loc[:, '2000':'2050'])]
    netz_crude_trn.loc[netz_crude_trn['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    netz_crude_ne = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['17_nonenergy_use'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                            .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_crude_ne = netz_crude_ne[['fuel_code', 'item_code_new'] + list(netz_crude_ne.loc[:, '2000':'2050'])]
    netz_crude_ne.loc[netz_crude_ne['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    netz_crude_ns = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['16_5_nonspecified_others'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['6_crude_oil_and_ngl']))]\
                                            .copy().replace(np.nan, 0).reset_index(drop = True)

    netz_crude_ns = netz_crude_ns[['fuel_code', 'item_code_new'] + list(netz_crude_ns.loc[:, '2000':'2050'])]
    netz_crude_ns.loc[netz_crude_ns['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    # Own-use
    netz_crude_own = netz_trans_df1[(netz_trans_df1['economy'] == economy) & 
                                  (netz_trans_df1['Sector'] == 'OWN') &
                                  (netz_trans_df1['FUEL'].isin(['6_1_crude_oil', '6_x_ngls']))]\
                                      .copy().reset_index(drop = True)

    netz_crude_own = netz_crude_own.groupby(['economy']).sum().copy().reset_index(drop = True)\
                        .assign(fuel_code = 'Crude oil & NGL', item_code_new = '10_losses_and_own_use')

    #################################################################################
    hist_ownoil = EGEDA_hist_own_oil[(EGEDA_hist_own_oil['economy'] == economy) &
                                     (EGEDA_hist_own_oil['FUEL'] == 'Crude oil & NGL')].copy().\
                                        iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_hist_own_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_crude_own = hist_ownoil.merge(netz_crude_own[['fuel_code', 'item_code_new'] + list(netz_crude_own.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_crude_own = netz_crude_own[['fuel_code', 'item_code_new'] + list(netz_crude_own.loc[:, '2000':'2050'])].copy().reset_index(drop = True)

    # Power
    netz_crude_power = netz_power_df1[(netz_power_df1['economy'] == economy) &
                                    (netz_power_df1['FUEL'].isin(['6_1_crude_oil', '6_x_ngls']))].copy().reset_index(drop = True)

    netz_crude_power = netz_crude_power.groupby(['economy']).sum().copy().reset_index(drop = True)\
                          .assign(fuel_code = 'Crude oil & NGL', item_code_new = 'Power')

    #################################################################################
    hist_poweroil = EGEDA_histpower_oil[(EGEDA_histpower_oil['economy'] == economy) &
                                        (EGEDA_histpower_oil['FUEL'] == 'Crude oil & NGL')].copy()\
                                            .iloc[:,:][['FUEL', 'item_code_new'] + list(EGEDA_histpower_oil.loc[:, '2000':'2018'])]\
                                            .rename(columns = {'FUEL': 'fuel_code'}).reset_index(drop = True)

    netz_crude_power = hist_poweroil.merge(netz_crude_power[['fuel_code', 'item_code_new'] + list(netz_crude_power.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_crude_power = netz_crude_power[['fuel_code', 'item_code_new'] + list(netz_crude_power.loc[:, '2000':'2050'])].copy().reset_index(drop = True)
    
    # Refining
    netz_crude_refinery = netz_refinery_1[netz_refinery_1['FUEL'].isin(['Crude oil', 'NGLs'])]\
                                .copy().groupby(['Transformation']).sum().reset_index(drop = True)\
                                    .assign(fuel_code = '6_crude_oil_and_ngl', item_code_new = '9_4_oil_refineries')

    hist_refine = EGEDA_hist_refining[EGEDA_hist_refining['economy'] == economy].copy()\
                     .iloc[:,:][['fuel_code', 'item_code_new'] + list(EGEDA_hist_refining.loc[:, '2000':'2018'])]\
                     .reset_index(drop = True)


    netz_crude_refinery = hist_refine.merge(netz_crude_refinery[['fuel_code', 'item_code_new'] + list(netz_crude_refinery.loc[:,'2019': '2050'])],\
        how = 'right', on = ['fuel_code', 'item_code_new']).replace(np.nan, 0)

    netz_crude_refinery = netz_crude_refinery[['fuel_code', 'item_code_new'] + list(netz_crude_refinery.loc[:, '2000': '2050'])]

    netz_crude_refinery.loc[netz_crude_refinery['fuel_code'] == '6_crude_oil_and_ngl', 'fuel_code'] = 'Crude oil & NGL'

    netz_crudecons_1 = netz_crude_ind.append([netz_crude_bld, netz_crude_ag, netz_crude_trn, netz_crude_ne, 
                                            netz_crude_ns, netz_crude_own, netz_crude_power, netz_crude_refinery])\
                                               .copy().reset_index(drop = True)

    netz_crudecons_1.loc[netz_crudecons_1['item_code_new'] == '10_losses_and_own_use', 'item_code_new'] = 'Own-use and losses'
    netz_crudecons_1.loc[netz_crudecons_1['item_code_new'] == '14_industry_sector', 'item_code_new'] = 'Industry'
    netz_crudecons_1.loc[netz_crudecons_1['item_code_new'] == '15_transport_sector', 'item_code_new'] = 'Transport'
    netz_crudecons_1.loc[netz_crudecons_1['item_code_new'] == '17_nonenergy_use', 'item_code_new'] = 'Non-energy'
    netz_crudecons_1.loc[netz_crudecons_1['item_code_new'] == '16_5_nonspecified_others', 'item_code_new'] = 'Non-specified'
    netz_crudecons_1.loc[netz_crudecons_1['item_code_new'] == '9_4_oil_refineries', 'item_code_new'] = 'Refining'

    netz_crudecons_1.loc['Total'] = netz_crudecons_1.sum(numeric_only = True)

    netz_crudecons_1.loc['Total', 'fuel_code'] = 'Crude oil & NGL'
    netz_crudecons_1.loc['Total', 'item_code_new'] = 'Total'

    netz_crudecons_1 = netz_crudecons_1.copy().reset_index(drop = True)

    netz_crudecons_1_rows = netz_crudecons_1.shape[0]
    netz_crudecons_1_cols = netz_crudecons_1.shape[1]

    
    # Liquid and solid renewables

    # Renew and petprod gone

    ##################################################################################################################

    # Emissions intensity

    # REFERENCE
    ref_co2int_1 = ref_emiss_fuel_1[ref_emiss_fuel_1['fuel_code'] == 'Total'].copy().reset_index(drop = True)

    ref_co2int_1 = ref_co2int_1.append(ref_tpes_1[ref_tpes_1['fuel_code'] == 'Total']).copy().reset_index(drop = True)

    ref_calc1 = ['Reference', 'CO2 intensity'] + list(ref_co2int_1.iloc[0, 2:] / ref_co2int_1.iloc[1, 2:])
    ref_series1 = pd.Series(ref_calc1, index = ref_co2int_1.columns)

    ref_co2int_2 = ref_co2int_1.append(ref_series1, ignore_index = True).reset_index(drop = True)

    ref_co2int_2_rows = ref_co2int_2.shape[0]
    ref_co2int_2_cols = ref_co2int_2.shape[1]

    # CARBON NEUTRALITY
    netz_co2int_1 = netz_emiss_fuel_1[netz_emiss_fuel_1['fuel_code'] == 'Total'].copy().reset_index(drop = True)

    netz_co2int_1 = netz_co2int_1.append(netz_tpes_1[netz_tpes_1['fuel_code'] == 'Total']).copy().reset_index(drop = True)

    netz_calc1 = ['Carbon Neutrality', 'CO2 intensity'] + list(netz_co2int_1.iloc[0, 2:] / netz_co2int_1.iloc[1, 2:])
    netz_series1 = pd.Series(netz_calc1, index = netz_co2int_1.columns)

    netz_co2int_2 = netz_co2int_1.append(netz_series1, ignore_index = True).reset_index(drop = True)

    # Remove 2000 to 2017 data from CN
    netz_co2int_2.loc[netz_co2int_2['fuel_code'] == 'Carbon Neutrality', '2000':'2017'] = np.nan

    netz_co2int_2_rows = netz_co2int_2.shape[0]
    netz_co2int_2_cols = netz_co2int_2.shape[1]

    # Electricity by sector

    ## REMOVED

    # Df builds are complete

    ################################################################################################################
    ################################################################################################################
    ############## NEW DATAFRAME CONSTRUCTION FOR PUBLISHED APPENDIX ###############################################
    ################################################################################################################
    ################################################################################################################
    
    ######################################### 
    ## MACRO
    # GDP
    # Population
    # GDP per capita

    macro_1 = macro_1[macro_1['Series'].isin(['GDP 2018 USD PPP', 'Population', 'GDP per capita'])].copy()\
        .iloc[:, 1:].reset_index(drop = True)

    macro_1.loc[macro_1['Series'] == 'GDP 2018 USD PPP', 'Series'] = 'GDP (2018 USD billion PPP)'
    macro_1.loc[macro_1['Series'] == 'Population', 'Series'] = 'Population (millions)'
    macro_1.loc[macro_1['Series'] == 'GDP per capita', 'Series'] = 'GDP per capita (2018 USD PPP)' 

    ## TPES
    # TPES per capita
    # TPES per GDP

    ref_tpes_3 = ref_tpes_1[ref_tpes_1['fuel_code'] == 'Total'].copy().\
        rename(columns = {'item_code_new': 'Series'}).iloc[:, 1:].reset_index(drop = True)

    ref_tpes_3.loc[ref_tpes_3['Series'] == '7_total_primary_energy_supply', 'Series'] = 'Total primary energy supply (PJ)'

    ref_tpes_calcs = ref_tpes_3.append(macro_1[macro_1['Series'].isin(['Population (millions)', 'GDP (2018 USD billion PPP)'])])\
        .copy().reset_index(drop = True)

    ref_tpes_pc = ['TPES per capita (GJ per person)'] + list(ref_tpes_calcs.iloc[0, 1:] / ref_tpes_calcs.iloc[2, 1:])
    ref_tpes_pc_series = pd.Series(ref_tpes_pc, index = ref_tpes_3.columns)

    ref_tpes_pGDP = ['TPES per GDP (GJ per thousand 2018 USD PPP)'] + list(ref_tpes_calcs.iloc[0, 1:] / ref_tpes_calcs.iloc[1, 1:])
    ref_tpes_pGDP_series = pd.Series(ref_tpes_pGDP, index = ref_tpes_3.columns)

    ref_tpes_3 = ref_tpes_3.append([ref_tpes_pc_series, ref_tpes_pGDP_series], ignore_index = True).reset_index(drop = True)

    ## FED
    # Final energy intensity per capita
    # Final energy intensity per GDP

    ref_tfec_1 = ref_tfec_1.copy().rename(columns = {'item_code_new': 'Series'}).iloc[:, 1:]

    ref_tfec_1 = ref_fedfuel_1[ref_fedfuel_1['fuel_code'] == 'Total'].copy()\
        .rename(columns = {'item_code_new': 'Series'}).iloc[:, 1:].append(ref_tfec_1)

    ref_tfec_1.loc[ref_tfec_1['Series'] == '12_total_final_consumption', 'Series'] = 'Total final consumption (PJ)'

    ref_tfec_1.loc[ref_tfec_1['Series'] == 'TFEC', 'Series'] = 'Final energy demand (PJ)'

    ref_tfec_calcs = ref_tfec_1.append(macro_1[macro_1['Series'].isin(['Population (millions)', 'GDP (2018 USD billion PPP)'])])\
        .copy().reset_index(drop = True)

    ref_tfec_pc = ['Final energy demand per capita (GJ per person)'] + list(ref_tfec_calcs.iloc[1, 1:] / ref_tfec_calcs.iloc[3, 1:])
    ref_tfec_pc_series = pd.Series(ref_tfec_pc, index = ref_tfec_1.columns)

    ref_tfec_pGDP = ['Final energy intensity (GJ per thousand 2018 USD PPP)'] + list(ref_tfec_calcs.iloc[1, 1:] / ref_tfec_calcs.iloc[2, 1:])
    ref_tfec_pGDP_series = pd.Series(ref_tfec_pGDP, index = ref_tfec_1.columns)

    ref_tfec_1 = ref_tfec_1.append([ref_tfec_pc_series, ref_tfec_pGDP_series], ignore_index = True).reset_index(drop = True)

    ########## CO2 intensity

    ref_co2int = ref_co2int_2.copy().drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'})

    ref_co2int = ref_co2int[ref_co2int['Series'].isin(['Emissions', 'CO2 intensity'])].reset_index(drop = True)

    ref_co2int_calc1 = ['CO2 intensity (tonnes per thousand 2018 USD PPP)'] + list(ref_co2int.iloc[0, 1:] / ref_tfec_calcs.iloc[2, 1:])
    ref_co2int_calc1_series = pd.Series(ref_co2int_calc1, index = ref_co2int.columns)

    ref_co2int_calc2 = ['CO2 emissions per capita (tonnes per person)'] + list(ref_co2int.iloc[0, 1:] / ref_tfec_calcs.iloc[3, 1:])
    ref_co2int_calc2_series = pd.Series(ref_co2int_calc2, index = ref_co2int.columns)

    ref_co2int = ref_co2int.append([ref_co2int_calc1_series, ref_co2int_calc2_series], ignore_index = True)\
        .reset_index(drop = True)

    ref_co2int.loc[ref_co2int['Series'] == 'Emissions', 'Series'] = 'CO2 emissions (million tonnes)'
    ref_co2int.loc[ref_co2int['Series'] == 'CO2 intensity', 'Series'] = 'CO2 intensity (tonnes per GJ of TPES)'

    ref_co2int = ref_co2int.iloc[[0, 3, 1, 2]].reset_index(drop = True)

    #########################################
    ### PRODUCTION, TRADE, AND SUPPLY
    ## PRODUCTION
    # Coal
    # Oil
    # Gas
    # Nuclear
    # Hydro
    # Non-hydro renewables (split-up?)
    # Other

    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    ref_prod_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = ref_prod_df[ref_prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = ref_prod_df[ref_prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = ref_prod_df[ref_prod_df['fuel_code'].isin(['11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', 
                                                            '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                            '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels'])]\
                                                                .groupby(['item_code_new']).sum().assign(fuel_code = 'Other renewables',
                                                                                                         item_code_new = '1_indigenous_production')
    
    others = ref_prod_df[ref_prod_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other',
                                                                                                     item_code_new = '1_indigenous_production')
    
    ref_prod_1 = ref_prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(ref_prod_df.loc[:, '2000':])].reset_index(drop = True)

    ref_prod_1.loc[ref_prod_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_prod_1.loc[ref_prod_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'
    ref_prod_1.loc[ref_prod_1['fuel_code'] == '10_hydro', 'fuel_code'] = 'Hydro'

    ref_prod_1 = ref_prod_1[ref_prod_1['fuel_code'].isin(['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Other'])]\
        .set_index('fuel_code').loc[['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Other']]\
            .reset_index().replace(np.nan, 0)

    ref_prod_1.loc['Total'] = ref_prod_1.sum(numeric_only = True)

    ref_prod_1.loc['Total', 'fuel_code'] = 'Production (PJ)'
    ref_prod_1.loc['Total', 'item_code_new'] = '1_indigenous_production'

    ref_prod_1 = ref_prod_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[7, 0, 1, 2, 3, 4, 5, 6], :].reset_index(drop = True)


    ## NET IMPORTS
    # Coal
    # Crude oil
    # Oil products
    # Gas
    # Bioenergy
    # Electricity

    ref_nettrade_1 = ref_nettrade_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[8, 0, 1, 6, 3, 7, 5, 2, 4]].reset_index(drop = True)

    ref_nettrade_1.loc[ref_nettrade_1['Series'] == 'Trade balance', 'Series'] = 'Net energy imports (PJ)'

    ## INTERNATIONAL TRANSPORT
    
    # Marine
    ref_bunkers_1.loc['Total'] = ref_bunkers_1.sum(numeric_only = True)
    ref_bunkers_1.loc['Total', 'fuel_code'] = 'Marine'
    ref_bunkers_1.loc['Total', 'item_code_new'] = '4_international_marine_bunkers'    
    
    ref_bunkers_1 = ref_bunkers_1.copy().reset_index(drop = True)

    ref_bunkers_marine = ref_bunkers_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})

    # Aviation
    ref_bunkers_2.loc['Total'] = ref_bunkers_2.sum(numeric_only = True)
    ref_bunkers_2.loc['Total', 'fuel_code'] = 'Aviation'
    ref_bunkers_2.loc['Total', 'item_code_new'] = '5_international_aviation_bunkers'    
    
    ref_bunkers_2 = ref_bunkers_2.copy().reset_index(drop = True)

    ref_bunkers_aviation = ref_bunkers_2.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})

    # aggregate

    ref_bunkers_3 = ref_bunkers_marine[ref_bunkers_marine['Series'] == 'Marine'].copy()\
        .append(ref_bunkers_aviation[ref_bunkers_aviation['Series'] == 'Aviation'].copy()).reset_index(drop = True)

    ref_bunkers_3.loc['Total'] = ref_bunkers_3.sum(numeric_only = True)

    ref_bunkers_3.loc['Total', 'Series'] = 'International transport (PJ)'

    negative = ref_bunkers_3.copy().set_index(['Series']) * -1

    ref_bunkers_3 = negative.reset_index().iloc[[2, 0, 1]].reset_index(drop = True)

    ## STOCK CHANGE (Only really historical but a little bit in early model years)
    # Coal
    # Oil
    # Gas
    # Other?

    ref_stock_1 = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                        (EGEDA_years_reference['item_code_new'].isin(['6_stock_change'])) &
                                        (EGEDA_years_reference['fuel_code'].isin(['1_coal', '2_coal_products', 
                                        '3_peat','4_peat_products', '6_crude_oil_and_ngl', '7_petroleum_products', 
                                        '8_gas']))].copy().replace(np.nan, 0)

    ref_stock_coal = ref_stock_1[ref_stock_1['fuel_code'].isin(['1_coal', '2_coal_products', '3_peat','4_peat_products'])]\
        .groupby(['economy', 'item_code_new']).sum().reset_index()

    ref_stock_coal['fuel_code'] = 'Coal'

    ref_stock_oil = ref_stock_1[ref_stock_1['fuel_code'].isin(['6_crude_oil_and_ngl', '7_petroleum_products'])]\
        .groupby(['economy', 'item_code_new']).sum().reset_index()

    ref_stock_oil['fuel_code'] = 'Oil'

    ref_stock_1 = ref_stock_1.append([ref_stock_coal, ref_stock_oil]).reset_index(drop = True)

    ref_stock_1 = ref_stock_1.drop(['economy', 'item_code_new'], axis = 1).rename(columns = {'fuel_code': 'Series'})

    ref_stock_1 = ref_stock_1[['Series'] + list(ref_stock_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ref_stock_1 = ref_stock_1.set_index('Series').loc[['Coal', 'Oil', '8_gas']].reset_index()

    ref_stock_1.loc[ref_stock_1['Series'] == '8_gas', 'Series'] = 'Gas'

    ref_stock_1.loc['Total'] = ref_stock_1.sum(numeric_only = True)

    ref_stock_1.loc['Total', 'Series'] = 'Stock change (PJ)'
    ref_stock_1 = ref_stock_1.copy().reset_index(drop = True).iloc[[3, 0, 1, 2]].reset_index(drop = True)    

    ## TOTAL PRIMARY ENERGY SUPPLY
    # Coal
    # Oil
    # Gas
    # Nuclear
    # Hydro
    # Non-hydro renewables (split-up?)
    # Other

    ref_tpes_df = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                          (EGEDA_years_reference['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_reference['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '7_total_primary_energy_supply')
    
    oil = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '7_total_primary_energy_supply')
    
    renewables = ref_tpes_df[ref_tpes_df['fuel_code'].isin(['11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', 
                                                            '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                            '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels'])].groupby(['item_code_new']).sum().assign(fuel_code = 'Other renewables',
                                                                                                              item_code_new = '7_total_primary_energy_supply')
    
    others = ref_tpes_df[ref_tpes_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other',
                                                                                                     item_code_new = '7_total_primary_energy_supply')
    
    ref_tpes_1 = ref_tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(ref_tpes_df.loc[:, '2000':])].reset_index(drop = True)

    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '10_hydro', 'fuel_code'] = 'Hydro'
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    ref_tpes_1.loc[ref_tpes_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    ref_tpes_1 = ref_tpes_1[ref_tpes_1['fuel_code'].isin(['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Electricity', 'Hydrogen', 'Other'])]\
        .set_index('fuel_code').loc[['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Electricity', 'Hydrogen', 'Other']].reset_index().replace(np.nan, 0)

    ref_tpes_1.loc['Total'] = ref_tpes_1.sum(numeric_only = True)

    ref_tpes_1.loc['Total', 'fuel_code'] = 'Total primary energy supply (PJ)'
    ref_tpes_1.loc['Total', 'item_code_new'] = '7_total_primary_energy_supply'

    ref_tpes_1 = ref_tpes_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[9, 0, 1, 2, 3, 4, 5, 6, 7, 8], :].reset_index(drop = True)

    ###################################################

    #### TRANSFORMATION
    ### ELECTRICITY and HEAT GENERATION
    ## INPUT FUEL
    # Coal
    # Oil 
    # Gas
    # Nuclear
    # Hydro
    # Non-hydro renewables
    # Other

    ref_powuse = ref_pow_use_2.copy()

    coal_pow = ref_powuse[ref_powuse['FUEL'].isin(['Coal', 'Lignite'])].groupby(['Transformation']).sum().assign(FUEL = 'Coal')

    renew_pow = ref_powuse[ref_powuse['FUEL'].isin(['Solar', 'Wind', 'Biomass', 'Geothermal', 'Other renewables'])].groupby(['Transformation']).sum().assign(FUEL = 'Other renewables')

    ref_powuse = ref_powuse[ref_powuse['FUEL'].isin(['Oil', 'Gas', 'Nuclear', 'Hydro', 'Other'])].copy()\
        .drop('Transformation', axis = 1).reset_index(drop = True)

    ref_powuse = ref_powuse.append([coal_pow, renew_pow]).rename(columns = {'FUEL': 'Series'}).reset_index(drop = True)

    ref_powuse = ref_powuse.iloc[[5, 0, 1, 2, 3, 6, 4]].reset_index(drop = True)

    pos_to_neg = ref_powuse.select_dtypes(include = [np.number]) * -1
    ref_powuse[pos_to_neg.columns] = pos_to_neg

    ref_powuse_rows = ref_powuse.shape[0]
    ref_powuse_cols = ref_powuse.shape[1]

    ## OUTPUT FUEL
    # Electricity
    # Heat

    ref_elecout = ref_elecgen_2[ref_elecgen_2['TECHNOLOGY'] == 'Total'].copy().drop('Generation', axis = 1)\
        .rename(columns = {'TECHNOLOGY': 'Series'}).reset_index(drop = True)

    s = ref_elecout.select_dtypes(include=[np.number]) * 3.6 
    ref_elecout[s.columns] = s

    ref_elecout.loc[ref_elecout['Series'] == 'Total', 'Series'] = 'Electricity'

    ref_heatout = ref_heatgen_2[ref_heatgen_2['TECHNOLOGY'] == 'Total'].copy().drop('Generation', axis = 1)\
        .rename(columns = {'TECHNOLOGY': 'Series'}).reset_index(drop = True)

    ref_heatout.loc[ref_heatout['Series'] == 'Total', 'Series'] = 'Heat'

    ref_elecheat = ref_elecout.append(ref_heatout).reset_index(drop = True)

    ref_elecheat_rows = ref_elecheat.shape[0]
    ref_elecheat_cols = ref_elecheat.shape[1]

    # Sum of input and output
    ref_powsum = ref_powuse.copy().append([pd.Series(ref_powuse.sum() + ref_elecheat.sum())], ignore_index = True).iloc[[7]]\
        .reset_index(drop = True)

    ref_powsum.iloc[0, 0] = 'Electricity and heat generation (PJ)'

    ### REFINERIES
    ## INPUT FUEL
    # Crude oil
    ## OUTPUT FUEL
    # Refined products

    ref_refineryin = ref_crudecons_1[ref_crudecons_1['item_code_new'] == 'Refining'].copy().drop('fuel_code', axis = 1)\
        .rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    ref_refineryin.loc[ref_refineryin['Series'] == 'Refining', 'Series'] = 'Crude oil'

    pos_to_neg = ref_refineryin.select_dtypes(include = [np.number]) * -1
    ref_refineryin[pos_to_neg.columns] = pos_to_neg

    # Output

    ref_refiningout = pd.concat([EGEDA_hist_refiningout[EGEDA_hist_refiningout['economy'] == economy].copy().reset_index(drop = True),\
        ref_refinery_2[ref_refinery_2['FUEL'] == 'Total'][list(ref_refinery_2.loc[:, '2019':'2050'])].copy().reset_index(drop = True)], axis = 1)

    ref_refiningout = ref_refiningout.drop(['economy', 'item_code_new'], axis = 1).rename(columns = {'fuel_code': 'Series'})

    ref_refiningout.loc[ref_refiningout['Series'] == '7_petroleum_products', 'Series'] = 'Petroleum products'

    ref_refining = ref_refineryin.append(ref_refiningout).reset_index(drop = True)

    ref_refining = ref_refining.copy().append(ref_refining.sum(numeric_only = True), ignore_index = True)

    ref_refining.iloc[2, 0] = 'Refineries (PJ)'

    ref_refining = ref_refining.iloc[[2, 0, 1]].reset_index(drop = True)

    ref_refining_1 = ref_refining.copy().iloc[[0]].reset_index(drop = True)
    ref_refining_2 = ref_refining.copy().iloc[[1]].reset_index(drop = True)
    ref_refining_3 = ref_refining.copy().iloc[[2]].reset_index(drop = True)

    ### ENERGY INDUSTRY OWN-USE
    ## INPUT FUEL
    # Coal 
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat 
    # Other

    ref_ownloss = ref_ownuse_1.copy().drop('Sector', axis = 1).rename(columns = {'FUEL': 'Series'}).reset_index(drop = True)

    ref_ownloss.loc[ref_ownloss['Series'] == 'Waste', 'Series'] = 'Other'
    ref_ownloss.loc[ref_ownloss['Series'] == 'Total', 'Series'] = 'Own-use and losses (PJ)'

    ref_ownloss = ref_ownloss.iloc[[7, 0, 1, 2, 3, 4, 5, 6]].reset_index(drop = True)

    pos_to_neg = ref_ownloss.select_dtypes(include = [np.number]) * -1
    ref_ownloss[pos_to_neg.columns] = pos_to_neg

    ref_ownloss_rows = ref_ownloss.shape[0]
    ref_ownloss_cols = ref_ownloss.shape[1]
    
    ### DISTRIBUTION LOSSES
    # Coal
    # Oil
    # Gas
    # Electricity
    # Heat

    ### TRANSFERS

    ref_transtat = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) &
                                         (EGEDA_years_reference['item_code_new'].isin(['8_transfers', '11_statistical_discrepancy'])) &
                                         (EGEDA_years_reference['fuel_code'] == '19_total')]\
                                             .replace(np.nan, 0).drop(['economy', 'fuel_code'], axis = 1)\
                                                 .rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    ref_transtat = ref_transtat[['Series'] + list(ref_transtat.loc[:, '2000': '2050'])].reset_index(drop = True)

    ref_transtat.loc[ref_transtat['Series'] == '8_transfers', 'Series'] = 'Transfers'
    ref_transtat.loc[ref_transtat['Series'] == '11_statistical_discrepancy', 'Series'] = 'Statistical discrepancy'

    ref_transtat_rows = ref_transtat.shape[0]
    ref_transtat_cols = ref_transtat.shape[1]

    ### STATISTICAL DISCREPANCIES (only historical in 7th)

    ######################################################

    ### DEMAND
    ## FED By Sector
    # Agriculture and non-specified (split)
    # Buildings
    # Industry
    # Transport (domestic)
    # Non-energy

    ref_fedsector = ref_fedsector_2.copy().drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'})\
        .iloc[[6, 0, 1, 2, 3, 4, 5]].reset_index(drop = True)

    ref_fedsector.loc[ref_fedsector['Series'] == 'Total', 'Series'] = 'Final energy demand by sector (PJ)'

    ## FED by fuel
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Cooling (lol?)
    # Hydrogen
    # Other

    ref_fedfuel = ref_fedfuel_1.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[9, 0, 1, 2, 4, 3, 5, 6, 7, 8]].reset_index(drop = True)

    ref_fedfuel.loc[ref_fedfuel['Series'] == 'Total', 'Series'] = 'Final energy demand by fuel (PJ)'
    ref_fedfuel.loc[ref_fedfuel['Series'] == 'Others', 'Series'] = 'Other'

    ## AGRICULTURE and NON-SPECIFIED (split?)
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ## BUILDINGS 
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ref_build = ref_bld_2.copy()

    renew = ref_build[ref_build['fuel_code'].isin(['Other renewables', 'Biomass'])].copy().groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Renewables').reset_index(drop = True)

    ref_build = ref_build[ref_build['fuel_code'].isin(['Coal', 'Oil', 'Gas', 'Hydrogen', 'Electricity', 'Heat', 'Others', 'Total'])]\
        .append([renew]).drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    ref_build.loc[ref_build['Series'] == 'Others', 'Series'] = 'Other'
    ref_build.loc[ref_build['Series'] == 'Total', 'Series'] = 'Buildings (PJ)'

    ref_build = ref_build.iloc[[7, 0, 1, 2, 8, 3, 4, 5, 6]].reset_index(drop = True)

    ## INDUSTRY
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ref_ind = ref_ind_2.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    ref_ind.loc[ref_ind['Series'] == 'Others', 'Series'] = 'Other'
    ref_ind.loc[ref_ind['Series'] == 'Total', 'Series'] = 'Industry (PJ)'
    ref_ind.loc[ref_ind['Series'] == 'Others', 'Series'] = 'Other'

    ref_ind = ref_ind.iloc[[8, 0, 1, 2, 3, 4, 5, 6, 7]].reset_index(drop = True)

    ## TRANSPORT
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ref_trn = ref_trn_1.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    ref_trn.loc[ref_trn['Series'] == 'Total', 'Series'] = 'Transport (PJ)'

    ref_trn = ref_trn.iloc[[9, 0, 1, 2, 3, 4, 5, 6, 7, 8]].reset_index(drop = True)

    ## NON-ENERGY
    # Coal
    # Oil
    # Gas

    ref_nonen = EGEDA_years_reference[(EGEDA_years_reference['economy'] == economy) & 
                                      (EGEDA_years_reference['item_code_new'] == '17_nonenergy_use') &
                                      (EGEDA_years_reference['fuel_code'].isin(['1_coal', '2_coal_products',
                                       '3_peat', '4_peat_products', '6_crude_oil_and_ngl',
                                       '7_petroleum_products', '8_gas', '19_total']))].loc[:, 'fuel_code':].copy().reset_index(drop = True)

    coal_ne = ref_nonen[ref_nonen['fuel_code'].isin(['1_coal', '2_coal_products', '3_peat', '4_peat_products'])].copy()\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Coal').reset_index(drop = True)

    oil_ne = ref_nonen[ref_nonen['fuel_code'].isin(['6_crude_oil_and_ngl', '7_petroleum_products'])].copy()\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Oil').reset_index(drop = True)

    gas_ne = ref_nonen[ref_nonen['fuel_code'].isin(['8_gas'])].copy()\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Gas').reset_index(drop = True)

    ref_nonen = ref_nonen.append([coal_ne, oil_ne, gas_ne]).reset_index(drop = True)

    ref_nonen = ref_nonen[ref_nonen['fuel_code'].isin(['Coal', 'Oil', 'Gas', '19_total'])].drop('item_code_new', axis = 1)\
        .rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    ref_nonen.loc[ref_nonen['Series'] == '19_total', 'Series'] = 'Non-energy (PJ)'

    ref_nonen = ref_nonen[['Series'] + list(ref_stock_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ###############################################################

    #### ELECTRICITY
    ### CAPACITY

    ref_powcap = pd.DataFrame(['Coal', 'Coal CCS',\
        'Lignite', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Bio', 'Wind', 'Solar', 'Geothermal',\
            'Waste', 'Storage', 'Other'], columns = ['TECHNOLOGY']).merge(ref_powcap_1.copy(),\
                on = 'TECHNOLOGY', how = 'outer').replace(np.nan, 0).rename(columns = {'TECHNOLOGY': 'Series'})\
                    .reset_index(drop = True)

    ref_powcap.loc[ref_powcap['Series'] == 'Total', 'Series'] = 'Total capacity (GW)' 

    ref_powcap = ref_powcap.iloc[[15, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]].reset_index(drop = True)

    ###################################
    ### GENERATION 

    ref_elecgen = ref_elecgen_2.copy().drop('Generation', axis = 1).rename(columns = {'TECHNOLOGY': 'Series'})

    ref_elecgen.loc[ref_elecgen['Series'] == 'Total', 'Series'] = 'Electricity generation (TWh)'

    ref_elecgen = ref_elecgen.iloc[[16, 0, 1, 2, 4, 5, 3, 7, 6, 10, 8, 9, 11, 12, 13, 14, 15]].reset_index(drop = True)

    ######################################
    ### EMISSIONS
    ## BY FUEL
    # Coal
    # Oil
    # Gas

    ref_emfuel = ref_emiss_fuel_1.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .reset_index(drop = True)

    ref_emfuel.loc[ref_emfuel['Series'] == 'Heat & others', 'Series'] = 'Other'
    ref_emfuel.loc[ref_emfuel['Series'] == 'Total', 'Series'] = 'Energy sector CO2 emissions (million tonnes)'

    ref_emfuel = ref_emfuel[ref_emfuel['Series']\
        .isin(['Coal', 'Oil', 'Gas', 'Other', 'Energy sector CO2 emissions (million tonnes)'])]\
            .reset_index(drop = True)

    ref_emfuel = ref_emfuel.iloc[[4, 0, 1, 2, 3]].reset_index(drop = True)

    ## BY SECTOR
    # Power
    # Own-use and losses
    # Agriculture 
    # Buildings
    # Industry
    # Transport (domestic)
    # Non-energy
    # Non-specified

    ref_emsector = ref_emiss_sector_1.copy().drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'})\
        .reset_index(drop = True)

    ref_emsector.loc[ref_emsector['Series'] == 'Total', 'Series'] = 'CO2 emissions by sector (million tonnes)'

    ref_emsector = ref_emsector.iloc[[7, 0, 1, 2, 3, 4, 5, 6]].reset_index(drop = True)

    ######################## Captured emissions

    # emissions factors
    hyd_coal = 0.094686
    hyd_gas = 0.056151

    ref_capemiss = ref_capemiss_df1[ref_capemiss_df1['REGION'] == economy].drop(['REGION', 'EMISSION'], axis = 1)\
        .reset_index(drop = True)

    # Coal
    ref_hydcoal = ref_hydrogen_1[ref_hydrogen_1['Technology'] == 'Coal gasification CCS'].copy().reset_index(drop = True)
    
    ref_hydcoal_captured = ref_hydcoal.select_dtypes(include = [np.number]) * hyd_coal

    ref_hydcoal[ref_hydcoal_captured.columns] = ref_hydcoal_captured
    ref_hydcoal = ref_hydcoal.rename(columns = {'Fuel': 'Series'})

    # Gas
    ref_hydgas = ref_hydrogen_1[ref_hydrogen_1['Technology'] == 'Steam methane reforming CCS'].copy().reset_index(drop = True)
    
    ref_hydgas_captured = ref_hydgas.select_dtypes(include = [np.number]) * hyd_gas

    ref_hydgas[ref_hydgas_captured.columns] = ref_hydgas_captured
    ref_hydgas = ref_hydgas.rename(columns = {'Fuel': 'Series'})

    ref_hydcap = pd.DataFrame(columns = ref_emsector.columns).append([ref_hydcoal, ref_hydgas])\
        .replace(np.nan, 0).drop('Technology', axis = 1).reset_index(drop = True)

    if ref_hydcap.empty:
        new_row = ['Hydrogen'] + [0] * 51
        new_series = pd.Series(new_row, index = ref_emsector.columns)
        ref_hydcap = ref_hydcap.append(new_series, ignore_index = True)

    else:
        ref_hydcap = ref_hydcap.copy().groupby('Series').sum().reset_index()

    ######################################################

    industry = ref_capemiss[ref_capemiss['TECHNOLOGY'].str.startswith('IND')].reset_index(drop = True)
    industry['Series'] = 'Industry'

    industry = industry.copy().drop('TECHNOLOGY', axis = 1).groupby('Series').sum().reset_index()

    ref_industry_cap = pd.DataFrame(columns = ref_emsector.columns)

    if industry.empty:
        new_row = ['Industry'] + [0] * 51
        new_series = pd.Series(new_row, index = ref_emsector.columns)
        ref_industry_cap = ref_industry_cap.append(new_series, ignore_index = True)

    else:
        ref_industry_cap = ref_industry_cap.append(industry).replace(np.nan, 0).reset_index(drop = True)

    power = ref_capemiss[ref_capemiss['TECHNOLOGY'].str.startswith('POW')].reset_index(drop = True)
    power['Series'] = 'Power'

    power = power.copy().drop('TECHNOLOGY', axis = 1).groupby('Series').sum().reset_index()

    ref_power_cap = pd.DataFrame(columns = ref_emsector.columns)

    if power.empty:
        new_row = ['Power'] + [0] * 51
        new_series = pd.Series(new_row, index = ref_emsector.columns)
        ref_power_cap = ref_power_cap.append(new_series, ignore_index = True)

    else:
        ref_power_cap = ref_power_cap.append(power).replace(np.nan, 0).reset_index(drop = True)
    
    ownuse = ref_capemiss[ref_capemiss['TECHNOLOGY'].str.startswith('OWN')].reset_index(drop = True)
    ownuse['Series'] = 'Own-use'

    ownuse = ownuse.copy().drop('TECHNOLOGY', axis = 1).groupby('Series').sum().reset_index()

    ref_ownuse_cap = pd.DataFrame(columns = ref_emsector.columns)

    if ownuse.empty:
        new_row = ['Own-use'] + [0] * 51
        new_series = pd.Series(new_row, index = ref_emsector.columns)
        ref_ownuse_cap = ref_ownuse_cap.append(new_series, ignore_index = True)

    else:
        ref_ownuse_cap = ref_ownuse_cap.append(ownuse).replace(np.nan, 0).reset_index(drop = True)

    # Captured emissions

    ref_captured = ref_industry_cap.append([ref_hydcap, ref_power_cap, ref_ownuse_cap]).reset_index(drop = True)

    cols = ref_captured.columns[1:]

    ref_captured[cols] = ref_captured[cols].apply(pd.to_numeric)

    ref_captured.loc['Total'] = ref_captured.sum(numeric_only = True)

    ref_captured.loc['Total', 'Series'] = 'Captured CO2 emissions (million tonnes)'
    ref_captured = ref_captured.copy().reset_index(drop = True).iloc[[4, 0, 1, 2, 3]].reset_index(drop = True)

    ref_captured_rows = ref_captured.shape[0]

    # Modern renewables

    ref_modren = ref_modren_4[ref_modren_4['item_code_new'].isin(['Total', 'Reference'])].copy()\
        .drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    ref_modren.loc[ref_modren['Series'] == 'Total', 'Series'] = 'Modern renewables in FED (PJ)'
    ref_modren.loc[ref_modren['Series'] == 'Reference', 'Series'] = 'Modern renewables share of final energy demand'

    # More comprehensive modern renewables breakdown

    ref_modren_breakdown = ref_modren_4.copy().iloc[[6, 0, 1, 2, 3, 4, 5, 9, 11]].drop('fuel_code', axis = 1)\
        .rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    ref_modren_breakdown.loc[ref_modren_breakdown['Series'] == 'Electricity and heat TFEC', 'Series'] = 'Electricity and heat (not including own-use and losses)'
    ref_modren_breakdown.loc[ref_modren_breakdown['Series'] == 'Total', 'Series'] = 'Modern renewables in FED (PJ)'
    ref_modren_breakdown.loc[ref_modren_breakdown['Series'] == 'TFEC', 'Series'] = 'Final energy demand (PJ)'
    ref_modren_breakdown.loc[ref_modren_breakdown['Series'] == 'Reference', 'Series'] = 'Modern renewables share of FED'

    ref_modren_gen = ref_modren_4.copy().iloc[7:9].drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .reset_index(drop = True)

    ref_modren_gen.loc[ref_modren_gen['Series'] == 'Modern renewables', 'Series'] = 'Renewable generation (TWh)'
    ref_modren_gen.loc[ref_modren_gen['Series'] == 'Total', 'Series'] = 'Total generation (TWh)'

    cols = ref_modren_gen.columns[1:]

    ref_modren_gen[cols] = ref_modren_gen[cols] / 3.6

    ref_modren_gencalc = ['Renewable generation share'] + list(ref_modren_gen.iloc[0, 1:] / ref_modren_gen.iloc[1, 1:])
    ref_modren_gencalc_series = pd.Series(ref_modren_gencalc, index = ref_modren_gen.columns)

    ref_modren_gen = ref_modren_gen.copy().append(ref_modren_gencalc_series, ignore_index = True).reset_index(drop = True)

    # Join the data
    ref_modren_A = ref_modren_breakdown.append(ref_modren_gen).reset_index(drop = True)

    ref_modren_A_rows = ref_modren_A.shape[0]

    # Join relevant dataframes together

    ref_top_df = macro_1.append([ref_tpes_3, ref_tfec_1, ref_co2int, ref_modren]).reset_index(drop = True)

    ref_top_rows = ref_top_df.shape[0]
    ref_top_cols = ref_top_df.shape[1]

    ref_supply_df = ref_prod_1.append([ref_nettrade_1, ref_bunkers_3, ref_stock_1, ref_tpes_1]).reset_index(drop = True)

    ref_supply_rows = ref_supply_df.shape[0]
    ref_supply_cols = ref_supply_df.shape[1]

    # ref_powsum
    # 'Input fuel'
    # ref_powuse 
    # 'Output fuel'
    # ref_elecheat
    # ref_refining_1
    # 'Input fuel'
    # ref_refining_2
    # 'Output fuel'
    # ref_refining_3
    # ref_ownloss
    # ref_transtat

    ref_demand_df = ref_fedsector.copy().append([ref_fedfuel, ref_ind, ref_trn, ref_build, ref_nonen]).reset_index(drop = True)

    ref_demand_rows = ref_demand_df.shape[0]
    ref_demand_cols = ref_demand_df.shape[1]

    ref_electricity = ref_elecgen.append(ref_powcap).reset_index(drop = True)
    # ref_powcap
    # ref_elecgen

    ref_electricity_rows = ref_electricity.shape[0]

    # ref_emfuel
    # 'By sector'
    # ref_emsector

    ref_emfuel_rows = ref_emfuel.shape[0]
    ref_emsector_rows = ref_emsector.shape[0]

    ################################################################################################################

    ## TPES
    # TPES per capita
    # TPES per GDP

    netz_tpes_3 = netz_tpes_1[netz_tpes_1['fuel_code'] == 'Total'].copy().\
        rename(columns = {'item_code_new': 'Series'}).iloc[:, 1:].reset_index(drop = True)

    netz_tpes_3.loc[netz_tpes_3['Series'] == '7_total_primary_energy_supply', 'Series'] = 'Total primary energy supply (PJ)'

    netz_tpes_calcs = netz_tpes_3.append(macro_1[macro_1['Series'].isin(['Population (millions)', 'GDP (2018 USD billion PPP)'])])\
        .copy().reset_index(drop = True)

    netz_tpes_pc = ['TPES per capita (GJ per person)'] + list(netz_tpes_calcs.iloc[0, 1:] / netz_tpes_calcs.iloc[2, 1:])
    netz_tpes_pc_series = pd.Series(netz_tpes_pc, index = netz_tpes_3.columns)

    netz_tpes_pGDP = ['TPES per GDP (GJ per thousand 2018 USD PPP)'] + list(netz_tpes_calcs.iloc[0, 1:] / netz_tpes_calcs.iloc[1, 1:])
    netz_tpes_pGDP_series = pd.Series(netz_tpes_pGDP, index = netz_tpes_3.columns)

    netz_tpes_3 = netz_tpes_3.append([netz_tpes_pc_series, netz_tpes_pGDP_series], ignore_index = True).reset_index(drop = True)

    ## FED
    # Final energy intensity per capita
    # Final energy intensity per GDP

    netz_tfec_1 = netz_tfec_1.copy().rename(columns = {'item_code_new': 'Series'}).iloc[:, 1:]

    netz_tfec_1 = netz_fedfuel_1[netz_fedfuel_1['fuel_code'] == 'Total'].copy()\
        .rename(columns = {'item_code_new': 'Series'}).iloc[:, 1:].append(netz_tfec_1)

    netz_tfec_1.loc[netz_tfec_1['Series'] == '12_total_final_consumption', 'Series'] = 'Total final consumption (PJ)'

    netz_tfec_1.loc[netz_tfec_1['Series'] == 'TFEC', 'Series'] = 'Final energy demand (PJ)'

    netz_tfec_calcs = netz_tfec_1.append(macro_1[macro_1['Series'].isin(['Population (millions)', 'GDP (2018 USD billion PPP)'])])\
        .copy().reset_index(drop = True)

    netz_tfec_pc = ['Final energy demand per capita (GJ per person)'] + list(netz_tfec_calcs.iloc[1, 1:] / netz_tfec_calcs.iloc[3, 1:])
    netz_tfec_pc_series = pd.Series(netz_tfec_pc, index = netz_tfec_1.columns)

    netz_tfec_pGDP = ['Final energy intensity (MJ per 2018 USD PPP)'] + list(netz_tfec_calcs.iloc[1, 1:] / netz_tfec_calcs.iloc[2, 1:])
    netz_tfec_pGDP_series = pd.Series(netz_tfec_pGDP, index = netz_tfec_1.columns)

    netz_tfec_1 = netz_tfec_1.append([netz_tfec_pc_series, netz_tfec_pGDP_series], ignore_index = True).reset_index(drop = True)

    ########## CO2 intensity

    netz_co2int = netz_co2int_2.copy().drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'})

    netz_co2int = netz_co2int[netz_co2int['Series'].isin(['Emissions', 'CO2 intensity'])].reset_index(drop = True)

    netz_co2int_calc1 = ['CO2 intensity (tonnes per thousand 2018 USD PPP)'] + list(netz_co2int.iloc[0, 1:] / netz_tfec_calcs.iloc[2, 1:])
    netz_co2int_calc1_series = pd.Series(netz_co2int_calc1, index = netz_co2int.columns)

    netz_co2int_calc2 = ['CO2 emissions per capita (tonnes per person)'] + list(netz_co2int.iloc[0, 1:] / netz_tfec_calcs.iloc[3, 1:])
    netz_co2int_calc2_series = pd.Series(netz_co2int_calc2, index = netz_co2int.columns)

    netz_co2int = netz_co2int.append([netz_co2int_calc1_series, netz_co2int_calc2_series], ignore_index = True)\
        .reset_index(drop = True)

    netz_co2int.loc[netz_co2int['Series'] == 'Emissions', 'Series'] = 'CO2 emissions (million tonnes)'
    netz_co2int.loc[netz_co2int['Series'] == 'CO2 intensity', 'Series'] = 'CO2 intensity (tonnes per GJ of TPES)'

    netz_co2int = netz_co2int.iloc[[0, 3, 1, 2]].reset_index(drop = True)

    #########################################
    ### PRODUCTION, TRADE, AND SUPPLY
    ## PRODUCTION
    # Coal
    # Oil
    # Gas
    # Nuclear
    # Hydro
    # Non-hydro renewables (split-up?)
    # Other

    # Second data frame: production (and also fifth and seventh data frames with slight tweaks)
    netz_prod_df = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'] == '1_indigenous_production') &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = netz_prod_df[netz_prod_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '1_indigenous_production')
    
    oil = netz_prod_df[netz_prod_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '1_indigenous_production')
    
    renewables = netz_prod_df[netz_prod_df['fuel_code'].isin(['11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', 
                                                            '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                            '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels'])]\
                                                                .groupby(['item_code_new']).sum().assign(fuel_code = 'Other renewables',
                                                                                                         item_code_new = '1_indigenous_production')
    
    others = netz_prod_df[netz_prod_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other',
                                                                                                     item_code_new = '1_indigenous_production')
    
    netz_prod_1 = netz_prod_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(netz_prod_df.loc[:, '2000':])].reset_index(drop = True)

    netz_prod_1.loc[netz_prod_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_prod_1.loc[netz_prod_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'
    netz_prod_1.loc[netz_prod_1['fuel_code'] == '10_hydro', 'fuel_code'] = 'Hydro'

    netz_prod_1 = netz_prod_1[netz_prod_1['fuel_code'].isin(['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Other'])]\
        .set_index('fuel_code').loc[['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Other']]\
            .reset_index().replace(np.nan, 0)

    netz_prod_1.loc['Total'] = netz_prod_1.sum(numeric_only = True)

    netz_prod_1.loc['Total', 'fuel_code'] = 'Production (PJ)'
    netz_prod_1.loc['Total', 'item_code_new'] = '1_indigenous_production'

    netz_prod_1 = netz_prod_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[7, 0, 1, 2, 3, 4, 5, 6], :].reset_index(drop = True)


    ## NET IMPORTS
    # Coal
    # Crude oil
    # Oil products
    # Gas
    # Bioenergy
    # Electricity

    netz_nettrade_1 = netz_nettrade_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[8, 0, 1, 6, 3, 7, 5, 2, 4]].reset_index(drop = True)

    netz_nettrade_1.loc[netz_nettrade_1['Series'] == 'Trade balance', 'Series'] = 'Net energy imports (PJ)'

    ## INTERNATIONAL TRANSPORT
    
    # Marine
    netz_bunkers_1.loc['Total'] = netz_bunkers_1.sum(numeric_only = True)
    netz_bunkers_1.loc['Total', 'fuel_code'] = 'Marine'
    netz_bunkers_1.loc['Total', 'item_code_new'] = '4_international_marine_bunkers'    
    
    netz_bunkers_1 = netz_bunkers_1.copy().reset_index(drop = True)

    netz_bunkers_marine = netz_bunkers_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})

    # Aviation
    netz_bunkers_2.loc['Total'] = netz_bunkers_2.sum(numeric_only = True)
    netz_bunkers_2.loc['Total', 'fuel_code'] = 'Aviation'
    netz_bunkers_2.loc['Total', 'item_code_new'] = '5_international_aviation_bunkers'    
    
    netz_bunkers_2 = netz_bunkers_2.copy().reset_index(drop = True)

    netz_bunkers_aviation = netz_bunkers_2.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})

    # aggregate

    netz_bunkers_3 = netz_bunkers_marine[netz_bunkers_marine['Series'] == 'Marine'].copy()\
        .append(netz_bunkers_aviation[netz_bunkers_aviation['Series'] == 'Aviation'].copy()).reset_index(drop = True)

    netz_bunkers_3.loc['Total'] = netz_bunkers_3.sum(numeric_only = True)

    netz_bunkers_3.loc['Total', 'Series'] = 'International transport (PJ)'

    negative = netz_bunkers_3.copy().set_index(['Series']) * -1

    netz_bunkers_3 = negative.reset_index().iloc[[2, 0, 1]].reset_index(drop = True)

    ## STOCK CHANGE (Only really historical but a little bit in early model years)
    # Coal
    # Oil
    # Gas
    # Other?

    netz_stock_1 = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                        (EGEDA_years_netzero['item_code_new'].isin(['6_stock_change'])) &
                                        (EGEDA_years_netzero['fuel_code'].isin(['1_coal', '2_coal_products', 
                                        '3_peat','4_peat_products', '6_crude_oil_and_ngl', '7_petroleum_products', 
                                        '8_gas']))].copy().replace(np.nan, 0)

    netz_stock_coal = netz_stock_1[netz_stock_1['fuel_code'].isin(['1_coal', '2_coal_products', '3_peat','4_peat_products'])]\
        .groupby(['economy', 'item_code_new']).sum().reset_index()

    netz_stock_coal['fuel_code'] = 'Coal'

    netz_stock_oil = netz_stock_1[netz_stock_1['fuel_code'].isin(['6_crude_oil_and_ngl', '7_petroleum_products'])]\
        .groupby(['economy', 'item_code_new']).sum().reset_index()

    netz_stock_oil['fuel_code'] = 'Oil'

    netz_stock_1 = netz_stock_1.append([netz_stock_coal, netz_stock_oil]).reset_index(drop = True)

    netz_stock_1 = netz_stock_1.drop(['economy', 'item_code_new'], axis = 1).rename(columns = {'fuel_code': 'Series'})

    netz_stock_1 = netz_stock_1[['Series'] + list(netz_stock_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    netz_stock_1 = netz_stock_1.set_index('Series').loc[['Coal', 'Oil', '8_gas']].reset_index()

    netz_stock_1.loc[netz_stock_1['Series'] == '8_gas', 'Series'] = 'Gas'

    netz_stock_1.loc['Total'] = netz_stock_1.sum(numeric_only = True)

    netz_stock_1.loc['Total', 'Series'] = 'Stock change (PJ)'
    netz_stock_1 = netz_stock_1.copy().reset_index(drop = True).iloc[[3, 0, 1, 2]].reset_index(drop = True)    

    ## TOTAL PRIMARY ENERGY SUPPLY
    # Coal
    # Oil
    # Gas
    # Nuclear
    # Hydro
    # Non-hydro renewables (split-up?)
    # Other

    netz_tpes_df = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                          (EGEDA_years_netzero['item_code_new'] == '7_total_primary_energy_supply') &
                          (EGEDA_years_netzero['fuel_code'].isin(Required_fuels))].loc[:, 'fuel_code':]

    coal = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Coal_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Coal',
                                                                                                  item_code_new = '7_total_primary_energy_supply')
    
    oil = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Oil_fuels)].groupby(['item_code_new']).sum().assign(fuel_code = 'Oil',
                                                                                                item_code_new = '7_total_primary_energy_supply')
    
    renewables = netz_tpes_df[netz_tpes_df['fuel_code'].isin(['11_geothermal', '12_solar', '13_tide_wave_ocean', '14_wind', '15_solid_biomass', 
                                                            '16_1_biogas', '16_3_municipal_solid_waste_renewable', '16_5_biogasoline', 
                                                            '16_6_biodiesel', '16_7_bio_jet_kerosene', '16_8_other_liquid_biofuels'])].groupby(['item_code_new']).sum().assign(fuel_code = 'Other renewables',
                                                                                                              item_code_new = '7_total_primary_energy_supply')
    
    others = netz_tpes_df[netz_tpes_df['fuel_code'].isin(Other_fuels_TPES)].groupby(['item_code_new']).sum().assign(fuel_code = 'Other',
                                                                                                     item_code_new = '7_total_primary_energy_supply')
    
    netz_tpes_1 = netz_tpes_df.append([coal, oil, renewables, others])[['fuel_code', 
                                                                'item_code_new'] + list(netz_tpes_df.loc[:, '2000':])].reset_index(drop = True)

    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '8_gas', 'fuel_code'] = 'Gas'
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '9_nuclear', 'fuel_code'] = 'Nuclear'
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '10_hydro', 'fuel_code'] = 'Hydro'
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '17_electricity', 'fuel_code'] = 'Electricity'
    netz_tpes_1.loc[netz_tpes_1['fuel_code'] == '16_x_hydrogen', 'fuel_code'] = 'Hydrogen'

    netz_tpes_1 = netz_tpes_1[netz_tpes_1['fuel_code'].isin(['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Electricity', 'Hydrogen', 'Other'])]\
        .set_index('fuel_code').loc[['Coal', 'Oil', 'Gas', 'Nuclear', 'Hydro', 'Other renewables', 'Electricity', 'Hydrogen', 'Other']].reset_index().replace(np.nan, 0)

    netz_tpes_1.loc['Total'] = netz_tpes_1.sum(numeric_only = True)

    netz_tpes_1.loc['Total', 'fuel_code'] = 'Total primary energy supply (PJ)'
    netz_tpes_1.loc['Total', 'item_code_new'] = '7_total_primary_energy_supply'

    netz_tpes_1 = netz_tpes_1.drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[9, 0, 1, 2, 3, 4, 5, 6, 7, 8], :].reset_index(drop = True)

    ###################################################

    #### TRANSFORMATION
    ### ELECTRICITY and HEAT GENERATION
    ## INPUT FUEL
    # Coal
    # Oil 
    # Gas
    # Nuclear
    # Hydro
    # Non-hydro renewables
    # Other

    netz_powuse = netz_pow_use_2.copy()

    coal_pow = netz_powuse[netz_powuse['FUEL'].isin(['Coal', 'Lignite'])].groupby(['Transformation']).sum().assign(FUEL = 'Coal')

    renew_pow = netz_powuse[netz_powuse['FUEL'].isin(['Solar', 'Wind', 'Biomass', 'Geothermal', 'Other renewables'])].groupby(['Transformation']).sum().assign(FUEL = 'Other renewables')

    netz_powuse = netz_powuse[netz_powuse['FUEL'].isin(['Oil', 'Gas', 'Nuclear', 'Hydro', 'Other'])].copy()\
        .drop('Transformation', axis = 1).reset_index(drop = True)

    netz_powuse = netz_powuse.append([coal_pow, renew_pow]).rename(columns = {'FUEL': 'Series'}).reset_index(drop = True)

    netz_powuse = netz_powuse.iloc[[5, 0, 1, 2, 3, 6, 4]].reset_index(drop = True)

    pos_to_neg = netz_powuse.select_dtypes(include = [np.number]) * -1
    netz_powuse[pos_to_neg.columns] = pos_to_neg

    netz_powuse_rows = netz_powuse.shape[0]
    netz_powuse_cols = netz_powuse.shape[1]

    ## OUTPUT FUEL
    # Electricity
    # Heat

    netz_elecout = netz_elecgen_2[netz_elecgen_2['TECHNOLOGY'] == 'Total'].copy().drop('Generation', axis = 1)\
        .rename(columns = {'TECHNOLOGY': 'Series'}).reset_index(drop = True)

    s = netz_elecout.select_dtypes(include=[np.number]) * 3.6 
    netz_elecout[s.columns] = s

    netz_elecout.loc[netz_elecout['Series'] == 'Total', 'Series'] = 'Electricity'

    netz_heatout = netz_heatgen_2[netz_heatgen_2['TECHNOLOGY'] == 'Total'].copy().drop('Generation', axis = 1)\
        .rename(columns = {'TECHNOLOGY': 'Series'}).reset_index(drop = True)

    netz_heatout.loc[netz_heatout['Series'] == 'Total', 'Series'] = 'Heat'

    netz_elecheat = netz_elecout.append(netz_heatout).reset_index(drop = True)

    netz_elecheat_rows = netz_elecheat.shape[0]
    netz_elecheat_cols = netz_elecheat.shape[1]

    # Sum of input and output
    netz_powsum = netz_powuse.copy().append([pd.Series(netz_powuse.sum() + netz_elecheat.sum())], ignore_index = True).iloc[[7]]\
        .reset_index(drop = True)

    netz_powsum.iloc[0, 0] = 'Electricity and heat generation (PJ)'

    ### REFINERIES
    ## INPUT FUEL
    # Crude oil
    ## OUTPUT FUEL
    # Refined products

    netz_refineryin = netz_crudecons_1[netz_crudecons_1['item_code_new'] == 'Refining'].copy().drop('fuel_code', axis = 1)\
        .rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    netz_refineryin.loc[netz_refineryin['Series'] == 'Refining', 'Series'] = 'Crude oil'

    pos_to_neg = netz_refineryin.select_dtypes(include = [np.number]) * -1
    netz_refineryin[pos_to_neg.columns] = pos_to_neg

    # Output

    netz_refiningout = pd.concat([EGEDA_hist_refiningout[EGEDA_hist_refiningout['economy'] == economy].copy().reset_index(drop = True),\
        netz_refinery_2[netz_refinery_2['FUEL'] == 'Total'][list(netz_refinery_2.loc[:, '2019':'2050'])].copy().reset_index(drop = True)], axis = 1)

    netz_refiningout = netz_refiningout.drop(['economy', 'item_code_new'], axis = 1).rename(columns = {'fuel_code': 'Series'})

    netz_refiningout.loc[netz_refiningout['Series'] == '7_petroleum_products', 'Series'] = 'Petroleum products'

    netz_refining = netz_refineryin.append(netz_refiningout).reset_index(drop = True)

    netz_refining = netz_refining.copy().append(netz_refining.sum(numeric_only = True), ignore_index = True)

    netz_refining.iloc[2, 0] = 'Refineries (PJ)'

    netz_refining = netz_refining.iloc[[2, 0, 1]].reset_index(drop = True)

    netz_refining_1 = netz_refining.copy().iloc[[0]].reset_index(drop = True)
    netz_refining_2 = netz_refining.copy().iloc[[1]].reset_index(drop = True)
    netz_refining_3 = netz_refining.copy().iloc[[2]].reset_index(drop = True)

    ### ENERGY INDUSTRY OWN-USE
    ## INPUT FUEL
    # Coal 
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat 
    # Other

    netz_ownloss = netz_ownuse_1.copy().drop('Sector', axis = 1).rename(columns = {'FUEL': 'Series'}).reset_index(drop = True)

    netz_ownloss.loc[netz_ownloss['Series'] == 'Waste', 'Series'] = 'Other'
    netz_ownloss.loc[netz_ownloss['Series'] == 'Total', 'Series'] = 'Own-use and losses (PJ)'

    netz_ownloss = netz_ownloss.iloc[[7, 0, 1, 2, 3, 4, 5, 6]].reset_index(drop = True)

    pos_to_neg = netz_ownloss.select_dtypes(include = [np.number]) * -1
    netz_ownloss[pos_to_neg.columns] = pos_to_neg

    netz_ownloss_rows = netz_ownloss.shape[0]
    netz_ownloss_cols = netz_ownloss.shape[1]
    
    ### DISTRIBUTION LOSSES
    # Coal
    # Oil
    # Gas
    # Electricity
    # Heat

    ### TRANSFERS

    netz_transtat = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) &
                                         (EGEDA_years_netzero['item_code_new'].isin(['8_transfers', '11_statistical_discrepancy'])) &
                                         (EGEDA_years_netzero['fuel_code'] == '19_total')]\
                                             .replace(np.nan, 0).drop(['economy', 'fuel_code'], axis = 1)\
                                                 .rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    netz_transtat = netz_transtat[['Series'] + list(netz_transtat.loc[:, '2000': '2050'])].reset_index(drop = True)

    netz_transtat.loc[netz_transtat['Series'] == '8_transfers', 'Series'] = 'Transfers'
    netz_transtat.loc[netz_transtat['Series'] == '11_statistical_discrepancy', 'Series'] = 'Statistical discrepancy'

    netz_transtat_rows = netz_transtat.shape[0]
    netz_transtat_cols = netz_transtat.shape[1]

    ### STATISTICAL DISCREPANCIES (only historical in 7th)

    ######################################################

    ### DEMAND
    ## FED By Sector
    # Agriculture and non-specified (split)
    # Buildings
    # Industry
    # Transport (domestic)
    # Non-energy

    netz_fedsector = netz_fedsector_2.copy().drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'})\
        .iloc[[6, 0, 1, 2, 3, 4, 5]].reset_index(drop = True)

    netz_fedsector.loc[netz_fedsector['Series'] == 'Total', 'Series'] = 'Final energy demand by sector (PJ)'

    ## FED by fuel
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Cooling (lol?)
    # Hydrogen
    # Other

    netz_fedfuel = netz_fedfuel_1.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .iloc[[9, 0, 1, 2, 4, 3, 5, 6, 7, 8]].reset_index(drop = True)

    netz_fedfuel.loc[netz_fedfuel['Series'] == 'Total', 'Series'] = 'Final energy demand by fuel (PJ)'
    netz_fedfuel.loc[netz_fedfuel['Series'] == 'Others', 'Series'] = 'Other'

    ## AGRICULTURE and NON-SPECIFIED (split?)
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    ## BUILDINGS 
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    netz_build = netz_bld_2.copy()

    renew = netz_build[netz_build['fuel_code'].isin(['Other renewables', 'Biomass'])].copy().groupby(['item_code_new'])\
        .sum().assign(fuel_code = 'Renewables').reset_index(drop = True)

    netz_build = netz_build[netz_build['fuel_code'].isin(['Coal', 'Oil', 'Gas', 'Hydrogen', 'Electricity', 'Heat', 'Others', 'Total'])]\
        .append([renew]).drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    netz_build.loc[netz_build['Series'] == 'Others', 'Series'] = 'Other'
    netz_build.loc[netz_build['Series'] == 'Total', 'Series'] = 'Buildings (PJ)'

    netz_build = netz_build.iloc[[7, 0, 1, 2, 8, 3, 4, 5, 6]].reset_index(drop = True)

    ## INDUSTRY
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    netz_ind = netz_ind_2.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    netz_ind.loc[netz_ind['Series'] == 'Others', 'Series'] = 'Other'
    netz_ind.loc[netz_ind['Series'] == 'Total', 'Series'] = 'Industry (PJ)'
    netz_ind.loc[netz_ind['Series'] == 'Others', 'Series'] = 'Other'

    netz_ind = netz_ind.iloc[[8, 0, 1, 2, 3, 4, 5, 6, 7]].reset_index(drop = True)

    ## TRANSPORT
    # Coal
    # Oil
    # Gas
    # Renewables
    # Electricity
    # Heat
    # Hydrogen
    # Other

    netz_trn = netz_trn_1.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    netz_trn.loc[netz_trn['Series'] == 'Total', 'Series'] = 'Transport (PJ)'

    netz_trn = netz_trn.iloc[[9, 0, 1, 2, 3, 4, 5, 6, 7, 8]].reset_index(drop = True)

    ## NON-ENERGY
    # Coal
    # Oil
    # Gas

    netz_nonen = EGEDA_years_netzero[(EGEDA_years_netzero['economy'] == economy) & 
                                      (EGEDA_years_netzero['item_code_new'] == '17_nonenergy_use') &
                                      (EGEDA_years_netzero['fuel_code'].isin(['1_coal', '2_coal_products',
                                       '3_peat', '4_peat_products', '6_crude_oil_and_ngl',
                                       '7_petroleum_products', '8_gas', '19_total']))].loc[:, 'fuel_code':].copy().reset_index(drop = True)

    coal_ne = netz_nonen[netz_nonen['fuel_code'].isin(['1_coal', '2_coal_products', '3_peat', '4_peat_products'])].copy()\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Coal').reset_index(drop = True)

    oil_ne = netz_nonen[netz_nonen['fuel_code'].isin(['6_crude_oil_and_ngl', '7_petroleum_products'])].copy()\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Oil').reset_index(drop = True)

    gas_ne = netz_nonen[netz_nonen['fuel_code'].isin(['8_gas'])].copy()\
        .groupby(['item_code_new']).sum().assign(fuel_code = 'Gas').reset_index(drop = True)

    netz_nonen = netz_nonen.append([coal_ne, oil_ne, gas_ne]).reset_index(drop = True)

    netz_nonen = netz_nonen[netz_nonen['fuel_code'].isin(['Coal', 'Oil', 'Gas', '19_total'])].drop('item_code_new', axis = 1)\
        .rename(columns = {'fuel_code': 'Series'}).reset_index(drop = True)

    netz_nonen.loc[netz_nonen['Series'] == '19_total', 'Series'] = 'Non-energy (PJ)'

    netz_nonen = netz_nonen[['Series'] + list(netz_stock_1.loc[:, '2000':'2050'])].reset_index(drop = True)

    ###############################################################

    #### ELECTRICITY
    ### CAPACITY

    netz_powcap = pd.DataFrame(['Coal', 'Coal CCS',\
        'Lignite', 'Gas', 'Gas CCS', 'Oil', 'Nuclear', 'Hydro', 'Bio', 'Wind', 'Solar', 'Geothermal',\
            'Waste', 'Storage', 'Other'], columns = ['TECHNOLOGY']).merge(netz_powcap_1.copy(),\
                on = 'TECHNOLOGY', how = 'outer').replace(np.nan, 0).rename(columns = {'TECHNOLOGY': 'Series'})\
                    .reset_index(drop = True)

    netz_powcap.loc[netz_powcap['Series'] == 'Total', 'Series'] = 'Total capacity (GW)' 

    netz_powcap = netz_powcap.iloc[[15, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]].reset_index(drop = True)

    ###################################
    ### GENERATION 

    netz_elecgen = netz_elecgen_2.copy().drop('Generation', axis = 1).rename(columns = {'TECHNOLOGY': 'Series'})

    netz_elecgen.loc[netz_elecgen['Series'] == 'Total', 'Series'] = 'Electricity generation (TWh)'

    netz_elecgen = netz_elecgen.iloc[[16, 0, 1, 2, 4, 5, 3, 7, 6, 10, 8, 9, 11, 12, 13, 14, 15]].reset_index(drop = True)

    ######################################
    ### EMISSIONS
    ## BY FUEL
    # Coal
    # Oil
    # Gas

    netz_emfuel = netz_emiss_fuel_1.copy().drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .reset_index(drop = True)

    netz_emfuel.loc[netz_emfuel['Series'] == 'Heat & others', 'Series'] = 'Other'
    netz_emfuel.loc[netz_emfuel['Series'] == 'Total', 'Series'] = 'Energy sector CO2 emissions (million tonnes)'

    netz_emfuel = netz_emfuel[netz_emfuel['Series']\
        .isin(['Coal', 'Oil', 'Gas', 'Other', 'Energy sector CO2 emissions (million tonnes)'])]\
            .reset_index(drop = True)

    netz_emfuel = netz_emfuel.iloc[[4, 0, 1, 2, 3]].reset_index(drop = True)

    ## BY SECTOR
    # Power
    # Own-use and losses
    # Agriculture 
    # Buildings
    # Industry
    # Transport (domestic)
    # Non-energy
    # Non-specified

    netz_emsector = netz_emiss_sector_1.copy().drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'})\
        .reset_index(drop = True)

    netz_emsector.loc[netz_emsector['Series'] == 'Total', 'Series'] = 'CO2 emissions by sector (million tonnes)'

    netz_emsector = netz_emsector.iloc[[7, 0, 1, 2, 3, 4, 5, 6]].reset_index(drop = True)

    ######################## Captured emissions

    # emissions factors
    hyd_coal = 0.094686
    hyd_gas = 0.056151

    netz_capemiss = netz_capemiss_df1[netz_capemiss_df1['REGION'] == economy].drop(['REGION', 'EMISSION'], axis = 1)\
        .reset_index(drop = True)

    # Coal
    netz_hydcoal = netz_hydrogen_1[netz_hydrogen_1['Technology'] == 'Coal gasification CCS'].copy().reset_index(drop = True)
    
    netz_hydcoal_captured = netz_hydcoal.select_dtypes(include = [np.number]) * hyd_coal

    netz_hydcoal[netz_hydcoal_captured.columns] = netz_hydcoal_captured
    netz_hydcoal = netz_hydcoal.rename(columns = {'Fuel': 'Series'})

    # Gas
    netz_hydgas = netz_hydrogen_1[netz_hydrogen_1['Technology'] == 'Steam methane reforming CCS'].copy().reset_index(drop = True)
    
    netz_hydgas_captured = netz_hydgas.select_dtypes(include = [np.number]) * hyd_gas

    netz_hydgas[netz_hydgas_captured.columns] = netz_hydgas_captured
    netz_hydgas = netz_hydgas.rename(columns = {'Fuel': 'Series'})

    netz_hydcap = pd.DataFrame(columns = netz_emsector.columns).append([netz_hydcoal, netz_hydgas])\
        .replace(np.nan, 0).drop('Technology', axis = 1).reset_index(drop = True)

    if netz_hydcap.empty:
        new_row = ['Hydrogen'] + [0] * 51
        new_series = pd.Series(new_row, index = netz_emsector.columns)
        netz_hydcap = netz_hydcap.append(new_series, ignore_index = True)

    else:
        netz_hydcap = netz_hydcap.copy().groupby('Series').sum().reset_index()

    ######################################################

    industry = netz_capemiss[netz_capemiss['TECHNOLOGY'].str.startswith('IND')].reset_index(drop = True)
    industry['Series'] = 'Industry'

    industry = industry.copy().drop('TECHNOLOGY', axis = 1).groupby('Series').sum().reset_index()

    netz_industry_cap = pd.DataFrame(columns = netz_emsector.columns)

    if industry.empty:
        new_row = ['Industry'] + [0] * 51
        new_series = pd.Series(new_row, index = netz_emsector.columns)
        netz_industry_cap = netz_industry_cap.append(new_series, ignore_index = True)

    else:
        netz_industry_cap = netz_industry_cap.append(industry).replace(np.nan, 0).reset_index(drop = True)

    power = netz_capemiss[netz_capemiss['TECHNOLOGY'].str.startswith('POW')].reset_index(drop = True)
    power['Series'] = 'Power'

    power = power.copy().drop('TECHNOLOGY', axis = 1).groupby('Series').sum().reset_index()

    netz_power_cap = pd.DataFrame(columns = netz_emsector.columns)

    if power.empty:
        new_row = ['Power'] + [0] * 51
        new_series = pd.Series(new_row, index = netz_emsector.columns)
        netz_power_cap = netz_power_cap.append(new_series, ignore_index = True)

    else:
        netz_power_cap = netz_power_cap.append(power).replace(np.nan, 0).reset_index(drop = True)
    
    ownuse = netz_capemiss[netz_capemiss['TECHNOLOGY'].str.startswith('OWN')].reset_index(drop = True)
    ownuse['Series'] = 'Own-use'

    ownuse = ownuse.copy().drop('TECHNOLOGY', axis = 1).groupby('Series').sum().reset_index()

    netz_ownuse_cap = pd.DataFrame(columns = netz_emsector.columns)

    if ownuse.empty:
        new_row = ['Own-use'] + [0] * 51
        new_series = pd.Series(new_row, index = netz_emsector.columns)
        netz_ownuse_cap = netz_ownuse_cap.append(new_series, ignore_index = True)

    else:
        netz_ownuse_cap = netz_ownuse_cap.append(ownuse).replace(np.nan, 0).reset_index(drop = True)

    # Captured emissions

    netz_captured = netz_industry_cap.append([netz_hydcap, netz_power_cap, netz_ownuse_cap]).reset_index(drop = True)

    cols = netz_captured.columns[1:]

    netz_captured[cols] = netz_captured[cols].apply(pd.to_numeric)

    netz_captured.loc['Total'] = netz_captured.sum(numeric_only = True)

    netz_captured.loc['Total', 'Series'] = 'Captured CO2 emissions (million tonnes)'
    netz_captured = netz_captured.copy().reset_index(drop = True).iloc[[4, 0, 1, 2, 3]].reset_index(drop = True)

    netz_captured_rows = netz_captured.shape[0]

    # Modern renewables

    netz_modren = netz_modren_4[netz_modren_4['item_code_new'].isin(['Total', 'Carbon Neutrality'])].copy()\
        .drop('fuel_code', axis = 1).rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    netz_modren.loc[netz_modren['Series'] == 'Total', 'Series'] = 'Modern renewables in FED (PJ)'
    netz_modren.loc[netz_modren['Series'] == 'Carbon Neutrality', 'Series'] = 'Modern renewables share of final energy demand'

    # More comprehensive modern renewables breakdown

    netz_modren_breakdown = netz_modren_4.copy().iloc[[6, 0, 1, 2, 3, 4, 5, 9, 11]].drop('fuel_code', axis = 1)\
        .rename(columns = {'item_code_new': 'Series'}).reset_index(drop = True)

    netz_modren_breakdown.loc[netz_modren_breakdown['Series'] == 'Electricity and heat TFEC', 'Series'] = 'Electricity and heat (not including own-use and losses)'
    netz_modren_breakdown.loc[netz_modren_breakdown['Series'] == 'Total', 'Series'] = 'Modern renewables in FED (PJ)'
    netz_modren_breakdown.loc[netz_modren_breakdown['Series'] == 'TFEC', 'Series'] = 'Final energy demand (PJ)'
    netz_modren_breakdown.loc[netz_modren_breakdown['Series'] == 'Carbon Neutrality', 'Series'] = 'Modern renewables share of FED'

    netz_modren_gen = netz_modren_4.copy().iloc[7:9].drop('item_code_new', axis = 1).rename(columns = {'fuel_code': 'Series'})\
        .reset_index(drop = True)

    netz_modren_gen.loc[netz_modren_gen['Series'] == 'Modern renewables', 'Series'] = 'Renewable generation (TWh)'
    netz_modren_gen.loc[netz_modren_gen['Series'] == 'Total', 'Series'] = 'Total generation (TWh)'

    # Convert power PJ to TWh
    cols = netz_modren_gen.columns[1:]

    netz_modren_gen[cols] = netz_modren_gen[cols] / 3.6

    netz_modren_gencalc = ['Renewable generation share'] + list(netz_modren_gen.iloc[0, 1:] / netz_modren_gen.iloc[1, 1:])
    netz_modren_gencalc_series = pd.Series(netz_modren_gencalc, index = netz_modren_gen.columns)

    netz_modren_gen = netz_modren_gen.copy().append(netz_modren_gencalc_series, ignore_index = True).reset_index(drop = True)

    # Join the data
    netz_modren_A = netz_modren_breakdown.append(netz_modren_gen).reset_index(drop = True)

    netz_modren_A_rows = netz_modren_A.shape[0]

    # Join relevant dataframes together

    netz_top_df = macro_1.append([netz_tpes_3, netz_tfec_1, netz_co2int, netz_modren]).reset_index(drop = True)

    netz_top_rows = netz_top_df.shape[0]
    netz_top_cols = netz_top_df.shape[1]

    netz_supply_df = netz_prod_1.append([netz_nettrade_1, netz_bunkers_3, netz_stock_1, netz_tpes_1]).reset_index(drop = True)

    netz_supply_rows = netz_supply_df.shape[0]
    netz_supply_cols = netz_supply_df.shape[1]

    # netz_powsum
    # 'Input fuel'
    # netz_powuse 
    # 'Output fuel'
    # netz_elecheat
    # netz_refining_1
    # 'Input fuel'
    # netz_refining_2
    # 'Output fuel'
    # netz_refining_3
    # netz_ownloss
    # netz_transtat

    netz_demand_df = netz_fedsector.copy().append([netz_fedfuel, netz_ind, netz_trn, netz_build, netz_nonen]).reset_index(drop = True)

    netz_demand_rows = netz_demand_df.shape[0]
    netz_demand_cols = netz_demand_df.shape[1]

    netz_electricity = netz_elecgen.append(netz_powcap).reset_index(drop = True)
    # netz_powcap
    # netz_elecgen

    netz_electricity_rows = netz_electricity.shape[0]

    # netz_emfuel
    # 'By sector'
    # netz_emsector

    netz_emfuel_rows = netz_emfuel.shape[0]
    netz_emsector_rows = netz_emsector.shape[0]

    # Define directory to save charts and tables workbook
    script_dir = './results/'
    results_dir = os.path.join(script_dir, 'Appendix')
    if not os.path.isdir(results_dir):
        os.makedirs(results_dir)
        
    # Create a Pandas excel writer workbook using xlsxwriter as the engine and save it in the directory created above
    writer = pd.ExcelWriter(results_dir + '/' + economy + '_appendix.xlsx', engine = 'xlsxwriter')
    workbook = writer.book
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Insert the various dataframes into different sheets of the workbook
    # REFERENCE and NETZERO

    # Data frames needed for appendix - Reference
    ref_top_df.to_excel(writer, sheet_name = economy, index = False, startrow = 5)
    ref_supply_df.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + 9)
    ref_powsum.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + ref_supply_rows + 13)
    ref_powuse.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + 16)
    ref_elecheat.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + 17)
    ref_refining_1.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + 17)
    ref_refining_2.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + 19)
    ref_refining_3.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + 21)
    ref_ownloss.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + 22)
    ref_transtat.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + 22)
    ref_demand_df.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + ref_transtat_rows + 25)
    ref_electricity.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + ref_transtat_rows + ref_demand_rows + 29)
    ref_emfuel.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + ref_transtat_rows + ref_demand_rows + ref_electricity_rows + 33)
    # Remove total first row of emissions by sector because already provided in emissions by fuel
    ref_emsector.iloc[1:,:].to_excel(writer, sheet_name = economy, index = False, header = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + ref_transtat_rows + ref_demand_rows + ref_electricity_rows + ref_emfuel_rows + 35)
    ref_captured.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + ref_transtat_rows + ref_demand_rows + ref_electricity_rows + ref_emfuel_rows + ref_emsector_rows + 37)
    ref_modren_A.to_excel(writer, sheet_name = economy, index = False, startrow = ref_top_rows + ref_supply_rows + ref_powuse_rows + ref_elecheat_rows + ref_ownloss_rows + ref_transtat_rows + ref_demand_rows + ref_electricity_rows + ref_emfuel_rows + ref_emsector_rows + ref_captured_rows + 41)

    # Reference rows
    ref_rows = 223 
    
    # Carbon neutrality
    # Data frames needed for appendix - Carbon neutrality
    netz_top_df.to_excel(writer, sheet_name = economy, index = False, startrow = 5 + ref_rows)
    netz_supply_df.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + 9 + ref_rows)
    netz_powsum.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + netz_supply_rows + 13 + ref_rows)
    netz_powuse.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + 16 + ref_rows)
    netz_elecheat.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + 17 + ref_rows)
    netz_refining_1.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + 17 + ref_rows)
    netz_refining_2.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + 19 + ref_rows)
    netz_refining_3.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + 21 + ref_rows)
    netz_ownloss.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + 22 + ref_rows)
    netz_transtat.to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + 22 + ref_rows)
    netz_demand_df.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + netz_transtat_rows + 25 + ref_rows)
    netz_electricity.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + netz_transtat_rows + netz_demand_rows + 29 + ref_rows)
    netz_emfuel.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + netz_transtat_rows + netz_demand_rows + netz_electricity_rows + 33 + ref_rows)
    # Remove total first row of emissions by sector because already provided in emissions by fuel
    netz_emsector.iloc[1:,:].to_excel(writer, sheet_name = economy, index = False, header = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + netz_transtat_rows + netz_demand_rows + netz_electricity_rows + netz_emfuel_rows + 35 + ref_rows)
    netz_captured.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + netz_transtat_rows + netz_demand_rows + netz_electricity_rows + netz_emfuel_rows + netz_emsector_rows + 37 + ref_rows)
    netz_modren_A.to_excel(writer, sheet_name = economy, index = False, startrow = netz_top_rows + netz_supply_rows + netz_powuse_rows + netz_elecheat_rows + netz_ownloss_rows + netz_transtat_rows + netz_demand_rows + netz_electricity_rows + netz_emfuel_rows + netz_emsector_rows + netz_captured_rows + 41 + ref_rows)

    worksheet1 = writer.sheets[economy]

    # Comma format and header format        
    space_format = workbook.add_format({'num_format': '# ### ### ##0.0;-# ### ### ##0.0;-'})
    percentage_format = workbook.add_format({'num_format': '0.0%'})
    header_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True})
    cell_format1 = workbook.add_format({'bold': True})
    cell_format2 = workbook.add_format({'font_size': 9})

    worksheet1.write(0, 0, APEC_economies[economy], cell_format1)
    worksheet1.write(2, 0, 'Reference scenario')
    worksheet1.write(4, 0, 'Summary - Reference')
    worksheet1.write(24, 0, 'Production, trade, and supply - Reference')
    worksheet1.write(62, 0, 'Transformation - Reference')
    worksheet1.write(65, 0, 'Input')
    worksheet1.write(73, 0, 'Output')
    worksheet1.write(77, 0, 'Input')
    worksheet1.write(79, 0, 'Output')
    worksheet1.write(93, 0, 'Demand - Reference')
    worksheet1.write(146, 0, 'Electricity - Reference')
    worksheet1.write(183, 0, 'CO2 emissions - Reference')
    worksheet1.write(190, 0, 'By sector')
    worksheet1.write(200, 0, 'Carbon, capture, and storage technologies - Reference')
    worksheet1.write(209, 0, 'Modern renewables breakdown - Reference')

    # worksheet1.write(0, 0, APEC_economies[economy], cell_format1)
    worksheet1.write(2 + ref_rows, 0, 'Carbon Neutrality scenario')
    worksheet1.write(4 + ref_rows, 0, 'Summary - Carbon Neutrality')
    worksheet1.write(24 + ref_rows, 0, 'Production, trade, and supply - Carbon Neutrality')
    worksheet1.write(62 + ref_rows, 0, 'Transformation - Carbon Neutrality')
    worksheet1.write(65 + ref_rows, 0, 'Input')
    worksheet1.write(73 + ref_rows, 0, 'Output')
    worksheet1.write(77 + ref_rows, 0, 'Input')
    worksheet1.write(79 + ref_rows, 0, 'Output')
    worksheet1.write(93 + ref_rows, 0, 'Demand - Carbon Neutrality')
    worksheet1.write(146 + ref_rows, 0, 'Electricity - Carbon Neutrality')
    worksheet1.write(183 + ref_rows, 0, 'CO2 emissions - Carbon Neutrality')
    worksheet1.write(190 + ref_rows, 0, 'By sector')
    worksheet1.write(200 + ref_rows, 0, 'Carbon, capture, and storage technologies - Carbon Neutrality')
    worksheet1.write(209 + ref_rows, 0, 'Modern renewables breakdown - Carbon Neutrality')
        
    writer.save()

print('Economy appendices are saved in the results folder specified.')

# Now consolidate the workbooks for each of the economies into one workbook

excel_filenames = glob.glob("./results/Appendix/*.xlsx")

writer3 = pd.ExcelWriter('./results/Appendix/Outlook_appendix.xlsx', engine = 'xlsxwriter')

for index, file in enumerate(excel_filenames):
    temp_1 = pd.read_excel(file)
    temp_1.to_excel(writer3, sheet_name = Economy_codes[index], index = False)

writer3.save()    

print('The individual appendix workbooks have been stitched together into one appendix file with no formatting.')



